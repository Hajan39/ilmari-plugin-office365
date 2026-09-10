// ilmari-plugin-office365 — Microsoft 365 (Graph) for ilmari.
//
// One file on purpose: ilmari hash-pins the entry file it approved, so keeping
// everything here means every line the daemon runs is covered by that pin.
//
// Auth is delegated (device code), not client credentials: acting *as the
// signed-in user* is the point — mail leaves from their own mailbox and lands
// in their Sent Items — and a public client needs no client secret to store.

import { createHash, randomBytes } from "node:crypto";

const PLUGIN = "office365";
const GRAPH = "https://graph.microsoft.com/v1.0";
const DEVICE_CODE_GRANT = "urn:ietf:params:oauth:grant-type:device_code";
/** Where the browser sign-in lands. Nothing listens there: the user copies the
 *  code out of the address bar. Must be registered on the app as a "Mobile and
 *  desktop" redirect (Microsoft's own public clients already have it). */
const BROWSER_REDIRECT = "http://localhost";

/** offline_access is what makes a sign-in survive a daemon restart; the rest
 *  are the delegated permissions this plugin's surfaces need. The user is
 *  asked to consent once, to whatever the app registration actually lists. */
export const SCOPE =
  "offline_access Mail.Send Calendars.Read Files.Read.All Chat.ReadWrite ChannelMessage.Send";

/** Graph refuses a sendMail request over ~4 MB, attachments included; this is
 *  the ceiling this plugin enforces first, with a readable reason. */
const ATTACHMENT_LIMIT = 3 * 1024 * 1024;
/** A file read into a prompt is charged as tokens — read the head of a big
 *  one rather than blowing up the context window. */
const FILE_READ_LIMIT = 200_000;

const configDecl = {
  clientId: {
    label: "Application (client) ID",
    description:
      "The app registration ilmari signs in through: Azure Portal -> Microsoft Entra ID -> App registrations -> your app -> Overview. It must be a public client with device code flow allowed. No client secret is needed.",
    env: "OFFICE365_CLIENT_ID",
  },
  signIn: {
    label: "Sign-in method",
    description:
      "device (default) or browser. Use browser when Conditional Access blocks the device-code flow: the task log shows a sign-in link to open in your own (managed) browser; it ends on an unreachable http://localhost page — paste that page's address, or just its code=..., into 'Authorization code' while the step is waiting.",
    env: "OFFICE365_SIGN_IN",
  },
  authCode: {
    label: "Authorization code",
    description:
      "Only for the browser sign-in: paste the code (or the whole http://localhost?code=... address) here while a step waits for it. Single-use; cleared automatically once exchanged.",
    secret: true,
  },
  scopes: {
    label: "Permissions",
    description:
      "Space-separated delegated scopes to ask for. Empty means all the plugin's surfaces. Files.Read.All, Chat.ReadWrite and ChannelMessage.Send need an administrator's consent — drop them here (leaving e.g. 'offline_access Mail.Send Calendars.Read') to sign in with user consent alone, at the cost of the file, chat and channel steps.",
    env: "OFFICE365_SCOPES",
  },
  tenantId: {
    label: "Directory (tenant) ID",
    description:
      "Which Entra tenant to sign in against — your company's tenant id, on the same Overview page. Empty means organizations, which lets any work or school account in.",
    env: "OFFICE365_TENANT_ID",
  },
  to: {
    label: "Default recipient",
    description:
      "Where the email channel sends task notifications, and the fallback recipient for an email step that names none. Several addresses separated by commas.",
    env: "OFFICE365_TO",
  },
  refreshToken: {
    label: "Sign-in token",
    description:
      "Filled in automatically after the device-code sign-in and rotated from then on. Clear this field to sign out and be asked to sign in again.",
    secret: true,
  },
};

// ---------------------------------------------------------------- auth

/** Access tokens live an hour; keeping the current one per account means a
 *  workflow of ten Graph steps signs in once. Never written to disk — only
 *  the refresh token is, and that goes to ilmari's encrypted config store. */
const tokenCache = new Map();

const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const authBase = (tenant) =>
  `https://login.microsoftonline.com/${encodeURIComponent(tenant || "organizations")}/oauth2/v2.0`;

async function postForm(url, form) {
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(form).toString(),
  });
  return await res.json();
}

/**
 * Runs the device-code flow to completion. Blocks until the user signs in or
 * the code expires (~15 minutes): node types are allowed to block, and this
 * happens once per sign-in, not once per call.
 */
export async function deviceCodeSignIn(clientId, tenantId, announce, sleep = wait, scope = SCOPE) {
  const base = authBase(tenantId);
  const res = await fetch(`${base}/devicecode`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: clientId, scope }).toString(),
  });
  const start = await res.json();
  if (!start.device_code || !start.user_code) {
    throw new Error(`device code request failed: ${start.error_description ?? res.status}`);
  }
  announce(
    start.message ??
      `Sign in to Microsoft at ${start.verification_uri} and enter the code ${start.user_code}`,
  );

  const deadline = Date.now() + (start.expires_in ?? 900) * 1000;
  let intervalMs = (start.interval ?? 5) * 1000;
  while (Date.now() < deadline) {
    await sleep(intervalMs);
    const token = await postForm(`${base}/token`, {
      grant_type: DEVICE_CODE_GRANT,
      client_id: clientId,
      device_code: start.device_code,
    });
    if (token.access_token && token.refresh_token) {
      return {
        accessToken: token.access_token,
        refreshToken: token.refresh_token,
        expiresIn: token.expires_in ?? 3600,
      };
    }
    // the two documented "keep waiting" answers; anything else is fatal
    if (token.error === "slow_down") intervalMs += 5000;
    else if (token.error !== "authorization_pending") {
      throw new Error(`sign-in failed: ${token.error_description ?? token.error}`);
    }
  }
  throw new Error("sign-in timed out — the device code expired before it was entered");
}

/**
 * Authorization code + PKCE, with the sign-in done in the user's own browser
 * and the code carried back by hand — for tenants whose Conditional Access
 * refuses device code (the token request comes from the daemon, not from a
 * compliant device). `readCode` is polled until it yields the pasted code.
 */
export async function browserSignIn(clientId, tenantId, announce, readCode, sleep = wait, scope = SCOPE) {
  const base = authBase(tenantId);
  const verifier = randomBytes(32).toString("base64url");
  const challenge = createHash("sha256").update(verifier).digest("base64url");
  const url = `${base}/authorize?${new URLSearchParams({
    client_id: clientId,
    response_type: "code",
    redirect_uri: BROWSER_REDIRECT,
    scope,
    code_challenge: challenge,
    code_challenge_method: "S256",
    prompt: "select_account",
  })}`;
  announce(
    `Sign in to Microsoft at ${url}
` +
      `The browser ends on an unreachable ${BROWSER_REDIRECT} page: paste its address (or the code=... part) into the office365 plugin's "Authorization code" field.`,
  );
  // a code left over from an earlier attempt is bound to that attempt's
  // verifier and would only fail the exchange — wait for a fresh paste
  const stale = readCode();
  const deadline = Date.now() + 900_000;
  while (Date.now() < deadline) {
    await sleep(3000);
    const pasted = readCode();
    if (!pasted || pasted === stale) continue;
    const code = pasted.includes("code=") ? new URL(pasted).searchParams.get("code") : pasted;
    const token = await postForm(`${base}/token`, {
      grant_type: "authorization_code",
      client_id: clientId,
      code,
      redirect_uri: BROWSER_REDIRECT,
      code_verifier: verifier,
      scope,
    });
    if (token.access_token && token.refresh_token) {
      return {
        accessToken: token.access_token,
        refreshToken: token.refresh_token,
        expiresIn: token.expires_in ?? 3600,
      };
    }
    throw new Error(`sign-in failed: ${token.error_description ?? token.error}`);
  }
  throw new Error("sign-in timed out — no authorization code was pasted within 15 minutes");
}

/**
 * A usable access token: the cached one, else a refresh, else a full sign-in
 * announced through `announce`. `interactive: false` — the email channel,
 * which must not hang a notification for fifteen minutes — turns a missing or
 * dead sign-in into null instead of a device-code prompt.
 */
export async function graphToken(ctx, { project, announce, interactive = true } = {}) {
  const cfg = ctx.pluginConfig(PLUGIN, project) ?? {};
  const clientId = (cfg.clientId ?? "").trim();
  if (!clientId) throw new Error("office365: no application (client) ID configured");
  const tenantId = (cfg.tenantId ?? "").trim();
  const scope = (cfg.scopes ?? "").trim() || SCOPE;
  const say = announce ?? (() => {});

  const key = `${clientId}@${project ?? ""}`;
  const hit = tokenCache.get(key);
  if (hit && hit.expiresAt > Date.now() + 60_000) return hit.token;

  const remember = ({ accessToken, expiresIn, refreshToken }) => {
    tokenCache.set(key, { token: accessToken, expiresAt: Date.now() + expiresIn * 1000 });
    // savePluginConfig is newer than the first plugin contract; without it the
    // sign-in still works, it just has to be repeated after a daemon restart,
    // which beats refusing to send at all
    ctx.savePluginConfig?.(PLUGIN, { refreshToken }, project);
    return accessToken;
  };

  const stored = (cfg.refreshToken ?? "").trim();
  if (stored) {
    const refreshed = await postForm(`${authBase(tenantId)}/token`, {
      grant_type: "refresh_token",
      client_id: clientId,
      refresh_token: stored,
      scope,
    });
    if (refreshed.access_token && refreshed.refresh_token) {
      return remember({
        accessToken: refreshed.access_token,
        refreshToken: refreshed.refresh_token,
        expiresIn: refreshed.expires_in ?? 3600,
      });
    }
    say(`Microsoft sign-in expired (${refreshed.error ?? "unknown"})`);
  }
  if (!interactive) return null;
  if ((cfg.signIn ?? "").trim() === "browser") {
    const readCode = () => (ctx.pluginConfig(PLUGIN, project)?.authCode ?? "").trim();
    const tokens = await browserSignIn(clientId, tenantId, say, readCode, wait, scope);
    ctx.savePluginConfig?.(PLUGIN, { authCode: "" }, project); // single-use
    return remember(tokens);
  }
  return remember(await deviceCodeSignIn(clientId, tenantId, say, wait, scope));
}

/** One Graph call as the signed-in user. Returns the parsed body, or
 *  undefined for the 202/204 answers sendMail and message posts give. */
export async function graph(ctx, method, path, body, opts = {}) {
  const token = await graphToken(ctx, opts);
  if (!token) return undefined;
  const res = await fetch(`${GRAPH}${path}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      ...(body !== undefined ? { "Content-Type": "application/json" } : {}),
    },
    ...(body !== undefined ? { body: JSON.stringify(body) } : {}),
  });
  // the path carries no token, so the status and body are safe to surface
  if (!res.ok) throw new Error(`Graph ${method} ${path}: ${res.status} ${await res.text()}`);
  if (res.status === 202 || res.status === 204) return undefined;
  return await res.json();
}

/** Follows @odata.nextLink until the collection ends or `max` items are in
 *  hand — a month of meetings is more than one page. */
async function graphAll(ctx, path, opts = {}, max = 500) {
  const items = [];
  let next = path;
  while (next && items.length < max) {
    const page = await graph(ctx, "GET", next, undefined, opts);
    if (!page) break;
    items.push(...(page.value ?? []));
    const link = page["@odata.nextLink"];
    next = link ? link.replace(GRAPH, "") : undefined;
  }
  return items.slice(0, max);
}

/** Drops every cached access token. Exported for tests and for the case where
 *  the user clears the sign-in token in the GUI. */
export function forgetTokens() {
  tokenCache.clear();
}

// ------------------------------------------------------------- helpers

const text = (v) => (v === undefined || v === null ? "" : String(v));
const trimmed = (v) => text(v).trim();
const list = (v) =>
  trimmed(v)
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
const recipients = (v) => list(v).map((address) => ({ emailAddress: { address } }));

/** Days back from today, today included, or an explicit ISO range. Dates are
 *  computed here so a workflow never has to shell out to `date`. */
export function dayRange({ days, from, to }, now = new Date()) {
  const start = trimmed(from);
  const end = trimmed(to);
  if (start || end) return { from: start || end, to: end || start };
  const back = Math.max(1, Math.floor(Number(days ?? 1)) || 1);
  const day = (d) => d.toISOString().slice(0, 10);
  return { from: day(new Date(now.getTime() - (back - 1) * 86_400_000)), to: day(now) };
}

/** Reads the named files from the working directory as Graph fileAttachments.
 *  Paths are resolved inside the workdir: a step is workflow-authored, but its
 *  parameters can come from an issue title, so "../../.ssh/id_rsa" must not
 *  become an attachment. */
async function readAttachments(workdir, paths) {
  if (paths.length === 0) return { attachments: [], skipped: [] };
  const { readFile } = await import("node:fs/promises");
  const { resolve, sep, basename } = await import("node:path");
  const root = resolve(workdir);
  const attachments = [];
  const skipped = [];
  let total = 0;
  for (const rel of paths) {
    const full = resolve(root, rel);
    if (full !== root && !full.startsWith(root + sep)) {
      skipped.push(`${rel} (outside the working directory)`);
      continue;
    }
    let bytes;
    try {
      bytes = await readFile(full);
    } catch (error) {
      skipped.push(`${rel} (${error.code ?? error})`);
      continue;
    }
    total += bytes.length;
    if (total > ATTACHMENT_LIMIT) {
      skipped.push(`${rel} (over the ${ATTACHMENT_LIMIT / 1024 / 1024} MB total limit)`);
      total -= bytes.length;
      continue;
    }
    attachments.push({
      "@odata.type": "#microsoft.graph.fileAttachment",
      name: basename(rel),
      contentBytes: bytes.toString("base64"),
    });
  }
  return { attachments, skipped };
}

/** The one place a mail is built and posted, shared by the node type, the
 *  agent tool and the notification channel. */
async function sendMail(ctx, { to, cc, subject, body, html, attachments = [] }, opts) {
  const toRecipients = recipients(to);
  if (toRecipients.length === 0) throw new Error("no recipient");
  await graph(
    ctx,
    "POST",
    "/me/sendMail",
    {
      message: {
        subject: subject || "(no subject)",
        body: { contentType: html ? "HTML" : "Text", content: body },
        toRecipients,
        ...(recipients(cc).length > 0 ? { ccRecipients: recipients(cc) } : {}),
        ...(attachments.length > 0 ? { attachments } : {}),
      },
      saveToSentItems: true,
    },
    opts,
  );
  return toRecipients.map((r) => r.emailAddress.address);
}

/** Turns a node's ctx into the sign-in announcement route: the task log and
 *  the GUI, since the daemon has no terminal to prompt on. */
const announceVia = (ctx) => (message) => {
  ctx.emit?.("office365_sign_in", { message });
};

const nodeOpts = (ctx) => ({ project: ctx.project, announce: announceVia(ctx) });

const jsonOutput = (value) => ({ ok: true, output: JSON.stringify(value) });

/** Every node type and tool funnels its failures through here: a readable
 *  reason, never the token, so catches and fallbacks can act on it. */
const failed = (what, error) => ({ ok: false, reason: `${what}: ${error?.message ?? error}` });

// -------------------------------------------------------------- surfaces

const calendarFields = "subject,start,end,isAllDay,organizer,location,categories,webLink";

async function calendarView(ctx, params, opts) {
  const { from, to } = dayRange(params);
  // calendarView expands recurring meetings into their actual occurrences,
  // which /me/events does not — a weekly stand-up should show up every week
  const path =
    `/me/calendarView?startDateTime=${from}T00:00:00&endDateTime=${to}T23:59:59` +
    `&$select=${calendarFields}&$orderby=start/dateTime&$top=100`;
  const events = await graphAll(ctx, path, opts);
  return events.map((e) => ({
    subject: e.subject ?? "",
    start: e.start?.dateTime ?? "",
    end: e.end?.dateTime ?? "",
    allDay: Boolean(e.isAllDay),
    organizer: e.organizer?.emailAddress?.address ?? "",
    location: e.location?.displayName ?? "",
    categories: e.categories ?? [],
    url: e.webLink ?? "",
  }));
}

async function findFiles(ctx, query, opts) {
  const path = `/me/drive/root/search(q='${encodeURIComponent(query).replace(/'/g, "''")}')?$top=25&$select=id,name,size,lastModifiedDateTime,webUrl,parentReference`;
  const found = await graphAll(ctx, path, opts, 25);
  return found.map((f) => ({
    id: f.id ?? "",
    name: f.name ?? "",
    size: Number(f.size ?? 0),
    modified: f.lastModifiedDateTime ?? "",
    path: f.parentReference?.path ?? "",
    url: f.webUrl ?? "",
  }));
}

/** Downloads one drive item as text. Binary formats (a .docx, a .pdf) come
 *  back as bytes Graph will not convert, so this is for text-shaped files. */
async function readFile(ctx, itemId, opts) {
  const token = await graphToken(ctx, opts);
  const res = await fetch(`${GRAPH}/me/drive/items/${encodeURIComponent(itemId)}/content`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Graph GET /me/drive/items/${itemId}/content: ${res.status}`);
  const body = await res.text();
  return body.length > FILE_READ_LIMIT
    ? `${body.slice(0, FILE_READ_LIMIT)}\n… truncated at ${FILE_READ_LIMIT} characters`
    : body;
}

/** A Teams message goes either into a 1:1/group chat or into a channel; the
 *  two have different endpoints and different permissions, so the caller says
 *  which by filling in `chat` or `team` + `channel`. */
async function postTeamsMessage(ctx, { chat, team, channel, message, html }, opts) {
  const content = text(message);
  if (!content.trim()) throw new Error("message is empty");
  const body = { body: { contentType: html ? "html" : "text", content } };
  if (trimmed(chat)) {
    await graph(ctx, "POST", `/chats/${encodeURIComponent(trimmed(chat))}/messages`, body, opts);
    return `chat ${trimmed(chat)}`;
  }
  if (trimmed(team) && trimmed(channel)) {
    const path = `/teams/${encodeURIComponent(trimmed(team))}/channels/${encodeURIComponent(trimmed(channel))}/messages`;
    await graph(ctx, "POST", path, body, opts);
    return `channel ${trimmed(channel)}`;
  }
  throw new Error("name either a chat id, or both a team id and a channel id");
}

// ---------------------------------------------------------------- plugin

const plugin = {
  name: "office365",
  version: "0.1.1",
  description:
    "Microsoft 365 for ilmari, as the signed-in user: send mail from your own mailbox (with attachments from the working directory), read your calendar, search and read OneDrive/SharePoint files, and post into a Teams chat or channel. Sign-in is the device-code flow — you approve it once in a browser and ilmari keeps a refresh token in its encrypted config.",
  setup:
    "Register a public client app in Microsoft Entra ID (App registrations -> New registration -> Accounts in this organizational directory; under Authentication turn on 'Allow public client flows'), add the delegated permissions you want (Mail.Send, Calendars.Read, Files.Read.All, Chat.ReadWrite, ChannelMessage.Send) and paste its Application (client) ID and Directory (tenant) ID below. The first send prints a code and a link into the task log — open it, sign in, and ilmari remembers you from then on.",
  // net: Graph and the sign-in endpoint. fs: attachments are read out of the
  // task's working directory. secrets: the refresh token is stored encrypted.
  capabilities: ["net", "fs", "secrets"],
  config: configDecl,

  channels: [
    {
      name: "email",
      async send(message, ctx) {
        if (!ctx?.pluginConfig) return; // no way to reach the config: stay silent
        const cfg = ctx.pluginConfig("office365") ?? {};
        if (!cfg.clientId || !cfg.to) return; // unconfigured — stay silent
        // interactive: false — a notification must never block on a sign-in
        // the user has to finish in a browser
        await sendMail(
          ctx,
          {
            to: cfg.to,
            subject: message.split("\n")[0].slice(0, 120) || "ilmari",
            body: message,
          },
          { interactive: false },
        );
      },
    },
  ],

  nodeTypes: [
    {
      type: "office365-email",
      kind: "output",
      description:
        "Sends an email from your own Microsoft 365 mailbox — how a task reports out to people who do not watch ilmari. It appears in your Sent Items like anything else you send. {{<nodeId>.result}} inserts an earlier step's output into the body.",
      params: {
        to: {
          type: "string",
          description:
            "Recipients, separated by commas. Empty uses the default recipient from the plugin's settings.",
          example: "team@corp.com, boss@corp.com",
        },
        cc: { type: "string", description: "Recipients in copy, separated by commas." },
        subject: {
          type: "string",
          required: true,
          description: "Subject line. {{task}} is the task title.",
          example: "ilmari finished {{task}}",
        },
        body: {
          type: "string",
          required: true,
          description:
            "The message. {{task}} and {{<nodeId>.result}} interpolate, so an earlier step's output can be the whole mail.",
          example: "{{deliver.result}}",
        },
        html: {
          type: "boolean",
          description: "Send the body as HTML instead of plain text. Empty = plain text.",
        },
        attach: {
          type: "string",
          description:
            "Files to attach, as paths inside the working directory, separated by commas. Together they must stay under 3 MB, which is what Graph accepts in one send; anything refused is named in the step's output.",
          example: "report.md, out/coverage.txt",
        },
      },
      async run(params, ctx) {
        const cfg = ctx.pluginConfig(PLUGIN, ctx.project) ?? {};
        const to = trimmed(params.to) || text(cfg.to);
        if (!to) return { ok: false, reason: "office365-email: no recipient" };
        try {
          const { attachments, skipped } = await readAttachments(ctx.workdir, list(params.attach));
          const sent = await sendMail(
            ctx,
            {
              to,
              cc: params.cc,
              subject: trimmed(params.subject),
              body: text(params.body),
              html: Boolean(params.html),
              attachments,
            },
            nodeOpts(ctx),
          );
          ctx.emit("office365_mail_sent", { to: sent, attachments: attachments.length });
          const note = skipped.length > 0 ? `\nskipped: ${skipped.join(", ")}` : "";
          return { ok: true, output: `sent to ${sent.join(", ")}${note}` };
        } catch (error) {
          return failed("office365-email", error);
        }
      },
    },
    {
      type: "office365-calendar",
      description:
        "Reads your Microsoft 365 calendar. The step's output is a JSON array of meetings — subject, start, end, organizer, location, categories — so {{<nodeId>.result}} can hand a day or a week to a later prompt. Recurring meetings are expanded into their actual occurrences. Nothing in the range is a success with an empty array.",
      params: {
        days: {
          type: "number",
          description:
            "How many days back from today to read, today included. Ignored when from/to are set. Empty = 1.",
          example: "7",
        },
        from: {
          type: "string",
          description:
            "Start of an explicit range as an ISO date (YYYY-MM-DD); wins over days. Setting only one end reads that single day.",
          example: "2026-09-01",
        },
        to: { type: "string", description: "End of the explicit range, inclusive.", example: "2026-09-30" },
      },
      async run(params, ctx) {
        try {
          return jsonOutput(await calendarView(ctx, params, nodeOpts(ctx)));
        } catch (error) {
          return failed("office365-calendar", error);
        }
      },
    },
    {
      type: "office365-teams-message",
      kind: "output",
      description:
        "Posts a message into a Teams chat or channel as you. Give it a chat id for a direct or group chat, or a team id and a channel id for a channel post. {{<nodeId>.result}} inserts an earlier step's output.",
      params: {
        chat: {
          type: "string",
          description:
            "Chat id for a direct or group chat. Copy it from the chat's link in Teams (the 19:...@thread.v2 part), or leave empty and use team + channel instead.",
        },
        team: { type: "string", description: "Team id, for posting into a channel." },
        channel: { type: "string", description: "Channel id inside that team." },
        message: {
          type: "string",
          required: true,
          description: "The message. {{task}} and {{<nodeId>.result}} interpolate.",
          example: "ilmari finished {{task}}: {{deliver.result}}",
        },
        html: { type: "boolean", description: "Send as HTML instead of plain text." },
      },
      async run(params, ctx) {
        try {
          const where = await postTeamsMessage(
            ctx,
            {
              chat: params.chat,
              team: params.team,
              channel: params.channel,
              message: params.message,
              html: Boolean(params.html),
            },
            nodeOpts(ctx),
          );
          ctx.emit("office365_teams_posted", { where });
          return { ok: true, output: `posted to ${where}` };
        } catch (error) {
          return failed("office365-teams-message", error);
        }
      },
    },
  ],

  tools: [
    {
      name: "office365_calendar",
      description:
        "Read the user's Microsoft 365 calendar for a range of days. Returns the meetings as JSON: subject, start, end, organizer, location, categories. Use it to see what the user actually spent a day on.",
      parameters: {
        days: { type: "number", description: "Days back from today, today included. Default 1." },
        from: { type: "string", description: "Start date (YYYY-MM-DD); overrides days." },
        to: { type: "string", description: "End date (YYYY-MM-DD), inclusive." },
      },
      async execute(input, ctx) {
        return JSON.stringify(await calendarView(ctx, input, { announce: announceVia(ctx) }));
      },
    },
    {
      name: "office365_find_files",
      description:
        "Search the user's OneDrive and the SharePoint files shared with them. Returns up to 25 matches as JSON with id, name, size, last modified and a link. Pass an id to office365_read_file to read one.",
      parameters: {
        query: { type: "string", required: true, description: "What to search for." },
      },
      async execute(input, ctx) {
        return JSON.stringify(
          await findFiles(ctx, trimmed(input.query), { announce: announceVia(ctx) }),
        );
      },
    },
    {
      name: "office365_read_file",
      description:
        "Read one OneDrive/SharePoint file as text, by the id office365_find_files returned. Text-shaped files only (Markdown, code, CSV, plain text) — a Word or PDF file comes back as unreadable bytes. Long files are truncated.",
      parameters: {
        id: { type: "string", required: true, description: "Drive item id from office365_find_files." },
      },
      async execute(input, ctx) {
        return await readFile(ctx, trimmed(input.id), { announce: announceVia(ctx) });
      },
    },
    {
      name: "office365_send_email",
      description:
        "Send an email from the user's own mailbox. Use it when the user asked for something to be mailed out; the workflow's own reporting should use the office365-email step instead.",
      parameters: {
        to: { type: "string", required: true, description: "Recipients, comma separated." },
        subject: { type: "string", required: true, description: "Subject line." },
        body: { type: "string", required: true, description: "Message text." },
      },
      async execute(input, ctx) {
        const sent = await sendMail(
          ctx,
          { to: input.to, subject: trimmed(input.subject), body: text(input.body) },
          { announce: announceVia(ctx) },
        );
        return `sent to ${sent.join(", ")}`;
      },
    },
    {
      name: "office365_send_teams_message",
      description:
        "Post a message into a Teams chat (chat id) or channel (team id + channel id) as the user.",
      parameters: {
        chat: { type: "string", description: "Chat id, for a direct or group chat." },
        team: { type: "string", description: "Team id, for a channel post." },
        channel: { type: "string", description: "Channel id inside that team." },
        message: { type: "string", required: true, description: "What to post." },
      },
      async execute(input, ctx) {
        const where = await postTeamsMessage(ctx, input, { announce: announceVia(ctx) });
        return `posted to ${where}`;
      },
    },
  ],
};

export default plugin;
