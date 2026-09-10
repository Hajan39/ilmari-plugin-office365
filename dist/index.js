// ilmari-plugin-office365 — Microsoft 365 (Graph) for ilmari.
//
// One file on purpose: ilmari hash-pins the entry file it approved, so keeping
// everything here means every line the daemon runs is covered by that pin.
//
// Auth is delegated (device code), not client credentials: acting *as the
// signed-in user* is the point — mail leaves from their own mailbox and lands
// in their Sent Items — and a public client needs no client secret to store.

import { Buffer } from "node:buffer";
import { createHash, randomBytes } from "node:crypto";
import { connect as netConnect } from "node:net";
import { connect as tlsConnect } from "node:tls";

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
/** An unanswered relay must not hang a workflow step for ever. */
const TIMEOUT_MS = 30_000;

const configDecl = {
  clientId: {
    label: "Application (client) ID",
    description:
      "The app registration ilmari signs in through: Azure Portal -> Microsoft Entra ID -> App registrations -> your app -> Overview. It must be a public client with device code flow allowed. No client secret is needed.",
    env: "OFFICE365_CLIENT_ID",
  },
  signIn: {
    options: ["device", "browser", "app", "smtp"],
    label: "Sign-in method",
    description:
      "device (default), browser, app or smtp. smtp skips Graph altogether and hands the mail to a relay, which is the way out of a tenant that consents to nothing — the email step and channel keep working, the calendar, file and Teams steps do not. app is the unattended one: ilmari signs in as the application itself with a client secret, once, with no browser and no per-workflow prompt — it needs application (not delegated) permissions consented by an administrator, and a mailbox to act as. Use browser when Conditional Access blocks the device-code flow: the task log shows a sign-in link to open in your own (managed) browser; it ends on an unreachable http://localhost page — paste that page's address, or just its code=..., into 'Authorization code' while the step is waiting.",
    env: "OFFICE365_SIGN_IN",
  },
  clientSecret: {
    label: "Client secret",
    description:
      "Only for the app sign-in: a client secret on the app registration (Certificates & secrets). Stored encrypted. Note it expires — Entra caps a secret at 24 months.",
    secret: true,
    env: "OFFICE365_CLIENT_SECRET",
  },
  mailbox: {
    label: "Act as mailbox",
    description:
      "Only for the app sign-in: the account whose mailbox, calendar and drive the steps act on (user@corp.com), since an application has no mailbox of its own. Have an administrator restrict the app to this one mailbox with an application access policy — Mail.Send as an application otherwise reaches every mailbox in the tenant.",
    env: "OFFICE365_MAILBOX",
  },
  authCode: {
    label: "Authorization code",
    description:
      "Only for the browser sign-in: paste the code (or the whole http://localhost?code=... address) here while a step waits for it. Single-use; cleared automatically once exchanged.",
    secret: true,
  },
  scopes: {
    label: "Permissions",
    options: [
      "offline_access",
      "Mail.Send",
      "Calendars.Read",
      "Files.Read.All",
      "Chat.ReadWrite",
      "ChannelMessage.Send",
    ],
    multiple: true,
    description:
      "The delegated permissions to ask for; none picked means all of them. Each one costs a surface when left out: Mail.Send the email step and channel, Calendars.Read the calendar step, Files.Read.All the file tools, Chat.ReadWrite and ChannelMessage.Send the Teams step. offline_access is what lets a sign-in survive a restart. Tenants that refuse user consent need an administrator to approve whatever is picked here.",
    env: "OFFICE365_SCOPES",
  },
  smtpHost: {
    label: "SMTP server",
    description:
      "Only for the smtp sign-in: the relay that accepts the mail, e.g. smtp.office365.com or a company relay. Microsoft 365 has SMTP AUTH switched off by default, so smtp.office365.com works only once an administrator turns it on for the mailbox.",
    env: "OFFICE365_SMTP_HOST",
  },
  smtpPort: {
    label: "SMTP port",
    description: "587 for STARTTLS (the default), 465 for implicit TLS, 25 for an internal relay.",
    env: "OFFICE365_SMTP_PORT",
  },
  smtpUser: {
    label: "SMTP username",
    description:
      "Usually the full email address. Leave empty for an internal relay that takes unauthenticated mail from your network.",
    env: "OFFICE365_SMTP_USER",
  },
  smtpPassword: {
    label: "SMTP password",
    description:
      "The password or app password for that user. Stored encrypted, and never sent over a connection the relay has not encrypted first.",
    secret: true,
    env: "OFFICE365_SMTP_PASSWORD",
  },
  smtpFrom: {
    label: "From address",
    description:
      "The sender the relay puts on every message. Most relays insist it match the authenticated user; empty falls back to the username.",
    env: "OFFICE365_SMTP_FROM",
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

  // the application signs in as itself: one consented registration, no browser
  // and no per-workflow prompt, so an unattended channel can send too
  if ((cfg.signIn ?? "").trim() === "app") {
    const clientSecret = (cfg.clientSecret ?? "").trim();
    if (!clientSecret) throw new Error("office365: the app sign-in needs a client secret");
    if (!tenantId) throw new Error("office365: the app sign-in needs a directory (tenant) ID");
    const token = await postForm(`${authBase(tenantId)}/token`, {
      grant_type: "client_credentials",
      client_id: clientId,
      client_secret: clientSecret,
      scope: "https://graph.microsoft.com/.default",
    });
    if (!token.access_token) {
      throw new Error(`app sign-in failed: ${token.error_description ?? token.error}`);
    }
    tokenCache.set(key, {
      token: token.access_token,
      expiresAt: Date.now() + (token.expires_in ?? 3600) * 1000,
    });
    return token.access_token;
  }

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

/** An application has no mailbox, calendar or drive of its own, so every /me
 *  path has to name the account it acts as instead. */
export function actingPath(cfg, path) {
  if ((cfg?.signIn ?? "").trim() !== "app") return path;
  const mailbox = (cfg.mailbox ?? "").trim();
  if (!mailbox) throw new Error("office365: the app sign-in needs a mailbox to act as");
  return path.replace(/^\/me\b/, `/users/${encodeURIComponent(mailbox)}`);
}

/** One Graph call as the signed-in user. Returns the parsed body, or
 *  undefined for the 202/204 answers sendMail and message posts give. */
export async function graph(ctx, method, path, body, opts = {}) {
  const token = await graphToken(ctx, opts);
  if (!token) return undefined;
  path = actingPath(ctx.pluginConfig(PLUGIN, opts.project), path);
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

/** Hands the same mail to a relay instead of to Graph. The attachments are
 *  already Graph fileAttachments by the time they get here, so their bytes
 *  come back out of the base64 the Graph shape carries them in. */
async function relayMail(cfg, { to, cc, subject, body, html, attachments }) {
  const host = text(cfg.smtpHost).trim();
  if (!host) throw new Error("the smtp sign-in needs an SMTP server");
  const user = text(cfg.smtpUser).trim();
  const from = text(cfg.smtpFrom).trim() || user;
  if (!from) throw new Error("the smtp sign-in needs a from address");
  const toList = list(to);
  const ccList = list(cc);
  const message = buildMessage({
    from,
    to: toList,
    cc: ccList,
    subject: subject || "(no subject)",
    body,
    html,
    attachments: attachments.map((a) => ({
      name: a.name,
      content: Buffer.from(a.contentBytes, "base64"),
    })),
  });
  return await smtpSend({
    host,
    port: Number(text(cfg.smtpPort).trim() || 587),
    user,
    password: text(cfg.smtpPassword),
    from,
    recipients: [...toList, ...ccList],
    message,
  });
}

/** The one place a mail is built and posted, shared by the node type, the
 *  agent tool and the notification channel. */
async function sendMail(ctx, { to, cc, subject, body, html, attachments = [] }, opts) {
  const cfg = ctx.pluginConfig(PLUGIN, opts?.project) ?? {};
  if (text(cfg.signIn).trim() === "smtp") {
    return await relayMail(cfg, { to, cc, subject, body, html, attachments });
  }
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
  const path = actingPath(
    ctx.pluginConfig(PLUGIN, opts?.project),
    `/me/drive/items/${encodeURIComponent(itemId)}/content`,
  );
  const res = await fetch(`${GRAPH}${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Graph GET ${path}: ${res.status}`);
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

// The SMTP client below is lifted from ilmari-plugin-smtp (Apache-2.0, same
// author): one file per plugin is the hash-pin rule, so it is copied rather
// than imported. Keep the two in sync when either is fixed.
// ------------------------------------------------------------ smtp client

/** One SMTP conversation over an already-open socket: read replies, write
 *  commands, and upgrade the socket in place when STARTTLS is used. */
function dialogue(socket) {
  let buffer = "";
  let waiter = null;

  const onData = (chunk) => {
    buffer += chunk.toString("utf8");
    deliver();
  };
  /** A reply ends at the line whose fourth character is a space ("250 ok"),
   *  not at the first newline — "250-SIZE" lines are continuations. */
  const deliver = () => {
    if (!waiter) return;
    const lines = buffer.split("\r\n");
    const end = lines.findIndex((l) => /^\d{3} /.test(l));
    if (end === -1) return;
    const reply = lines.slice(0, end + 1).join("\n");
    buffer = lines.slice(end + 1).join("\r\n");
    const { resolve: done } = waiter;
    waiter = null;
    done({ code: Number(reply.slice(0, 3)), text: reply });
  };

  let current = socket;
  current.on("data", onData);

  return {
    get socket() {
      return current;
    },
    /** Hands the data listener to the TLS socket STARTTLS produced. */
    upgrade(next) {
      current.removeListener("data", onData);
      buffer = "";
      current = next;
      current.on("data", onData);
    },
    read() {
      return new Promise((res, rej) => {
        waiter = { resolve: res };
        deliver();
        const timer = setTimeout(() => rej(new Error("SMTP timed out waiting for a reply")), TIMEOUT_MS);
        const clear = () => clearTimeout(timer);
        const wrapped = waiter;
        if (wrapped) {
          const original = wrapped.resolve;
          wrapped.resolve = (value) => {
            clear();
            original(value);
          };
        }
      });
    },
    write(line) {
      current.write(`${line}\r\n`);
    },
    end() {
      current.end();
    },
  };
}

async function expect(chat, codes, what) {
  const reply = await chat.read();
  if (!codes.includes(reply.code)) throw new Error(`${what}: ${reply.text.trim()}`);
  return reply;
}

async function command(chat, line, codes, what) {
  chat.write(line);
  return await expect(chat, codes, what);
}

function openSocket({ host, port, implicitTls, rejectUnauthorized }) {
  return new Promise((res, rej) => {
    const socket = implicitTls
      ? tlsConnect({ host, port, servername: host, rejectUnauthorized }, () => res(socket))
      : netConnect({ host, port }, () => res(socket));
    socket.setTimeout(TIMEOUT_MS);
    socket.once("error", rej);
    socket.once("timeout", () => {
      socket.destroy();
      rej(new Error(`SMTP connection to ${host}:${port} timed out`));
    });
  });
}

/**
 * Delivers one already-built message. Returns the accepted recipients.
 *
 * `requireTls` is what keeps a password off the wire in the clear: unless the
 * connection is already TLS, the server must offer STARTTLS before AUTH.
 */
export async function smtpSend(options) {
  const {
    host,
    port,
    user,
    password,
    from,
    recipients,
    message,
    rejectUnauthorized = true,
    requireTls = true,
  } = options;
  const implicitTls = Number(port) === 465;
  const socket = await openSocket({ host, port: Number(port), implicitTls, rejectUnauthorized });
  let chat = dialogue(socket);
  let secure = implicitTls;
  try {
    await expect(chat, [220], `SMTP ${host} greeting`);
    let greeting = await command(chat, `EHLO ${clientName(from)}`, [250], "EHLO");

    if (!secure && /STARTTLS/i.test(greeting.text)) {
      await command(chat, "STARTTLS", [220], "STARTTLS");
      const upgraded = await new Promise((res, rej) => {
        const t = tlsConnect({ socket, servername: host, rejectUnauthorized }, () => res(t));
        t.once("error", rej);
      });
      chat.upgrade(upgraded);
      secure = true;
      // the capability list is renegotiated over the encrypted channel
      greeting = await command(chat, `EHLO ${clientName(from)}`, [250], "EHLO after STARTTLS");
    }

    if (user) {
      if (!secure && requireTls) {
        throw new Error(
          `${host}:${port} offered no STARTTLS — refusing to send the password over an unencrypted connection`,
        );
      }
      await authenticate(chat, greeting.text, user, password ?? "");
    }

    await command(chat, `MAIL FROM:<${from}>`, [250], "MAIL FROM");
    const accepted = [];
    for (const rcpt of recipients) {
      const reply = await command(chat, `RCPT TO:<${rcpt}>`, [250, 251], `RCPT TO <${rcpt}>`);
      if (reply.code === 250 || reply.code === 251) accepted.push(rcpt);
    }
    await command(chat, "DATA", [354], "DATA");
    chat.write(`${dotStuff(message)}\r\n.`);
    await expect(chat, [250], "message body");
    chat.write("QUIT");
    return accepted;
  } finally {
    chat.end();
  }
}

/** AUTH PLAIN when offered (one round trip), else AUTH LOGIN, which is what
 *  older Exchange and many internal relays speak. */
async function authenticate(chat, capabilities, user, password) {
  const b64 = (s) => Buffer.from(s, "utf8").toString("base64");
  if (/AUTH[ =-][^\n]*PLAIN/i.test(capabilities)) {
    await command(chat, `AUTH PLAIN ${b64(`\0${user}\0${password}`)}`, [235], "AUTH PLAIN");
    return;
  }
  await command(chat, "AUTH LOGIN", [334], "AUTH LOGIN");
  await command(chat, b64(user), [334], "AUTH LOGIN username");
  await command(chat, b64(password), [235], "AUTH LOGIN password");
}

/** A bare "." on its own line ends the DATA block, so any line that already
 *  starts with one has to be doubled. */
const dotStuff = (body) => body.replace(/\r\n\./g, "\r\n..");

const clientName = (from) => from.split("@")[1] || "localhost";

/** RFC 2047, so a Czech subject does not arrive as mojibake. */
const encodeHeader = (value) =>
  /^[\x20-\x7e]*$/.test(value)
    ? value
    : `=?UTF-8?B?${Buffer.from(value, "utf8").toString("base64")}?=`;

const base64Lines = (buf) => (buf.toString("base64").match(/.{1,76}/g) ?? []).join("\r\n");

/** Builds the RFC 5322 message: plain single part, or multipart/mixed when
 *  there are attachments. Everything is base64 so no relay has to be trusted
 *  with 8-bit or long lines. */
export function buildMessage({ from, to, cc, subject, body, html, attachments = [], date }) {
  const boundary = `ilmari-${Math.random().toString(36).slice(2)}-${Date.now().toString(36)}`;
  const headers = [
    `From: ${from}`,
    `To: ${to.join(", ")}`,
    ...(cc.length > 0 ? [`Cc: ${cc.join(", ")}`] : []),
    `Subject: ${encodeHeader(subject)}`,
    `Date: ${(date ?? new Date()).toUTCString()}`,
    `Message-ID: <${boundary}@${clientName(from)}>`,
    "MIME-Version: 1.0",
  ];
  const bodyPart = [
    `Content-Type: text/${html ? "html" : "plain"}; charset=UTF-8`,
    "Content-Transfer-Encoding: base64",
    "",
    base64Lines(Buffer.from(body, "utf8")),
  ];

  if (attachments.length === 0) {
    return [...headers, ...bodyPart].join("\r\n");
  }
  const parts = [
    `--${boundary}`,
    ...bodyPart,
    ...attachments.flatMap((a) => [
      `--${boundary}`,
      `Content-Type: application/octet-stream; name="${encodeHeader(a.name)}"`,
      "Content-Transfer-Encoding: base64",
      `Content-Disposition: attachment; filename="${encodeHeader(a.name)}"`,
      "",
      base64Lines(a.content),
    ]),
    `--${boundary}--`,
  ];
  return [...headers, `Content-Type: multipart/mixed; boundary="${boundary}"`, "", ...parts].join(
    "\r\n",
  );
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
        // the smtp sign-in reaches no Graph app at all, so it is the relay
        // that has to be configured, not a client id
        const configured = text(cfg.signIn).trim() === "smtp" ? cfg.smtpHost : cfg.clientId;
        if (!configured || !cfg.to) return; // unconfigured — stay silent
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
