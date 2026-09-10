// Every test stubs global fetch: no network, no real tenant. The plugin reads
// its config through ctx.pluginConfig and writes the rotated refresh token
// through ctx.savePluginConfig, so a fake ctx is all the wiring needed.
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, test } from "node:test";
import plugin, { browserSignIn, dayRange, deviceCodeSignIn, forgetTokens } from "../dist/index.js";

const realFetch = globalThis.fetch;
afterEach(() => {
  globalThis.fetch = realFetch;
  forgetTokens();
});

/** Records every request and answers from `routes`, matched by URL substring. */
function stubFetch(routes) {
  const seen = [];
  globalThis.fetch = async (url, init = {}) => {
    const body = init.body ?? "";
    seen.push({ url: String(url), method: init.method ?? "GET", body, headers: init.headers ?? {} });
    for (const [match, answer] of routes) {
      if (String(url).includes(match)) {
        const res = typeof answer === "function" ? answer(String(url), body) : answer;
        return {
          ok: res.status === undefined || res.status < 400,
          status: res.status ?? 200,
          json: async () => res.json ?? {},
          text: async () => res.text ?? JSON.stringify(res.json ?? {}),
        };
      }
    }
    throw new Error(`unstubbed request: ${url}`);
  };
  return seen;
}

const TOKEN_OK = {
  json: { access_token: "at-1", refresh_token: "rt-2", expires_in: 3600 },
};

function ctxFor(overrides = {}) {
  const saved = [];
  const emitted = [];
  return {
    saved,
    emitted,
    taskId: "t1",
    taskTitle: "check it",
    workdir: overrides.workdir ?? process.cwd(),
    outputs: {},
    project: "CRV",
    emit: (type, data) => emitted.push({ type, data }),
    waitForEvent: async () => null,
    pluginConfig: () => ({
      clientId: "app-1",
      tenantId: "tenant-1",
      to: "team@corp.com",
      refreshToken: "rt-1",
      ...overrides.config,
    }),
    savePluginConfig: (name, values, project) => saved.push({ name, values, project }),
  };
}

const nodeOf = (type) => {
  const node = plugin.nodeTypes.find((n) => n.type === type);
  if (!node) throw new Error(`node type ${type} missing`);
  return node;
};
const toolOf = (name) => {
  const tool = plugin.tools.find((t) => t.name === name);
  if (!tool) throw new Error(`tool ${name} missing`);
  return tool;
};

test("the plugin declares what ilmari's static inspect reads", () => {
  assert.equal(plugin.name, "office365");
  assert.match(plugin.version, /^\d+\.\d+\.\d+$/);
  assert.deepEqual(plugin.capabilities, ["net", "fs", "secrets"]);
  assert.deepEqual(
    plugin.nodeTypes.map((n) => n.type),
    ["office365-email", "office365-calendar", "office365-teams-message"],
  );
  // every node type and tool the GUI renders needs its user-facing copy
  for (const node of plugin.nodeTypes) assert.ok(node.description, node.type);
  for (const tool of plugin.tools) assert.ok(tool.description, tool.name);
});

test("dayRange counts days back from today and lets from/to win", () => {
  const now = new Date("2026-09-10T09:00:00Z");
  assert.deepEqual(dayRange({}, now), { from: "2026-09-10", to: "2026-09-10" });
  assert.deepEqual(dayRange({ days: 7 }, now), { from: "2026-09-04", to: "2026-09-10" });
  assert.deepEqual(dayRange({ days: 30, from: "2026-09-01" }, now), {
    from: "2026-09-01",
    to: "2026-09-01",
  });
});

test("a stored refresh token is exchanged and the rotated one is saved", async () => {
  const seen = stubFetch([
    ["/oauth2/v2.0/token", TOKEN_OK],
    ["/me/sendMail", { status: 202 }],
  ]);
  const ctx = ctxFor();
  const out = await nodeOf("office365-email").run(
    { to: "a@corp.com", subject: "done", body: "all green" },
    ctx,
  );

  assert.deepEqual(out, { ok: true, output: "sent to a@corp.com" });
  assert.ok(seen[0].body.includes("grant_type=refresh_token"));
  // the rotated token goes back into the encrypted store, scoped to the project
  assert.deepEqual(ctx.saved, [
    { name: "office365", values: { refreshToken: "rt-2" }, project: "CRV" },
  ]);
  const mail = JSON.parse(seen[1].body);
  assert.equal(mail.message.subject, "done");
  assert.equal(mail.message.body.contentType, "Text");
  assert.deepEqual(mail.message.toRecipients, [{ emailAddress: { address: "a@corp.com" } }]);
  assert.equal(mail.saveToSentItems, true);
  assert.equal(seen[1].headers.Authorization, "Bearer at-1");
});

test("the access token is cached across steps", async () => {
  const seen = stubFetch([
    ["/oauth2/v2.0/token", TOKEN_OK],
    ["/me/sendMail", { status: 202 }],
  ]);
  const ctx = ctxFor();
  const send = () =>
    nodeOf("office365-email").run({ to: "a@corp.com", subject: "s", body: "b" }, ctx);
  await send();
  await send();
  assert.equal(seen.filter((r) => r.url.includes("/oauth2/v2.0/token")).length, 1);
});

test("an email step with no recipient and no default fails instead of sending", async () => {
  stubFetch([["/oauth2/v2.0/token", TOKEN_OK]]);
  const ctx = ctxFor({ config: { to: "" } });
  const out = await nodeOf("office365-email").run({ subject: "s", body: "b" }, ctx);
  assert.equal(out.ok, false);
  assert.match(out.reason, /no recipient/);
});

test("attachments are read from the workdir, and a path escaping it is refused", async () => {
  const wd = mkdtempSync(join(tmpdir(), "o365-"));
  writeFileSync(join(wd, "report.md"), "# report");
  const seen = stubFetch([
    ["/oauth2/v2.0/token", TOKEN_OK],
    ["/me/sendMail", { status: 202 }],
  ]);
  const ctx = ctxFor({ workdir: wd });
  const out = await nodeOf("office365-email").run(
    { to: "a@corp.com", subject: "s", body: "b", attach: "report.md, ../../secrets.txt" },
    ctx,
  );

  assert.equal(out.ok, true);
  assert.match(out.output, /skipped: \.\.\/\.\.\/secrets\.txt \(outside the working directory\)/);
  const mail = JSON.parse(seen[1].body);
  assert.equal(mail.message.attachments.length, 1);
  assert.equal(mail.message.attachments[0].name, "report.md");
  assert.equal(
    Buffer.from(mail.message.attachments[0].contentBytes, "base64").toString(),
    "# report",
  );
});

test("the calendar step expands pages into one JSON array", async () => {
  const page2 = "/me/calendarView?page=2";
  stubFetch([
    ["/oauth2/v2.0/token", TOKEN_OK],
    [
      "/me/calendarView",
      (url) =>
        url.includes("page=2")
          ? { json: { value: [event("Retro")] } }
          : {
              json: {
                value: [event("Stand-up")],
                "@odata.nextLink": `https://graph.microsoft.com/v1.0${page2}`,
              },
            },
    ],
  ]);
  const out = await nodeOf("office365-calendar").run({ days: 7 }, ctxFor());
  assert.equal(out.ok, true);
  assert.deepEqual(
    JSON.parse(out.output).map((e) => e.subject),
    ["Stand-up", "Retro"],
  );
});

function event(subject) {
  return {
    subject,
    start: { dateTime: "2026-09-10T09:00:00" },
    end: { dateTime: "2026-09-10T09:15:00" },
    organizer: { emailAddress: { address: "boss@corp.com" } },
    location: { displayName: "Teams" },
    webLink: "https://outlook.office.com/x",
  };
}

test("an empty calendar range is a success with an empty array", async () => {
  stubFetch([
    ["/oauth2/v2.0/token", TOKEN_OK],
    ["/me/calendarView", { json: { value: [] } }],
  ]);
  const out = await nodeOf("office365-calendar").run({}, ctxFor());
  assert.deepEqual(out, { ok: true, output: "[]" });
});

test("a Graph error fails the step readably, without echoing the token", async () => {
  stubFetch([
    ["/oauth2/v2.0/token", TOKEN_OK],
    ["/me/sendMail", { status: 403, text: '{"error":{"message":"Access denied"}}' }],
  ]);
  const out = await nodeOf("office365-email").run(
    { to: "a@corp.com", subject: "s", body: "b" },
    ctxFor(),
  );
  assert.equal(out.ok, false);
  assert.match(out.reason, /office365-email: Graph POST \/me\/sendMail: 403/);
  assert.ok(!out.reason.includes("at-1"));
});

test("a Teams message goes to a chat, or to a channel, or is refused", async () => {
  const seen = stubFetch([
    ["/oauth2/v2.0/token", TOKEN_OK],
    ["/messages", { status: 201, json: { id: "m1" } }],
  ]);
  const node = nodeOf("office365-teams-message");

  const chat = await node.run({ chat: "19:abc@thread.v2", message: "hi" }, ctxFor());
  assert.deepEqual(chat, { ok: true, output: "posted to chat 19:abc@thread.v2" });
  assert.ok(seen[1].url.endsWith("/chats/19%3Aabc%40thread.v2/messages"));

  const channel = await node.run({ team: "t1", channel: "c1", message: "hi" }, ctxFor());
  assert.equal(channel.output, "posted to channel c1");

  const neither = await node.run({ message: "hi" }, ctxFor());
  assert.equal(neither.ok, false);
  assert.match(neither.reason, /name either a chat id/);
});

test("the email channel stays silent when unconfigured and never blocks on sign-in", async () => {
  const channel = plugin.channels.find((c) => c.name === "email");
  // no config at all: nothing is sent, nothing throws
  await channel.send("task finished", { pluginConfig: () => ({}) });

  // configured but signed out: the refresh fails and no device code is started
  const seen = stubFetch([["/oauth2/v2.0/token", { json: { error: "invalid_grant" } }]]);
  await channel.send("task finished\nsecond line", {
    pluginConfig: () => ({ clientId: "app-1", to: "team@corp.com", refreshToken: "dead" }),
  });
  assert.ok(!seen.some((r) => r.url.includes("/devicecode")));
});

test("the channel sends the first line as the subject once signed in", async () => {
  const seen = stubFetch([
    ["/oauth2/v2.0/token", TOKEN_OK],
    ["/me/sendMail", { status: 202 }],
  ]);
  const channel = plugin.channels.find((c) => c.name === "email");
  await channel.send("task finished: PROJ-42\nbody line", {
    pluginConfig: () => ({ clientId: "app-1", to: "team@corp.com", refreshToken: "rt-1" }),
  });
  const mail = JSON.parse(seen[1].body);
  assert.equal(mail.message.subject, "task finished: PROJ-42");
  assert.equal(mail.message.body.content, "task finished: PROJ-42\nbody line");
});

test("device code sign-in waits through authorization_pending and slow_down", async () => {
  let call = 0;
  stubFetch([
    ["/devicecode", { json: { device_code: "dc", user_code: "ABCD", interval: 1, expires_in: 60 } }],
    [
      "/oauth2/v2.0/token",
      () => {
        call++;
        if (call === 1) return { json: { error: "authorization_pending" } };
        if (call === 2) return { json: { error: "slow_down" } };
        return TOKEN_OK;
      },
    ],
  ]);
  const said = [];
  const out = await deviceCodeSignIn("app-1", "tenant-1", (m) => said.push(m), async () => {});
  assert.equal(out.refreshToken, "rt-2");
  assert.equal(call, 3);
  assert.match(said[0], /ABCD/);
});

test("browser sign-in announces a PKCE link, waits for a fresh paste and exchanges it", async () => {
  const seen = stubFetch([["/oauth2/v2.0/token", TOKEN_OK]]);
  const said = [];
  let pasted = "old-code"; // left over from an earlier attempt: must be ignored
  let polls = 0;
  const readCode = () => {
    if (++polls === 3) pasted = "http://localhost/?code=fresh-code&session_state=x";
    return pasted;
  };
  const out = await browserSignIn("app-1", "tenant-1", (m) => said.push(m), readCode, async () => {});
  assert.equal(out.refreshToken, "rt-2");
  const link = new URL(said[0].match(/https:\S+/)[0]);
  assert.equal(link.searchParams.get("code_challenge_method"), "S256");
  assert.equal(link.searchParams.get("redirect_uri"), "http://localhost");
  const body = new URLSearchParams(seen.at(-1).body);
  assert.equal(seen.length, 1);
  assert.equal(body.get("grant_type"), "authorization_code");
  assert.equal(body.get("code"), "fresh-code");
  assert.equal(body.get("redirect_uri"), "http://localhost");
  const challenge = createHash("sha256").update(body.get("code_verifier")).digest("base64url");
  assert.equal(challenge, link.searchParams.get("code_challenge"));
});

test("device code sign-in surfaces a fatal error instead of looping", async () => {
  stubFetch([
    ["/devicecode", { json: { device_code: "dc", user_code: "ABCD", interval: 1, expires_in: 60 } }],
    ["/oauth2/v2.0/token", { json: { error: "expired_token", error_description: "code expired" } }],
  ]);
  await assert.rejects(
    () => deviceCodeSignIn("app-1", "tenant-1", () => {}, async () => {}),
    /code expired/,
  );
});

test("the file tools search and read as the signed-in user", async () => {
  stubFetch([
    ["/oauth2/v2.0/token", TOKEN_OK],
    [
      "/search(q=",
      {
        json: {
          value: [
            {
              id: "01ABC",
              name: "spec.md",
              size: 12,
              lastModifiedDateTime: "2026-09-01T10:00:00Z",
              webUrl: "https://sharepoint/spec.md",
              parentReference: { path: "/drive/root:/docs" },
            },
          ],
        },
      },
    ],
    ["/content", { text: "# spec" }],
  ]);
  const ctx = ctxFor();
  const found = JSON.parse(await toolOf("office365_find_files").execute({ query: "spec" }, ctx));
  assert.deepEqual(found[0], {
    id: "01ABC",
    name: "spec.md",
    size: 12,
    modified: "2026-09-01T10:00:00Z",
    path: "/drive/root:/docs",
    url: "https://sharepoint/spec.md",
  });
  assert.equal(await toolOf("office365_read_file").execute({ id: "01ABC" }, ctx), "# spec");
});

test("configured permissions replace the full scope set", async () => {
  const seen = stubFetch([
    ["/oauth2/v2.0/token", TOKEN_OK],
    ["/me/sendMail", { status: 202 }],
  ]);
  const ctx = ctxFor({ config: { scopes: "offline_access Mail.Send" } });
  await nodeOf("office365-email").run({ to: "a@corp.com", subject: "s", body: "b" }, ctx);
  assert.equal(new URLSearchParams(seen[0].body).get("scope"), "offline_access Mail.Send");
});

test("the app sign-in uses a client secret and acts on the configured mailbox", async () => {
  const seen = stubFetch([
    ["/oauth2/v2.0/token", { json: { access_token: "at-app", expires_in: 3600 } }],
    ["/sendMail", { status: 202 }],
  ]);
  const ctx = ctxFor({
    config: {
      signIn: "app",
      clientSecret: "sec-1",
      mailbox: "bot@corp.com",
      refreshToken: "", // no delegated sign-in to fall back on
    },
  });
  const out = await nodeOf("office365-email").run({ to: "a@corp.com", subject: "s", body: "b" }, ctx);

  assert.equal(out.ok, true);
  const auth = new URLSearchParams(seen[0].body);
  assert.equal(auth.get("grant_type"), "client_credentials");
  assert.equal(auth.get("client_secret"), "sec-1");
  assert.equal(auth.get("scope"), "https://graph.microsoft.com/.default");
  assert.match(seen[1].url, /\/users\/bot%40corp\.com\/sendMail$/);
  // an application token is not a sign-in to remember
  assert.deepEqual(ctx.saved, []);
});
