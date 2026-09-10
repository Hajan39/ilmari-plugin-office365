# ilmari-plugin-office365

Microsoft 365 for [ilmari](https://github.com/wraithyy/ilmari), acting **as the
signed-in user**: mail leaves from your own mailbox and lands in your Sent
Items, calendar and files are yours, Teams messages are posted under your name.

Sign-in is the OAuth **device code** flow — no client secret to store, no
service mailbox, no admin-consented app impersonating people. The first Graph
call prints a code and a link into the task log; you approve it once in a
browser and ilmari keeps a rotating refresh token in its encrypted config.

## What it contributes

| Kind | Name | Does |
|---|---|---|
| node | `office365-email` | Sends an email, optionally with files from the working directory attached |
| node | `office365-calendar` | Reads your calendar for a range of days as JSON |
| node | `office365-teams-message` | Posts into a Teams chat or channel |
| tool | `office365_calendar` | Same calendar read, callable by an agent mid-run |
| tool | `office365_find_files` | Searches OneDrive and the SharePoint files shared with you |
| tool | `office365_read_file` | Reads one of those files as text |
| tool | `office365_send_email` | Sends a mail the agent composed |
| tool | `office365_send_teams_message` | Posts a Teams message the agent composed |
| channel | `email` | Task notifications and agent questions, mailed to the default recipient |

### Capabilities

- `net` — Microsoft Graph (`graph.microsoft.com`) and the sign-in endpoint
  (`login.microsoftonline.com`). Nothing else is contacted.
- `fs` — reads the files an `office365-email` step is told to attach, and only
  from inside the task's working directory: a path that escapes it is refused
  and reported in the step's output rather than attached.
- `secrets` — the rotating refresh token is stored in ilmari's encrypted plugin
  config (never in the repo, never in an environment variable).

## Setup

### 1. Register an app in Microsoft Entra ID

Azure Portal → **Microsoft Entra ID** → **App registrations** → **New
registration**:

- Name: anything (`ilmari` does).
- Supported account types: *Accounts in this organizational directory only*.
- Redirect URI: leave empty.

Then, in the new app:

- **Authentication** → *Advanced settings* → **Allow public client flows: Yes**.
  Without this the device-code sign-in is rejected.
- **API permissions** → *Add a permission* → *Microsoft Graph* → **Delegated
  permissions**, and add what you actually want to use:

  | Permission | Needed for |
  |---|---|
  | `Mail.Send` | the email step, tool and channel |
  | `Calendars.Read` | the calendar step and tool |
  | `Files.Read.All` | the file search and read tools |
  | `Chat.ReadWrite` | posting into a Teams chat |
  | `ChannelMessage.Send` | posting into a Teams channel |

  Delegated permissions never grant more than the signed-in user already has.
  Some tenants require an admin to consent to them once.

- **Overview** → copy the **Application (client) ID** and **Directory (tenant)
  ID**.

### 2. Install the plugin

```bash
ilmari plugin add https://github.com/Hajan39/ilmari-plugin-office365
```

Approve the capabilities when asked. To test a local checkout instead:

```bash
ilmari plugin add /abs/path/ilmari-plugin-office365/dist/index.js
```

### 3. Configure it

In the GUI's **Plugins** screen, on **office365**:

| Field | Value |
|---|---|
| Application (client) ID | from the app registration's Overview |
| Directory (tenant) ID | from the same page; empty means "any work account" |
| Default recipient | where the `email` channel sends notifications |
| Sign-in method | `device` (default) or `browser`, see below |
| Authorization code | only for `browser`: pasted while a step waits for it |
| Sign-in token | leave empty — it fills itself in |

Every field has an environment fallback (`OFFICE365_CLIENT_ID`,
`OFFICE365_TENANT_ID`, `OFFICE365_TO`) for a headless install. Configuration can
also be scoped to one project, which needs ilmari with the project-aware node
context (`ctx.project`).

### 4. Sign in

Run any workflow with an office365 step. The task log shows:

```
Sign in to Microsoft at https://microsoft.com/devicelogin and enter the code F7XK3PQ2
```

Open it, sign in, and the step continues on its own. From then on the refresh
token is used silently; clear the **Sign-in token** field to sign out.

#### When Conditional Access blocks device code

Some tenants answer the device-code sign-in with *"You cannot access this right
now"*: the token request comes from the daemon, which is not a compliant
device. Set **Sign-in method** to `browser`. The step then logs a sign-in link
(authorization code + PKCE) to open in your own managed browser; after
signing in, the browser lands on an unreachable `http://localhost/?code=...`
page. Paste that address (or just the code) into **Authorization code** while
the step is still waiting — it polls for up to 15 minutes, exchanges the code
once and clears the field.

Prerequisite: `http://localhost` registered on the app as a *Mobile and
desktop applications* redirect URI. When you cannot register an app at all,
Microsoft's own public clients already have it — e.g. Microsoft Graph Command
Line Tools, client ID `14d82eec-204b-4c2f-b7e8-296a70dab67e` — and the
delegated scopes are then consented per user, if the tenant allows user
consent.

#### Signing in once, for the whole install

The two flows above sign in *a person*, in a browser, per install. A tenant
that refuses user consent, or an unattended daemon that must send mail with
nobody watching, wants the third: **Sign-in method** `app`. ilmari then signs
in as the application itself (client credentials), with no browser and no
prompt — one registration, consented once by an administrator.

It needs, on the app registration:

- a **client secret** (Certificates & secrets), pasted into **Client secret** —
  Entra expires it after at most 24 months, and a step then fails until it is
  replaced
- **application** permissions (not delegated), admin-consented: `Mail.Send`,
  `Calendars.Read`, `Files.Read.All`. Teams has no usable application
  equivalent here, so the Teams step stays on a delegated sign-in.
- **Act as mailbox** — an application owns no mailbox, calendar or drive, so
  every step acts on the account named here

`Mail.Send` as an application reaches *every* mailbox in the tenant. Have the
administrator scope the app to the one mailbox with an application access
policy (`New-ApplicationAccessPolicy`) — otherwise a bug, or anyone who can
edit a workflow, can send as anybody.

The **Permissions** field is ignored in this mode: application permissions are
whatever the registration was consented for, and the token is asked for with
`.default`.

## Using it

Report a task's result by mail:

```json
{
  "id": "mail",
  "type": "office365-email",
  "needs": ["deliver"],
  "to": "team@corp.com",
  "subject": "ilmari finished {{task}}",
  "body": "{{deliver.result}}",
  "attach": "report.md"
}
```

A timesheet workflow — what the calendar says next to what was reported:

```json
{ "id": "cal", "type": "office365-calendar", "days": 7 }
```

`{{cal.result}}` is a JSON array of `{ subject, start, end, allDay, organizer,
location, categories, url }`, with recurring meetings expanded into their
actual occurrences. An empty range is a success with `[]`, not a failure, so
`catches` only fires on a real error.

Post into Teams:

```json
{
  "id": "notify",
  "type": "office365-teams-message",
  "chat": "19:...@thread.v2",
  "message": "ilmari finished {{task}}: {{deliver.result}}"
}
```

Use `team` + `channel` instead of `chat` for a channel post. Chat and channel
ids come out of the Teams link for that conversation.

## Limits worth knowing

- Attachments: 3 MB in total per mail. Graph refuses a `sendMail` request over
  roughly 4 MB, so anything larger is skipped and named in the step's output
  rather than failing the send.
- `office365_read_file` reads text-shaped files. A `.docx` or `.pdf` comes back
  as bytes Graph will not convert; long files are truncated at 200 000
  characters so a file cannot blow up an agent's context window.
- The `email` channel never starts an interactive sign-in — a notification must
  not block for fifteen minutes. If nobody has signed in yet, it stays silent
  until a workflow step signs in.
- Delegated only: Graph does not let a signed-in user forge another sender, so
  there is no "send as somebody else".

## Development

```bash
node --test "test/*.test.mjs"
```

The tests stub `fetch`, so they need no tenant and no network. `dist/index.js`
is the shipped file and is committed — there is no build step, and ilmari
hash-pins exactly the file you approved.

Releases are git tags `vX.Y.Z`; `ilmari plugin outdated` compares against the
newest tag.

## License

Apache-2.0
