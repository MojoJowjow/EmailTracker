# EmailTracker

Continuously monitors a shared Microsoft 365 mailbox and displays tracked correspondence in a local web dashboard.

- Tracks **sender, recipients, subject, timestamp** for every inbound and outbound message in the shared mailbox.
- Links inbound messages to the replies and forwards that go out afterward via Graph `conversationId` — click a row to see the entire back-and-forth.
- **User-defined filter rules** by sender pattern and/or subject pattern. Rules are applied at display time (history is never lost).
- No LLM / cloud AI — purely a deterministic tracker running locally.

## Requirements

- Windows / macOS / Linux
- Python 3.11+
- A shared mailbox on Microsoft 365 that the signed-in user has access to
- An Entra ID (Azure AD) **app registration** in the tenant — see setup below

## First-time setup

### 1. Register an app in Microsoft Entra ID

1. Sign in to <https://entra.microsoft.com> and go to **Identity → Applications → App registrations → New registration**.
2. Name: `EmailTracker` (anything is fine).
3. **Supported account types:** *Accounts in this organizational directory only*.
4. **Redirect URI:** leave blank for now (device-code flow doesn't need one).
5. Click **Register**.
6. On the overview page, copy the **Application (client) ID** and **Directory (tenant) ID** — these go into `.env`.
7. Go to **Authentication** → enable **Allow public client flows: Yes** → Save. (Required for device-code sign-in.)
8. Go to **API permissions → Add a permission → Microsoft Graph → Delegated permissions**, and add:
   - `Mail.Read`
   - `Mail.Read.Shared`
   - `offline_access` (added automatically by MSAL)
9. If you're a tenant admin, click **Grant admin consent** so users don't have to consent individually.

> If your IT team manages app registrations, send them this section and ask for the Client ID + Tenant ID in return.

### 2. Confirm access to the shared mailbox

The signed-in user must have **Read** permission on the shared mailbox in Exchange (the "Add another mailbox" access you normally need to open it in Outlook). Without that, Graph will return 403 even with `Mail.Read.Shared`.

### 3. Install and configure

```bash
# Clone / open the project folder
cd "ERC Project/EmailTracker"

# Create a virtualenv and install
py -m venv .venv
.venv/Scripts/python.exe -m pip install --upgrade pip
.venv/Scripts/python.exe -m pip install -e .

# Create your local config
cp .env.example .env
```

Edit `.env` and fill in:

- `TENANT_ID` — the Directory (tenant) ID from step 1
- `CLIENT_ID` — the Application (client) ID from step 1
- `SHARED_MAILBOX` — the UPN of the shared mailbox, e.g. `support@erc.ph`

Leave the other values at their defaults unless you need to change them.

### 4. Run the service

```bash
.venv/Scripts/python.exe -m uvicorn emailtracker.web.app:app --host 127.0.0.1 --port 8000
```

**First run only:** the terminal will print a Microsoft sign-in URL and a device code, e.g.

```
EmailTracker: sign-in required
To sign in, use a web browser to open the page https://microsoft.com/devicelogin and enter the code ABCD-1234 to authenticate.
```

Open that URL in a browser, enter the code, and sign in as a user who has access to the shared mailbox. After you finish, the poller starts immediately and subsequent runs won't prompt again (credentials are cached in `.token_cache.json`).

Open <http://127.0.0.1:8000> in your browser — the dashboard should load and populate within ~60 seconds.

## Using the dashboard

- **Live feed** (left panel) — newest messages first. Direction `IN` = arrived in the shared inbox, `OUT` = sent from the shared inbox. Auto-refreshes every 10 seconds. Click any row to open the full conversation.
- **Ad-hoc search box** — type to filter by sender, subject, or recipient. Not saved.
- **Filter rules panel** (right panel) — create saved rules to focus the feed:
  - *Sender contains* — substring, or use `*` as a wildcard, e.g. `*@acme.com`
  - *Subject contains* — substring, or with `*`, e.g. `*invoice*`
  - Matching is case-insensitive. Within a rule, sender and subject are AND'd. Across rules, matches are OR'd (a message shows if any enabled rule matches it).
  - Disable a rule with the **Disable** button; delete with **Delete**. With no rules enabled, the feed shows every tracked message.
- **Status bar** — poller health, tracked message count, last successful sync timestamp.

## How it works

- A single process runs FastAPI + a background asyncio task (the poller) together.
- The poller calls Microsoft Graph **delta queries** against the shared mailbox's Inbox and Sent Items folders every `POLL_INTERVAL_SECONDS` (default 60s). Delta queries return only what changed since the last sync token, so normal operation is very cheap — most ticks download zero messages.
- Each tracked message is stored in a local SQLite database (`emailtracker.db`). Only metadata is persisted — bodies and attachments are never downloaded.
- If the delta token expires (Graph returns 410), the poller automatically does a full resync of that folder on the next tick.
- On Graph errors, the poller logs the error to the status bar, backs off exponentially (capped at 5 minutes), and keeps running.

## Data stored

SQLite file `emailtracker.db`, three tables:

- `messages` — one row per tracked email (id, conversationId, direction, sender, recipients, subject, timestamps)
- `sync_state` — per-folder delta link, last successful sync, last error
- `filter_rules` — your saved focus rules

All writes happen inside transactions. Back up the file any time the service is stopped — it's the entire state.

## Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| Poller status is red, `last_error` mentions `AADSTS` / `invalid_client` | `.env` has the wrong `TENANT_ID` / `CLIENT_ID`, or "Allow public client flows" is not enabled in the app registration. |
| `403 Forbidden` from Graph | The signed-in user doesn't have mailbox access. Have an admin add them in Exchange, or grant `Mail.Read.Shared` admin consent. |
| Feed stays empty after first run | The poller hasn't completed its first tick yet (wait ~60s), or the shared mailbox is genuinely empty. Check `GET /status`. |
| `.token_cache.json` deleted or corrupted | The next run will re-prompt for device-code sign-in. Safe to delete and retry. |
| Port 8000 already in use | Change `WEB_PORT` in `.env`, or pass `--port NNNN` to `uvicorn`. |

## Stopping the service

`Ctrl+C` in the terminal. The poller's current tick finishes (up to a few seconds), then the DB and HTTP client close cleanly.
