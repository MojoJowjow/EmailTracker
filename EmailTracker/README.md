# EmailTracker

Continuously monitors a shared Microsoft Outlook mailbox and displays tracked correspondence in a local web dashboard.

- Tracks **sender, recipients, subject, timestamp** for every inbound and outbound message in the shared mailbox.
- Links inbound messages to the replies and forwards that go out afterward — click a row to see the entire back-and-forth.
- **User-defined filter rules** by sender pattern and/or subject pattern. Rules are applied at display time (history is never lost).
- Reads directly from your local Outlook client via COM — **no Azure AD, no app registration, no Microsoft Graph tokens**. If the mailbox shows up in your Outlook folder list, EmailTracker can read it.

## Requirements

- Windows 10/11
- Microsoft Outlook (Classic — "New Outlook" doesn't expose COM), already signed in with the shared mailbox visible in your folder list
- Python 3.11+

## How it works

Outlook exposes its running instance as a COM automation server. EmailTracker dispatches `Outlook.Application`, resolves the shared mailbox by its SMTP address, and enumerates the **Inbox** and **Sent Items** folders on a background thread every `POLL_INTERVAL_SECONDS` (default 60s). Only metadata is stored — bodies and attachments are never persisted.

This means **Outlook has to be running on the same machine as EmailTracker** while the service is polling. If you close Outlook, the poller will error and the dashboard will show a red status; reopen Outlook and it recovers on the next tick.

## Setup

```bash
# Open the project folder
cd "ERC Project/EmailTracker"

# Create a virtualenv and install
py -m venv .venv
.venv/Scripts/python.exe -m pip install --upgrade pip
.venv/Scripts/python.exe -m pip install -e .

# Create your local config
cp .env.example .env
```

Edit `.env` and set:

- `SHARED_MAILBOX` — the **primary SMTP address** of the shared mailbox, e.g. `support@erc.ph`. This must match what Outlook knows it as. (In Outlook, right-click the shared mailbox in the folder pane → Data File Properties to confirm.)

Optional tweaks:

- `INITIAL_SYNC_DAYS` — on first run, how far back to import messages (default 30 days). Increase if you want more history up front; decrease to start lean.
- `POLL_INTERVAL_SECONDS` — polling cadence (default 60).
- `WEB_HOST` / `WEB_PORT` — where the dashboard binds (default `127.0.0.1:8000`).

## Running

**Step 1: Make sure Outlook is open** on the same machine and the shared mailbox is visible in the folder list.

**Step 2: Start the service:**

```bash
.venv/Scripts/python.exe -m uvicorn emailtracker.web.app:app --host 127.0.0.1 --port 8000
```

On first start, Windows may pop up a security prompt asking whether to allow programmatic access to Outlook — click **Allow**. After that, open <http://127.0.0.1:8000> in your browser. The dashboard should populate within ~60 seconds.

## Using the dashboard

- **Live feed** (left panel) — newest messages first. Direction `IN` = arrived in the shared inbox, `OUT` = sent from the shared inbox. Auto-refreshes every 10 seconds. Click any row to open the full conversation.
- **Ad-hoc search box** — type to filter by sender, subject, or recipient. Not saved.
- **Filter rules panel** (right panel) — create saved rules to focus the feed:
  - *Sender contains* — substring, or use `*` as a wildcard, e.g. `*@acme.com`
  - *Subject contains* — substring, or with `*`, e.g. `*invoice*`
  - Matching is case-insensitive. Within a rule, sender and subject are AND'd. Across rules, matches are OR'd (a message shows if any enabled rule matches it).
  - Disable a rule with the **Disable** button; delete with **Delete**. With no rules enabled, the feed shows every tracked message.
- **Status bar** — poller health, tracked message count, last successful sync timestamp.

## Data stored

SQLite file `emailtracker.db`, three tables:

- `messages` — one row per tracked email (Outlook EntryID, ConversationID, direction, sender, recipients, subject, received time)
- `sync_state` — per-folder last-successful-sync watermark and last error
- `filter_rules` — your saved focus rules

All writes happen inside transactions. Back up the file any time the service is stopped — it's the entire state.

## Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| Status bar red, `last_error` says "Could not attach to Outlook" | Outlook isn't running. Open the Classic Outlook desktop client. |
| Status bar red, `last_error` says "could not resolve" | `SHARED_MAILBOX` in `.env` doesn't match the primary SMTP of the shared mailbox — or the mailbox isn't added to your Outlook profile. Verify via right-click → Data File Properties in Outlook. |
| Status bar red, `last_error` says "Read permission" | Your Outlook account can see the mailbox but doesn't have read rights on Inbox / Sent Items. Ask the mailbox owner to add you in Exchange. |
| Feed stays empty after first run | The poller hasn't completed its first tick yet (wait ~60s), or the shared mailbox has no mail in the last `INITIAL_SYNC_DAYS`. Check `GET /status`. |
| Outlook prompts "Allow programmatic access?" on every start | Happens when antivirus is managing Outlook's trust center. In Outlook → File → Options → Trust Center → Programmatic Access, adjust the setting (requires admin for the "never warn" option). |
| Port 8000 already in use | Change `WEB_PORT` in `.env`, or pass `--port NNNN` to `uvicorn`. |
| I'm on **New Outlook** and it doesn't work | New Outlook is a different app that doesn't expose COM. Switch back to Classic Outlook via the "New Outlook" toggle in the title bar, or use File → Options. |

## Stopping the service

`Ctrl+C` in the terminal. The poller's current tick finishes (a few seconds at most), then the DB closes cleanly. Your data stays in `emailtracker.db` for the next run.
