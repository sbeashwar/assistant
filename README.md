# Personal Assistant — out-of-repo deps

Source for the local services that surround the OneDrive personal-assistant workspace and the MySuperApp PWA. **No secrets in this repo** — token caches, push tokens, VAPID keys, and `.env.local` are excluded by `.gitignore` and must be supplied separately when bootstrapping a new machine.

## Architecture

The personal assistant is general-purpose: the chat surface (PWA at `https://assistant.tail4621a0.ts.net`) spawns `claude -p` against the OneDrive workspace, with the `outlook-email` MCP attached. Anything the user asks — travel, calendar, email triage, web research, code lookup — flows through that one chat or its async siblings:

| Surface | Trigger | Notification |
|---|---|---|
| PWA chat | user message | push-on-completion deep-links to `/assistant?c=<conv>` |
| `RemindersScanner` (hourly) | trigger condition met (judged by `claude -p`) | Web Push to subscribed devices |
| Direct `claude` CLI on PC | terminal | none |

There are **no** standing per-domain scrapers. The travel skill, for instance, keeps a local cache under `notes/travel/` and uses a freshness rule (24 h for departures within 14 days, 7 d otherwise) — the assistant refreshes the Outlook **Travel** folder on demand when a query needs newer data.

## Bootstrap on a new machine

```powershell
# 1. Toolchain (winget will UAC-prompt on first install)
winget install -e --id OpenJS.NodeJS.LTS
winget install -e --id Python.Python.3.10
winget install -e --id GitHub.cli
winget install -e --id Tailscale.Tailscale
npm install -g pnpm
npm install -g @anthropic-ai/claude-code

# 2. Clone this repo + the PWA
git clone https://github.com/sbeashwar/assistant.git    C:\git\assistant
git clone https://github.com/sbeashwar/aisuperapp.git   C:\git\MySuperApp

# 3. Python venv for the email MCP + reminders scanner
python -m venv C:\git\venvs\outlook-email
C:\git\venvs\outlook-email\Scripts\python.exe -m pip install msal requests mcp pydantic httpx python-dotenv

# 4. Wait for OneDrive to sync C:\Users\<you>\OneDrive\Assistant\

# 5. Restore secrets (NOT in this repo — keep a separate encrypted bundle)
#    - apps/web/.env.local                  (VAPID keys, AUTH_SECRET, PUSH_INTERNAL_TOKEN)
#    - mcp-servers/outlook-email/.token_cache/token_cache.json   (MSAL refresh token)
#    - mcp-servers/outlook-email/push_config.json                (PUSH_INTERNAL_TOKEN + fire URLs)

# 6. Copy runtime locations
robocopy C:\git\assistant\mcp-servers       C:\git\mcp-servers       /E
robocopy C:\git\assistant\local-tool-server C:\Users\<you>\OneDrive\Assistant\local-tool-server /E

# 7. Install + build
cd C:\Users\<you>\OneDrive\Assistant\local-tool-server && pnpm install
cd C:\git\MySuperApp && pnpm install
pnpm --filter @mysuperapp/web build

# 8. Join Tailscale (preserves https://assistant.tail4621a0.ts.net)
tailscale up
tailscale set --hostname=assistant
tailscale serve --bg --https 443 http://localhost:3000

# 9. Register scheduled tasks (see scheduled-tasks/README.md for the registration script)
powershell -File scheduled-tasks\register.ps1

# 10. Verify
powershell C:\Users\<you>\OneDrive\Assistant\local-tool-server\keep-alive.ps1
curl http://localhost:3000/
```

## What's in this repo

| Path                            | What it is                                              |
| ------------------------------- | ------------------------------------------------------- |
| `mcp-servers/outlook-email/`    | Python MSAL+Graph MCP server: mail tools (list_inbox, search_email, send_email, …) + calendar tools (list_calendar_events, search_calendar_events, get_calendar_event). config.json holds public client_id only. |
| `mcp-servers/reminders/`        | Hourly proactive-reminder scanner. For each `notes/reminders/*.md` with `status: active`, spawns `claude -p` to judge the trigger condition using all available tools (email/calendar/web/files/etc.). Fires Web Push on `fired=true`. |
| `local-tool-server/`            | Express server on :3100 (file/command tools) + `keep-alive.ps1` that restarts Tailscale, the tool server, and Next.js every 30 minutes. |
| `scheduled-tasks/`              | PowerShell scripts to register `AssistantKeepAlive` and `RemindersScanner`. |

## What is NOT in this repo

| File                                                              | Where it lives                                                   |
| ----------------------------------------------------------------- | ---------------------------------------------------------------- |
| `apps/web/.env.local` (VAPID, AUTH_SECRET, PUSH_INTERNAL_TOKEN)   | Old machine; keep in an encrypted bundle.                        |
| `mcp-servers/outlook-email/.token_cache/*.json`                   | MSAL refresh tokens. Re-auth via device-code if lost.            |
| `mcp-servers/outlook-email/push_config.json`                      | Same bearer as `PUSH_INTERNAL_TOKEN`. Recreate from `.env.local`.|
| `C:\Users\<you>\OneDrive\Assistant\`                              | OneDrive auto-sync. CLAUDE.md, notes, reminders, skills, workspace context. |
| `C:\git\MySuperApp\`                                              | github.com/sbeashwar/aisuperapp                                  |

## MCP scopes (Outlook personal account)

`mcp-servers/outlook-email/server.py` requests these Graph scopes via MSAL device-code flow:

- `Mail.Read`, `Mail.ReadWrite`, `Mail.Send` — inbox + folders + send/reply
- `Calendars.Read` — list/search/read calendar events

When scopes change, the first call after restart re-prompts for device-code consent. The cached token at `.token_cache/token_cache.json` then covers all approved scopes.

## Related repos

- [sbeashwar/aisuperapp](https://github.com/sbeashwar/aisuperapp) — Next.js 14 PWA. Has its own setup README.
- OneDrive workspace at `OneDrive\Assistant\` (synced, not a repo) — has its own `CLAUDE.md` and skill docs.
