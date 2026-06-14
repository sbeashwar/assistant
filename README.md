# Personal Assistant — out-of-repo deps

Source for the local services that surround the OneDrive personal-assistant workspace and the MySuperApp PWA. **No secrets in this repo** — token caches, push tokens, VAPID keys, and `.env.local` are excluded by `.gitignore` and must be supplied separately when bootstrapping a new machine.

## Bootstrap on a new machine

```powershell
# 1. Toolchain (run as administrator the first time so winget can elevate)
winget install -e --id OpenJS.NodeJS.LTS
winget install -e --id Python.Python.3.10
winget install -e --id GitHub.cli
winget install -e --id Tailscale.Tailscale
npm install -g pnpm
npm install -g @anthropic-ai/claude-code

# 2. Clone this repo + MySuperApp
git clone https://github.com/sbeashwar/assistant.git C:\git\assistant
git clone https://github.com/sbeashwar/aisuperapp.git C:\git\MySuperApp

# 3. Python venv for MCP servers
python -m venv C:\git\venvs\outlook-email
C:\git\venvs\outlook-email\Scripts\python.exe -m pip install msal requests mcp pydantic httpx python-dotenv

# 4. Wait for OneDrive to sync C:\Users\<you>\OneDrive\Assistant\

# 5. Restore secrets (NOT in this repo — keep a separate encrypted bundle)
#    - apps/web/.env.local                  (VAPID keys, AUTH_SECRET, PUSH_INTERNAL_TOKEN)
#    - mcp-servers/outlook-email/.token_cache/token_cache.json           (MSAL refresh token, read scope)
#    - mcp-servers/outlook-email/.token_cache/sunnapana_token_cache.json (MSAL refresh token, send scope)
#    - mcp-servers/outlook-email/push_config.json                        (PUSH_INTERNAL_TOKEN + fire URLs)

# 6. Copy mcp-servers + local-tool-server into the runtime locations
robocopy C:\git\assistant\mcp-servers       C:\git\mcp-servers       /E
robocopy C:\git\assistant\local-tool-server C:\Users\<you>\OneDrive\Assistant\local-tool-server /E

# 7. Install local-tool-server deps + build the web app
cd C:\Users\<you>\OneDrive\Assistant\local-tool-server && pnpm install
cd C:\git\MySuperApp && pnpm install
pnpm --filter @mysuperapp/web build

# 8. Join Tailscale, set hostname (preserves https://assistant.tail4621a0.ts.net)
tailscale up
tailscale set --hostname=assistant
tailscale serve --bg --https 443 http://localhost:3000

# 9. Register scheduled tasks
schtasks /create /tn AssistantKeepAlive   /xml scheduled-tasks\AssistantKeepAlive.xml
schtasks /create /tn TravelEmailScanner   /xml scheduled-tasks\TravelEmailScanner.xml

# 10. Verify
powershell C:\Users\<you>\OneDrive\Assistant\local-tool-server\keep-alive.ps1
curl http://localhost:3000/
```

## What's in this repo

| Path                       | What it is                                              |
| -------------------------- | ------------------------------------------------------- |
| `mcp-servers/outlook-email/` | Python MSAL+Graph MCP server (read inbox, send mail). config.json holds public client_id only. |
| `mcp-servers/sunnapana/`     | Email reply bot (`sunnapana.py`) + proactive reminders scanner (`reminders.py`). 30-min scheduled task. |
| `mcp-servers/travel-scanner/`| Background extractor turning inbox bookings into JSON under `notes/travel/bookings/`. Hourly scheduled task. |
| `local-tool-server/`         | Express server on :3100 (file/command tools) + `keep-alive.ps1` that restarts Tailscale, the tool server, and Next.js every 30 minutes. |
| `scheduled-tasks/`           | XML exports of the Task Scheduler entries (with user-specific SIDs stripped or templated). |

## What is NOT in this repo

| File                                                              | Where it lives                                                   |
| ----------------------------------------------------------------- | ---------------------------------------------------------------- |
| `apps/web/.env.local` (VAPID, AUTH_SECRET, PUSH_INTERNAL_TOKEN)   | Old machine; back up to an encrypted bundle.                     |
| `mcp-servers/outlook-email/.token_cache/*.json`                   | MSAL refresh tokens. Re-auth via device-code flow if lost.       |
| `mcp-servers/outlook-email/push_config.json`                      | Same bearer as `PUSH_INTERNAL_TOKEN`. Recreate from `.env.local`.|
| `mcp-servers/travel-scanner/state.json`                           | Per-machine LRU of processed message IDs. Auto-rebuilds.         |
| `C:\Users\<you>\OneDrive\Assistant\`                              | OneDrive auto-sync. Contains worklog, notes, reminders, skills.  |
| `C:\git\MySuperApp\`                                              | github.com/sbeashwar/aisuperapp                                  |

## Related repos

- [sbeashwar/aisuperapp](https://github.com/sbeashwar/aisuperapp) — the Next.js PWA (MySuperApp / SuNaPaNa). Has its own setup README.
- The personal OneDrive workspace at `OneDrive\Assistant\` syncs automatically; it has its own `CLAUDE.md`.
