# Personal Assistant — out-of-repo deps

Source for the local services that surround the OneDrive personal-assistant workspace and the [aisuperapp](https://github.com/sbeashwar/aisuperapp) PWA. **No secrets in this repo** — token caches, push tokens, VAPID keys, and `.env.local` are excluded by `.gitignore` and supplied separately when bootstrapping a new machine.

---

## Architecture (TL;DR)

The personal assistant is general-purpose. The chat surface (PWA at `https://assistant.tail4621a0.ts.net`) spawns `claude -p` against the OneDrive workspace with the `outlook-email` MCP attached. Email/calendar/web/files are tools the assistant reaches for; **there are no standing per-domain scrapers**.

| Surface | Trigger | Notification path |
|---|---|---|
| PWA chat | user message | push-on-completion deep-links to `/assistant?c=<conv>` |
| `RemindersScanner` (hourly) | judge says `fired=true` | Web Push to subscribed devices |
| Direct `claude` CLI on PC | terminal | none |

Travel data uses a local cache under `notes/travel/` with a tiered freshness rule (24 h for departures within 14 days, 7 d otherwise). The assistant refreshes from the Outlook **Travel** folder on demand when a query needs newer data.

---

## Bootstrap on a new Windows machine

Read [BOOTSTRAP-GOTCHAS.md](BOOTSTRAP-GOTCHAS.md) first — concrete pitfalls discovered the last time this was done. The script below assumes you've read it.

### Phase 1 — Toolchain (~15 min, needs UAC prompts)

```powershell
# In an elevated PowerShell window, or just accept UAC prompts as they come
winget install -e --id OpenJS.NodeJS.LTS
winget install -e --id Python.Python.3.10           # NOT 3.14 — match this exactly
winget install -e --id GitHub.cli
winget install -e --id Tailscale.Tailscale

# pnpm + Claude CLI via npm (user scope, no UAC)
npm install -g pnpm
npm install -g @anthropic-ai/claude-code

# Verify
node --version       # >= v20
pnpm --version       # >= 9
claude --version
tailscale --version
gh --version
```

### Phase 2 — Git auth (one-time, blocks Entra Conditional Access)

**Do NOT use `gh auth login --web`** — `sbeashwar` is SSO-enforced through Microsoft's 1ES open-source enterprise org. The browser flow hits AADSTS53003 ("device unregistered") on any non-Entra-joined machine. See [BOOTSTRAP-GOTCHAS.md](BOOTSTRAP-GOTCHAS.md#entra-conditional-access-blocks-gh-auth-login).

Use SSH instead — bypasses Entra entirely, never expires:

```powershell
# Generate keypair
mkdir C:\Users\$env:USERNAME\.ssh -ErrorAction SilentlyContinue
ssh-keygen -t ed25519 -f "C:\Users\$env:USERNAME\.ssh\id_ed25519" -N '""' `
  -C "sbeashwar@$($env:COMPUTERNAME)-$(Get-Date -Format yyyyMMdd)"
Get-Content "C:\Users\$env:USERNAME\.ssh\id_ed25519.pub"     # copy this
```

Add the public key at https://github.com/settings/ssh/new (title: `<hostname> bootstrap <date>`), then verify:

```powershell
ssh -T git@github.com    # expect "Hi sbeashwar! You've successfully authenticated"
```

`gh` CLI is **optional** — only needed for `gh repo create`, `gh pr view`, etc. Skip it unless you'll do GitHub API work from this machine. If you do need it, generate a classic PAT with `repo` AND `read:org` scopes at https://github.com/settings/tokens and `gh auth login --with-token`.

### Phase 3 — Clone (use SSH remotes)

```powershell
mkdir C:\git -ErrorAction SilentlyContinue
cd C:\git
git clone git@github.com:sbeashwar/assistant.git
git clone git@github.com:sbeashwar/aisuperapp.git MySuperApp
git clone git@github.com:sbeashwar/music.git       # optional
```

### Phase 4 — Python venv (recreate, don't copy)

**Never copy a venv from another machine** — `pyvenv.cfg` hard-codes the Python executable path and `Scripts\python.exe` is a thin wrapper around it. Build fresh:

```powershell
python -m venv C:\git\venvs\outlook-email
C:\git\venvs\outlook-email\Scripts\python.exe -m pip install --upgrade pip
C:\git\venvs\outlook-email\Scripts\python.exe -m pip install `
  msal requests mcp pydantic httpx python-dotenv
```

### Phase 5 — Restore secrets (NOT in this repo)

Keep these in an encrypted bundle (1Password / Bitwarden / etc.). They are the ONLY pieces a clone-from-scratch can't recreate:

| File | Contents | Recovery if lost |
|---|---|---|
| `C:\git\MySuperApp\apps\web\.env.local` | VAPID public/private, AUTH_SECRET, PUSH_INTERNAL_TOKEN, FINNHUB_API_KEY | Regenerate VAPID via `npx web-push generate-vapid-keys` (**re-invalidates all phone subs**), regen AUTH_SECRET and PUSH_INTERNAL_TOKEN with `openssl rand`. |
| `C:\git\mcp-servers\outlook-email\.token_cache\token_cache.json` | MSAL refresh tokens for Outlook (mail + calendar scopes) | Delete; first MCP call triggers device-code reauth — visible on stderr. |
| `C:\git\mcp-servers\outlook-email\push_config.json` | Same bearer as `PUSH_INTERNAL_TOKEN` + fire URLs (`localhost:3000/api/push/fire`, `/api/assistant/fire`) | Hand-write — see [push_config.example.json](mcp-servers/outlook-email/push_config.example.json). |

### Phase 6 — Wire MCP into runtime location

```powershell
robocopy C:\git\assistant\mcp-servers       C:\git\mcp-servers       /E /XO
robocopy C:\git\assistant\local-tool-server C:\Users\$env:USERNAME\OneDrive\Assistant\local-tool-server /E /XO
```

Wait for OneDrive to finish syncing `C:\Users\<you>\OneDrive\Assistant\` — CLAUDE.md, notes, reminders, skills, slash commands all live there.

### Phase 7 — Install + build

```powershell
cd C:\Users\$env:USERNAME\OneDrive\Assistant\local-tool-server
pnpm install

cd C:\git\MySuperApp
pnpm install   # better-sqlite3 build will fail — IGNORE IT, app uses @libsql/client
pnpm --filter @mysuperapp/web build
```

### Phase 8 — Tailscale (preserves existing PWA URL)

```powershell
tailscale up                           # browser opens for SSO
tailscale set --hostname=assistant     # preserves https://assistant.tail4621a0.ts.net cert
tailscale serve --bg --https 443 http://localhost:3000
```

### Phase 9 — Scheduled tasks

```powershell
powershell -ExecutionPolicy Bypass -File C:\git\assistant\scheduled-tasks\register.ps1
```

Registers `AssistantKeepAlive` (every 30 min) and `RemindersScanner` (hourly), both user-scope (no admin).

### Phase 10 — First start + verify

```powershell
# First run of keep-alive starts everything
powershell C:\Users\$env:USERNAME\OneDrive\Assistant\local-tool-server\keep-alive.ps1

# Smoke
curl http://localhost:3000/                                 # 200
curl http://localhost:3100/health                           # {"status":"ok"}
curl https://assistant.tail4621a0.ts.net/                   # 200 (from this machine)

# Phone: open https://assistant.tail4621a0.ts.net (existing PWA install will keep working
# IF you preserved the VAPID keys in Phase 5; otherwise re-subscribe in /settings)
```

### Phase 11 — Trigger Calendar scope reconsent (one-time)

The first call to a calendar tool (`list_calendar_events` etc.) will trigger a fresh MSAL device-code flow because the cached token doesn't include `Calendars.Read`. Run it interactively first so it doesn't hang inside `RemindersScanner`:

```powershell
C:\git\venvs\outlook-email\Scripts\python.exe -c @'
import asyncio
import sys; sys.path.insert(0, r"C:\git\mcp-servers\outlook-email")
from server import list_calendar_events
print(asyncio.run(list_calendar_events(days_ahead=1)))
'@
# Enter the device code shown on stderr at https://microsoft.com/devicelogin
```

After this completes once, the cache covers all 4 scopes and `RemindersScanner` runs unattended.

### Phase 12 — Append migration entry to worklog

```powershell
notepad C:\Users\$env:USERNAME\OneDrive\Assistant\notes\worklog.md
# Add a dated entry summarizing the bootstrap.
```

---

## What's in this repo

| Path                            | What it is                                              |
| ------------------------------- | ------------------------------------------------------- |
| `mcp-servers/outlook-email/`    | Python MSAL+Graph MCP server: mail tools (list_inbox, search_email, send_email, …) + calendar tools (list_calendar_events, search_calendar_events, get_calendar_event). 11 tools total. config.json holds public client_id only. |
| `mcp-servers/reminders/`        | Hourly proactive-reminder scanner. For each `notes/reminders/*.md` with `status: active`, spawns `claude -p` to judge whether the trigger condition is met. Fires Web Push on `fired=true`. |
| `local-tool-server/`            | Express server on :3100 (file/command tools) + `keep-alive.ps1` that restarts Tailscale, the tool server, and Next.js every 30 minutes. |
| `scheduled-tasks/register.ps1`  | Idempotent user-scope task registration (AssistantKeepAlive + RemindersScanner). |
| `BOOTSTRAP-GOTCHAS.md`          | Pitfalls encountered last time someone bootstrapped this from scratch. **Read before starting.** |

## What is NOT in this repo (per `.gitignore`)

| File                                                              | Where it lives                                                   |
| ----------------------------------------------------------------- | ---------------------------------------------------------------- |
| `apps/web/.env.local`                                             | VAPID keys, AUTH_SECRET, PUSH_INTERNAL_TOKEN. Keep encrypted.    |
| `mcp-servers/outlook-email/.token_cache/*.json`                   | MSAL refresh tokens.                                             |
| `mcp-servers/outlook-email/push_config.json`                      | Bearer + push fire URLs.                                         |
| `C:\Users\<you>\OneDrive\Assistant\`                              | OneDrive auto-sync. CLAUDE.md, notes, reminders, skills.          |
| `C:\git\MySuperApp\`                                              | github.com/sbeashwar/aisuperapp                                  |

## Related repos

- [sbeashwar/aisuperapp](https://github.com/sbeashwar/aisuperapp) — Next.js 14 PWA. Has its own setup README.
- OneDrive workspace at `OneDrive\Assistant\` (synced, not a repo) — has its own `CLAUDE.md` and skill docs.
