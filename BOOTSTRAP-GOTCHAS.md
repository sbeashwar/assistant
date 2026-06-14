# Bootstrap gotchas

Concrete pitfalls hit during the 2026-06-13 new-machine bootstrap. Most are not in the main README because they'd dominate it — but reading these in advance saves hours.

---

## Entra Conditional Access blocks `gh auth login` (AADSTS53003)

**Symptom:** `gh auth login --web` opens the Microsoft sign-in page, your sign-in succeeds, then you see:

> You don't have access to this resource.
> Error Code: 53003 · App name: Microsoft GitHub for Open Source Enterprise Cloud Access (1ES) · Device state: Unregistered

**Cause:** The `sbeashwar` GitHub account is SSO-enforced through Microsoft's 1ES open-source enterprise org. The `repo` scope route goes through Entra, which requires the device to be Entra-registered (Intune-enrolled). A fresh personal Windows machine isn't enrolled, so the SSO step denies the OAuth grant for `gh`.

**Workaround for git itself (long-term, recommended):** SSH keys bypass Entra entirely.

```powershell
ssh-keygen -t ed25519 -f "C:\Users\$env:USERNAME\.ssh\id_ed25519" -N '""'
# Add the .pub to https://github.com/settings/ssh/new
# Use git@github.com: remotes
```

**Workaround for `gh` CLI specifically (if you need it):** classic PAT.
- Generate at https://github.com/settings/tokens with scopes: `repo` AND `read:org` (both are required — read:org alone won't work, repo alone fails with "missing required scope 'read:org'").
- `echo ghp_… | gh auth login --hostname github.com --git-protocol https --with-token`
- Set a short expiry (1–7 days) and revoke after the bootstrap session.

**To upload your SSH key via `gh` instead of the web UI:** needs `admin:public_key` scope on the PAT. Probably easier to use the web UI than to chase that extra scope.

---

## PowerShell 5.1 chokes on UTF-8 without BOM

**Symptom:** A `.ps1` script runs fine in PowerShell 7 but errors out in 5.1 (which is what Windows uses by default for scheduled tasks) with confusing parser errors like `The string is missing the terminator: "` and `Missing closing '}'` — citing line numbers that look syntactically fine.

**Cause:** Without a UTF-8 BOM, PowerShell 5.1 reads `.ps1` files as Windows-1252. Multi-byte UTF-8 characters (em-dash `—`, smart quotes, anything non-ASCII) become garbage bytes that break string parsing.

**Fix (both belt and suspenders):**
1. Avoid non-ASCII in `.ps1` files. Use `-` not `—`.
2. Save `.ps1` files as **UTF-8 with BOM**. In VS Code: bottom-right "UTF-8" → "Save with Encoding" → "UTF-8 with BOM". From bash:
   ```bash
   printf '\xef\xbb\xbf' | cat - script.ps1 > tmp && mv tmp script.ps1
   ```

The `local-tool-server/keep-alive.ps1` in this repo is already saved with BOM. Preserve it.

---

## Python venvs are NOT portable across machines

**Symptom:** You copy `C:\git\venvs\outlook-email\` from another machine. `Scripts/` directory is missing, or `python.exe` errors immediately, or imports fail.

**Cause:** Windows venvs hard-code the Python executable path in `pyvenv.cfg`. `Scripts\python.exe` is a thin wrapper around that path. Symlinks don't survive most sync tools. The Python version must match exactly between source and target machines.

**Recipe:** Always recreate the venv on each machine. Save the package list, not the venv:

```powershell
# On the old machine:
C:\git\venvs\outlook-email\Scripts\python.exe -m pip freeze > requirements.txt

# On the new machine:
python -m venv C:\git\venvs\outlook-email
C:\git\venvs\outlook-email\Scripts\python.exe -m pip install -r requirements.txt
```

For this repo's outlook-email MCP, the minimum deps are: `msal requests mcp pydantic httpx python-dotenv`. The full transitive set is larger, but pip will resolve it.

---

## MSAL token cache IS portable (and the real prize)

In contrast to the venv, `.token_cache/token_cache.json` is plain JSON with refresh tokens — no machine-specific bits. Copy it verbatim into `C:\git\mcp-servers\outlook-email\.token_cache\` and the first call gets a fresh access token via the cached refresh token. **No device-code re-auth needed.**

This is THE most valuable file to preserve when migrating. Lose it and you re-auth with a device-code prompt; lose `apps/web/.env.local` and you can regenerate everything except (optionally) the VAPID keys.

---

## Adding an MSAL scope triggers a fresh device-code prompt

If you change `SCOPES` in `outlook-email/server.py` (e.g., adding `Calendars.Read`), the cached token still works for the old scopes. **First call needing the new scope** triggers a fresh device-code flow on stderr. Not an error, just:

```
SIGN IN REQUIRED
Go to: https://microsoft.com/devicelogin
Enter code: ABCD-EFGH
```

The MCP server hangs until you complete the flow. To avoid that hang firing *inside* the scheduled `RemindersScanner` (where you can't see stderr and it'll just timeout), **run the new scope's tool interactively first** to clear the prompt:

```powershell
C:\git\venvs\outlook-email\Scripts\python.exe -c "
import asyncio, sys
sys.path.insert(0, r'C:\git\mcp-servers\outlook-email')
from server import list_calendar_events
print(asyncio.run(list_calendar_events(days_ahead=1)))
"
```

Enter the code at https://microsoft.com/devicelogin. After it returns, the cache covers all approved scopes going forward.

---

## `better-sqlite3` build failure during `pnpm install` is benign

**Symptom:** `pnpm install` in MySuperApp prints walls of `gyp ERR! find VS Could not find any Visual Studio installation` errors, taking the exit code to look scary at the end.

**Reality:** `better-sqlite3` is an OPTIONAL peer dep of `drizzle-orm` / `drizzle-kit`. The web app uses `@libsql/client`, which is precompiled and doesn't need a C++ compiler. The build succeeds, the tests pass — it's noise.

**Do not** install Visual Studio Build Tools to "fix" it. You'll waste an hour and gain nothing.

---

## Git Bash mangles `/foo` args to Windows tools

**Symptom:** From Git Bash, `schtasks /query /tn AssistantKeepAlive` errors with `ERROR: Invalid argument/option - 'C:/Program Files/Git/query'`.

**Cause:** MSYS path conversion. `/query` looks like a Unix path, Bash rewrites it to a Windows path.

**Workaround:** invoke through cmd.exe with double-slash to escape the conversion:

```bash
cmd.exe //c 'schtasks /query /tn AssistantKeepAlive'
```

Or use PowerShell `Get-ScheduledTask` cmdlets (named params, no slashes):

```powershell
Get-ScheduledTask -TaskName AssistantKeepAlive | Get-ScheduledTaskInfo
```

The second is cleaner and what `scheduled-tasks/register.ps1` uses throughout.

---

## Bash here-doc → PowerShell `-Command` is fragile

Don't try `powershell.exe -Command @'long script with $vars and (Get-Date)...'@` from Bash. Bash treats `@'` as array start, `(Get-Date)` as subshell, etc. Mangled before PowerShell sees it.

Write the script to a file with your editor or `Write`/`cat <<EOF`, then:

```bash
powershell.exe -NoProfile -ExecutionPolicy Bypass -File script.ps1
```

Works first try every time.

---

## Tailscale node naming

When you rename a machine via `tailscale set --hostname=<name>`, MagicDNS issues a fresh cert in seconds. The phone PWA keeps working with no re-install IF the URL matches.

**Choose hostname carefully on first setup.** Don't use machine-specific names like `sbook` or `dell-xps`. Use abstract names: `assistant`, `home-pc`, `prod`. Everything in the workspace will hard-code that name across many files; renaming later is mechanical but tedious.

This repo's PWA URL is `https://assistant.tail4621a0.ts.net` — preserve it.

---

## VAPID keys regeneration kills all phone subscriptions

If you regenerate VAPID keys (e.g., lost `apps/web/.env.local`), every device that already subscribed to push notifications gets `410 Gone` from the push service and is silently pruned. The user has to re-subscribe from `/settings` on each device.

This is harmless but surprising. Mention it before regenerating. To preserve, keep the original VAPID keypair in your encrypted bundle.

---

## OneDrive sync race on workspace read

Right after install, OneDrive may still be syncing the `Assistant/` folder. If you start the PWA before sync completes, `claude -p` will see a partial workspace (missing skills, missing notes) and give weird answers.

**Check sync is done before starting services:**
- OneDrive system tray icon shows green checkmark (not blue rotating arrows).
- `ls C:\Users\<you>\OneDrive\Assistant\` shows all expected dirs (`.claude`, `.github`, `notes`, `templates`, `local-tool-server`).

---

## Reminders log location

Scheduled tasks run with no console attached. `RemindersScanner` writes everything to `C:\git\mcp-servers\reminders\logs\reminders.log`. To watch live:

```powershell
Get-Content C:\git\mcp-servers\reminders\logs\reminders.log -Tail 50 -Wait
```

The judge prompt + Claude's reasoning + the final JSON verdict all land there. First place to look if a reminder isn't firing as expected.
