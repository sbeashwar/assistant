# Scheduled tasks

`register.ps1` registers two user-scope scheduled tasks (no admin required):

| Task | Cadence | What it runs |
|---|---|---|
| `AssistantKeepAlive` | Logon + every 30 min | `keep-alive.ps1` — ensures Tailscale, tool server (:3100), and Next.js (:3000) are up. |
| `RemindersScanner` | Hourly | `reminders.py` — judges every active `notes/reminders/*.md` trigger via `claude -p`. Fires Web Push when met. |

Run from the repo root:

```powershell
powershell -ExecutionPolicy Bypass -File scheduled-tasks\register.ps1
```

The script is idempotent (`-Force` replaces existing tasks).

## To remove

```powershell
Unregister-ScheduledTask -TaskName "AssistantKeepAlive","RemindersScanner" -Confirm:$false
```
