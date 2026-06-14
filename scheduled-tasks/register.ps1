# scheduled-tasks/register.ps1
# Registers all assistant scheduled tasks under the current user (no admin needed).
# Idempotent: re-running replaces existing tasks.

$ErrorActionPreference = "Stop"

$here = Split-Path -Parent $MyInvocation.MyCommand.Path

# === AssistantKeepAlive ===
# Runs the keep-alive script that restarts Tailscale Serve, tool server (:3100),
# and Next.js (:3000) if any are down. Logon + every 30 min.
$keepAliveScript = "C:\Users\$env:USERNAME\OneDrive\Assistant\local-tool-server\keep-alive.ps1"

$action1 = New-ScheduledTaskAction `
  -Execute "powershell.exe" `
  -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$keepAliveScript`""

$trigger1a = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$trigger1b = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
  -RepetitionInterval (New-TimeSpan -Minutes 30)

$settings1 = New-ScheduledTaskSettingsSet `
  -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
  -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

$principal1 = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName "AssistantKeepAlive" `
  -Action $action1 -Trigger @($trigger1a, $trigger1b) `
  -Settings $settings1 -Principal $principal1 `
  -Description "Restart Tailscale Serve, tool server (:3100), and Next.js (:3000) if down." `
  -Force | Out-Null

Write-Host "Registered AssistantKeepAlive"

# === RemindersScanner ===
# Hourly: judges each active reminder's trigger via claude -p, fires Web Push on hit.
$reminderScript = "C:\git\mcp-servers\reminders\reminders.py"
$pythonExe = "C:\git\venvs\outlook-email\Scripts\python.exe"

$action2 = New-ScheduledTaskAction `
  -Execute $pythonExe `
  -Argument "`"$reminderScript`"" `
  -WorkingDirectory (Split-Path $reminderScript)

$trigger2 = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(4) `
  -RepetitionInterval (New-TimeSpan -Hours 1)

$settings2 = New-ScheduledTaskSettingsSet `
  -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
  -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 20)

$principal2 = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName "RemindersScanner" `
  -Action $action2 -Trigger $trigger2 `
  -Settings $settings2 -Principal $principal2 `
  -Description "Hourly: for each active reminder, ask claude -p whether its trigger condition has been met. Fires Web Push on hit." `
  -Force | Out-Null

Write-Host "Registered RemindersScanner"

Write-Host ""
Get-ScheduledTask -TaskName "AssistantKeepAlive","RemindersScanner" |
  Select-Object TaskName, State, @{N="NextRun";E={(Get-ScheduledTaskInfo $_).NextRunTime}} |
  Format-Table -AutoSize
