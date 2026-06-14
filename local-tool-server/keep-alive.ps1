# keep-alive.ps1 - Ensure local tool server + Next.js app are running
# Runs via Windows Task Scheduler every 30 minutes
# Architecture: Phone → Tailscale → PC (Next.js on :3000 + tool server on :3100)

$ErrorActionPreference = "SilentlyContinue"
$LogFile = "$PSScriptRoot\keep-alive.log"
$ToolServerDir = $PSScriptRoot
$ToolApiSecret = "msa-local-tools-103400"
$NextJsDir = "C:\git\MySuperApp\apps\web"

function Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$ts  $msg" | Out-File -Append -FilePath $LogFile -Encoding utf8
}

# Trim log to last 200 lines
if (Test-Path $LogFile) {
    $lines = Get-Content $LogFile -Tail 200
    $lines | Set-Content $LogFile -Encoding utf8
}

Log "--- keep-alive check ---"

# 1. Check Tailscale
$tailscaleOk = $false
try {
    $tsStatus = & "C:\Program Files\Tailscale\tailscale.exe" status 2>&1 | Out-String
    if ($tsStatus -notmatch "Logged out" -and $tsStatus -notmatch "stopped") {
        $tailscaleOk = $true
    }
} catch {}

if (-not $tailscaleOk) {
    Log "Tailscale DOWN or logged out - attempting restart..."
    Start-Process -FilePath "C:\Program Files\Tailscale\tailscale.exe" -ArgumentList "up" `
        -WindowStyle Hidden -PassThru | Out-Null
    Start-Sleep -Seconds 5
    try {
        $tsStatus = & "C:\Program Files\Tailscale\tailscale.exe" status 2>&1 | Out-String
        if ($tsStatus -notmatch "Logged out") {
            Log "Tailscale restarted OK"
        } else {
            Log "Tailscale still down - may need manual login"
        }
    } catch {
        Log "Tailscale check failed: $_"
    }
} else {
    Log "Tailscale OK"
}

# 1b. Check Tailscale Serve (HTTPS proxy for PWA)
$serveStatus = & "C:\Program Files\Tailscale\tailscale.exe" serve status 2>&1 | Out-String
if ($serveStatus -match "No serve config") {
    Log "Tailscale Serve not configured - enabling HTTPS proxy..."
    & "C:\Program Files\Tailscale\tailscale.exe" serve --bg --https 443 http://localhost:3000 2>&1 | Out-Null
    Log "Tailscale Serve enabled at https://assistant.tail4621a0.ts.net"
} else {
    Log "Tailscale Serve OK"
}

# 2. Check local tool server
$serverOk = $false
try {
    $r = Invoke-RestMethod -Uri "http://localhost:3100/health" -TimeoutSec 5
    if ($r.status -eq "ok") { $serverOk = $true }
} catch {}

if (-not $serverOk) {
    Log "Tool server DOWN - restarting..."
    # Kill any orphan node processes on port 3100
    $proc = Get-NetTCPConnection -LocalPort 3100 -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($proc) {
        Stop-Process -Id $proc.OwningProcess -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    }
    # Start tool server in background
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "node"
    $psi.Arguments = "server.js"
    $psi.WorkingDirectory = $ToolServerDir
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.EnvironmentVariables["TOOL_API_SECRET"] = $ToolApiSecret
    $p = [System.Diagnostics.Process]::Start($psi)
    Start-Sleep -Seconds 3
    # Verify
    try {
        $r = Invoke-RestMethod -Uri "http://localhost:3100/health" -TimeoutSec 5
        if ($r.status -eq "ok") { Log "Tool server restarted OK (PID: $($p.Id))" }
        else { Log "Tool server restart FAILED (bad health)" }
    } catch {
        Log "Tool server restart FAILED: $_"
    }
} else {
    Log "Tool server OK"
}

# 3. Check Next.js app server
$nextOk = $false
try {
    $r = Invoke-WebRequest -Uri "http://localhost:3000/" -TimeoutSec 5 -UseBasicParsing
    if ($r.StatusCode -eq 200) { $nextOk = $true }
} catch {}

if (-not $nextOk) {
    Log "Next.js app DOWN - restarting..."
    # Kill any orphan processes on port 3000
    $proc = Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($proc) {
        Stop-Process -Id $proc.OwningProcess -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }
    # Start Next.js production server bound to 0.0.0.0 for Tailscale access
    # Production server starts in seconds (vs 30s for dev). Build with: pnpm --filter @mysuperapp/web build
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "cmd.exe"
    $psi.Arguments = "/c npx next start --hostname 0.0.0.0"
    $psi.WorkingDirectory = $NextJsDir
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $false
    $psi.RedirectStandardError = $false
    $p = [System.Diagnostics.Process]::Start($psi)
    Log "Next.js starting (PID: $($p.Id)), waiting up to 30s..."
    # Wait up to 30s for Next.js to be ready
    $started = $false
    for ($i = 0; $i -lt 6; $i++) {
        Start-Sleep -Seconds 5
        try {
            $r = Invoke-WebRequest -Uri "http://localhost:3000/" -TimeoutSec 3 -UseBasicParsing
            if ($r.StatusCode -eq 200) { $started = $true; break }
        } catch {}
    }
    if ($started) {
        Log "Next.js app restarted OK (PID: $($p.Id))"
    } else {
        Log "Next.js app restart FAILED — process may still be compiling"
    }
} else {
    Log "Next.js app OK"
}

Log "--- check complete ---"
