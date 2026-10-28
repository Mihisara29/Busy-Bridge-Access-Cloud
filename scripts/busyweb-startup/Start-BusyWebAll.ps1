# One-command BusyWeb backend launcher. Windows PowerShell 5.1 on Windows 64-bit.
# Launches Main 8081, Reports 8085, Executive worker, and a portless WhatsApp worker.
# Optionally runs Vite with -IncludeFrontend.
param(
    [string]$BackendRoot = 'C:\Busy-Bridge-Access-Cloud',
    [switch]$IncludeFrontend,
    [string]$FrontendRoot = 'C:\busy-bevy - Local',
    [switch]$ProcessPendingWhatsAppJobs
)
$ErrorActionPreference = 'Stop'

function Test-ListeningPort([int]$Port) {
    try {
        $listeners = @(Get-NetTCPConnection -State Listen -ErrorAction Stop | Where-Object { $_.LocalPort -eq $Port })
        return ($listeners.Count -gt 0)
    }
    catch {
        # No best-effort launch against an unknown port state: avoid collisions.
        throw "Could not inspect TCP port $Port. $($_.Exception.Message)"
    }
}

function Test-RunningProcess([string]$Pattern, [string[]]$Names) {
    foreach ($name in $Names) {
        try {
            $items = @(Get-CimInstance Win32_Process -Filter "Name='$name'" -ErrorAction Stop)
            foreach ($item in $items) {
                if ([string]$item.CommandLine -match $Pattern) { return $true }
            }
        }
        catch { throw "Unable to inspect existing processes: $($_.Exception.Message)" }
    }
    return $false
}

function Start-Component([string]$Component, [string]$Exe, [string]$Runner, [string]$BackendRoot, [string]$FrontendRoot, [string]$ChromePath) {
    # Quoting is necessary because paths may contain spaces (for example busy-bevy - Local).
    foreach ($value in @($Runner,$BackendRoot,$FrontendRoot,$ChromePath)) {
        if ([string]$value -match '"') { throw 'Double quotes in filesystem paths are unsupported.' }
    }
    $command = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -Component {1} -BackendRoot "{2}" -FrontendRoot "{3}" -ChromePath "{4}"' -f $Runner, $Component, $BackendRoot, $FrontendRoot, $ChromePath
    $proc = Start-Process -FilePath $Exe -ArgumentList $command -WorkingDirectory $BackendRoot -PassThru -ErrorAction Stop
    Write-Host ("[STARTED] {0}, PID={1} (new console; check its output for readiness)" -f $Component,$proc.Id) -ForegroundColor Green
}

if (-not (Test-Path -LiteralPath $BackendRoot -PathType Container)) { throw "Backend folder not found: $BackendRoot" }
$BackendRoot = (Resolve-Path -LiteralPath $BackendRoot).ProviderPath
$runner = Join-Path $PSScriptRoot 'BusyWeb-Process.ps1'
if (-not (Test-Path -LiteralPath $runner -PathType Leaf)) { throw "Missing process runner: $runner" }
if (-not (Test-Path -LiteralPath (Join-Path $BackendRoot 'busy_api.ps1') -PathType Leaf)) { throw 'Missing busy_api.ps1' }
if (-not (Test-Path -LiteralPath (Join-Path $BackendRoot 'workers\ExecutiveReportWorker.ps1') -PathType Leaf)) { throw 'Missing ExecutiveReportWorker.ps1' }
if (-not (Test-Path -LiteralPath (Join-Path $BackendRoot 'report-jobs') -PathType Container)) { throw 'Missing report-jobs folder' }
if (-not (Test-Path -LiteralPath (Join-Path $BackendRoot 'whatsapp-worker\worker.js') -PathType Leaf)) { throw 'Missing whatsapp-worker\worker.js' }
if (-not (Test-Path -LiteralPath (Join-Path $BackendRoot 'whatsapp-worker\node_modules\whatsapp-web.js') -PathType Container)) { throw 'WhatsApp npm dependencies not installed' }
if (-not (Get-Command node.exe -ErrorAction SilentlyContinue)) { throw 'node.exe not found in PATH' }
if ($IncludeFrontend) {
    if (-not (Test-Path -LiteralPath $FrontendRoot -PathType Container)) { throw "Frontend folder not found: $FrontendRoot" }
    $FrontendRoot = (Resolve-Path -LiteralPath $FrontendRoot).ProviderPath
    if (-not (Get-Command npm.cmd -ErrorAction SilentlyContinue)) { throw 'npm.cmd not found in PATH' }
}

$chrome = @(
    'C:\Program Files\Google\Chrome\Application\chrome.exe',
    'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    (Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe')
) | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
if (-not $chrome) { throw 'Could not find Google Chrome. Set up Chrome before starting WhatsApp worker.' }

$ps32 = Join-Path $env:WINDIR 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
if (-not (Test-Path -LiteralPath $ps32 -PathType Leaf)) { throw 'Required 32-bit Windows PowerShell not found.' }

Write-Host "[BUSYWEB] Backend: $BackendRoot" -ForegroundColor Cyan
Write-Host '[BUSYWEB] Safe normal startup. Exposure maintenance remains OFF.' -ForegroundColor Cyan
$gate = Join-Path $BackendRoot 'report-jobs\enable-heavy-exposure.flag'
if (Test-Path -LiteralPath $gate) {
    Write-Warning 'Heavy Exposure API gate flag still exists. This launcher leaves it unchanged, but normal worker will NOT process heavy jobs. Remove the flag after maintenance using your established procedure.'
}

# Never start duplicate API listeners; an occupied port may also be an unrelated program.
if (Test-ListeningPort 8081) { Write-Warning '8081 already has a listener: skipped Main. Confirm it is your BUSY API.' }
else { Start-Component 'Main' $ps32 $runner $BackendRoot $FrontendRoot $chrome }
Start-Sleep -Milliseconds 250
if (Test-ListeningPort 8085) { Write-Warning '8085 already has a listener: skipped Reports. Confirm it is your BUSY API.' }
else { Start-Component 'Reports' $ps32 $runner $BackendRoot $FrontendRoot $chrome }

if (Test-RunningProcess '(ExecutiveReportWorker\.ps1|BusyWeb-Process\.ps1.*-Component Executive)' @('powershell.exe','pwsh.exe')) {
    Write-Host '[SKIP] Executive worker already appears active.' -ForegroundColor Yellow
}
else { Start-Component 'Executive' $ps32 $runner $BackendRoot $FrontendRoot $chrome }

if (Test-RunningProcess '(whatsapp-worker[\\/]worker\.js|BusyWeb-Process\.ps1.*-Component WhatsApp)' @('node.exe','powershell.exe','pwsh.exe')) {
    Write-Host '[SKIP] WhatsApp worker already appears active.' -ForegroundColor Yellow
}
else {
    $statusFile = Join-Path $BackendRoot 'whatsapp\connection-status.json'
    $recentHeartbeat = $false
    if (Test-Path -LiteralPath $statusFile -PathType Leaf) {
        try {
            $st = Get-Content -LiteralPath $statusFile -Raw | ConvertFrom-Json
            $age = ([datetime]::UtcNow - [datetime]::Parse([string]$st.updatedAt).ToUniversalTime()).TotalSeconds
            $recentHeartbeat = ($age -ge 0 -and $age -lt 30)
        }
        catch { Write-Warning 'Unreadable WhatsApp heartbeat. Process list will be used.' }
    }
    if ($recentHeartbeat) {
        Write-Warning 'WhatsApp heartbeat is fresh. Not starting a second worker even though process matching found none.'
    }
    else {
        $pendingDir = Join-Path $BackendRoot 'whatsapp\pending'
        $pendingCount = 0
        if (Test-Path -LiteralPath $pendingDir -PathType Container) {
            $pendingCount = @(Get-ChildItem -LiteralPath $pendingDir -Filter '*.json' -File).Count
        }
        if ($pendingCount -gt 0 -and -not $ProcessPendingWhatsAppJobs) {
            Write-Warning ("WhatsApp worker NOT started: {0} pending PDF job(s) could be sent immediately. Review the queue; then rerun with -ProcessPendingWhatsAppJobs." -f $pendingCount)
        }
        else {
            if ($pendingCount -gt 0) { Write-Warning "Starting WhatsApp with $pendingCount pending job(s) as explicitly requested." }
            Start-Component 'WhatsApp' $ps32 $runner $BackendRoot $FrontendRoot $chrome
        }
    }
}

if ($IncludeFrontend) {
    if (Test-ListeningPort 8080) { Write-Warning '8080 already has a listener: skipped Vite.' }
    else { Start-Component 'Frontend' $ps32 $runner $BackendRoot $FrontendRoot $chrome }
}

Write-Host ''
Write-Host '[BUSYWEB] Launch commands issued. Read each window for READY / server started.' -ForegroundColor Cyan
Write-Host '[BUSYWEB] Closing this launcher window does not stop the child consoles.' -ForegroundColor Gray
Write-Host '[BUSYWEB] WhatsApp still has no HTTP listener; use Ctrl+C in its window to stop it.' -ForegroundColor Gray
