param(
    [int]$Port = 8081
)

$ErrorActionPreference = "Stop"

$BackendRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$WorkerDir = Join-Path $BackendRoot "push-worker"
$ConfigPath = Join-Path $BackendRoot "data\push\push_config.json"
$WorkerJs = Join-Path $WorkerDir "worker.js"
$NodeModules = Join-Path $WorkerDir "node_modules\web-push"
$BusyApi = Join-Path $BackendRoot "busy_api.ps1"

$PowerShell32 = "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"

if (-not (Test-Path -LiteralPath $PowerShell32)) {
    throw "32-bit Windows PowerShell was not found: $PowerShell32"
}

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    throw "Node.js was not found. Install Node.js 18+."
}

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    throw "Push configuration is missing. Run push-worker\setup_push.ps1 first."
}

if (-not (Test-Path -LiteralPath $NodeModules)) {
    throw "Push worker dependencies are missing. Run push-worker\setup_push.ps1 first."
}

if (-not (Test-Path -LiteralPath $BusyApi)) {
    throw "busy_api.ps1 was not found: $BusyApi"
}

$LogDir = Join-Path $BackendRoot "data\push"
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

$StdOut = Join-Path $LogDir "worker.stdout.log"
$StdErr = Join-Path $LogDir "worker.stderr.log"

Write-Host "Starting BUSY Cloud Web Push worker..." -ForegroundColor Cyan

$worker = Start-Process `
    -FilePath "node" `
    -ArgumentList "`"$WorkerJs`"" `
    -WorkingDirectory $WorkerDir `
    -WindowStyle Hidden `
    -RedirectStandardOutput $StdOut `
    -RedirectStandardError $StdErr `
    -PassThru

Start-Sleep -Milliseconds 500

if ($worker.HasExited) {
    throw "Push worker exited immediately. Check $StdErr"
}

Write-Host "Push worker PID: $($worker.Id)" -ForegroundColor Green
Write-Host "Starting 32-bit BUSY PowerShell bridge on port $Port..." -ForegroundColor Cyan

try {
    & $PowerShell32 `
        -ExecutionPolicy Bypass `
        -File $BusyApi `
        -Port $Port
}
finally {
    if ($worker -and -not $worker.HasExited) {
        Write-Host "Stopping Web Push worker..." -ForegroundColor DarkYellow
        try { Stop-Process -Id $worker.Id -Force } catch {}
    }
}
