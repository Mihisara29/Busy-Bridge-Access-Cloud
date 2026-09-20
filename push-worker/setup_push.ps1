param(
    [string]$Subject = "mailto:admin@example.com"
)

$ErrorActionPreference = "Stop"

$WorkerDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BackendRoot = Split-Path -Parent $WorkerDir
$ConfigPath = Join-Path $BackendRoot "data\push\push_config.json"

Write-Host "BUSY Cloud Push Setup" -ForegroundColor Cyan
Write-Host "Worker folder: $WorkerDir" -ForegroundColor DarkGray

$node = Get-Command node -ErrorAction SilentlyContinue
$npm = Get-Command npm -ErrorAction SilentlyContinue

if (-not $node) {
    throw "Node.js was not found. Install Node.js 18+ first."
}

if (-not $npm) {
    throw "npm was not found. Install npm/Node.js first."
}

Push-Location $WorkerDir

try {
    if (-not (Test-Path (Join-Path $WorkerDir "node_modules\web-push"))) {
        Write-Host "Installing Web Push worker dependency..." -ForegroundColor Yellow
        & npm install

        if ($LASTEXITCODE -ne 0) {
            throw "npm install failed."
        }
    }
    else {
        Write-Host "web-push dependency already installed." -ForegroundColor Green
    }

    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        Write-Host "Generating VAPID keys + worker secret..." -ForegroundColor Yellow
        & node ".\generate_config.js" $Subject

        if ($LASTEXITCODE -ne 0) {
            throw "VAPID configuration generation failed."
        }
    }
    else {
        Write-Host "Existing push_config.json preserved." -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "Push setup completed." -ForegroundColor Green
    Write-Host "IMPORTANT: Back up this file:" -ForegroundColor Yellow
    Write-Host "  $ConfigPath" -ForegroundColor White
    Write-Host ""
    Write-Host "Production Web Push requires HTTPS for the frontend origin." -ForegroundColor Cyan
}
finally {
    Pop-Location
}
