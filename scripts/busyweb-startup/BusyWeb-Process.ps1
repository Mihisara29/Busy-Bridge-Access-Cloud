# BusyWeb process host - Windows PowerShell 5.1 compatible.
# This script runs in its own console; launched by Start-BusyWebAll.ps1.
param(
    [Parameter(Mandatory=$true)][ValidateSet('Main','Reports','Executive','WhatsApp','Frontend')][string]$Component,
    [string]$BackendRoot = 'C:\Busy-Bridge-Access-Cloud',
    [string]$FrontendRoot = 'C:\busy-bevy - Local',
    [string]$ChromePath = ''
)
$ErrorActionPreference = 'Stop'
try {
    $Host.UI.RawUI.WindowTitle = "BusyWeb - $Component"
    Write-Host "[BUSYWEB] Starting $Component" -ForegroundColor Cyan
    switch ($Component) {
        'Main' {
            Set-Location -LiteralPath $BackendRoot
            & (Join-Path $BackendRoot 'busy_api.ps1') -Port 8081 -ServiceMode Main
            break
        }
        'Reports' {
            Set-Location -LiteralPath $BackendRoot
            & (Join-Path $BackendRoot 'busy_api.ps1') -Port 8085 -ServiceMode Reports
            break
        }
        'Executive' {
            # Do not silently activate multi-minute Exposure calculations.
            $env:BUSYWEB_EXECUTIVE_ENABLE_EXPOSURE = '0'
            Set-Location -LiteralPath $BackendRoot
            & (Join-Path $BackendRoot 'workers\ExecutiveReportWorker.ps1') -BridgeRoot $BackendRoot
            break
        }
        'WhatsApp' {
            if (-not (Test-Path -LiteralPath $ChromePath -PathType Leaf)) {
                throw 'Google Chrome path is invalid. Edit ChromePath or reinstall Chrome.'
            }
            $env:PUPPETEER_EXECUTABLE_PATH = $ChromePath
            $env:PUPPETEER_SKIP_DOWNLOAD = 'true'
            $workerRoot = Join-Path $BackendRoot 'whatsapp-worker'
            Set-Location -LiteralPath $workerRoot
            $node = Get-Command 'node.exe' -ErrorAction Stop
            Write-Host '[BUSYWEB] WhatsApp: no HTTP port. Keep this window open.' -ForegroundColor Yellow
            & ($node.Source) (Join-Path $workerRoot 'worker.js')
            if ($LASTEXITCODE -ne 0) { throw "WhatsApp worker exited with code $LASTEXITCODE" }
            break
        }
        'Frontend' {
            Set-Location -LiteralPath $FrontendRoot
            $npm = Get-Command 'npm.cmd' -ErrorAction Stop
            & ($npm.Source) run dev
            if ($LASTEXITCODE -ne 0) { throw "Vite exited with code $LASTEXITCODE" }
            break
        }
    }
}
catch {
    Write-Host "[BUSYWEB] $Component stopped with error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'See above for the cause. No other component was stopped.' -ForegroundColor Yellow
    [void](Read-Host 'Press Enter to close this window')
    exit 1
}
