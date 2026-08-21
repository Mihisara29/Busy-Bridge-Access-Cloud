Set-StrictMode -Version Latest

function Write-OfflineSyncLog {
    param(
        [Parameter(Mandatory = $true)] [string] $Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string] $Level = 'INFO'
    )

    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$timestamp] [OFFLINE-SYNC] [$Level] $Message"

    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        default { Write-Host $line -ForegroundColor Cyan }
    }

    try {
        $projectRoot = Split-Path $PSScriptRoot -Parent
        $logDirectory = Join-Path $projectRoot 'logs'
        $logPath = Join-Path $logDirectory 'offline_sync.log'

        if (-not (Test-Path $logDirectory)) {
            New-Item `
                -ItemType Directory `
                -Path $logDirectory `
                -Force | Out-Null
        }

        Add-Content `
            -Path $logPath `
            -Value $line `
            -Encoding UTF8
    }
    catch {
        Write-Host "  Could not write offline sync log file: $($_.Exception.Message)" -ForegroundColor DarkYellow
    }
}
