#requires -Version 5.1
[CmdletBinding()]
param([string]$BackendRoot = (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent))
$ErrorActionPreference='Continue'
Write-Host "Backend folder: $BackendRoot" -ForegroundColor Cyan
foreach($port in @(8081,8085)) {
    try {
        $r=Invoke-WebRequest -UseBasicParsing -Uri "http://127.0.0.1:$port/health" -TimeoutSec 5
        Write-Host "Port $port /health: HTTP $($r.StatusCode)" -ForegroundColor Green
    } catch { Write-Warning "Port $port /health failed: $($_.Exception.Message)" }
}
$workers=@(Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue | Where-Object {
    $_.CommandLine -and
    $_.CommandLine.IndexOf('ExecutiveReportWorker.ps1',[StringComparison]::OrdinalIgnoreCase) -ge 0 -and
    $_.CommandLine.IndexOf($BackendRoot,[StringComparison]::OrdinalIgnoreCase) -ge 0
})
Write-Host "Executive workers found: $($workers.Count)"
foreach($worker in $workers){Write-Host "  Worker PID: $($worker.ProcessId)"}
if($workers.Count -ne 1){Write-Warning 'Exactly one Executive worker is expected.'}
$reports=Join-Path $BackendRoot 'modules\reports.ps1'
$insights=Join-Path $BackendRoot 'modules\executive_insights.ps1'
if((Test-Path -LiteralPath $reports) -and (Test-Path -LiteralPath $insights)) {
    $rt=[IO.File]::ReadAllText($reports)
    $it=[IO.File]::ReadAllText($insights)
    $scoped=($rt.Contains('BUSYWEB_SCOPED_OUTSTANDING_V1') -and $it.Contains('BUSYWEB_SCOPED_CUSTOMER_BALANCES_V5'))
    Write-Host "Scoped customer balance code present: $scoped"
    if(-not $scoped){Write-Warning 'The scoped customer balance fix is not present in the active files.'}
}
Write-Host 'A healthy listener and present code do not mean a fresh Insights job completed.' -ForegroundColor Yellow
