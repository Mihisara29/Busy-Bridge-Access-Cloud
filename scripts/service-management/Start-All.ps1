#requires -Version 5.1
[CmdletBinding()]
param([string]$BackendRoot = (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent))
$ErrorActionPreference='Stop'
$ps32='C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
$api=Join-Path $BackendRoot 'busy_api.ps1'
$worker=Join-Path $BackendRoot 'workers\ExecutiveReportWorker.ps1'
foreach($f in @($ps32,$api,$worker)) {
    if(-not (Test-Path -LiteralPath $f -PathType Leaf)){ throw "Required file not found: $f" }
}

function Test-Listener([int]$Port) {
    $client=New-Object Net.Sockets.TcpClient
    try {
        $async=$client.BeginConnect('127.0.0.1',$Port,$null,$null)
        try { return ($async.AsyncWaitHandle.WaitOne(800) -and $client.Connected) }
        finally { $async.AsyncWaitHandle.Close() }
    } catch { return $false }
    finally { $client.Close() }
}
function Test-Health([int]$Port) {
    try {
        $resp=Invoke-WebRequest -UseBasicParsing -Uri "http://127.0.0.1:$Port/health" -TimeoutSec 5
        return ($resp.StatusCode -eq 200)
    } catch { return $false }
}
function Start-Api([int]$Port,[string]$Mode) {
    if(Test-Listener $Port) {
        if(Test-Health $Port){ Write-Host "$Mode API port $Port already healthy; not starting duplicate." -ForegroundColor Green; return }
        throw "Port $Port has a listener but is not healthy. Investigate its owner; not starting a duplicate."
    }
    # -NoExit keeps error output visible after a startup failure.
    $arguments='-NoProfile -ExecutionPolicy Bypass -NoExit -File "' + $api + '" -Port ' + $Port + ' -ServiceMode ' + $Mode
    $p=Start-Process -FilePath $ps32 -ArgumentList $arguments -PassThru
    Write-Host "Started $Mode API on port $Port (PID $($p.Id))." -ForegroundColor Green
}
function Get-Workers {
    @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction Stop | Where-Object {
        $_.CommandLine -and
        $_.CommandLine.IndexOf('ExecutiveReportWorker.ps1',[StringComparison]::OrdinalIgnoreCase) -ge 0 -and
        $_.CommandLine.IndexOf($BackendRoot,[StringComparison]::OrdinalIgnoreCase) -ge 0
    })
}
Write-Host "Single source: $BackendRoot" -ForegroundColor Cyan
Start-Api -Port 8081 -Mode 'Main'
Start-Api -Port 8085 -Mode 'Reports'
$workers=@(Get-Workers)
if($workers.Count -gt 1){
    Write-Warning ("Multiple workers already running: " + (($workers | ForEach-Object {$_.ProcessId}) -join ', '))
    throw 'Do not run multiple ExecutiveReportWorker instances. Stop duplicates manually after verifying PIDs.'
}
if($workers.Count -eq 1) {
    Write-Host "Executive worker already running (PID $($workers[0].ProcessId))." -ForegroundColor Green
} else {
    $arguments='-NoProfile -ExecutionPolicy Bypass -NoExit -File "' + $worker + '" -BridgeRoot "' + $BackendRoot + '"'
    $p=Start-Process -FilePath $ps32 -ArgumentList $arguments -PassThru
    Write-Host "Started Executive worker (PID $($p.Id))." -ForegroundColor Green
}
Write-Host 'Waiting briefly for API initialization...' -ForegroundColor Cyan
Start-Sleep -Seconds 8
foreach($port in @(8081,8085)) {
    if(Test-Health $port){Write-Host "Port $port /health: OK" -ForegroundColor Green}
    else { Write-Warning "Port $port is not healthy yet. Its own PowerShell window contains startup details; initialization can take longer." }
}
Write-Host 'These checks do not confirm dashboard financial results. Run Check-Services.ps1 for details.' -ForegroundColor Yellow
