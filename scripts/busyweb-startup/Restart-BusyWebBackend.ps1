# BusyWeb phased backend-only clean restart v2.
# Windows PowerShell 5.1, 32-bit BUSY COM hosts. NEVER starts or stops Vite/frontend.
# Safe default: preserve queued WhatsApp PDFs in whatsapp\held, do not send old jobs.
[CmdletBinding()]
param(
    [string]$BackendRoot = 'C:\Busy-Bridge-Access-Cloud',
    [ValidateSet('Hold','Send','Abort')][string]$PendingJobPolicy = 'Hold',
    [ValidateRange(20,240)][int]$ApiReadyTimeoutSeconds = 90,
    [ValidateRange(30,240)][int]$WhatsAppReadyTimeoutSeconds = 90,
    [ValidateRange(1,3)][int]$WhatsAppMaxAttempts = 2,
    [switch]$CheckOnly
)
$ErrorActionPreference = 'Stop'

function Say([string]$Message, [string]$Color='Cyan') {
    $line = '[BUSYWEB-PHASED] ' + $Message
    Write-Host $line -ForegroundColor $Color
    if ($script:LogFile) { Add-Content -LiteralPath $script:LogFile -Value ('{0} {1}' -f (Get-Date -Format 's'),$line) -Encoding UTF8 }
}
function Assert-EmptyProcessing {
    $folder = Join-Path $script:WaRoot 'processing'
    if (Test-Path -LiteralPath $folder -PathType Container) {
        $items = @(Get-ChildItem -LiteralPath $folder -Filter '*.json' -File -ErrorAction Stop)
        if ($items.Count -gt 0) {
            throw "WhatsApp has $($items.Count) in-progress/uncertain job(s). Do not interrupt or retry sends; check delivery manually first."
        }
    }
}
function Get-Owned {
    $all = @(Get-CimInstance Win32_Process -ErrorAction Stop | Where-Object {
        $_.Name -in @('powershell.exe','pwsh.exe','node.exe')
    })
    $result = @()
    foreach ($p in $all) {
        if ([int]$p.ProcessId -eq $PID) { continue }
        $cmd = [string]$p.CommandLine
        if ([string]::IsNullOrWhiteSpace($cmd)) { continue }
        if ($cmd -notmatch [regex]::Escape($script:BackendRoot)) { continue }
        $kind = ''
        if ($p.Name -eq 'node.exe') {
            if ($cmd -match 'whatsapp-worker[\\/]worker\.js') { $kind = 'WhatsAppNode' }
            elseif ($cmd -match '(push-worker|web-push)[\\/].*\.js') { $kind = 'PushNode' }
        } else {
            if ($cmd -match 'BusyWeb-Process\.ps1') {
                if ($cmd -match '(?i)-Component\s+WhatsApp\b') { $kind = 'WhatsAppHost' }
                elseif ($cmd -match '(?i)-Component\s+Executive\b') { $kind = 'ExecutiveHost' }
                elseif ($cmd -match '(?i)-Component\s+Reports\b') { $kind = 'ReportsHost' }
                elseif ($cmd -match '(?i)-Component\s+Main\b') { $kind = 'MainHost' }
            } elseif ($cmd -match 'ExecutiveReportWorker\.ps1|Run-NormalExecutive\.ps1|Run-ExposureWorker\.ps1') {
                $kind = 'ExecutiveHost'
            } elseif ($cmd -match 'busy_api\.ps1') {
                $kind = 'ApiHost'
            }
        }
        if ($kind) {
            $result += [PSCustomObject]@{
                ProcessId = [int]$p.ProcessId; Name = [string]$p.Name
                ParentProcessId = [int]$p.ParentProcessId; Kind = $kind; CommandLine = $cmd
            }
        }
    }
    return @($result)
}
function Get-Listeners {
    return @(Get-NetTCPConnection -State Listen -ErrorAction Stop | Where-Object {
        $_.LocalPort -in @(8081,8085)
    })
}
function Assert-PortOwners([object[]]$owned) {
    $known = @($owned | ForEach-Object { [int]$_.ProcessId })
    foreach ($l in @(Get-Listeners)) {
        $pidNumber = [int]$l.OwningProcess
        if ($pidNumber -in $known) { continue }
        # HTTP.sys may report PID 4 for HttpListener. Allow only when we have a
        # corresponding identifiable BUSY host, then REQUIRE its port to close
        # after that host exits. Never terminate PID 4.
        if ($pidNumber -eq 4) {
            $matching = if ([int]$l.LocalPort -eq 8081) {
                @($owned | Where-Object { $_.Kind -in @('MainHost','ApiHost') })
            } else {
                @($owned | Where-Object { $_.Kind -in @('ReportsHost','ApiHost') })
            }
            if ($matching.Count -gt 0) { continue }
        }
        throw "Port $($l.LocalPort) is held by unrecognized PID $pidNumber. Refusing to stop an unrelated process."
    }
}
function Get-Worker([string]$kind) {
    return @(Get-Owned | Where-Object { $_.Kind -eq $kind })
}
function Assert-PushPidSafe {
    $pidFile = Join-Path $script:BackendRoot 'data\push\worker.pid'
    if (-not (Test-Path -LiteralPath $pidFile -PathType Leaf)) { return }
    $raw = (Get-Content -LiteralPath $pidFile -Raw).Trim()
    $pushPid = 0
    if (-not [int]::TryParse($raw,[ref]$pushPid)) { throw "Invalid Web Push PID file: $pidFile" }
    $live = @(Get-CimInstance Win32_Process -Filter "ProcessId=$pushPid" -ErrorAction Stop)
    if ($live.Count -gt 0) {
        $cmd = [string]$live[0].CommandLine
        if ($cmd -notmatch [regex]::Escape($script:BackendRoot) -or $live[0].Name -ne 'node.exe') {
            throw "Web Push PID marker ($pushPid) refers to an unrecognized live process; refusing to kill it."
        }
        if ($cmd -notmatch '(push|worker\.js)') {
            throw "Web Push PID marker ($pushPid) is ambiguous; inspect the process before restart."
        }
        Say "Stopping identified Web Push child PID=$pushPid" 'Yellow'
        Stop-Process -Id $pushPid -Force -ErrorAction Stop
    }
    # Only remove this marker after verifying its referenced process was absent or stopped.
    Remove-Item -LiteralPath $pidFile -Force -ErrorAction Stop
}
function Wait-PortsFree([int]$timeoutSeconds) {
    $deadline = [datetime]::UtcNow.AddSeconds($timeoutSeconds)
    do {
        if (@(Get-Listeners).Count -eq 0) { return }
        Start-Sleep -Seconds 1
    } while ([datetime]::UtcNow -lt $deadline)
    throw 'Port 8081 and/or 8085 stayed in LISTEN state after stopping identified BusyWeb processes. Refusing to start duplicate servers.'
}
function Get-PendingJobs {
    $folder = Join-Path $script:WaRoot 'pending'
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) { return @() }
    $items = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $folder -Filter '*.json' -File -ErrorAction Stop)) {
        if ($f.BaseName -notmatch '^[a-f0-9]{32}$') { throw "Unexpected pending job filename: $($f.Name)" }
        $j = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        if ([string]$j.jobId -cne [string]$f.BaseName) { throw "Queue job ID mismatch: $($f.Name)" }
        $pdf = Join-Path $folder ($f.BaseName+'.pdf')
        if (-not (Test-Path -LiteralPath $pdf -PathType Leaf)) { throw "Pending job PDF missing: $pdf" }
        $phone = [string]$j.phone
        $last4 = if ($phone.Length -ge 4) { $phone.Substring($phone.Length-4) } else { '?' }
        $items += [PSCustomObject]@{JobId=$f.BaseName;Voucher=[string]$j.vchNo;Format=[string]$j.invoiceFormat;PhoneLast4=$last4;Json=$f.FullName;Pdf=$pdf}
    }
    return @($items)
}
function Hold-QueuedJobs {
    $jobs = @(Get-PendingJobs)
    if ($jobs.Count -eq 0) { return }
    $base = Join-Path $script:WaRoot 'held'
    foreach ($j in $jobs) {
        if (Test-Path -LiteralPath (Join-Path $base $j.JobId)) {
            throw "Held destination already exists for job $($j.JobId). Resolve this conflict before restart."
        }
    }
    foreach ($j in $jobs) {
        $dst = Join-Path $base $j.JobId
        New-Item -Path $dst -ItemType Directory -Force -ErrorAction Stop | Out-Null
        # PDF moved first; JSON last, so an incomplete hold cannot be processed.
        Move-Item -LiteralPath $j.Pdf -Destination $dst -ErrorAction Stop
        Move-Item -LiteralPath $j.Json -Destination $dst -ErrorAction Stop
        Say "HELD job=$($j.JobId) voucher=$($j.Voucher) (not sent; preserved for manual review)" 'Yellow'
    }
}
function Archive-WaStatus {
    $file = Join-Path $script:WaRoot 'connection-status.json'
    if (Test-Path -LiteralPath $file -PathType Leaf) {
        $archived = Join-Path $script:WaRoot ('connection-status.before-phased-'+(Get-Date -Format 'yyyyMMdd-HHmmss-fff')+'.json')
        Move-Item -LiteralPath $file -Destination $archived -ErrorAction Stop
    }
}
function Get-WaStatus([datetime]$startedUtc) {
    $nodeWorkers = @(Get-Worker 'WhatsAppNode')
    if ($nodeWorkers.Count -ne 1) { return $null }
    $file = Join-Path $script:WaRoot 'connection-status.json'
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { return $null }
    try {
        $s = Get-Content -LiteralPath $file -Raw -Encoding UTF8 | ConvertFrom-Json
        $stamp = [datetime]::Parse([string]$s.updatedAt).ToUniversalTime()
        $age = ([datetime]::UtcNow - $stamp).TotalSeconds
        if ($age -lt -5 -or $age -ge 25 -or $stamp -lt $startedUtc.AddSeconds(-3)) { return $null }
        return $s
    } catch { return $null }
}
function Test-ApiHealth([int]$port) {
    $listener = @(Get-Listeners | Where-Object { [int]$_.LocalPort -eq $port })
    if ($listener.Count -eq 0) { return $false }
    try {
        $r = Invoke-WebRequest -Uri ("http://127.0.0.1:$port/health") -UseBasicParsing -TimeoutSec 3 -ErrorAction Stop
        return ([int]$r.StatusCode -eq 200)
    } catch { return $false }
}
function Wait-ApiHealth([int]$port,[int]$timeout) {
    $until = [datetime]::UtcNow.AddSeconds($timeout)
    do {
        if (Test-ApiHealth $port) { Say "API ${port}: HEALTHY" 'Green'; return }
        Start-Sleep -Seconds 2
    } while ([datetime]::UtcNow -lt $until)
    throw "API $port never returned HTTP 200 on /health. See its BusyWeb console. WhatsApp was NOT started."
}
function Get-Chrome {
    $options = @(
        'C:\Program Files\Google\Chrome\Application\chrome.exe',
        'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
        (Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe')
    )
    return @($options | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1)[0]
}
function Start-BackendComponent([string]$component) {
    $exe = Join-Path $env:WINDIR 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
    $runner = Join-Path $script:BackendRoot 'scripts\busyweb-startup\BusyWeb-Process.ps1'
    $args = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -Component {1} -BackendRoot "{2}" -ChromePath "{3}"' -f $runner,$component,$script:BackendRoot,$script:Chrome
    $p = Start-Process -FilePath $exe -ArgumentList $args -WorkingDirectory $script:BackendRoot -PassThru -ErrorAction Stop
    Say "Launched $component PID=$($p.Id)" 'Green'
    return $p.Id
}
function Stop-WaWorkers([string]$reason) {
    Assert-EmptyProcessing
    $p = @(Get-Owned | Where-Object { $_.Kind -in @('WhatsAppNode','WhatsAppHost') })
    $p = @($p | Sort-Object @{Expression={ if ($_.Kind -eq 'WhatsAppNode') { 0 } else { 1 } }})
    foreach ($item in $p) {
        # A ready worker may begin processing at any time. Never kill if a job
        # was claimed between our initial check and the termination phase.
        Assert-EmptyProcessing
        $live = Get-Process -Id $item.ProcessId -ErrorAction SilentlyContinue
        if ($null -ne $live) {
            Stop-Process -Id $item.ProcessId -Force -ErrorAction Stop
            Say "${reason}: stopped $($item.Kind) PID=$($item.ProcessId)" 'Yellow'
        }
    }
    Start-Sleep -Seconds 2
    if (@(Get-Owned | Where-Object { $_.Kind -in @('WhatsAppNode','WhatsAppHost') }).Count -gt 0) {
        throw 'WhatsApp process remained after attempted shutdown; refusing to launch a second worker.'
    }
}

if (-not (Test-Path -LiteralPath $BackendRoot -PathType Container)) { throw "Backend root missing: $BackendRoot" }
$BackendRoot = (Resolve-Path -LiteralPath $BackendRoot).ProviderPath.TrimEnd('\')
$script:BackendRoot = $BackendRoot
$script:WaRoot = Join-Path $BackendRoot 'whatsapp'
$script:LogFile = $null
foreach ($f in @(
    (Join-Path $BackendRoot 'busy_api.ps1'),
    (Join-Path $BackendRoot 'workers\ExecutiveReportWorker.ps1'),
    (Join-Path $BackendRoot 'whatsapp-worker\worker.js'),
    (Join-Path $BackendRoot 'scripts\busyweb-startup\BusyWeb-Process.ps1')
)) { if (-not (Test-Path -LiteralPath $f -PathType Leaf)) { throw "Missing prerequisite: $f" } }
foreach ($folder in @($script:WaRoot,(Join-Path $BackendRoot 'report-jobs'))) {
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) { throw "Missing prerequisite folder: $folder" }
}
foreach ($name in @('Get-CimInstance','Get-NetTCPConnection','node.exe','Invoke-WebRequest')) {
    if (-not (Get-Command $name -ErrorAction SilentlyContinue)) { throw "Missing Windows tool: $name" }
}
$ps32 = Join-Path $env:WINDIR 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
if (-not (Test-Path -LiteralPath $ps32 -PathType Leaf)) { throw '32-bit Windows PowerShell not installed.' }
$script:Chrome = Get-Chrome
if (-not $script:Chrome) { throw 'Google Chrome not found.' }
if (Test-Path -LiteralPath (Join-Path $BackendRoot 'report-jobs\enable-heavy-exposure.flag') -PathType Leaf) {
    throw 'Exposure maintenance flag is active; complete Exposure maintenance before a clean restart.'
}
$mutex = New-Object System.Threading.Mutex($false,'Local\BusyWebBackendCleanRestart')
$locked = $false
try {
    $locked = $mutex.WaitOne(0)
    if (-not $locked) { throw 'Another BusyWeb restart is in progress.' }
    $known = @(Get-Owned)
    Assert-PortOwners $known
    Assert-EmptyProcessing
    $queued = @(Get-PendingJobs)
    Say "Existing owned processes: $($known.Count); queued WhatsApp PDFs: $($queued.Count)"
    if ($queued.Count -gt 0) {
        $queued | Select-Object Voucher,Format,PhoneLast4,JobId | Format-Table -AutoSize | Out-Host
        if ($PendingJobPolicy -eq 'Abort') { throw 'Pending WhatsApp jobs exist. Policy Abort prevents restart.' }
        if ($PendingJobPolicy -eq 'Send') {
            Say 'DANGER: -PendingJobPolicy Send explicitly authorizes all shown queued PDFs.' 'Yellow'
        } else {
            Say 'Safe default: existing queued PDFs will be moved to whatsapp\held (not sent).' 'Yellow'
        }
    }
    if ($CheckOnly) {
        Say 'CheckOnly completed; no services or queue files were changed.' 'Green'
        return
    }
    $logDir = Join-Path $BackendRoot 'scripts\busyweb-startup\logs'
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    $script:LogFile = Join-Path $logDir ('phased-restart-'+(Get-Date -Format 'yyyyMMdd-HHmmss-fff')+'.log')
    Say 'Starting phased backend-only clean restart. Frontend excluded.'
    # Stop only identified backend PowerShell/Node processes, never all processes
    # named node.exe / powershell.exe, and never interfere with unknown 8081/8085 owner.
    $old = @(Get-Owned | Sort-Object @{Expression={if ($_.Name -eq 'node.exe'){0}else{1}}})
    foreach ($p in $old) {
        Assert-EmptyProcessing
        if (Get-Process -Id $p.ProcessId -ErrorAction SilentlyContinue) {
            Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop
            Say "Stopped owned $($p.Kind) PID=$($p.ProcessId)" 'Yellow'
        }
    }
    Assert-PushPidSafe
    # A job may have been claimed during the stop window. Never restart the
    # worker in that case: markInterrupted would make its outcome uncertain.
    Assert-EmptyProcessing
    Wait-PortsFree 20
    Start-Sleep -Seconds 2
    if (@(Get-Owned).Count -gt 0) {
        throw 'An old BusyWeb process is still running. New services were NOT started.'
    }
    # Queue policy applied only AFTER the API and WA processes are stopped.
    if ($PendingJobPolicy -eq 'Hold') { Hold-QueuedJobs }
    elseif ($PendingJobPolicy -eq 'Abort' -and @(Get-PendingJobs).Count -gt 0) {
        throw 'New queued WhatsApp jobs appeared during shutdown; aborting before launch.'
    }
    Archive-WaStatus

    Say 'PHASE 1/4: Starting Main API 8081...'
    [void](Start-BackendComponent 'Main')
    Wait-ApiHealth 8081 $ApiReadyTimeoutSeconds

    Say 'PHASE 2/4: Starting Reports API 8085...'
    [void](Start-BackendComponent 'Reports')
    Wait-ApiHealth 8085 $ApiReadyTimeoutSeconds

    Say 'PHASE 3/4: Starting Executive Worker (normal mode, heavy Exposure OFF)...'
    [void](Start-BackendComponent 'Executive')
    $execReady = $false
    for ($i=0; $i -lt 15; $i++) {
        Start-Sleep -Seconds 2
        if (@(Get-Worker 'ExecutiveHost').Count -eq 1) { $execReady = $true; break }
    }
    if (-not $execReady) { throw 'Executive worker process did not remain running. WhatsApp was NOT started.' }
    Say 'Executive worker process running (its computation readiness is not asserted).' 'Green'

    Say 'PHASE 4/4: Starting WhatsApp after APIs and Executive Worker...'
    $waReady = $false
    for ($attempt = 1; $attempt -le $WhatsAppMaxAttempts; $attempt++) {
        Assert-EmptyProcessing
        $existingWa = @(Get-Owned | Where-Object { $_.Kind -in @('WhatsAppNode','WhatsAppHost') })
        if ($existingWa.Count -gt 0) {
            throw 'WhatsApp worker already exists before attempt. Refusing duplicate Chrome sessions.'
        }
        Archive-WaStatus
        $started = [datetime]::UtcNow
        Say "WhatsApp initialization attempt $attempt/$WhatsAppMaxAttempts"
        [void](Start-BackendComponent 'WhatsApp')
        $deadline = [datetime]::UtcNow.AddSeconds($WhatsAppReadyTimeoutSeconds)
        $consecutive = 0
        $pairing = $false
        do {
            Start-Sleep -Seconds 2
            $s = Get-WaStatus $started
            if ($null -ne $s -and $s.state -eq 'ready' -and $s.ready -eq $true) {
                $consecutive++
                if ($consecutive -ge 2) { $waReady = $true; break }
            } else { $consecutive = 0 }
            if ($null -ne $s -and $s.state -eq 'pairing') { $pairing = $true }
            Assert-EmptyProcessing
        } while ([datetime]::UtcNow -lt $deadline)
        if ($waReady) { Say "WhatsApp READY with live heartbeat on attempt $attempt" 'Green'; break }
        if ($pairing) {
            Say 'WhatsApp is waiting for QR pairing. Leaving worker running; no automatic retry during pairing.' 'Yellow'
            break
        }
        # A late READY may have appeared just as our timeout expired.
        $late = Get-WaStatus $started
        if ($null -ne $late -and $late.ready -eq $true -and $late.state -eq 'ready') {
            Say 'WhatsApp became READY at timeout boundary; preserving session rather than restarting.' 'Yellow'
            $waReady = $true
            break
        }
        Assert-EmptyProcessing
        # Check again before terminating; a late ready event must not be
        # interrupted and potentially cause an accidental duplicate send.
        $beforeStop = Get-WaStatus $started
        if ($null -ne $beforeStop -and $beforeStop.ready -eq $true -and $beforeStop.state -eq 'ready') {
            $waReady = $true
            Say 'Late WhatsApp READY observed; leaving the worker untouched.' 'Yellow'
            break
        }
        Say "WhatsApp did not reach READY within $WhatsAppReadyTimeoutSeconds seconds. Safely restarting WhatsApp only." 'Yellow'
        Stop-WaWorkers 'stalled initialization'
    }
    Say '------------------------------------------'
    Say ('Main API: ' + $(if (Test-ApiHealth 8081) {'HEALTHY'} else {'UNHEALTHY'}))
    Say ('Reports API: ' + $(if (Test-ApiHealth 8085) {'HEALTHY'} else {'UNHEALTHY'}))
    Say ('Executive: ' + $(if (@(Get-Worker 'ExecutiveHost').Count -eq 1) {'RUNNING'} else {'NOT RUNNING'}))
    Say ('WhatsApp: ' + $(if ($waReady) {'READY'} else {'NOT READY'})) $(if ($waReady) {'Green'} else {'Yellow'})
    Say 'Frontend was NOT stopped or started. Heavy dashboard Exposure was NOT enabled.'
    Say "Diagnostic log: $script:LogFile"
    if (-not $waReady) {
        Write-Warning 'WhatsApp is not READY (or requires QR pairing). Main/Reports/Executive remain running. No invoice job was automatically retried.'
    }
} finally {
    if ($locked) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
