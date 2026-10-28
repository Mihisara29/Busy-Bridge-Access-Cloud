# BusyWeb local full-backend launcher (Windows PowerShell 5.1, 64-bit Windows)
# Install at C:\Busy-Bridge-Access-Cloud\Start-BusyWeb-Local.ps1
# Components: Main 8081, Reports 8085, Executive (core/insights, heavy OFF),
# WhatsApp (with pending-job safeguard), dedicated Exposure, snapshot scheduler.
# Optionally start Vite frontend with -IncludeFrontend.
# This is an interactive local launcher, NOT a Windows VPS service installer.
[CmdletBinding()]
param(
    [string]$BackendRoot = '',
    [string]$FrontendRoot = 'C:\busy-bevy - Local',
    [switch]$IncludeFrontend,
    [switch]$SkipWhatsApp,
    [switch]$ProcessPendingWhatsAppJobs,
    [switch]$CheckOnly
)
$ErrorActionPreference = 'Stop'

function Write-Step([string]$Message, [string]$Color = 'Cyan') {
    Write-Host ("[BUSYWEB-LOCAL] " + $Message) -ForegroundColor $Color
}
function Require-File([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Required file missing: $Path" }
}
function Test-Listener([int]$Port) {
    # Fail closed rather than starting on a port whose owner cannot be inspected.
    $all = @(Get-NetTCPConnection -State Listen -ErrorAction Stop | Where-Object { $_.LocalPort -eq $Port })
    return ($all.Count -gt 0)
}
function Test-ApiHealthy([int]$Port) {
    try {
        $h = Invoke-RestMethod -Uri ("http://127.0.0.1:$Port/health") -TimeoutSec 4 -ErrorAction Stop
        return ($h.success -eq $true -and [string]$h.serverId -eq 'busy-server-native-unified')
    } catch { return $false }
}
function Test-FileLockHeld([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    $stream = $null
    try {
        $stream = [System.IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
        return $false
    }
    catch [System.IO.IOException] { return $true }
    finally { if ($null -ne $stream) { $stream.Dispose() } }
}
function Test-ExecutiveMutexHeld {
    $m = New-Object System.Threading.Mutex($false, 'Local\BusyWebExecutiveExposureWorker')
    $locked = $false
    try {
        try { $locked = $m.WaitOne(0) }
        catch [System.Threading.AbandonedMutexException] { $locked = $true }
        return (-not $locked)
    }
    finally {
        if ($locked) { $m.ReleaseMutex() }
        $m.Dispose()
    }
}
function Test-ComponentPresent([string]$Component) {
    # Use title AND command-line; the old BusyWeb launcher sometimes hides its path.
    $windowTitle = switch ($Component) {
        'Exposure' { 'BusyWeb - Exposure' }
        'Scheduler' { 'BusyWeb - Exposure Scheduler' }
        default { "BusyWeb - $Component" }
    }
    $namedHosts = @(Get-Process -Name powershell,pwsh -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowTitle -eq $windowTitle })
    if ($namedHosts.Count -gt 0) { return $true }

    $items = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe' OR Name='node.exe'" -ErrorAction Stop)
    foreach ($item in $items) {
        $cmd = [string]$item.CommandLine
        if ([string]::IsNullOrWhiteSpace($cmd)) { continue }
        switch ($Component) {
            'Executive' {
                if ($cmd -match 'ExecutiveReportWorker\.ps1' -or
                    ($cmd -match 'BusyWeb-Process\.ps1' -and $cmd -match '(?i)-Component\s+"?Executive\b')) { return $true }
            }
            'WhatsApp' {
                if ($cmd -match 'whatsapp-worker[\\/]worker\.js' -or
                    ($cmd -match 'BusyWeb-Process\.ps1' -and $cmd -match '(?i)-Component\s+"?WhatsApp\b')) { return $true }
            }
            'Exposure' { if ($cmd -match 'ExecutiveExposureWorker\.ps1') { return $true } }
            'Scheduler' { if ($cmd -match 'ExecutiveSnapshotScheduler\.ps1') { return $true } }
            default {
                if ($cmd -match 'BusyWeb-Process\.ps1' -and
                    $cmd -match ('(?i)-Component\s+"?' + [regex]::Escape($Component) + '\b')) { return $true }
            }
        }
    }
    return $false
}
function Start-BaseComponent([string]$Component, [string]$Runner, [string]$Ps32, [string]$Root, [string]$Frontend, [string]$Chrome) {
    foreach ($p in @($Runner,$Root,$Frontend,$Chrome)) {
        if ($p.Contains('"')) { throw 'Paths containing quotation marks are not supported.' }
    }
    $argsText = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -Component {1} -BackendRoot "{2}" -FrontendRoot "{3}"' -f $Runner,$Component,$Root,$Frontend
    if ($Component -eq 'WhatsApp') { $argsText += (' -ChromePath "{0}"' -f $Chrome) }
    $p = Start-Process -FilePath $Ps32 -ArgumentList $argsText -WorkingDirectory $Root -PassThru -ErrorAction Stop
    Write-Step ("Launched $Component PID=$($p.Id); its own window displays startup errors.") 'Green'
}
function Get-ActiveExposureProcessing([string]$Jobs) {
    $found = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $Jobs -Filter '*.processing.json' -File)) {
        try {
            $j = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            if ([string]$j.section -eq 'exposure') { $found += $f.Name }
        }
        catch {
            # A processing file we cannot classify is not safe to ignore.
            throw "Cannot classify processing job '$($f.Name)': $($_.Exception.Message)"
        }
    }
    return @($found)
}
function Start-AutomationWorker([string]$Name,[string]$ScriptName,[string]$LockName,[string]$Root,[string]$Ps32,[string]$Jobs,[string]$LogDir) {
    $lockPath = Join-Path $Jobs $LockName
    if (Test-FileLockHeld $lockPath) {
        Write-Step ("$Name already owns the exclusive lock; skipped.") 'Green'
        return
    }
    if (Test-ComponentPresent $Name) {
        Write-Step ("$Name process exists but its lock is free; not launching a duplicate. Inspect it.") 'Yellow'
        return
    }
    if ($Name -eq 'Exposure') {
        $inflight = @(Get-ActiveExposureProcessing $Jobs)
        if ($inflight.Count -gt 0) {
            Write-Step ("Exposure job(s) already marked processing while the worker lock is free: $($inflight -join ', '). Refusing unsafe restart.") 'Yellow'
            return
        }
    }
    $scriptFile = Join-Path $Root ("workers\" + $ScriptName)
    Require-File $scriptFile
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
    $stdout = Join-Path $LogDir ("$($ScriptName.Replace('.ps1',''))-$stamp.out.log")
    $stderr = Join-Path $LogDir ("$($ScriptName.Replace('.ps1',''))-$stamp.err.log")
    $argsText = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -BridgeRoot "{1}"' -f $scriptFile,$Root
    # Log capture means the worker can be diagnosed even if its console is hidden.
    $p = Start-Process -FilePath $Ps32 -ArgumentList $argsText -WorkingDirectory $Root `
        -RedirectStandardOutput $stdout -RedirectStandardError $stderr -WindowStyle Hidden -PassThru -ErrorAction Stop
    Write-Step ("Launched $Name PID=$($p.Id); stdout=$stdout ; stderr=$stderr") 'Green'
}
function Show-SnapshotSummary([string]$Root,[object]$Settings) {
    $jobs = Join-Path $Root 'report-jobs'
    $tz = [string]$Settings.timeZoneId
    if ([string]::IsNullOrWhiteSpace($tz)) { $tz = 'Sri Lanka Standard Time' }
    $zone = [TimeZoneInfo]::FindSystemTimeZoneById($tz)
    $asOf = [TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow,$zone).ToString('yyyy-MM-dd')
    Write-Step ("Current company date ($tz): $asOf")
    foreach ($company in @($Settings.companies)) {
        $instance = [string]$company.instanceId
        $code = [string]$company.companyCode
        $identity = ("$instance|$code|$asOf|exposure").ToLowerInvariant()
        $sha = [Security.Cryptography.SHA256]::Create()
        try { $key = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($identity))).Replace('-','').ToLowerInvariant() }
        finally { $sha.Dispose() }
        $pointerFile = Join-Path $jobs ("$key.active.json")
        if (-not (Test-Path -LiteralPath $pointerFile)) {
            Write-Step ("$instance/$code : no snapshot pointer for $asOf yet (scheduler may enqueue shortly).") 'Yellow'
            continue
        }
        try {
            $pointer = Get-Content -LiteralPath $pointerFile -Raw -Encoding UTF8 | ConvertFrom-Json
            $id = [string]$pointer.jobId
            if ($id -notmatch '^[a-f0-9]{64}$') { throw 'Invalid snapshot job ID.' }
            $statusFile = Join-Path $jobs ("$id.status.json")
            $resultFile = Join-Path $jobs ("$id.result.json")
            $status = if (Test-Path -LiteralPath $statusFile) { Get-Content -LiteralPath $statusFile -Raw -Encoding UTF8 | ConvertFrom-Json } else { $null }
            $verified = $false
            if (([string]$status.state -eq 'completed') -and (Test-Path -LiteralPath $resultFile)) {
                $r = Get-Content -LiteralPath $resultFile -Raw -Encoding UTF8 | ConvertFrom-Json
                $warns = @($r.warnings | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                $verified = ($r.success -eq $true -and [string]$r.section -eq 'exposure' -and
                    [string]$r.instanceId -eq $instance -and [string]$r.companyCode -eq $code -and
                    [string]$r.asOf -eq $asOf -and $null -ne $r.stock -and $null -ne $r.outstanding -and $warns.Count -eq 0)
            }
            Write-Step ("$instance/$code : state=$($status.state) verified=$verified job=$id")
        }
        catch { Write-Step ("$instance/$code : snapshot read error: $($_.Exception.Message)") 'Yellow' }
    }
}

# ===== STATUS / READINESS REPORT (no report jobs are created here) =====
function Get-DetectedComponentPids([string]$Component) {
    $title = switch ($Component) {
        'Exposure' { 'BusyWeb - Exposure' }
        'Scheduler' { 'BusyWeb - Exposure Scheduler' }
        default { "BusyWeb - $Component" }
    }
    $ids = New-Object 'System.Collections.Generic.List[int]'
    foreach ($p in @(Get-Process -Name powershell,pwsh -ErrorAction SilentlyContinue)) {
        if ($p.MainWindowTitle -eq $title -and -not $ids.Contains([int]$p.Id)) {
            [void]$ids.Add([int]$p.Id)
        }
    }
    # A background PowerShell worker usually has no window. Match its script.
    # Output only its PID; never print command lines (which may contain secrets).
    try {
        $all = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe' OR Name='node.exe'" -ErrorAction Stop)
        foreach ($item in $all) {
            $cmd = [string]$item.CommandLine
            $matchFound = $false
            switch ($Component) {
                'Executive' {
                    $matchFound = ($cmd -match 'ExecutiveReportWorker\.ps1' -or
                        ($cmd -match 'BusyWeb-Process\.ps1' -and $cmd -match '(?i)-Component\s+"?Executive\b'))
                }
                'Exposure' { $matchFound = ($cmd -match 'ExecutiveExposureWorker\.ps1') }
                'Scheduler' { $matchFound = ($cmd -match 'ExecutiveSnapshotScheduler\.ps1') }
                'WhatsApp' {
                    $matchFound = ($cmd -match 'whatsapp-worker[\\/]worker\.js' -or
                        ($cmd -match 'BusyWeb-Process\.ps1' -and $cmd -match '(?i)-Component\s+"?WhatsApp\b'))
                }
                'Frontend' {
                    $matchFound = ($cmd -match '(?i)[\\/]vite[\\/]' -or
                        ($cmd -match 'BusyWeb-Process\.ps1' -and $cmd -match '(?i)-Component\s+"?Frontend\b'))
                }
            }
            if ($matchFound -and -not $ids.Contains([int]$item.ProcessId)) {
                [void]$ids.Add([int]$item.ProcessId)
            }
        }
    }
    catch {
        Write-Step "Process ID scan unavailable: $($_.Exception.Message)" 'Yellow'
    }
    return @($ids.ToArray() | Sort-Object -Unique)
}
function Get-ListenerIds([int]$Port) {
    try {
        return @(Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction Stop |
            Select-Object -ExpandProperty OwningProcess -Unique | Sort-Object -Unique)
    }
    catch { return @() }
}
function Test-FrontendHttp {
    try {
        $r = Invoke-WebRequest -Uri 'http://127.0.0.1:8080/' -UseBasicParsing -TimeoutSec 4 -ErrorAction Stop
        return ([int]$r.StatusCode -ge 200 -and [int]$r.StatusCode -lt 400)
    } catch { return $false }
}
function Test-RecentWhatsAppHeartbeat([string]$Root) {
    $heartbeat = Join-Path $Root 'whatsapp\connection-status.json'
    if (-not (Test-Path -LiteralPath $heartbeat -PathType Leaf)) { return $false }
    try {
        $data = Get-Content -LiteralPath $heartbeat -Raw -Encoding UTF8 | ConvertFrom-Json
        $timestamp = [datetime]::Parse([string]$data.updatedAt).ToUniversalTime()
        $age = ([datetime]::UtcNow - $timestamp).TotalSeconds
        return ($age -ge 0 -and $age -lt 30)
    } catch { return $false }
}
function New-StatusRow([string]$Name,[string]$Port,[string]$State,[string]$ProcessIds,[string]$Note,[bool]$Required) {
    return [pscustomobject]@{Name=$Name;Port=$Port;State=$State;Pid=$ProcessIds;Note=$Note;Required=$Required}
}
function Get-BackendStatusRows([string]$Root,[bool]$AutomationEnabled,[bool]$WantFrontend,[bool]$WhatsAppSkipped) {
    $rows = New-Object System.Collections.ArrayList
    foreach ($entry in @(@('Main API',8081),@('Reports API',8085))) {
        $label = [string]$entry[0]
        $port = [int]$entry[1]
        $ids = @(Get-ListenerIds $port)
        $pidLabel = if ($ids.Count -gt 0) { $ids -join ',' } else { '-' }
        if ($ids.Count -eq 0) {
            [void]$rows.Add((New-StatusRow $label ([string]$port) 'STOPPED' $pidLabel 'No listening process' $true))
        }
        elseif (Test-ApiHealthy $port) {
            [void]$rows.Add((New-StatusRow $label ([string]$port) 'HEALTHY' $pidLabel 'BUSY /health verified' $true))
        }
        else {
            [void]$rows.Add((New-StatusRow $label ([string]$port) 'UNHEALTHY' $pidLabel 'Listener exists; BUSY health failed' $true))
        }
    }

    $exIds = @(Get-DetectedComponentPids 'Executive')
    $exPid = if ($exIds.Count -gt 0) { $exIds -join ',' } else { '-' }
    $exLock = Test-ExecutiveMutexHeld
    if ($exLock -and $exIds.Count -gt 0) {
        $exState='RUNNING';$exNote='Process found; exclusive mutex held'
    }
    elseif ($exLock) {
        $exState='LOCK HELD';$exNote='Mutex held; worker PID not identified'
    }
    elseif ($exIds.Count -gt 0) {
        $exState='STARTING?';$exNote='Process found; mutex not yet held'
    }
    else { $exState='STOPPED';$exNote='No worker or mutex detected' }
    [void]$rows.Add((New-StatusRow 'Executive worker' '-' $exState $exPid $exNote $true))

    $waIds = @(Get-DetectedComponentPids 'WhatsApp')
    $waPid = if ($waIds.Count -gt 0) { $waIds -join ',' } else { '-' }
    $waHb = Test-RecentWhatsAppHeartbeat $Root
    if ($waIds.Count -gt 0 -and $waHb) { $waState='RUNNING';$waNote='Process + fresh heartbeat (login not checked)' }
    elseif ($waIds.Count -gt 0) { $waState='DETECTED';$waNote='Process found; heartbeat not fresh' }
    elseif ($waHb) { $waState='HEARTBEAT';$waNote='Recent heartbeat; PID not identified' }
    elseif ($WhatsAppSkipped) { $waState='SKIPPED';$waNote='Requested -SkipWhatsApp' }
    else { $waState='STOPPED';$waNote='No process or recent heartbeat' }
    [void]$rows.Add((New-StatusRow 'WhatsApp worker' '-' $waState $waPid $waNote (-not $WhatsAppSkipped)))

    $jobsDir = Join-Path $Root 'report-jobs'
    foreach ($entry in @(@('Exposure worker','Exposure','automated-exposure-worker.lock'),@('Snapshot scheduler','Scheduler','automated-exposure-scheduler.lock'))) {
        $name=[string]$entry[0];$type=[string]$entry[1]
        $workerIds = @(Get-DetectedComponentPids $type)
        $pids = if ($workerIds.Count -gt 0) { $workerIds -join ',' } else { '-' }
        $locked = Test-FileLockHeld (Join-Path $jobsDir ([string]$entry[2]))
        if (-not $AutomationEnabled) { $state='DISABLED';$note='Automation configuration is disabled' }
        elseif ($locked -and $workerIds.Count -gt 0) { $state='RUNNING';$note='Worker process + exclusive lock' }
        elseif ($locked) { $state='LOCK HELD';$note='Exclusive lock held; PID not identified' }
        elseif ($workerIds.Count -gt 0) { $state='STARTING?';$note='Process exists; lock free' }
        else { $state='STOPPED';$note='No lock holder or worker detected' }
        [void]$rows.Add((New-StatusRow $name '-' $state $pids $note $AutomationEnabled))
    }

    $frontPids = @(Get-ListenerIds 8080)
    $frontendPid = if ($frontPids.Count -gt 0) { $frontPids -join ',' } else { '-' }
    if ($frontPids.Count -eq 0) { $frontState='STOPPED';$frontNote='Port 8080 not listening' }
    elseif (Test-FrontendHttp) { $frontState='HTTP OK';$frontNote='Port responds to HTTP; app identity not verified' }
    else { $frontState='PORT IN USE';$frontNote='Port listening, HTTP response failed' }
    [void]$rows.Add((New-StatusRow 'Frontend (Vite)' '8080' $frontState $frontendPid $frontNote $WantFrontend))
    return @($rows.ToArray())
}
function Show-BackendStatusReport([string]$Root,[bool]$AutomationEnabled,[bool]$WantFrontend,[bool]$WhatsAppSkipped) {
    Write-Host ''
    Write-Host '================= BUSYWEB STATUS =================' -ForegroundColor Cyan
    Write-Host ('Checked at: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    Write-Host ('Backend: ' + $Root)
    $rows = @(Get-BackendStatusRows $Root $AutomationEnabled $WantFrontend $WhatsAppSkipped)
    $fmt = '{0,-20} {1,-5} {2,-12} {3,-12} {4}'
    Write-Host ($fmt -f 'COMPONENT','PORT','STATUS','PID','DETAIL') -ForegroundColor Cyan
    Write-Host ('-' * 91)
    foreach ($row in $rows) {
        $color = switch ($row.State) {
            'HEALTHY' { 'Green' }
            'RUNNING' { 'Green' }
            'HTTP OK' { 'Green' }
            'DISABLED' { 'DarkGray' }
            'SKIPPED' { 'DarkGray' }
            'STOPPED' { 'Red' }
            'UNHEALTHY' { 'Red' }
            default { 'Yellow' }
        }
        Write-Host ($fmt -f $row.Name,$row.Port,$row.State,$row.Pid,$row.Note) -ForegroundColor $color
    }
    Write-Host ('-' * 91)
    $bad = @($rows | Where-Object {
        if (-not $_.Required) { return $false }
        if ($_.Name -eq 'Frontend (Vite)') { return $_.State -ne 'HTTP OK' }
        if ($_.Name -match 'API$') { return $_.State -ne 'HEALTHY' }
        return $_.State -ne 'RUNNING'
    })
    if ($bad.Count -eq 0) {
        Write-Host 'OVERALL: REQUIRED COMPONENTS READY' -ForegroundColor Green
    } else {
        Write-Host ('OVERALL: NEEDS ATTENTION (' + ($bad.Name -join ', ') + ')') -ForegroundColor Yellow
    }
    Write-Host 'Status meanings: RUNNING = process and readiness signal; HTTP OK = responsive port only.' -ForegroundColor DarkGray
    Write-Host '================ END STATUS REPORT ================' -ForegroundColor Cyan
    return ($bad.Count -eq 0)
}

# The final table is printed even when startup encounters an error.
$script:startupFailed = $false
$script:startupMessage = ''
$script:reportRoot = ''
$script:reportAutomation = $false
try {
# Resolve the backend root after parameter binding: Windows PowerShell 5.1 can
# leave $PSScriptRoot empty when used as a parameter default expression.
if ([string]::IsNullOrWhiteSpace($BackendRoot)) {
    $launcherPath = [string]$PSCommandPath
    if ([string]::IsNullOrWhiteSpace($launcherPath)) {
        $launcherPath = [string]$MyInvocation.MyCommand.Path
    }
    if (-not [string]::IsNullOrWhiteSpace($launcherPath)) {
        $BackendRoot = Split-Path -Path $launcherPath -Parent
    }
}
if ([string]::IsNullOrWhiteSpace($BackendRoot)) {
    throw 'Cannot determine backend folder. Pass -BackendRoot explicitly.'
}
if (-not (Test-Path -LiteralPath $BackendRoot -PathType Container)) { throw "Backend root missing: $BackendRoot" }
$BackendRoot = (Resolve-Path -LiteralPath $BackendRoot).ProviderPath
$script:reportRoot = $BackendRoot
$ps32 = Join-Path $env:WINDIR 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
$runner = Join-Path $BackendRoot 'scripts\busyweb-startup\BusyWeb-Process.ps1'
$jobs = Join-Path $BackendRoot 'report-jobs'
$autoConfig = Join-Path $BackendRoot 'config\executive-automation.json'
Require-File $ps32
Require-File $runner
Require-File (Join-Path $BackendRoot 'busy_api.ps1')
Require-File (Join-Path $BackendRoot 'workers\ExecutiveReportWorker.ps1')
if (-not (Test-Path -LiteralPath $jobs -PathType Container)) { throw "Missing job folder: $jobs" }

Write-Step "Backend root: $BackendRoot"
Write-Step 'Existing healthy processes will be kept. No jobs will be manually created or deleted.'

$settings = $null
if (Test-Path -LiteralPath $autoConfig -PathType Leaf) {
    $settings = Get-Content -LiteralPath $autoConfig -Raw -Encoding UTF8 | ConvertFrom-Json
}
$automationOn = ($null -ne $settings -and $settings.enabled -eq $true)
$script:reportAutomation = $automationOn
if ($automationOn) {
    Require-File (Join-Path $BackendRoot 'workers\ExecutiveExposureWorker.ps1')
    Require-File (Join-Path $BackendRoot 'workers\ExecutiveSnapshotScheduler.ps1')
    Require-File (Join-Path $BackendRoot 'modules\executive_automation.ps1')
    Write-Step 'Automated Exposure is enabled in the installed configuration.' 'Green'
} else {
    Write-Step 'Automated Exposure is disabled or not installed; no automated Exposure workers will start.' 'Yellow'
}

# Ports must be either free or owned by a healthy BusyWeb API.
foreach ($port in @(8081,8085)) {
    if (Test-Listener $port) {
        if (-not (Test-ApiHealthy $port)) { throw "Port $port is occupied by a listener that does not return a valid BUSY health response. No additional processes started." }
        Write-Step "API $port already healthy; skipped." 'Green'
    }
    elseif (-not $CheckOnly) {
        $component = if ($port -eq 8081) { 'Main' } else { 'Reports' }
        Start-BaseComponent $component $runner $ps32 $BackendRoot $FrontendRoot ''
    }
    else { Write-Step "API $port not running." 'Yellow' }
}

if (-not $CheckOnly) {
    # Wait for API readiness before starting workers. Don't restart or stop them on failure.
    foreach ($port in @(8081,8085)) {
        $healthy = $false
        for ($n=0;$n -lt 20;$n++) {
            if (Test-ApiHealthy $port) { $healthy = $true; break }
            Start-Sleep -Seconds 3
        }
        if (-not $healthy) { throw "API $port did not become healthy. Inspect the $port console; workers were not started." }
        Write-Step "API $port healthy." 'Green'
    }
}

# The normal Executive worker is exclusively for Core/Insights (heavy Exposure OFF).
if (Test-ExecutiveMutexHeld) {
    Write-Step 'An Executive worker holds the named mutex; leaving it untouched. Confirm its heavy mode separately.' 'Green'
}
elseif (Test-ComponentPresent 'Executive') {
    Write-Step 'Executive process is present but mutex not acquired yet. Skipping duplicate.' 'Yellow'
}
elseif (-not $CheckOnly) {
    Start-BaseComponent 'Executive' $runner $ps32 $BackendRoot $FrontendRoot ''
    # Process creation is not readiness. Give the normal worker time to load modules.
    for ($n=0; $n -lt 12; $n++) {
        if (Test-ExecutiveMutexHeld) { break }
        Start-Sleep -Seconds 2
    }
}
else { Write-Step 'Normal Executive worker missing.' 'Yellow' }

# WhatsApp worker has no HTTP port; preserve existing heartbeat/pending-job safeguards.
if ($SkipWhatsApp) { Write-Step 'WhatsApp startup explicitly skipped.' 'Yellow' }
elseif (Test-ComponentPresent 'WhatsApp') { Write-Step 'WhatsApp process already present; skipped.' 'Green' }
else {
    $recentHeartbeat = $false
    $heartbeat = Join-Path $BackendRoot 'whatsapp\connection-status.json'
    if (Test-Path -LiteralPath $heartbeat) {
        try {
            $info = Get-Content -LiteralPath $heartbeat -Raw -Encoding UTF8 | ConvertFrom-Json
            $ageSeconds = ([datetime]::UtcNow - [datetime]::Parse([string]$info.updatedAt).ToUniversalTime()).TotalSeconds
            $recentHeartbeat = ($ageSeconds -ge 0 -and $ageSeconds -lt 30)
        }
        catch { Write-Step 'WhatsApp heartbeat could not be parsed; process scan used instead.' 'Yellow' }
    }
    if ($recentHeartbeat) { Write-Step 'Fresh WhatsApp heartbeat found; not launching a duplicate worker.' 'Yellow' }
    else {
        $pendingDir = Join-Path $BackendRoot 'whatsapp\pending'
        $pendingCount = if (Test-Path -LiteralPath $pendingDir -PathType Container) {
            @(Get-ChildItem -LiteralPath $pendingDir -Filter '*.json' -File).Count
        } else { 0 }
        if ($pendingCount -gt 0 -and -not $ProcessPendingWhatsAppJobs) {
            Write-Step "WhatsApp NOT started: $pendingCount pending job(s) may be sent on startup. Review first; use -ProcessPendingWhatsAppJobs only if intended." 'Yellow'
        }
        elseif ($CheckOnly) { Write-Step 'WhatsApp worker not running.' 'Yellow' }
        else {
            $chrome = @(
                'C:\Program Files\Google\Chrome\Application\chrome.exe',
                'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
                (Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe')
            ) | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
            if (-not $chrome) { Write-Step 'WhatsApp not started: Google Chrome executable not found.' 'Yellow' }
            elseif (-not (Get-Command node.exe -ErrorAction SilentlyContinue)) { Write-Step 'WhatsApp not started: node.exe missing from PATH.' 'Yellow' }
            elseif (-not (Test-Path -LiteralPath (Join-Path $BackendRoot 'whatsapp-worker\node_modules\whatsapp-web.js') -PathType Container)) {
                Write-Step 'WhatsApp not started: npm dependencies missing.' 'Yellow'
            }
            else { Start-BaseComponent 'WhatsApp' $runner $ps32 $BackendRoot $FrontendRoot $chrome }
        }
    }
}

if ($automationOn) {
    $workerLock = Join-Path $jobs 'automated-exposure-worker.lock'
    $schedulerLock = Join-Path $jobs 'automated-exposure-scheduler.lock'
    if (-not $CheckOnly) {
        $logDir = Join-Path $BackendRoot 'logs\automation-startup'
        if (-not (Test-Path -LiteralPath $logDir -PathType Container)) {
            New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop | Out-Null
        }
        Start-AutomationWorker 'Exposure' 'ExecutiveExposureWorker.ps1' 'automated-exposure-worker.lock' $BackendRoot $ps32 $jobs $logDir
        Start-AutomationWorker 'Scheduler' 'ExecutiveSnapshotScheduler.ps1' 'automated-exposure-scheduler.lock' $BackendRoot $ps32 $jobs $logDir
        for ($n=0;$n -lt 15;$n++) {
            if ((Test-FileLockHeld $workerLock) -and (Test-FileLockHeld $schedulerLock)) { break }
            Start-Sleep -Seconds 3
        }
    }
    $workerReady = Test-FileLockHeld $workerLock
    $schedulerReady = Test-FileLockHeld $schedulerLock
    Write-Step "Exposure worker lock held: $workerReady" $(if($workerReady){'Green'}else{'Yellow'})
    Write-Step "Scheduler lock held: $schedulerReady" $(if($schedulerReady){'Green'}else{'Yellow'})
    Show-SnapshotSummary $BackendRoot $settings
    if (-not $workerReady -or -not $schedulerReady) {
        Write-Step 'One or more automation workers are NOT ready. Inspect logs\automation-startup; do not launch duplicates blindly.' 'Yellow'
    }
}

if ($IncludeFrontend) {
    if (-not (Test-Path -LiteralPath $FrontendRoot -PathType Container)) { throw "Frontend root not found: $FrontendRoot" }
    if (Test-Listener 8080) { Write-Step 'Frontend 8080 already has a listener; skipped.' 'Green' }
    elseif ($CheckOnly) { Write-Step 'Frontend 8080 not running.' 'Yellow' }
    elseif (-not (Get-Command npm.cmd -ErrorAction SilentlyContinue)) { Write-Step 'Frontend not started: npm.cmd not in PATH.' 'Yellow' }
    else {
        Start-BaseComponent 'Frontend' $runner $ps32 $BackendRoot $FrontendRoot ''
        for ($n=0; $n -lt 12; $n++) {
            if (Test-FrontendHttp) { break }
            Start-Sleep -Seconds 2
        }
    }
}
Write-Step 'Startup/check actions completed; verifying actual component readiness below.' 'Cyan'
}
catch {
    $script:startupFailed = $true
    $script:startupMessage = $_.Exception.Message
    Write-Step ("STARTUP ERROR: " + $script:startupMessage) 'Red'
}
finally {
    $rootForReport = if (-not [string]::IsNullOrWhiteSpace($script:reportRoot)) {
        $script:reportRoot
    } else {
        'C:\Busy-Bridge-Access-Cloud'
    }
    if (Test-Path -LiteralPath $rootForReport -PathType Container) {
        try {
            $configPath = Join-Path $rootForReport 'config\executive-automation.json'
            $enabledForReport = $script:reportAutomation
            if (Test-Path -LiteralPath $configPath -PathType Leaf) {
                $latest = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
                $enabledForReport = ($latest.enabled -eq $true)
            }
            $ready = Show-BackendStatusReport $rootForReport $enabledForReport ([bool]$IncludeFrontend) ([bool]$SkipWhatsApp)
            if (-not $ready) { $script:startupFailed = $true }
        }
        catch {
            $script:startupFailed = $true
            Write-Step ("Could not finish status report: " + $_.Exception.Message) 'Red'
        }
    }
    else {
        $script:startupFailed = $true
        Write-Step 'No valid backend root available for final status inspection.' 'Red'
    }
    Write-Step 'The report is a point-in-time check; re-run with -CheckOnly any time.' 'Cyan'
    Write-Step 'Local interactive startup only (not a supervised VPS service).' 'Yellow'
}
if ($script:startupFailed) { exit 1 }
exit 0


