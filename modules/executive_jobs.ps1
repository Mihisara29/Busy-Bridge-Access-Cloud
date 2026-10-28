# BUSYWEB_STABLE_REPORT_QUEUE_V3 - PowerShell 5.1, company scoped disk queue.
# IMPORTANT: No financial work, COM calls, or expensive queries run in HTTP handler.
$script:ExecutiveJobRoot = Join-Path (Split-Path $PSScriptRoot -Parent) 'report-jobs'
# Completed snapshots are time-bounded. Expiration only queues work when a client asks.
# This catches external BUSY Desktop edits without database triggers or write hooks.
$script:ExecutiveJobTtlSeconds = 900  # 15 minutes
$script:ExecutiveExposureTtlSeconds = 7200  # 2 hours; full stock+outstanding can take 8+ minutes
$script:ExecutiveInsightsTtlSeconds = 1800  # 30 minutes; avoids repeated Insights rebuilds

function Get-ExecutiveJobPaths {
    param([string]$Id)
    if ($Id -notmatch '^[a-f0-9]{64}$') { throw 'Invalid job identifier' }
    return @{
        request=(Join-Path $script:ExecutiveJobRoot "$Id.request.json")
        processing=(Join-Path $script:ExecutiveJobRoot "$Id.processing.json")
        status=(Join-Path $script:ExecutiveJobRoot "$Id.status.json")
        result=(Join-Path $script:ExecutiveJobRoot "$Id.result.json")
    }
}
function Save-ExecutiveJobJson {
    param([string]$Path, $Value)
    $tmp = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    $Value | ConvertTo-Json -Depth 16 -Compress | Set-Content -LiteralPath $tmp -Encoding UTF8
    Move-Item -LiteralPath $tmp -Destination $Path -Force
}
function Read-ExecutiveJobJson {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try { return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { return $null }
}
function Get-ExecutiveDashboardAsync {
    param([string]$InstanceId,[string]$CompanyCode,[string]$AsOf,[ValidateSet("core","exposure","insights_today","insights_mtd")][string]$Section="exposure",[bool]$ForceRefresh=$false)
    # Pilot only: allowlisted Exposure GETs use the dedicated safe queue lane.
    # Other companies and Fast/Core/Insights retain the original implementation.
    if ($Section -eq 'exposure' -and (Get-Command Test-BusyWebExposureAutomationCompany -CommandType Function -ErrorAction SilentlyContinue)) {
        if (Test-BusyWebExposureAutomationCompany -InstanceId $InstanceId -CompanyCode $CompanyCode) {
            return (Get-BusyWebAutomatedExposure -InstanceId $InstanceId -CompanyCode $CompanyCode -AsOf $AsOf)
        }
    }
    Get-ExecutiveDashboardJob -InstanceId $InstanceId -CompanyCode $CompanyCode -AsOf $AsOf -Section $Section -ForceRefresh $ForceRefresh
}
function Get-ExecutiveExposureAsync {
    param([string]$InstanceId,[string]$CompanyCode,[string]$AsOf,[bool]$ForceRefresh=$false)
    Get-ExecutiveDashboardJob -InstanceId $InstanceId -CompanyCode $CompanyCode -AsOf $AsOf -Section "exposure" -ForceRefresh $ForceRefresh
}
# Each generation has a unique immutable job ID. The logical key holds only a pointer.
# No refresh can overwrite another generation's status, input, or result.
function Get-ExecutiveJobPointerPath {
    param([string]$LogicalKey)
    if($LogicalKey -notmatch '^[a-f0-9]{64}$'){ throw 'Invalid logical key' }
    return Join-Path $script:ExecutiveJobRoot "$LogicalKey.active.json"
}
function New-ExecutiveJobId {
    param([string]$LogicalKey)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($LogicalKey + '|' + [guid]::NewGuid().ToString('N'))
        return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-','').ToLowerInvariant()
    } finally { $sha.Dispose() }
}
function Test-ExecutiveJobResult {
    param($Value,[string]$Section)
    if($null -eq $Value -or $Value.success -ne $true -or [string]$Value.section -ne $Section) { return $false }
    if($Section -eq 'core') { return ($null -ne $Value.periods -and $null -ne $Value.trend) }
    if($Section -in @('insights_today','insights_mtd')) { return ($null -ne $Value.topCustomers -and $null -ne $Value.bestSellingItems) }
    return ($null -ne $Value.stock -or $null -ne $Value.outstanding -or ($Value.PSObject.Properties.Name -contains 'stock' -and $Value.PSObject.Properties.Name -contains 'outstanding'))
}
function Get-ExecutiveDashboardJob {
    param([string]$InstanceId,[string]$CompanyCode,[string]$AsOf,[ValidateSet('core','exposure','insights_today','insights_mtd')][string]$Section='exposure',[bool]$ForceRefresh=$false)
    if ([string]::IsNullOrWhiteSpace($InstanceId) -or [string]::IsNullOrWhiteSpace($CompanyCode)) {
        return @{success=$false;httpStatus=400;error='Select a company.'}
    }
    $today=(Get-Date).Date
    if([string]::IsNullOrWhiteSpace($AsOf)) { $asOfDate=$today }
    else {
        if($AsOf -notmatch '^\d{4}-\d{2}-\d{2}$'){return @{success=$false;httpStatus=400;error='asOf must be yyyy-MM-dd.'}}
        try {$asOfDate=[datetime]::ParseExact($AsOf,'yyyy-MM-dd',[System.Globalization.CultureInfo]::InvariantCulture)}
        catch {return @{success=$false;httpStatus=400;error='Invalid asOf date.'}}
    }
    $currentFy=if($today.Month -lt 4){$today.Year-1}else{$today.Year}
    $requestedFy=if($asOfDate.Month -lt 4){$asOfDate.Year-1}else{$asOfDate.Year}
    if($asOfDate -gt $today -or $requestedFy -ne $currentFy){return @{success=$false;httpStatus=400;error='Only dates in the active fiscal year through today are supported.'}}
    if(-not (Test-Path -LiteralPath $script:ExecutiveJobRoot -PathType Container)){
        return @{success=$false;httpStatus=503;error='Report worker storage is missing.'}
    }
    $dateText=$asOfDate.ToString('yyyy-MM-dd')
    $identity=($InstanceId+'|'+$CompanyCode+'|'+$dateText+'|'+$Section).ToLowerInvariant()
    $sha=[System.Security.Cryptography.SHA256]::Create()
    try {$key=([BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($identity)))).Replace('-','').ToLowerInvariant()}
    finally {$sha.Dispose()}
    $pointerPath=Get-ExecutiveJobPointerPath -LogicalKey $key
    $pointer=Read-ExecutiveJobJson -Path $pointerPath
    $now=[datetime]::UtcNow
    $previous=$null
    $previousId=''
    $activeId=''
    $activeStatus=$null
    if($null -ne $pointer -and [string]$pointer.jobId -match '^[a-f0-9]{64}$') {
        $activeId=[string]$pointer.jobId
        $activePaths=Get-ExecutiveJobPaths -Id $activeId
        $activeStatus=Read-ExecutiveJobJson -Path $activePaths.status
        if($null -ne $activeStatus -and [string]$activeStatus.state -eq 'completed') {
            $candidate=Read-ExecutiveJobJson -Path $activePaths.result
            if(Test-ExecutiveJobResult -Value $candidate -Section $Section) {
                $previous=$candidate
                $previousId=$activeId
            }
        }
        if($null -eq $previous -and [string]$pointer.previousJobId -match '^[a-f0-9]{64}$') {
            $candidateId=[string]$pointer.previousJobId
            $candidatePaths=Get-ExecutiveJobPaths -Id $candidateId
            $candidateStatus=Read-ExecutiveJobJson -Path $candidatePaths.status
            if($null -ne $candidateStatus -and [string]$candidateStatus.state -eq 'completed') {
                $candidate=Read-ExecutiveJobJson -Path $candidatePaths.result
                if(Test-ExecutiveJobResult -Value $candidate -Section $Section) {
                    $previous=$candidate
                    $previousId=$candidateId
                }
            }
        }
    }
    # Every response containing stale data explicitly says so. Never call it 'up to date'.
    function Send-PreviousSnapshot {
        param($Snapshot,[string]$State)
        $copy=@{}
        foreach($prop in $Snapshot.PSObject.Properties) {$copy[$prop.Name]=$prop.Value}
        # Hashtable results also supported for local test harnesses.
        if($Snapshot -is [System.Collections.IDictionary]){
            foreach($k in $Snapshot.Keys){$copy[[string]$k]=$Snapshot[$k]}
        }
        $copy['snapshotStale']=$true
        $copy['refreshState']=$State
        $copy['snapshotServedUtc']=$now.ToString('o')
        return @{success=$true;httpStatus=200;data=$copy}
    }
    # V3: Executive Exposure is the expensive, company-wide stock/outstanding report.
    # Serving a cached snapshot is safe. Browser traffic MUST NOT automatically
    # trigger a new full valuation, even if the original cache TTL has expired.
    # The worker runs heavy Exposure only when an operator explicitly opts in.
    # Full-company stock and outstanding remain intentionally protected: even a
    # browser refresh must not queue an unprocessable 9-minute Exposure job.
    # Administrator opt-in requires BOTH a local gate file AND a worker started
    # with BUSYWEB_EXECUTIVE_ENABLE_EXPOSURE=1. Never enable by public request.
    $exposureGate = Join-Path $script:ExecutiveJobRoot 'enable-heavy-exposure.flag'
    if ($Section -eq 'exposure' -and -not (Test-Path -LiteralPath $exposureGate)) {
        if ($null -ne $previous) {
            $fresh=$false
            if ($null -ne $activeStatus -and [string]$activeStatus.state -eq 'completed') {
                try { $fresh=(($now-[datetime]::Parse([string]$activeStatus.finishedUtc).ToUniversalTime()).TotalSeconds -lt $script:ExecutiveExposureTtlSeconds) }
                catch { $fresh=$false }
            }
            if ($fresh -and -not $ForceRefresh) { return @{success=$true;httpStatus=200;data=$previous} }
            return (Send-PreviousSnapshot -Snapshot $previous -State 'heavy-exposure-paused')
        }
        return @{success=$false;httpStatus=503;error='Exposure has no verified cached snapshot. Expensive Stock and Outstanding calculations are paused. Have an administrator enable maintenance recalculation.'}
    }
    if($null -ne $activeStatus) {
        $state=[string]$activeStatus.state
        if($state -in @('queued','running')) {
            if($null -ne $previous) { return (Send-PreviousSnapshot -Snapshot $previous -State $state) }
            # Actual worker can spend minutes on stock and outstanding calculations.
            # No busy-wait at 2s. Frontend backs off further and can cancel on unmount.
            try {$age=($now-[datetime]::Parse([string]$activeStatus.createdUtc).ToUniversalTime()).TotalMinutes}
            catch {$age=999}
            if($age -lt 30){ return @{success=$true;httpStatus=202;data=@{success=$true;status=$state;jobId=$activeId;retryAfterMs=15000}} }
            # A job can sit queued behind multi-minute financial reports. Do not
            # manufacture another job while its request/processing file exists.
            $hasRequest=(Test-Path -LiteralPath $activePaths.request)
            $hasProcessing=(Test-Path -LiteralPath $activePaths.processing)
            if ($hasRequest -or $hasProcessing) {
                return @{success=$false;httpStatus=503;error='Report job is delayed in the worker queue. Check ExecutiveReportWorker logs; no duplicate was enqueued.'}
            }
            # If a job has no queue/processing file and no valid result, it was
            # interrupted or orphaned. Recover ONE new generation for insights.
            # Old generations remain available for diagnostics and audit.
            if ($Section -notlike 'insights_*' -or $null -ne $previous) {
                return @{success=$false;httpStatus=503;error='Report job exceeded 30 minutes. Check report worker; manual refresh after diagnosis.'}
            }
            Write-Warning ('[EXEC-JOBS] Recovering orphaned '+$Section+' report, job '+$activeId)
            # Fall through to create a new immutable generation.
        }
        if($state -eq 'failed') {
            if($null -ne $previous -and -not $ForceRefresh) { return (Send-PreviousSnapshot -Snapshot $previous -State 'failed') }
            # A failed query should not be repeatedly requeued by automatic polling.
            if(-not $ForceRefresh) {
                if ($Section -notlike 'insights_*') {
                    return @{success=$false;httpStatus=503;error='Last report job failed. Check report worker log; retry manually.'}
                }
                try { $failedAge=($now-[datetime]::Parse([string]$activeStatus.finishedUtc).ToUniversalTime()).TotalMinutes }
                catch { $failedAge=0 }
                if ($failedAge -lt 20) {
                    return @{success=$false;httpStatus=503;error='Insights calculation failed; paused for 20 minutes to avoid repeated SQL load. Check worker log.'}
                }
                Write-Warning ('[EXEC-JOBS] Re-queuing failed '+$Section+' after cooldown.')
            }
        }
        if($state -eq 'completed' -and $null -ne $previous) {
            $ttl=if($Section -eq 'exposure'){$script:ExecutiveExposureTtlSeconds}elseif($Section -like 'insights_*'){$script:ExecutiveInsightsTtlSeconds}else{$script:ExecutiveJobTtlSeconds}
            try {$ageSeconds=($now-[datetime]::Parse([string]$activeStatus.finishedUtc).ToUniversalTime()).TotalSeconds}
            catch {$ageSeconds=[double]::PositiveInfinity}
            if(-not $ForceRefresh -and $ageSeconds -ge 0 -and $ageSeconds -lt $ttl){
                return @{success=$true;httpStatus=200;data=$previous}
            }
        }
    }
    # Limit queue buildup across companies and multiple browser tabs.
    # Paused heavyweight Exposure files do not count against Core/Insights
    # fast-lane capacity. They will remain queued for an opt-in maintenance run.
    $queued=0
    foreach($rq in @(Get-ChildItem -LiteralPath $script:ExecutiveJobRoot -Filter '*.request.json' -File -ErrorAction SilentlyContinue)) {
        $q=Read-ExecutiveJobJson -Path $rq.FullName
        if($null -ne $q -and [string]$q.section -ne 'exposure') { $queued++ }
    }
    if($queued -ge 12) {
        if($null -ne $previous) {return (Send-PreviousSnapshot -Snapshot $previous -State 'queue-busy')}
        return @{success=$false;httpStatus=503;error='Report queue is busy. Try later; voucher operations have priority.'}
    }
    $id=New-ExecutiveJobId -LogicalKey $key
    $paths=Get-ExecutiveJobPaths -Id $id
    $created=$now.ToString('o')
    $job=@{jobId=$id;instanceId=$InstanceId;companyCode=$CompanyCode;asOf=$dateText;section=$Section;createdUtc=$created}
    # Atomic immutable results. Persist before advancing active pointer.
    Save-ExecutiveJobJson -Path $paths.status -Value @{state='queued';createdUtc=$created;jobId=$id}
    Save-ExecutiveJobJson -Path $paths.request -Value $job
    Save-ExecutiveJobJson -Path $pointerPath -Value @{jobId=$id;previousJobId=$previousId;section=$Section;createdUtc=$created}
    if($null -ne $previous){return (Send-PreviousSnapshot -Snapshot $previous -State 'queued')}
    return @{success=$true;httpStatus=202;data=@{success=$true;status='queued';jobId=$id;retryAfterMs=15000}}

}
