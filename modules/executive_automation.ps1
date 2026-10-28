# BUSYWEB_AUTOMATED_EXPOSURE_PILOT_V1
# Add-on to executive_jobs.ps1. Enabled only for explicitly allowlisted companies.
# No COM/SQL/valuation code executes in these functions or in an HTTP request.
$script:BusyWebAutomationConfigPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'config\executive-automation.json'

function Get-BusyWebAutomationSettings {
    $path = $script:BusyWebAutomationConfigPath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
    $value = Read-ExecutiveJobJson -Path $path
    if ($null -eq $value -or $value.enabled -ne $true) { return $null }
    if ([string]$value.timeZoneId -ne 'Sri Lanka Standard Time') { throw 'Automation pilot requires Sri Lanka Standard Time.' }
    if ([int]$value.refreshMinutes -lt 60 -or [int]$value.refreshMinutes -gt 1440) { throw 'refreshMinutes must be 60-1440.' }
    if ([int]$value.maxOutstandingJobs -lt 1 -or [int]$value.maxOutstandingJobs -gt 5) { throw 'maxOutstandingJobs must be 1-5.' }
        $configuredCompanies = @(
        $value.companies | Where-Object { $null -ne $_ }
    )

    if ([string]$value.mode -eq 'registered-on-demand') {
        if ($configuredCompanies.Count -ne 0) {
            throw 'On-demand mode cannot contain fixed scheduler targets.'
        }
    }
    elseif ([string]::IsNullOrWhiteSpace([string]$value.mode)) {
        if ($configuredCompanies.Count -lt 1 -or
            $configuredCompanies.Count -gt 10) {
            throw 'Pilot requires 1-10 company entries.'
        }
    }
    else {
        throw 'Unsupported automation mode.'
    }
    foreach ($company in @($value.companies)) {
        if ([string]$company.instanceId -notmatch '^[A-Za-z0-9_-]{1,80}$' -or [string]$company.companyCode -notmatch '^[A-Za-z0-9_-]{1,80}$') {
            throw 'Invalid company allowlist identity.'
        }
    }
    return $value
}
function Test-BusyWebExposureAutomationCompany {
    param(
        [string]$InstanceId,
        [string]$CompanyCode
    )

    $settings = Get-BusyWebAutomationSettings

    if ($null -eq $settings) {
        return $false
    }

    if ($InstanceId -notmatch '^[A-Za-z0-9_-]{1,80}$' -or
        $CompanyCode -notmatch '^[A-Za-z0-9_-]{1,80}$') {
        return $false
    }

    if ([string]$settings.mode -eq 'registered-on-demand') {

        $backendRoot = Split-Path -Parent (
            Split-Path -Parent $script:BusyWebAutomationConfigPath
        )

        $registryPath = Join-Path $backendRoot 'instances.json'

        if (-not (Test-Path -LiteralPath $registryPath -PathType Leaf)) {
            return $false
        }

        try {
            $registry = Get-Content -LiteralPath $registryPath `
                -Raw -Encoding UTF8 | ConvertFrom-Json

            if ($registry -is [array]) {
                $instances = @($registry)
            }
            elseif ($null -ne $registry.instances) {
                $instances = @($registry.instances)
            }
            else {
                return $false
            }

            foreach ($instance in $instances) {
                if ($null -eq $instance) {
                    continue
                }

                if ([string]$instance.id -ine $InstanceId) {
                    continue
                }

                foreach ($company in @($instance.companies)) {
                    if ($null -eq $company) {
                        continue
                    }

                    if ([string]$company.code -ieq $CompanyCode) {
                        return $true
                    }
                }
            }
        }
        catch {
            Write-Warning (
                '[AUTO-EXPOSURE] Company registry validation failed: ' +
                $_.Exception.Message
            )
        }

        return $false
    }

    # Preserve existing pilot behavior for installations
    # that have not enabled dynamic routing.
    foreach ($company in @($settings.companies)) {
        if ([string]$company.instanceId -ieq $InstanceId -and
            [string]$company.companyCode -ieq $CompanyCode) {
            return $true
        }
    }

    return $false
}
function Get-BusyWebCompanyToday {
    $timezone = [System.TimeZoneInfo]::FindSystemTimeZoneById('Sri Lanka Standard Time')
    return ([System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow,$timezone)).Date
}
function Test-BusyWebVerifiedExposure {
    param($Value,[string]$InstanceId,[string]$CompanyCode,[string]$AsOf)
    if ($null -eq $Value -or $Value.success -ne $true -or [string]$Value.section -ne 'exposure') {return $false}
    if ([string]$Value.instanceId -ine $InstanceId -or [string]$Value.companyCode -ine $CompanyCode -or [string]$Value.asOf -ne $AsOf) {return $false}
    if ($null -eq $Value.stock -or $null -eq $Value.outstanding) {return $false}
    if (@($Value.warnings | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0) {return $false}
    if ([string]$Value.stock.asOf -ne $AsOf -or [string]$Value.outstanding.asOf -ne $AsOf) {return $false}
    return $true
}
function Get-BusyWebExposureCachedResult {
    param([string]$Id,[string]$InstanceId,[string]$CompanyCode,[string]$AsOf)
    if ($Id -notmatch '^[a-f0-9]{64}$') {return $null}
    $paths=Get-ExecutiveJobPaths -Id $Id
    $status=Read-ExecutiveJobJson -Path $paths.status
    if ($null -eq $status -or [string]$status.state -ne 'completed') {return $null}
    $value=Read-ExecutiveJobJson -Path $paths.result
    if (-not (Test-BusyWebVerifiedExposure -Value $value -InstanceId $InstanceId -CompanyCode $CompanyCode -AsOf $AsOf)) {return $null}
    return $value
}
function New-BusyWebExposureResponse {
    param($Snapshot,[string]$RefreshState)
    $copy=@{}
    foreach($p in $Snapshot.PSObject.Properties) {$copy[$p.Name]=$p.Value}
    if ($Snapshot -is [System.Collections.IDictionary]) {foreach($key in $Snapshot.Keys){$copy[[string]$key]=$Snapshot[$key]}}
    if (-not [string]::IsNullOrWhiteSpace($RefreshState)) {
        $copy.snapshotStale=$true
        $copy.refreshState=$RefreshState
        $copy.snapshotServedUtc=([datetime]::UtcNow.ToString('o'))
    }
    return @{success=$true;httpStatus=200;data=$copy}
}

# Shared between the authenticated Reports endpoint and the background scheduler.
# RefreshToday is scheduler-only; public HTTP paths never pass it.
function Get-BusyWebAutomatedExposure {
    param([string]$InstanceId,[string]$CompanyCode,[string]$AsOf,[switch]$RefreshToday)
    $settings=Get-BusyWebAutomationSettings
    if ($null -eq $settings -or -not (Test-BusyWebExposureAutomationCompany -InstanceId $InstanceId -CompanyCode $CompanyCode)) {
        return @{success=$false;httpStatus=503;error='Automated financial snapshots are not enabled for this company.'}
    }
    if (-not (Test-Path -LiteralPath $script:ExecutiveJobRoot -PathType Container)) {
        return @{success=$false;httpStatus=503;error='Executive report queue is not available.'}
    }
    $today=Get-BusyWebCompanyToday
    $date=$today
    if (-not [string]::IsNullOrWhiteSpace($AsOf)) {
        if ($AsOf -notmatch '^\d{4}-\d{2}-\d{2}$') {return @{success=$false;httpStatus=400;error='asOf must be yyyy-MM-dd.'}}
        try {$date=[datetime]::ParseExact($AsOf,'yyyy-MM-dd',[System.Globalization.CultureInfo]::InvariantCulture)}
        catch {return @{success=$false;httpStatus=400;error='Invalid asOf date.'}}
    }
    $fyToday=if($today.Month -lt 4){$today.Year-1}else{$today.Year}
    $fyDate=if($date.Month -lt 4){$date.Year-1}else{$date.Year}
    if ($date -gt $today -or $fyToday -ne $fyDate) {
        return @{success=$false;httpStatus=400;error='Only dates in the active fiscal year through today are supported.'}
    }
    $dateText=$date.ToString('yyyy-MM-dd')
    $identity=($InstanceId+'|'+$CompanyCode+'|'+$dateText+'|exposure').ToLowerInvariant()
    $sha=[System.Security.Cryptography.SHA256]::Create()
    try{$key=([BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($identity)))).Replace('-','').ToLowerInvariant()}
    finally{$sha.Dispose()}
    $pointerPath=Get-ExecutiveJobPointerPath -LogicalKey $key
    # Serialize the read/compare/enqueue/pointer operation across API and scheduler
    # processes, including different Windows logon sessions.
    $lockFile=Join-Path $script:ExecutiveJobRoot 'automated-exposure-enqueue.lock'
    $lock=$null
    for($tryCount=0;$tryCount -lt 15 -and $null -eq $lock;$tryCount++) {
        try{$lock=[System.IO.File]::Open($lockFile,[System.IO.FileMode]::OpenOrCreate,[System.IO.FileAccess]::ReadWrite,[System.IO.FileShare]::None)}
        catch [System.IO.IOException]{Start-Sleep -Milliseconds 150}
    }
    if ($null -eq $lock) {return @{success=$false;httpStatus=503;error='Report scheduler is busy. Try again shortly.'}}
    try {
        $pointer=Read-ExecutiveJobJson -Path $pointerPath
        $activeId=if($null -ne $pointer){[string]$pointer.jobId}else{''}
        $priorId=if($null -ne $pointer){[string]$pointer.previousJobId}else{''}
        $activeStatus=$null
        $activePaths=$null
        $verified=$null
        if($activeId -match '^[a-f0-9]{64}$') {
            $activePaths=Get-ExecutiveJobPaths -Id $activeId
            $activeStatus=Read-ExecutiveJobJson -Path $activePaths.status
            $verified=Get-BusyWebExposureCachedResult -Id $activeId -InstanceId $InstanceId -CompanyCode $CompanyCode -AsOf $dateText
        }
        if($null -eq $verified -and $priorId -match '^[a-f0-9]{64}$') {
            $verified=Get-BusyWebExposureCachedResult -Id $priorId -InstanceId $InstanceId -CompanyCode $CompanyCode -AsOf $dateText
        }
        $state=if($null -ne $activeStatus){[string]$activeStatus.state}else{''}
        if ($state -in @('queued','running')) {
            if($null -ne $verified){return (New-BusyWebExposureResponse -Snapshot $verified -RefreshState $state)}
            # Never create duplicates, including orphaned processes: operator must inspect.
            return @{success=$true;httpStatus=202;data=@{success=$true;status=$state;jobId=$activeId;retryAfterMs=15000}}
        }
        $shouldCreate=($null -eq $verified)
        if ($null -ne $verified -and $RefreshToday -and $date -eq $today) {
            $finish=$null
            if($state -eq 'completed') {
                try{$finish=[datetime]::Parse([string]$activeStatus.finishedUtc).ToUniversalTime()}catch{}
            }
            if($null -eq $finish){$shouldCreate=$true}
            elseif((([datetime]::UtcNow)-$finish).TotalMinutes -ge [int]$settings.refreshMinutes){$shouldCreate=$true}
        }
        if (-not $shouldCreate) {return (New-BusyWebExposureResponse -Snapshot $verified -RefreshState '')}
        if($state -eq 'failed') {
            $failureTime=$null
            try{$failureTime=[datetime]::Parse([string]$activeStatus.finishedUtc).ToUniversalTime()}catch{}
            if($null -eq $failureTime -or (([datetime]::UtcNow)-$failureTime).TotalMinutes -lt 60) {
                if($null -ne $verified){return (New-BusyWebExposureResponse -Snapshot $verified -RefreshState 'failed')}
                return @{success=$false;httpStatus=503;error='Financial snapshot failed. Automatic retry cooldown is active.'}
            }
        }
        # An incomplete generation with a live queue/processing file must never be replaced.
        if($null -ne $activePaths -and ((Test-Path -LiteralPath $activePaths.request) -or (Test-Path -LiteralPath $activePaths.processing))) {
            return @{success=$false;httpStatus=503;error='An existing financial report is awaiting operator recovery.'}
        }
        $count=0
        foreach($file in @(Get-ChildItem -LiteralPath $script:ExecutiveJobRoot -Filter '*.request.json' -File -ErrorAction SilentlyContinue)) {
            $queuedJob=Read-ExecutiveJobJson -Path $file.FullName
            if($null -ne $queuedJob -and [string]$queuedJob.section -eq 'exposure'){$count++}
        }
        foreach($file in @(Get-ChildItem -LiteralPath $script:ExecutiveJobRoot -Filter '*.processing.json' -File -ErrorAction SilentlyContinue)) {
            $queuedJob=Read-ExecutiveJobJson -Path $file.FullName
            if($null -ne $queuedJob -and [string]$queuedJob.section -eq 'exposure'){$count++}
        }
        if ($count -ge [int]$settings.maxOutstandingJobs) {
            if($null -ne $verified){return (New-BusyWebExposureResponse -Snapshot $verified -RefreshState 'queue-busy')}
            return @{success=$false;httpStatus=503;error='Financial report queue is busy. Try again later.'}
        }
        $id=New-ExecutiveJobId -LogicalKey $key
        $paths=Get-ExecutiveJobPaths -Id $id
        $created=[datetime]::UtcNow.ToString('o')
        $job=@{jobId=$id;instanceId=$InstanceId;companyCode=$CompanyCode;asOf=$dateText;section='exposure';createdUtc=$created}
        Save-ExecutiveJobJson -Path $paths.status -Value @{state='queued';createdUtc=$created;jobId=$id}
        Save-ExecutiveJobJson -Path $paths.request -Value $job
        $lastVerifiedId=''
        if($null -ne $verified){
            if($activeId -match '^[a-f0-9]{64}$' -and $state -eq 'completed'){$lastVerifiedId=$activeId}
            elseif($priorId -match '^[a-f0-9]{64}$'){$lastVerifiedId=$priorId}
        }
        Save-ExecutiveJobJson -Path $pointerPath -Value @{jobId=$id;previousJobId=$lastVerifiedId;section='exposure';createdUtc=$created}
        Write-Host ('[AUTO-EXPOSURE] Queued '+$InstanceId+'/'+$CompanyCode+' '+$dateText+' job='+$id)
        if($null -ne $verified){return (New-BusyWebExposureResponse -Snapshot $verified -RefreshState 'queued')}
        return @{success=$true;httpStatus=202;data=@{success=$true;status='queued';jobId=$id;retryAfterMs=15000}}
    }
    finally {$lock.Dispose()}
}
