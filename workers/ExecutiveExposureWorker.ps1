# Separate dedicated process. Run with Windows PowerShell 5.1 (32-bit for BUSY COM).
param([string]$BridgeRoot=(Split-Path $PSScriptRoot -Parent),[int]$PollMs=1000)
$ErrorActionPreference='Stop'
try { $Host.UI.RawUI.WindowTitle = 'BusyWeb - Exposure' } catch {}
$moduleDir=Join-Path $BridgeRoot 'modules'
$required=@('config.ps1','utils.ps1','connection.ps1','auth_native.ps1','items.ps1','accounts.ps1','salesman.ps1','vouchers.ps1','boms.ps1','vch_setting.ps1','web_approval.ps1','reports.ps1','account_ledger.ps1','dashboard.ps1','executive_insights.ps1','executive_jobs.ps1','executive_automation.ps1')
foreach($name in $required) {
    $file=Join-Path $moduleDir $name
    if(-not (Test-Path -LiteralPath $file)){throw "Missing worker module: $file"}
    . $file
}
Set-StrictMode -Off
$jobRoot=Join-Path $BridgeRoot 'report-jobs'
if(-not (Test-Path -LiteralPath $jobRoot)){throw 'Run installation setup for report-jobs first.'}
# Prevent concurrent worker processes from duplicating financial calculations.
# Dedicated Exposure worker: file lock works across Windows sessions and accounts.
# Pilot never modifies BUSY source data; uses verified existing calculation functions.
$settings=Get-BusyWebAutomationSettings
if($null -eq $settings){throw 'Automated Exposure is disabled in config; refusing worker startup.'}
$lockFile=Join-Path $jobRoot 'automated-exposure-worker.lock'
try{$exclusive=[System.IO.File]::Open($lockFile,[System.IO.FileMode]::OpenOrCreate,[System.IO.FileAccess]::ReadWrite,[System.IO.FileShare]::None)}
catch{throw 'An Automated Exposure worker already owns the cross-session lock.'}
try {
    Write-Host '[AUTO-EXPOSURE-WORKER] Dedicated lane ready. Normal Executive/8085 unaffected.' -ForegroundColor Green
    while($true) {
        # BUSYWEB_WORKER_FAST_FIRST_V3: Core/Insights before multi-minute Exposure.
        # Order by priority then creation time. Reports are still strictly serial.
        $ranked=@(Get-ChildItem -LiteralPath $jobRoot -Filter '*.request.json' -File | ForEach-Object {
            $p=Read-ExecutiveJobJson -Path $_.FullName
            $priority=switch ([string]$p.section) {
                'core' { 0; break }
                'insights_today' { 1; break }
                'insights_mtd' { 1; break }
                'exposure' { 2; break }
                default { 3 }
            }
            [pscustomobject]@{File=$_;Priority=$priority;Created=$_.CreationTimeUtc}
        } | Sort-Object Priority,Created)
        foreach($entry in $ranked) {
            $file=$entry.File
            if($file.Name -notmatch '^[a-f0-9]{64}\.request\.json$'){continue}
            $id=$file.Name.Substring(0,64)
            $paths=Get-ExecutiveJobPaths -Id $id
            $queuedJob=Read-ExecutiveJobJson -Path $paths.request
            if($null -eq $queuedJob){ continue }
            if ([string]$queuedJob.section -ne 'exposure') { continue }
            if (-not (Test-BusyWebExposureAutomationCompany -InstanceId ([string]$queuedJob.instanceId) -CompanyCode ([string]$queuedJob.companyCode))) {
                # A disabled/non-allowlisted tenant never enters expensive calculations.
                continue
            }
            $identity=([string]$queuedJob.instanceId+'|'+[string]$queuedJob.companyCode+'|'+[string]$queuedJob.asOf+'|'+[string]$queuedJob.section).ToLowerInvariant()
            $sha=[System.Security.Cryptography.SHA256]::Create()
            try {$logical=([BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($identity)))).Replace('-','').ToLowerInvariant()}
            finally {$sha.Dispose()}
            $active=Read-ExecutiveJobJson -Path (Get-ExecutiveJobPointerPath -LogicalKey $logical)
            if($null -eq $active) {
                # A partially-written enqueue has not published its active pointer.
                # Leave it in the queue for the next scan; never compute an orphan.
                continue
            }
            if([string]$active.jobId -ne $id) {
                # Superseded queue entry must never consume minutes of SQL time.
                Remove-Item -LiteralPath $paths.request -Force -ErrorAction SilentlyContinue
                continue
            }
            try { Move-Item -LiteralPath $paths.request -Destination $paths.processing -ErrorAction Stop }
            catch {continue}
            $job=Read-ExecutiveJobJson -Path $paths.processing
            if($null -eq $job){
                Save-ExecutiveJobJson -Path $paths.status -Value @{state='failed';createdUtc=([datetime]::UtcNow.ToString('o'));error='Invalid job payload.'}
                Remove-Item -LiteralPath $paths.processing -Force -ErrorAction SilentlyContinue
                continue
            }
            $created=[string]$job.createdUtc
            try {
                # Validate before computing; worker never trusts a client-provided script or SQL query.
                if([string]$job.jobId -ne $id -or [string]$job.instanceId -notmatch '^[A-Za-z0-9_-]{1,80}$' -or [string]$job.companyCode -notmatch '^[A-Za-z0-9_-]{1,80}$' -or [string]$job.asOf -notmatch '^\d{4}-\d{2}-\d{2}$') {throw 'Invalid job identity.'}
                $section = if ([string]::IsNullOrWhiteSpace([string]$job.section)) { 'exposure' } else { [string]$job.section }
                if ($section -notin @('core','exposure','insights_today','insights_mtd')) { throw 'Unsupported dashboard job section.' }
                Save-ExecutiveJobJson -Path $paths.status -Value @{state='running';jobId=$id;createdUtc=$created;startedUtc=([datetime]::UtcNow.ToString('o'))}
                Write-Host ('[EXEC-WORKER] Starting '+$section+' '+[string]$job.instanceId+'/'+[string]$job.companyCode+' job='+$id) -ForegroundColor Cyan
                if ($section -in @('insights_today','insights_mtd')) {
                    $period = $section.Substring(9)
                    $data = Get-ExecutiveSalesInsights -InstanceId ([string]$job.instanceId) -CompanyCode ([string]$job.companyCode) -AsOf ([string]$job.asOf) -Period $period
                } else {
                    $data = Get-ExecutiveDashboard -InstanceId ([string]$job.instanceId) -CompanyCode ([string]$job.companyCode) -AsOf ([string]$job.asOf) -Section $section -ForceRefresh $true
                }
                if(-not $data.success){throw ('Dashboard calculation failed: '+[string]$data.error)}
                if (-not (Test-BusyWebVerifiedExposure -Value $data -InstanceId ([string]$job.instanceId) -CompanyCode ([string]$job.companyCode) -AsOf ([string]$job.asOf))) {
                    throw 'Calculated Exposure failed strict Stock/Outstanding/date/warnings verification.'
                }
                Save-ExecutiveJobJson -Path $paths.result -Value $data
                # Confirm the result can be read before declaring the job complete.
                $persisted = Read-ExecutiveJobJson -Path $paths.result
                if (-not (Test-BusyWebVerifiedExposure -Value $persisted -InstanceId ([string]$job.instanceId) -CompanyCode ([string]$job.companyCode) -AsOf ([string]$job.asOf))) {
                    throw 'Report result persistence verification failed.'
                }
                Save-ExecutiveJobJson -Path $paths.status -Value @{state='completed';jobId=$id;createdUtc=$created;finishedUtc=([datetime]::UtcNow.ToString('o'))}
                Write-Host "[EXEC-WORKER] Completed $section $id" -ForegroundColor Green
            } catch {
                Write-Warning "[EXEC-WORKER] Job $id failed: $($_.Exception.Message)"
                Save-ExecutiveJobJson -Path $paths.status -Value @{state='failed';jobId=$id;createdUtc=$created;finishedUtc=([datetime]::UtcNow.ToString('o'));error='Report calculation failed. Check worker logs.'}
            } finally {
                Remove-Item -LiteralPath $paths.processing -Force -ErrorAction SilentlyContinue
            }
        }
        Start-Sleep -Milliseconds ([Math]::Max(250,$PollMs))
    }
} finally { $exclusive.Dispose() }
