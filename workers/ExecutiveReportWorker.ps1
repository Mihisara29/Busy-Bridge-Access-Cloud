# Separate dedicated process. Run with Windows PowerShell 5.1 (32-bit for BUSY COM).
param([string]$BridgeRoot=(Split-Path $PSScriptRoot -Parent),[int]$PollMs=1000)
$ErrorActionPreference='Stop'
$moduleDir=Join-Path $BridgeRoot 'modules'
$required=@('config.ps1','utils.ps1','connection.ps1','auth_native.ps1','items.ps1','accounts.ps1','salesman.ps1','vouchers.ps1','boms.ps1','vch_setting.ps1','web_approval.ps1','reports.ps1','account_ledger.ps1','dashboard.ps1','executive_insights.ps1','executive_jobs.ps1')
foreach($name in $required) {
    $file=Join-Path $moduleDir $name
    if(-not (Test-Path -LiteralPath $file)){throw "Missing worker module: $file"}
    . $file
}
Set-StrictMode -Off
$jobRoot=Join-Path $BridgeRoot 'report-jobs'
if(-not (Test-Path -LiteralPath $jobRoot)){throw 'Run installation setup for report-jobs first.'}
# Prevent concurrent worker processes from duplicating financial calculations.
$mutex = New-Object System.Threading.Mutex($false,'Local\BusyWebExecutiveExposureWorker')
if(-not $mutex.WaitOne(0)){throw 'Another executive report worker is already running in this Windows session.'}
try {
    # Heavy jobs take 5-9 minutes in the observed workload. Default them OFF.
    # When explicitly needed (prefer outside voucher-entry hours), launch this
    # worker with BUSYWEB_EXECUTIVE_ENABLE_EXPOSURE=1.
    $allowHeavyExposure=([string]$env:BUSYWEB_EXECUTIVE_ENABLE_EXPOSURE -eq '1')
    Write-Host ('[EXEC-WORKER] Ready. Heavy Exposure enabled: '+$allowHeavyExposure+'. Main API unaffected.') -ForegroundColor Green
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
            $id=$file.Name.Substring(0,64)
            if($id -notmatch '^[a-f0-9]{64}$'){continue}
            $paths=Get-ExecutiveJobPaths -Id $id
            $queuedJob=Read-ExecutiveJobJson -Path $paths.request
            if($null -eq $queuedJob){ continue }
            if (-not $allowHeavyExposure -and [string]$queuedJob.section -eq 'exposure') {
                # Leave on disk for explicitly scheduled future processing.
                # Never let browser traffic start a minutes-long stock valuation.
                continue
            }
            $identity=([string]$queuedJob.instanceId+'|'+[string]$queuedJob.companyCode+'|'+[string]$queuedJob.asOf+'|'+[string]$queuedJob.section).ToLowerInvariant()
            $sha=[System.Security.Cryptography.SHA256]::Create()
            try {$logical=([BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($identity)))).Replace('-','').ToLowerInvariant()}
            finally {$sha.Dispose()}
            $active=Read-ExecutiveJobJson -Path (Get-ExecutiveJobPointerPath -LogicalKey $logical)
            if($null -ne $active -and [string]$active.jobId -ne $id) {
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
                Save-ExecutiveJobJson -Path $paths.result -Value $data
                # Confirm the result can be read before declaring the job complete.
                $persisted = Read-ExecutiveJobJson -Path $paths.result
                if ($null -eq $persisted -or $persisted.success -ne $true -or [string]$persisted.section -ne $section) {
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
} finally { $mutex.ReleaseMutex();$mutex.Dispose() }
