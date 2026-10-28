# BUSYWEB_AUTOMATED_EXPOSURE_PILOT_V1
# Scheduler is COM/SQL-free: it only enqueues dated, deduplicated Exposure jobs.
param([string]$BridgeRoot=(Split-Path $PSScriptRoot -Parent),[int]$PollSeconds=60)
$ErrorActionPreference='Stop'
try { $Host.UI.RawUI.WindowTitle = 'BusyWeb - Exposure Scheduler' } catch {}
$modules=Join-Path $BridgeRoot 'modules'
. (Join-Path $modules 'executive_jobs.ps1')
. (Join-Path $modules 'executive_automation.ps1')
$jobRoot=Join-Path $BridgeRoot 'report-jobs'
if(-not (Test-Path -LiteralPath $jobRoot -PathType Container)){throw 'report-jobs missing'}
if($null -eq (Get-BusyWebAutomationSettings)){throw 'Automation pilot disabled in configuration.'}
$lockPath=Join-Path $jobRoot 'automated-exposure-scheduler.lock'
try{$exclusive=[System.IO.File]::Open($lockPath,[System.IO.FileMode]::OpenOrCreate,[System.IO.FileAccess]::ReadWrite,[System.IO.FileShare]::None)}
catch{throw 'Another automatic Exposure scheduler holds the cross-session lock.'}
try {
    Write-Host '[AUTO-SCHEDULER] Running. SQL/COM calculations never execute here.' -ForegroundColor Green
    $lastMessage=''
    while($true) {
        try {
            $settings=Get-BusyWebAutomationSettings
            if($null -eq $settings) {
                Write-Warning '[AUTO-SCHEDULER] Pilot disabled. Exiting; no queue changes.'
                break
            }
            $asOf=(Get-BusyWebCompanyToday).ToString('yyyy-MM-dd')
            # BUSYWEB-ONDEMAND-SCHEDULER-V1
            $targets = @($settings.companies | Where-Object { $null -ne $_ })

            if ([string]$settings.mode -eq 'registered-on-demand') {
                $targets = @()
                $registryPath = Join-Path $BridgeRoot 'instances.json'
                $registry = Get-Content -LiteralPath $registryPath -Raw -Encoding UTF8 | ConvertFrom-Json

                if ($registry -is [array]) {
                    $registeredInstances = @($registry)
                }
                elseif ($null -ne $registry.instances) {
                    $registeredInstances = @($registry.instances)
                }
                else {
                    throw 'Invalid instances.json structure.'
                }

                $seen = @{}

                foreach ($registeredInstance in $registeredInstances) {
                    $instanceId = [string]$registeredInstance.id
                    if ($instanceId -notmatch '^[A-Za-z0-9_-]{1,80}$') {
                        continue
                    }

                    foreach ($registeredCompany in @($registeredInstance.companies)) {
                        $companyCode = [string]$registeredCompany.code
                        if ($companyCode -notmatch '^[A-Za-z0-9_-]{1,80}$') {
                            continue
                        }

                        $identity = (
                            $instanceId + '|' + $companyCode + '|' +
                            $asOf + '|exposure'
                        ).ToLowerInvariant()

                        $sha = [System.Security.Cryptography.SHA256]::Create()
                        try {
                            $key = [BitConverter]::ToString(
                                $sha.ComputeHash(
                                    [System.Text.Encoding]::UTF8.GetBytes($identity)
                                )
                            ).Replace('-', '').ToLowerInvariant()
                        }
                        finally {
                            $sha.Dispose()
                        }

                        if ($seen.ContainsKey($key)) { continue }

                        $pointerPath = Get-ExecutiveJobPointerPath -LogicalKey $key
                        if (-not (Test-Path -LiteralPath $pointerPath -PathType Leaf)) {
                            continue
                        }

                        $pointer = Read-ExecutiveJobJson -Path $pointerPath
                        if ($null -eq $pointer -or
                            [string]$pointer.section -ne 'exposure' -or
                            [string]$pointer.jobId -notmatch '^[a-f0-9]{64}$') {
                            continue
                        }

                        $seen[$key] = $true
                        $targets += [pscustomobject]@{
                            instanceId = $instanceId
                            companyCode = $companyCode
                        }
                    }
                }
            }

            if ($targets.Count -eq 0) {
                $message = '[AUTO-SCHEDULER] No existing Exposure targets for ' + $asOf
                if ($message -ne $lastMessage) {
                    Write-Host $message
                    $lastMessage = $message
                }
            }

            foreach($company in $targets) {
                $instance=[string]$company.instanceId
                $code=[string]$company.companyCode
                # RefreshToday queues a new generation only if stale, within capacity.
                $r=Get-BusyWebAutomatedExposure -InstanceId $instance -CompanyCode $code -AsOf $asOf -RefreshToday
                $message=('[AUTO-SCHEDULER] '+$asOf+' '+$instance+'/'+$code+' HTTP '+[string]$r.httpStatus)
                if([int]$r.httpStatus -eq 202) {$message+=' '+[string]$r.data.status}
                if([int]$r.httpStatus -ge 400) {$message+=' '+[string]$r.error}
                if($message -ne $lastMessage) {Write-Host $message; $lastMessage=$message}
            }
        } catch {Write-Warning ('[AUTO-SCHEDULER] Safe iteration failure: '+$_.Exception.Message)}
        Start-Sleep -Seconds ([Math]::Max(30,[Math]::Min(300,$PollSeconds)))
    }
} finally {$exclusive.Dispose()}
