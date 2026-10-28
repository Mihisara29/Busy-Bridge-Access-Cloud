# BUSYWEB_WHATSAPP_PORTLESS_V1. Loaded only by Main API.
# Auth is performed by routes.ps1; direct callers must never bypass it.
$script:BusyWhatsAppRoot = Join-Path (Split-Path $PSScriptRoot -Parent) 'whatsapp'

function Read-BusyWhatsAppBoundedBody {
    param($Request, [int]$MaxBytes = 12000000)
    if ($Request.ContentLength64 -gt $MaxBytes) { throw 'Request too large' }
    $bytes=New-Object byte[] 16384
    $stream=New-Object System.IO.MemoryStream
    try {
        while (($n=$Request.InputStream.Read($bytes,0,$bytes.Length)) -gt 0) {
            if ($stream.Length + $n -gt $MaxBytes) {throw 'Request too large'}
            $stream.Write($bytes,0,$n)
        }
        if ($stream.Length -lt 1) {throw 'Request body required'}
        return [Text.Encoding]::UTF8.GetString($stream.ToArray())
    } finally {$stream.Dispose()}
}
function Initialize-BusyWhatsAppQueue {
    foreach ($folder in @('pending','processing','completed','failed','session')) {
        $path = Join-Path $script:BusyWhatsAppRoot $folder
        if (-not (Test-Path -LiteralPath $path)) {
            [void][System.IO.Directory]::CreateDirectory($path)
        }
    }
}
function Get-BusyWhatsAppStoredStatus {
    param([string]$JobId, [string]$InstanceId, [string]$CompanyCode, [string]$UserName)
    if ($JobId -notmatch '^[0-9a-f]{32}$') { return @{success=$false;error='Invalid job ID';httpStatus=400} }
    foreach ($folder in @('completed','failed','processing','pending')) {
        $file = Join-Path (Join-Path $script:BusyWhatsAppRoot $folder) ($JobId+'.json')
        if (Test-Path -LiteralPath $file) {
            $j = Get-Content -LiteralPath $file -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($j.instanceId -cne $InstanceId -or $j.companyCode -cne $CompanyCode -or $j.requestedBy -cne $UserName) {
                return @{success=$false;error='Job not found';httpStatus=404}
            }
            return @{success=$true;data=@{jobId=$JobId;status=[string]$j.status;error=[string]$j.error;updatedAt=[string]$j.updatedAt;messageId=[string]$j.messageId}}
        }
    }
    return @{success=$false;error='Job not found';httpStatus=404}
}

# WHATSAPP_LIVE_RECIPIENT_V1
# Direct, current (un-cached) party master read from the same fiscal database
# used by the voucher lookup. Never fall back to a potentially stale COM copy.
function Get-BusyWhatsAppLiveMaster {
    param(
        [string]$Name,
        [string]$InstanceId,
        [string]$CompanyCode
    )
    $resolver = Get-Command Get-BusyCloudFastConfigDbContext -ErrorAction SilentlyContinue
    if ($null -eq $resolver) {
        throw 'Live BUSY Account Master lookup is unavailable; WhatsApp sending is blocked.'
    }
    $ctx = $null
    $conn = $null
    $cmd = $null
    $reader = $null
    try {
        $ctx = Get-BusyCloudFastConfigDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        if ($null -eq $ctx -or $null -eq $ctx.connection) {
            throw 'Could not open current BUSY fiscal database; WhatsApp sending is blocked.'
        }
        $conn = $ctx.connection
        $dbType = [int]$ctx.dbType
        $placeholder = if ($dbType -eq 1) { '@partyName' } else { '?' }
        # No NOLOCK: a phone number must be a committed master-data value.
        $cmd = $conn.CreateCommand()
        $cmd.CommandTimeout = 10
        $cmd.CommandText = @"
SELECT M.Code AS AccountCode, M.Name AS AccountName, A.WhatsAppNo AS WhatsAppNo
FROM Master1 AS M
LEFT JOIN MasterAddressInfo AS A ON A.MasterCode = M.Code
WHERE M.MasterType = 2 AND M.Name = $placeholder
"@
        $p = $cmd.CreateParameter()
        $p.ParameterName = '@partyName'
        $p.DbType = [System.Data.DbType]::String
        $p.Value = $Name
        [void]$cmd.Parameters.Add($p)
        $reader = $cmd.ExecuteReader()
        $found = $null
        while ($reader.Read()) {
            if ($null -ne $found) {
                throw 'Ambiguous BUSY customer master records; WhatsApp sending is blocked.'
            }
            $nameValue = if ($reader.IsDBNull(1)) { '' } else { ([string]$reader.GetValue(1)).Trim() }
            $phoneValue = if ($reader.IsDBNull(2)) { '' } else { ([string]$reader.GetValue(2)).Trim() }
            $codeValue = if ($reader.IsDBNull(0)) { 0 } else { [int]$reader.GetValue(0) }
            $found = @{name=$nameValue; whatsappNo=$phoneValue; code=$codeValue}
        }
        if ($null -eq $found -or $found.code -le 0 -or
            -not $found.name.Equals($Name,[System.StringComparison]::OrdinalIgnoreCase)) {
            throw 'Customer was not found in the current BUSY fiscal database; WhatsApp sending is blocked.'
        }
        return $found
    } catch {
        # Fail closed: never reuse an old number if the current master cannot be read.
        throw ('Live BUSY WhatsApp number verification failed: ' + $_.Exception.Message)
    } finally {
        if ($null -ne $reader) {
            try { $reader.Close() } catch {}
            try { $reader.Dispose() } catch {}
        }
        if ($null -ne $cmd) { try { $cmd.Dispose() } catch {} }
        if ($null -ne $conn) {
            try { $conn.Close() } catch {}
            try { $conn.Dispose() } catch {}
        }
    }
}

function Resolve-BusyWhatsAppVoucher {
    param($Body, [string]$InstanceId, [string]$CompanyCode)
    $kind = 0
    if (-not [int]::TryParse([string]$Body.vchType, [ref]$kind) -or $kind -notin @(2,3,4,9,10,11,12,13,14,19)) {
        throw 'Unsupported voucher type'
    }
    $number = ([string]$Body.vchNo).Trim()
    $series = ([string]$Body.vchSeries).Trim(); if (-not $series) {$series='Main'}
    $date = ([string]$Body.vchDate).Trim()
    if (-not $number -or $number.Length -gt 120 -or -not $date -or $date.Length -gt 35 -or $series.Length -gt 120) {
        throw 'Valid posted voucher number, series and date required'
    }
    # BUSY is the source of truth: never trust browser-supplied party/recipient.
    # Account vouchers use an independent detail endpoint and must be posted in BUSY.
    if ($kind -in @(14,19)) {
        $found = Get-AccountVoucherDetail -VchType $kind -VchNo $number -VchSeries $series -VchDate $date -InstanceId $InstanceId -CompanyCode $CompanyCode
        if (-not $found.success -or $null -eq $found.data) {
            throw 'Posted account voucher not found (pending Web Approvals cannot be sent)'
        }
        $row = $found.data
        if (([string]$row.vchNo).Trim() -cne $number) { throw 'Account voucher identity mismatch' }
        $requiredDc = if ($kind -eq 14) { 'C' } else { 'D' }
        $candidateNames = @(
            @($row.accounts) | Where-Object {
                ([string]$_.dc).Trim().ToUpperInvariant() -eq $requiredDc -and
                -not [string]::IsNullOrWhiteSpace([string]$_.accountName) -and
                [double]$_.amount -gt 0
            } | ForEach-Object { ([string]$_.accountName).Trim() } | Select-Object -Unique
        )
        if ($candidateNames.Count -eq 0) {
            throw "No posted $requiredDc account found for this voucher; WhatsApp sending is blocked."
        }
        # No client-side account choice. If a double-entry voucher has more than
        # one distinct account on the intended side, never guess a recipient.
        if ($candidateNames.Count -ne 1) {
            throw "This voucher has $($candidateNames.Count) posted $requiredDc accounts. Automatic WhatsApp recipient is ambiguous; sending is blocked."
        }
        $party = [string]$candidateNames[0]
        # A browser may include the displayed account for a consistency check,
        # but it cannot change the server-selected posted counterparty.
        $requestedAccount = ([string]$Body.recipientAccountName).Trim()
        if ($requestedAccount -and
            -not $party.Equals($requestedAccount,[System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Requested account does not match the posted $requiredDc counterparty."
        }
    } else {
        $found = Get-VoucherDetail -VchType $kind -VchNo $number -VchSeries $series -VchDate $date -InstanceId $InstanceId -CompanyCode $CompanyCode
        if (-not $found.success -or $null -eq $found.data) { throw 'Posted voucher not found (pending approvals cannot be sent)' }
        $row = $found.data
        if (([string]$row.vchNo).Trim() -cne $number) { throw 'Voucher identity mismatch' }
        $party = ([string]$row.party).Trim()
        if (-not $party) { throw 'Voucher has no party' }
    }
    # Always read the latest committed customer number. Never reuse the
    # COM GetMasterXML response or the number previously shown in React.
    $account = Get-BusyWhatsAppLiveMaster -Name $party -InstanceId $InstanceId -CompanyCode $CompanyCode
    $raw = ([string]$account.whatsappNo).Trim()
    if (-not $raw) {throw 'No WhatsApp number is saved for this party in BUSY'}
    $digits = ($raw -replace '[^0-9]','')
    # Default country is Sri Lanka only when BUSY supplies a local 0XXXXXXXXX number.
    if ($digits -match '^0[0-9]{9}$') {$digits = '94' + $digits.Substring(1)}
    if ($digits -notmatch '^[1-9][0-9]{7,14}$') { throw 'Invalid WhatsApp number in BUSY master; correct it before sending' }
    return @{vchType=$kind;vchNo=$number;vchSeries=$series;vchDate=$date;party=$party;phone=$digits;
        recipientAccountName=$party}
}
function Get-BusyWhatsAppRecipient {
    param($Body,[string]$InstanceId,[string]$CompanyCode)
    $v=Resolve-BusyWhatsAppVoucher -Body $Body -InstanceId $InstanceId -CompanyCode $CompanyCode
    return @{success=$true;data=$v}
}
function Add-BusyWhatsAppInvoiceJob {
    param($Body,[string]$InstanceId,[string]$CompanyCode,[string]$UserName)
    Initialize-BusyWhatsAppQueue
    $v = Resolve-BusyWhatsAppVoucher -Body $Body -InstanceId $InstanceId -CompanyCode $CompanyCode
    if ($v.vchType -in @(14,19) -and -not $v.phone) {
        throw 'No valid WhatsApp recipient found for the posted account voucher.'
    }
    $confirm = ([string]$Body.confirmPhone -replace '[^0-9]','')
    if ($confirm -cne $v.phone) { throw 'Recipient has changed in BUSY. Refresh and confirm the latest number.' }
    if ($Body.consentConfirmed -ne $true) { throw 'Recipient opt-in confirmation required' }
    $format=([string]$Body.invoiceFormat).ToUpperInvariant()
    if ($format -notin @('A4','A5','TAX') -or ($format -eq 'TAX' -and $v.vchType -ne 9)) { throw 'Unsupported PDF invoice format' }
    $encoded=[string]$Body.pdfBase64
    if (-not $encoded -or $encoded.Length -gt 11500000) {throw 'PDF is missing or over 8 MB'}
    try { $bytes=[Convert]::FromBase64String($encoded) } catch {throw 'Invalid PDF encoding'}
    if ($bytes.Length -lt 100 -or $bytes.Length -gt 8MB) { throw 'PDF must be smaller than 8 MB' }
    if ([Text.Encoding]::ASCII.GetString($bytes,0,[Math]::Min(5,$bytes.Length)) -ne '%PDF-') { throw 'Uploaded file is not a PDF' }
    # Deduplicate accidental double-clicks and repeated requests for two minutes.
    $notBefore=[DateTime]::UtcNow.AddMinutes(-2)
    foreach($folder in @('pending','processing','completed')) {
        $dir=Join-Path $script:BusyWhatsAppRoot $folder
        $recent=@(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 150)
        foreach($file in $recent) {
            if ($file.LastWriteTimeUtc -lt $notBefore) { break }
            try {$old=Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8|ConvertFrom-Json} catch {continue}
            if ($old.instanceId -ceq $InstanceId -and $old.companyCode -ceq $CompanyCode -and
                $old.requestedBy -ceq $UserName -and $old.vchNo -ceq $v.vchNo -and
                $old.vchSeries -ceq $v.vchSeries -and [int]$old.vchType -eq $v.vchType -and
                $old.phone -ceq $v.phone -and $old.party -ceq $v.party -and $old.invoiceFormat -ceq $format -and
                $old.status -in @('queued','sending','submitted')) {
                return @{success=$true;data=@{jobId=[string]$old.jobId;status=[string]$old.status;
                    recipient=$v.phone;party=$v.party;duplicateSuppressed=$true}}
            }
        }
    }
    # The browser-generated PDF is untrusted content. Restrict sending to explicitly confirmed superadmins.
    $jobId=[Guid]::NewGuid().ToString('N')
    $pending=Join-Path $script:BusyWhatsAppRoot 'pending'
    $pdf=Join-Path $pending ($jobId+'.pdf')
    $json=Join-Path $pending ($jobId+'.json')
    $utc=[DateTime]::UtcNow.ToString('o')
    $job=@{jobId=$jobId;instanceId=$InstanceId;companyCode=$CompanyCode;requestedBy=$UserName;
        vchType=$v.vchType;vchNo=$v.vchNo;vchSeries=$v.vchSeries;vchDate=$v.vchDate;
        party=$v.party;phone=$v.phone;invoiceFormat=$format;status='queued';createdAt=$utc;updatedAt=$utc;
        error='';messageId='';attempts=0;pdfSha256=[BitConverter]::ToString(([Security.Cryptography.SHA256]::Create()).ComputeHash($bytes)).Replace('-','').ToLowerInvariant()}
    # Atomic publication: worker only sees a JSON file after PDF is safely written.
    $tmpPdf=$pdf+'.tmp'; $tmpJson=$json+'.tmp'
    try {
        [IO.File]::WriteAllBytes($tmpPdf,$bytes)
        Move-Item -LiteralPath $tmpPdf -Destination $pdf -ErrorAction Stop
        [IO.File]::WriteAllText($tmpJson,($job|ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
        Move-Item -LiteralPath $tmpJson -Destination $json -ErrorAction Stop
    } catch {
        foreach($p in @($tmpPdf,$tmpJson,$pdf,$json)) {Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue}
        throw
    }
    return @{success=$true;data=@{jobId=$jobId;status='queued';recipient=$v.phone;party=$v.party}}
}
function Get-BusyWhatsAppConnectionStatus {
    $file=Join-Path $script:BusyWhatsAppRoot 'connection-status.json'
    if(-not(Test-Path -LiteralPath $file)){return @{success=$true;data=@{ready=$false;state='not-started'}}}
    try {
        $s=Get-Content -LiteralPath $file -Raw -Encoding UTF8|ConvertFrom-Json
        $age=([DateTime]::UtcNow - ([DateTime]$s.updatedAt).ToUniversalTime()).TotalSeconds
        return @{success=$true;data=@{ready=([bool]$s.ready -and $age -lt 45);state=if($age -lt 45){[string]$s.state}else{'offline'};updatedAt=[string]$s.updatedAt}}
    } catch {return @{success=$true;data=@{ready=$false;state='offline'}}}
}
