# modules/offline_voucher_sync.ps1
# -----------------------------------------------------------------------------
# BusyCloud Local Voucher Synchronization (NEW implementation)
#
# This module is intentionally independent from the old offline_sync_*.ps1
# implementation. It provides:
#   - durable idempotency in the selected BUSY company database
#   - SQL Server + Access support
#   - safe retry / crash-window handling
#   - current numbering lookup immediately before posting
#   - reuse of the existing Create-Voucher pipeline
#
# Current scope: CREATE Sale vouchers only (BUSY VchType = 9).
# PowerShell 5.1 compatible.
# -----------------------------------------------------------------------------

Set-StrictMode -Version Latest

$script:OfflineSyncLedgerTable = 'BusyCloudOfflineVoucherSync'
$script:OfflineSyncStaleMinutes = 10

# -----------------------------------------------------------------------------
# Generic helpers
# -----------------------------------------------------------------------------

function Get-OfflineSyncPropertyValue {
    param(
        $Object,
        [Parameter(Mandatory = $true)][string[]]$Names,
        $DefaultValue = $null
    )

    if ($null -eq $Object) {
        return $DefaultValue
    }

    foreach ($name in $Names) {
        try {
            if ($Object -is [System.Collections.IDictionary]) {
                if ($Object.Contains($name)) {
                    return $Object[$name]
                }
            }

            $property = $Object.PSObject.Properties[$name]
            if ($null -ne $property) {
                return $property.Value
            }
        }
        catch {
        }
    }

    return $DefaultValue
}

function ConvertTo-OfflineSyncBool {
    param($Value)

    if ($Value -is [bool]) {
        return [bool]$Value
    }

    if ($null -eq $Value) {
        return $false
    }

    $text = ([string]$Value).Trim().ToLowerInvariant()
    return ($text -in @('1', 'true', 'yes', 'y'))
}

function Test-OfflineSyncDateText {
    param(
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    $parsed = [datetime]::MinValue
    $formats = @(
        'dd-MM-yyyy',
        'yyyy-MM-dd',
        'dd/MM/yyyy',
        'yyyy/MM/dd'
    )

    foreach ($format in $formats) {
        if (
            [datetime]::TryParseExact(
                $Value.Trim(),
                $format,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None,
                [ref]$parsed
            )
        ) {
            return $true
        }
    }

    return [datetime]::TryParse(
        $Value.Trim(),
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::AllowWhiteSpaces,
        [ref]$parsed
    )
}


function ConvertTo-OfflineSyncNumberingDate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $clean = $Value.Trim()

    if ([string]::IsNullOrWhiteSpace($clean)) {
        throw 'Voucher date is required.'
    }

    $parsed = [datetime]::MinValue

    $formats = @(
        'dd-MM-yyyy',
        'yyyy-MM-dd',
        'dd/MM/yyyy',
        'yyyy/MM/dd'
    )

    foreach ($format in $formats) {
        if (
            [datetime]::TryParseExact(
                $clean,
                $format,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None,
                [ref]$parsed
            )
        ) {
            # The Web/BUSY numbering layer is safest with the ISO form.
            return $parsed.ToString('yyyy-MM-dd')
        }
    }

    if (
        [datetime]::TryParse(
            $clean,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::AllowWhiteSpaces,
            [ref]$parsed
        )
    ) {
        return $parsed.ToString('yyyy-MM-dd')
    }

    throw "Voucher date '$Value' is invalid."
}

function Copy-OfflineSyncObject {
    param($Object)

    if ($null -eq $Object) {
        return $null
    }

    return (
        $Object |
        ConvertTo-Json -Depth 100 -Compress |
        ConvertFrom-Json
    )
}

function Get-OfflineSyncRequestHash {
    param(
        [Parameter(Mandatory = $true)]$Data
    )

    $hashSource = [ordered]@{
        schemaVersion      = Get-OfflineSyncPropertyValue -Object $Data -Names @('schemaVersion') -DefaultValue 1
        localId            = ([string](Get-OfflineSyncPropertyValue -Object $Data -Names @('localId') -DefaultValue '')).Trim()
        localReference     = ([string](Get-OfflineSyncPropertyValue -Object $Data -Names @('localReference') -DefaultValue '')).Trim()
        vchType            = [int](Get-OfflineSyncPropertyValue -Object $Data -Names @('vchType') -DefaultValue 0)
        requestedVchNo     = ([string](Get-OfflineSyncPropertyValue -Object $Data -Names @('requestedVchNo') -DefaultValue '')).Trim()
        requestedVchSeries = ([string](Get-OfflineSyncPropertyValue -Object $Data -Names @('requestedVchSeries') -DefaultValue '')).Trim()
        voucherDate        = ([string](Get-OfflineSyncPropertyValue -Object $Data -Names @('voucherDate') -DefaultValue '')).Trim()
        payload            = Get-OfflineSyncPropertyValue -Object $Data -Names @('payload') -DefaultValue $null
    }

    $json = $hashSource | ConvertTo-Json -Depth 100 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $sha = [System.Security.Cryptography.SHA256]::Create()

    try {
        $hashBytes = $sha.ComputeHash($bytes)
        return -join ($hashBytes | ForEach-Object { $_.ToString('x2') })
    }
    finally {
        $sha.Dispose()
    }
}

function New-OfflineSyncDbCommand {
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][bool]$IsSqlServer,
        [Parameter(Mandatory = $true)][string]$Sql,
        [hashtable]$Parameters = @{}
    )

    $command = $Connection.CreateCommand()

    if ($IsSqlServer) {
        $command.CommandText = $Sql

        $parameterNames = @(
            [regex]::Matches($Sql, '@([A-Za-z_][A-Za-z0-9_]*)') |
            ForEach-Object { $_.Groups[1].Value } |
            Select-Object -Unique
        )

        foreach ($name in $parameterNames) {
            $parameter = $command.CreateParameter()
            $parameter.ParameterName = "@$name"

            $value = $null
            if ($Parameters.ContainsKey($name)) {
                $value = $Parameters[$name]
            }

            $parameter.Value = if ($null -eq $value) {
                [DBNull]::Value
            }
            else {
                $value
            }

            [void]$command.Parameters.Add($parameter)
        }
    }
    else {
        # OleDb parameters are positional even when SQL text uses @names.
        # Replace each occurrence with '?' and add the matching value in exactly
        # the same order in which it appears in the SQL statement.
        $matches = [regex]::Matches($Sql, '@([A-Za-z_][A-Za-z0-9_]*)')
        $command.CommandText = [regex]::Replace(
            $Sql,
            '@([A-Za-z_][A-Za-z0-9_]*)',
            '?'
        )

        foreach ($match in $matches) {
            $name = $match.Groups[1].Value
            $parameter = $command.CreateParameter()

            $value = $null
            if ($Parameters.ContainsKey($name)) {
                $value = $Parameters[$name]
            }

            $parameter.Value = if ($null -eq $value) {
                [DBNull]::Value
            }
            else {
                $value
            }

            [void]$command.Parameters.Add($parameter)
        }
    }

    return $command
}

function Invoke-OfflineSyncNonQuery {
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][bool]$IsSqlServer,
        [Parameter(Mandatory = $true)][string]$Sql,
        [hashtable]$Parameters = @{}
    )

    $command = New-OfflineSyncDbCommand `
        -Connection $Connection `
        -IsSqlServer $IsSqlServer `
        -Sql $Sql `
        -Parameters $Parameters

    try {
        return [int]$command.ExecuteNonQuery()
    }
    finally {
        $command.Dispose()
    }
}

function Invoke-OfflineSyncQuerySingle {
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][bool]$IsSqlServer,
        [Parameter(Mandatory = $true)][string]$Sql,
        [hashtable]$Parameters = @{}
    )

    $command = New-OfflineSyncDbCommand `
        -Connection $Connection `
        -IsSqlServer $IsSqlServer `
        -Sql $Sql `
        -Parameters $Parameters

    $reader = $null

    try {
        $reader = $command.ExecuteReader()

        if (-not $reader.Read()) {
            return $null
        }

        $row = [ordered]@{}

        for ($i = 0; $i -lt $reader.FieldCount; $i++) {
            $name = $reader.GetName($i)
            $row[$name] = if ($reader.IsDBNull($i)) {
                $null
            }
            else {
                $reader.GetValue($i)
            }
        }

        return [pscustomobject]$row
    }
    finally {
        if ($null -ne $reader) {
            try { $reader.Close() } catch {}
            try { $reader.Dispose() } catch {}
        }

        $command.Dispose()
    }
}

# -----------------------------------------------------------------------------
# Ledger storage
# -----------------------------------------------------------------------------

function Test-OfflineSyncAccessTableExists {
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][string]$TableName
    )

    try {
        $schema = $Connection.GetSchema('Tables')

        foreach ($row in $schema.Rows) {
            $name = ''
            try { $name = ([string]$row['TABLE_NAME']).Trim() } catch {}

            if ($name -ieq $TableName) {
                return $true
            }
        }
    }
    catch {
    }

    return $false
}

function Ensure-OfflineSyncLedger {
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][bool]$IsSqlServer
    )

    $table = $script:OfflineSyncLedgerTable

    if ($IsSqlServer) {
        $sql = @"
IF OBJECT_ID('dbo.$table', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.$table (
        [Id] BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [LocalId] VARCHAR(100) NOT NULL,
        [InstanceId] VARCHAR(100) NOT NULL,
        [CompanyCode] VARCHAR(50) NOT NULL,
        [UserName] NVARCHAR(150) NULL,
        [VchType] INT NOT NULL,
        [VchSeries] NVARCHAR(150) NULL,
        [RequestedVchNo] NVARCHAR(150) NULL,
        [PostingVchNo] NVARCHAR(150) NULL,
        [FinalVchNo] NVARCHAR(150) NULL,
        [FinalVchCode] INT NULL,
        [Status] VARCHAR(30) NOT NULL,
        [Phase] VARCHAR(30) NOT NULL,
        [ProcessingToken] VARCHAR(64) NULL,
        [RequestHash] VARCHAR(64) NOT NULL,
        [RequestPayload] NVARCHAR(MAX) NULL,
        [ErrorMessage] NVARCHAR(MAX) NULL,
        [CanRetry] INT NOT NULL CONSTRAINT [DF_${table}_CanRetry] DEFAULT(1),
        [AttemptCount] INT NOT NULL CONSTRAINT [DF_${table}_AttemptCount] DEFAULT(0),
        [CreatedAt] DATETIME2 NOT NULL,
        [UpdatedAt] DATETIME2 NOT NULL,
        [SyncedAt] DATETIME2 NULL,
        CONSTRAINT [UQ_${table}_Scope] UNIQUE ([LocalId], [InstanceId], [CompanyCode])
    );
END
"@

        [void](Invoke-OfflineSyncNonQuery `
            -Connection $Connection `
            -IsSqlServer $true `
            -Sql $sql)

        return
    }

    if (Test-OfflineSyncAccessTableExists -Connection $Connection -TableName $table) {
        return
    }

    $createSql = @"
CREATE TABLE [$table] (
    [Id] AUTOINCREMENT CONSTRAINT [PK_${table}] PRIMARY KEY,
    [LocalId] TEXT(100) NOT NULL,
    [InstanceId] TEXT(100) NOT NULL,
    [CompanyCode] TEXT(50) NOT NULL,
    [UserName] TEXT(150),
    [VchType] INTEGER NOT NULL,
    [VchSeries] TEXT(150),
    [RequestedVchNo] TEXT(150),
    [PostingVchNo] TEXT(150),
    [FinalVchNo] TEXT(150),
    [FinalVchCode] LONG,
    [Status] TEXT(30) NOT NULL,
    [Phase] TEXT(30) NOT NULL,
    [ProcessingToken] TEXT(64),
    [RequestHash] TEXT(64) NOT NULL,
    [RequestPayload] MEMO,
    [ErrorMessage] MEMO,
    [CanRetry] BYTE,
    [AttemptCount] LONG,
    [CreatedAt] DATETIME,
    [UpdatedAt] DATETIME,
    [SyncedAt] DATETIME
)
"@

    try {
        [void](Invoke-OfflineSyncNonQuery `
            -Connection $Connection `
            -IsSqlServer $false `
            -Sql $createSql)
    }
    catch {
        # Another request/process may have created it after our schema check.
        if (-not (Test-OfflineSyncAccessTableExists -Connection $Connection -TableName $table)) {
            throw
        }
    }

    $indexSql = @"
CREATE UNIQUE INDEX [UQ_${table}_Scope]
ON [$table] ([LocalId], [InstanceId], [CompanyCode])
"@

    try {
        [void](Invoke-OfflineSyncNonQuery `
            -Connection $Connection `
            -IsSqlServer $false `
            -Sql $indexSql)
    }
    catch {
        # If the index already exists because another process won the race,
        # the table is still usable.
    }
}

function Get-OfflineSyncLedgerRecord {
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][bool]$IsSqlServer,
        [Parameter(Mandatory = $true)][string]$LocalId,
        [Parameter(Mandatory = $true)][string]$InstanceId,
        [Parameter(Mandatory = $true)][string]$CompanyCode
    )

    $table = $script:OfflineSyncLedgerTable

    $sql = @"
SELECT TOP 1 *
FROM [$table]
WHERE [LocalId] = @LocalId
  AND [InstanceId] = @InstanceId
  AND [CompanyCode] = @CompanyCode
"@

    return Invoke-OfflineSyncQuerySingle `
        -Connection $Connection `
        -IsSqlServer $IsSqlServer `
        -Sql $sql `
        -Parameters @{
            LocalId = $LocalId
            InstanceId = $InstanceId
            CompanyCode = $CompanyCode
        }
}

function Add-OfflineSyncLedgerRecord {
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][bool]$IsSqlServer,
        [Parameter(Mandatory = $true)]$Data,
        [Parameter(Mandatory = $true)][string]$InstanceId,
        [Parameter(Mandatory = $true)][string]$CompanyCode,
        [Parameter(Mandatory = $true)][string]$UserName,
        [Parameter(Mandatory = $true)][string]$RequestHash,
        [Parameter(Mandatory = $true)][string]$ProcessingToken
    )

    $table = $script:OfflineSyncLedgerTable
    $now = Get-Date
    $payloadJson = $Data | ConvertTo-Json -Depth 100 -Compress

    $vchType = [int](Get-OfflineSyncPropertyValue -Object $Data -Names @('vchType') -DefaultValue 0)
    $series = ([string](Get-OfflineSyncPropertyValue -Object $Data -Names @('requestedVchSeries') -DefaultValue '')).Trim()
    $requestedNo = ([string](Get-OfflineSyncPropertyValue -Object $Data -Names @('requestedVchNo') -DefaultValue '')).Trim()

    $sql = @"
INSERT INTO [$table] (
    [LocalId], [InstanceId], [CompanyCode], [UserName],
    [VchType], [VchSeries], [RequestedVchNo],
    [Status], [Phase], [ProcessingToken], [RequestHash],
    [RequestPayload], [ErrorMessage], [CanRetry], [AttemptCount],
    [CreatedAt], [UpdatedAt]
)
VALUES (
    @LocalId, @InstanceId, @CompanyCode, @UserName,
    @VchType, @VchSeries, @RequestedVchNo,
    'PROCESSING', 'VALIDATING', @ProcessingToken, @RequestHash,
    @RequestPayload, NULL, 1, 1,
    @CreatedAt, @UpdatedAt
)
"@

    return Invoke-OfflineSyncNonQuery `
        -Connection $Connection `
        -IsSqlServer $IsSqlServer `
        -Sql $sql `
        -Parameters @{
            LocalId = ([string]$Data.localId).Trim()
            InstanceId = $InstanceId
            CompanyCode = $CompanyCode
            UserName = $UserName
            VchType = $vchType
            VchSeries = $series
            RequestedVchNo = $requestedNo
            ProcessingToken = $ProcessingToken
            RequestHash = $RequestHash
            RequestPayload = $payloadJson
            CreatedAt = $now
            UpdatedAt = $now
        }
}

function Set-OfflineSyncLedgerPhase {
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][bool]$IsSqlServer,
        [Parameter(Mandatory = $true)]$Data,
        [Parameter(Mandatory = $true)][string]$InstanceId,
        [Parameter(Mandatory = $true)][string]$CompanyCode,
        [Parameter(Mandatory = $true)][string]$ProcessingToken,
        [Parameter(Mandatory = $true)][string]$Phase,
        [string]$PostingVchNo = ''
    )

    $table = $script:OfflineSyncLedgerTable
    $now = Get-Date

    $postingValue = if ([string]::IsNullOrWhiteSpace($PostingVchNo)) {
        $null
    }
    else {
        $PostingVchNo.Trim()
    }

    $parameters = @{
        Phase = $Phase
        UpdatedAt = $now
        LocalId = ([string]$Data.localId).Trim()
        InstanceId = $InstanceId
        CompanyCode = $CompanyCode
        ProcessingToken = $ProcessingToken
    }

    if ($null -eq $postingValue) {
        $sql = @"
UPDATE [$table]
SET [Phase] = @Phase,
    [UpdatedAt] = @UpdatedAt
WHERE [LocalId] = @LocalId
  AND [InstanceId] = @InstanceId
  AND [CompanyCode] = @CompanyCode
  AND [Status] = 'PROCESSING'
  AND [ProcessingToken] = @ProcessingToken
"@
    }
    else {
        $parameters['PostingVchNo'] = $postingValue
        $sql = @"
UPDATE [$table]
SET [Phase] = @Phase,
    [PostingVchNo] = @PostingVchNo,
    [UpdatedAt] = @UpdatedAt
WHERE [LocalId] = @LocalId
  AND [InstanceId] = @InstanceId
  AND [CompanyCode] = @CompanyCode
  AND [Status] = 'PROCESSING'
  AND [ProcessingToken] = @ProcessingToken
"@
    }

    return Invoke-OfflineSyncNonQuery `
        -Connection $Connection `
        -IsSqlServer $IsSqlServer `
        -Sql $sql `
        -Parameters $parameters
}

function Set-OfflineSyncLedgerFailed {
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][bool]$IsSqlServer,
        [Parameter(Mandatory = $true)]$Data,
        [Parameter(Mandatory = $true)][string]$InstanceId,
        [Parameter(Mandatory = $true)][string]$CompanyCode,
        [Parameter(Mandatory = $true)][string]$ProcessingToken,
        [Parameter(Mandatory = $true)][string]$ErrorMessage,
        [bool]$CanRetry = $true
    )

    $table = $script:OfflineSyncLedgerTable

    $sql = @"
UPDATE [$table]
SET [Status] = 'FAILED',
    [Phase] = 'FAILED',
    [ErrorMessage] = @ErrorMessage,
    [CanRetry] = @CanRetry,
    [ProcessingToken] = NULL,
    [UpdatedAt] = @UpdatedAt
WHERE [LocalId] = @LocalId
  AND [InstanceId] = @InstanceId
  AND [CompanyCode] = @CompanyCode
  AND [ProcessingToken] = @ProcessingToken
"@

    return Invoke-OfflineSyncNonQuery `
        -Connection $Connection `
        -IsSqlServer $IsSqlServer `
        -Sql $sql `
        -Parameters @{
            ErrorMessage = $ErrorMessage
            CanRetry = if ($CanRetry) { 1 } else { 0 }
            UpdatedAt = Get-Date
            LocalId = ([string]$Data.localId).Trim()
            InstanceId = $InstanceId
            CompanyCode = $CompanyCode
            ProcessingToken = $ProcessingToken
        }
}

function Set-OfflineSyncLedgerReviewRequired {
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][bool]$IsSqlServer,
        [Parameter(Mandatory = $true)]$Data,
        [Parameter(Mandatory = $true)][string]$InstanceId,
        [Parameter(Mandatory = $true)][string]$CompanyCode,
        [string]$ProcessingToken = '',
        [Parameter(Mandatory = $true)][string]$ErrorMessage
    )

    $table = $script:OfflineSyncLedgerTable

    $whereToken = ''
    $parameters = @{
        ErrorMessage = $ErrorMessage
        UpdatedAt = Get-Date
        LocalId = ([string]$Data.localId).Trim()
        InstanceId = $InstanceId
        CompanyCode = $CompanyCode
    }

    if (-not [string]::IsNullOrWhiteSpace($ProcessingToken)) {
        $whereToken = ' AND [ProcessingToken] = @ProcessingToken'
        $parameters['ProcessingToken'] = $ProcessingToken
    }

    $sql = @"
UPDATE [$table]
SET [Status] = 'REVIEW_REQUIRED',
    [Phase] = 'REVIEW_REQUIRED',
    [ErrorMessage] = @ErrorMessage,
    [CanRetry] = 0,
    [ProcessingToken] = NULL,
    [UpdatedAt] = @UpdatedAt
WHERE [LocalId] = @LocalId
  AND [InstanceId] = @InstanceId
  AND [CompanyCode] = @CompanyCode$($whereToken)
"@

    return Invoke-OfflineSyncNonQuery `
        -Connection $Connection `
        -IsSqlServer $IsSqlServer `
        -Sql $sql `
        -Parameters $parameters
}

function Set-OfflineSyncLedgerSynced {
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][bool]$IsSqlServer,
        [Parameter(Mandatory = $true)]$Data,
        [Parameter(Mandatory = $true)][string]$InstanceId,
        [Parameter(Mandatory = $true)][string]$CompanyCode,
        [string]$ProcessingToken = '',
        [Parameter(Mandatory = $true)][string]$FinalVchNo,
        [int]$FinalVchCode = 0,
        [string]$WarningMessage = ''
    )

    $table = $script:OfflineSyncLedgerTable
    $now = Get-Date

    $whereToken = ''
    $parameters = @{
        FinalVchNo = $FinalVchNo
        FinalVchCode = if ($FinalVchCode -gt 0) { $FinalVchCode } else { $null }
        ErrorMessage = if ([string]::IsNullOrWhiteSpace($WarningMessage)) { $null } else { $WarningMessage }
        UpdatedAt = $now
        SyncedAt = $now
        LocalId = ([string]$Data.localId).Trim()
        InstanceId = $InstanceId
        CompanyCode = $CompanyCode
    }

    if (-not [string]::IsNullOrWhiteSpace($ProcessingToken)) {
        $whereToken = ' AND [ProcessingToken] = @ProcessingToken'
        $parameters['ProcessingToken'] = $ProcessingToken
    }

    $sql = @"
UPDATE [$table]
SET [Status] = 'SYNCED',
    [Phase] = 'COMPLETED',
    [FinalVchNo] = @FinalVchNo,
    [FinalVchCode] = @FinalVchCode,
    [ErrorMessage] = @ErrorMessage,
    [CanRetry] = 0,
    [ProcessingToken] = NULL,
    [UpdatedAt] = @UpdatedAt,
    [SyncedAt] = @SyncedAt
WHERE [LocalId] = @LocalId
  AND [InstanceId] = @InstanceId
  AND [CompanyCode] = @CompanyCode$($whereToken)
"@

    return Invoke-OfflineSyncNonQuery `
        -Connection $Connection `
        -IsSqlServer $IsSqlServer `
        -Sql $sql `
        -Parameters $parameters
}

function Claim-OfflineSyncExistingRecord {
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][bool]$IsSqlServer,
        [Parameter(Mandatory = $true)]$Record,
        [Parameter(Mandatory = $true)]$Data,
        [Parameter(Mandatory = $true)][string]$InstanceId,
        [Parameter(Mandatory = $true)][string]$CompanyCode,
        [Parameter(Mandatory = $true)][string]$RequestHash,
        [Parameter(Mandatory = $true)][string]$ProcessingToken
    )

    $status = ([string]$Record.Status).Trim().ToUpperInvariant()
    $phase = ([string]$Record.Phase).Trim().ToUpperInvariant()
    $storedHash = ([string]$Record.RequestHash).Trim().ToLowerInvariant()

    if ($storedHash -ne $RequestHash.ToLowerInvariant()) {
        return @{
            success = $false
            conflict = $true
            httpStatus = 409
            errorCode = 'IDEMPOTENCY_PAYLOAD_MISMATCH'
            error = 'This localId was already used for different voucher data. A localId must never be reused for another transaction.'
        }
    }

    if ($status -eq 'SYNCED') {
        $syncedAt = $null
        if ($null -ne $Record.SyncedAt) {
            try { $syncedAt = ([datetime]$Record.SyncedAt).ToString('o') } catch {}
        }

        return @{
            success = $true
            alreadyProcessed = $true
            localId = ([string]$Data.localId).Trim()
            finalVchNo = [string]$Record.FinalVchNo
            finalVchCode = if ($null -ne $Record.FinalVchCode) { [int]$Record.FinalVchCode } else { $null }
            syncedAt = $syncedAt
        }
    }

    if ($status -eq 'REVIEW_REQUIRED') {
        return @{
            success = $false
            conflict = $true
            reviewRequired = $true
            httpStatus = 409
            errorCode = 'OFFLINE_SYNC_REVIEW_REQUIRED'
            error = if (-not [string]::IsNullOrWhiteSpace([string]$Record.ErrorMessage)) {
                [string]$Record.ErrorMessage
            }
            else {
                'This transaction needs reconciliation before another create attempt can be made.'
            }
        }
    }

    if ($status -eq 'PROCESSING') {
        $updatedAt = [datetime]::MinValue
        try { $updatedAt = [datetime]$Record.UpdatedAt } catch {}

        $isFresh = (
            $updatedAt -ne [datetime]::MinValue -and
            $updatedAt -gt (Get-Date).AddMinutes(-1 * $script:OfflineSyncStaleMinutes)
        )

        if ($isFresh) {
            return @{
                success = $false
                conflict = $true
                httpStatus = 409
                errorCode = 'OFFLINE_SYNC_IN_PROGRESS'
                error = 'This local voucher is already being synchronized.'
            }
        }

        # We deliberately never blindly retry after entering the BUSY posting
        # phase. The previous process could have created the voucher and crashed
        # before updating the ledger.
        if ($phase -in @('POSTING_TO_BUSY', 'COMPLETING')) {
            return @{
                success = $false
                conflict = $true
                stalePosting = $true
                httpStatus = 409
                errorCode = 'OFFLINE_SYNC_STALE_POSTING'
                error = 'A previous synchronization stopped during BUSY posting. The server must reconcile the possible BUSY voucher before another create attempt.'
            }
        }

        $table = $script:OfflineSyncLedgerTable
        $oldToken = ([string]$Record.ProcessingToken).Trim()

        if ([string]::IsNullOrWhiteSpace($oldToken)) {
            return @{
                success = $false
                conflict = $true
                reviewRequired = $true
                httpStatus = 409
                errorCode = 'OFFLINE_SYNC_REVIEW_REQUIRED'
                error = 'A stale PROCESSING record has no ownership token. Automatic retry is blocked for safety.'
            }
        }

        $rows = Invoke-OfflineSyncNonQuery `
            -Connection $Connection `
            -IsSqlServer $IsSqlServer `
            -Sql @"
UPDATE [$table]
SET [Status] = 'PROCESSING',
    [Phase] = 'VALIDATING',
    [ProcessingToken] = @ProcessingToken,
    [ErrorMessage] = NULL,
    [CanRetry] = 1,
    [AttemptCount] = [AttemptCount] + 1,
    [UpdatedAt] = @UpdatedAt
WHERE [LocalId] = @LocalId
  AND [InstanceId] = @InstanceId
  AND [CompanyCode] = @CompanyCode
  AND [Status] = 'PROCESSING'
  AND [ProcessingToken] = @OldProcessingToken
"@ `
            -Parameters @{
                ProcessingToken = $ProcessingToken
                UpdatedAt = Get-Date
                LocalId = ([string]$Data.localId).Trim()
                InstanceId = $InstanceId
                CompanyCode = $CompanyCode
                OldProcessingToken = $oldToken
            }

        if ($rows -eq 1) {
            return @{ success = $true; claimed = $true }
        }

        return @{
            success = $false
            conflict = $true
            httpStatus = 409
            errorCode = 'OFFLINE_SYNC_IN_PROGRESS'
            error = 'Another synchronization process claimed this local voucher.'
        }
    }

    if ($status -eq 'FAILED') {
        $canRetry = 0
        try { $canRetry = [int]$Record.CanRetry } catch {}

        if ($canRetry -ne 1) {
            return @{
                success = $false
                conflict = $true
                reviewRequired = $true
                httpStatus = 409
                errorCode = 'OFFLINE_SYNC_REVIEW_REQUIRED'
                error = if (-not [string]::IsNullOrWhiteSpace([string]$Record.ErrorMessage)) {
                    [string]$Record.ErrorMessage
                }
                else {
                    'The previous failure cannot be retried automatically.'
                }
            }
        }

        $table = $script:OfflineSyncLedgerTable
        $rows = Invoke-OfflineSyncNonQuery `
            -Connection $Connection `
            -IsSqlServer $IsSqlServer `
            -Sql @"
UPDATE [$table]
SET [Status] = 'PROCESSING',
    [Phase] = 'VALIDATING',
    [ProcessingToken] = @ProcessingToken,
    [ErrorMessage] = NULL,
    [CanRetry] = 1,
    [AttemptCount] = [AttemptCount] + 1,
    [UpdatedAt] = @UpdatedAt
WHERE [LocalId] = @LocalId
  AND [InstanceId] = @InstanceId
  AND [CompanyCode] = @CompanyCode
  AND [Status] = 'FAILED'
  AND [CanRetry] = 1
"@ `
            -Parameters @{
                ProcessingToken = $ProcessingToken
                UpdatedAt = Get-Date
                LocalId = ([string]$Data.localId).Trim()
                InstanceId = $InstanceId
                CompanyCode = $CompanyCode
            }

        if ($rows -eq 1) {
            return @{ success = $true; claimed = $true }
        }

        return @{
            success = $false
            conflict = $true
            httpStatus = 409
            errorCode = 'OFFLINE_SYNC_IN_PROGRESS'
            error = 'Another synchronization process claimed this retry first.'
        }
    }

    return @{
        success = $false
        conflict = $true
        httpStatus = 409
        errorCode = 'OFFLINE_SYNC_INVALID_SERVER_STATE'
        error = "Offline synchronization ledger has an unsupported state: $status."
    }
}

# -----------------------------------------------------------------------------
# Validation / numbering / reconciliation
# -----------------------------------------------------------------------------

function Test-OfflineVoucherSyncRequest {
    param(
        [Parameter(Mandatory = $true)]$Data
    )

    $localId = ([string](Get-OfflineSyncPropertyValue -Object $Data -Names @('localId') -DefaultValue '')).Trim()
    if ([string]::IsNullOrWhiteSpace($localId)) {
        return @{ success = $false; httpStatus = 400; errorCode = 'LOCAL_ID_REQUIRED'; error = 'localId is required.' }
    }

    $vchType = 0
    [void][int]::TryParse(
        [string](Get-OfflineSyncPropertyValue -Object $Data -Names @('vchType') -DefaultValue 0),
        [ref]$vchType
    )

    if ($vchType -ne 9) {
        return @{
            success = $false
            httpStatus = 400
            errorCode = 'OFFLINE_VOUCHER_TYPE_NOT_SUPPORTED'
            error = 'This first synchronization version supports new Sale vouchers only (vchType 9).'
        }
    }

    $payload = Get-OfflineSyncPropertyValue -Object $Data -Names @('payload') -DefaultValue $null
    if ($null -eq $payload) {
        return @{ success = $false; httpStatus = 400; errorCode = 'PAYLOAD_REQUIRED'; error = 'Voucher payload is required.' }
    }

    $payloadVchType = 0
    [void][int]::TryParse([string]$payload.vchType, [ref]$payloadVchType)
    if ($payloadVchType -ne $vchType) {
        return @{ success = $false; httpStatus = 400; errorCode = 'VOUCHER_TYPE_MISMATCH'; error = 'payload.vchType must match vchType.' }
    }

    $series = ([string](Get-OfflineSyncPropertyValue -Object $Data -Names @('requestedVchSeries') -DefaultValue '')).Trim()
    if ([string]::IsNullOrWhiteSpace($series)) {
        $series = ([string]$payload.vchSeries).Trim()
    }

    if ([string]::IsNullOrWhiteSpace($series)) {
        return @{ success = $false; httpStatus = 400; errorCode = 'VOUCHER_SERIES_REQUIRED'; error = 'Voucher series is required.' }
    }

    $dateText = ([string](Get-OfflineSyncPropertyValue -Object $Data -Names @('voucherDate') -DefaultValue '')).Trim()
    if ([string]::IsNullOrWhiteSpace($dateText)) {
        $dateText = ([string]$payload.date).Trim()
    }

    if (-not (Test-OfflineSyncDateText -Value $dateText)) {
        return @{ success = $false; httpStatus = 400; errorCode = 'INVALID_VOUCHER_DATE'; error = 'Voucher date is invalid.' }
    }

    $party = ([string](Get-OfflineSyncPropertyValue -Object $payload -Names @('party') -DefaultValue '')).Trim()
    if ([string]::IsNullOrWhiteSpace($party)) {
        return @{ success = $false; httpStatus = 400; errorCode = 'PARTY_REQUIRED'; error = 'Sale Party is required.' }
    }

    $items = @()
    try { $items = @($payload.items) } catch {}
    if ($items.Count -eq 0) {
        return @{ success = $false; httpStatus = 400; errorCode = 'ITEMS_REQUIRED'; error = 'At least one Sale item is required.' }
    }

    return @{
        success = $true
        vchType = $vchType
        series = $series
        voucherDate = $dateText
    }
}

function Resolve-OfflineVoucherPostingNumber {
    param(
        [Parameter(Mandatory = $true)]$Data,
        [Parameter(Mandatory = $true)][string]$InstanceId,
        [Parameter(Mandatory = $true)][string]$CompanyCode,
        [Parameter(Mandatory = $true)][string]$SeriesName,
        [Parameter(Mandatory = $true)][string]$VoucherDate
    )

    if (-not (Get-Command Get-EffectiveNumberingConfig -ErrorAction SilentlyContinue)) {
        return @{
            success = $false
            httpStatus = 500
            errorCode = 'NUMBERING_ENGINE_NOT_LOADED'
            error = 'Current BUSY/Web voucher numbering engine is not loaded.'
        }
    }

    # Local vouchers can contain dd-MM-yyyy (for example 16-09-2026),
    # while parts of the Web numbering engine parse dates using a culture
    # that may not accept that representation. Normalize only the date passed
    # to the numbering engine. The original voucher payload date is preserved.
    $numberingDate = ''

    try {
        $numberingDate = ConvertTo-OfflineSyncNumberingDate `
            -Value $VoucherDate
    }
    catch {
        return @{
            success = $false
            httpStatus = 400
            errorCode = 'INVALID_VOUCHER_DATE'
            error = $_.Exception.Message
        }
    }

    Write-Host (
        "  [OFFLINE-SYNC NUMBERING] type=9 series='{0}' voucherDate='{1}' numberingDate='{2}'" -f `
        $SeriesName,
        $VoucherDate,
        $numberingDate
    ) -ForegroundColor DarkCyan

    $config = Get-EffectiveNumberingConfig `
        -VchType 9 `
        -SeriesName $SeriesName `
        -VoucherDate $numberingDate `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if ($null -eq $config -or $config.success -eq $false) {
        return @{
            success = $false
            httpStatus = 400
            errorCode = 'NUMBERING_CONFIG_FAILED'
            error = if ($null -ne $config -and $config.error) {
                [string]$config.error
            }
            else {
                'Could not resolve the current voucher numbering configuration.'
            }
        }
    }

    $configData = Get-OfflineSyncPropertyValue `
        -Object $config `
        -Names @('data') `
        -DefaultValue $config

    $isAuto = ConvertTo-OfflineSyncBool (
        Get-OfflineSyncPropertyValue -Object $configData -Names @('is_auto', 'isAuto') -DefaultValue $false
    )

    if ($isAuto) {
        $nextNumber = ([string](
            Get-OfflineSyncPropertyValue `
                -Object $configData `
                -Names @('next_vch_no', 'nextVchNo') `
                -DefaultValue ''
        )).Trim()

        if ([string]::IsNullOrWhiteSpace($nextNumber)) {
            return @{
                success = $false
                httpStatus = 400
                errorCode = 'NEXT_VOUCHER_NUMBER_UNAVAILABLE'
                error = 'The series uses automatic numbering, but a current next voucher number could not be calculated.'
            }
        }

        return @{
            success = $true
            isAuto = $true
            vchNo = $nextNumber
            config = $configData
        }
    }

    # Manual-number series: preserve the number entered when the local voucher
    # was created. Create-Voucher remains responsible for its normal duplicate
    # checks against the current BUSY database.
    $requestedNumber = ([string](
        Get-OfflineSyncPropertyValue -Object $Data -Names @('requestedVchNo') -DefaultValue ''
    )).Trim()

    if ([string]::IsNullOrWhiteSpace($requestedNumber)) {
        try { $requestedNumber = ([string]$Data.payload.vchNo).Trim() } catch {}
    }

    if ([string]::IsNullOrWhiteSpace($requestedNumber)) {
        return @{
            success = $false
            httpStatus = 400
            errorCode = 'MANUAL_VOUCHER_NUMBER_REQUIRED'
            error = 'This series uses manual numbering, but the local voucher has no voucher number.'
        }
    }

    return @{
        success = $true
        isAuto = $false
        vchNo = $requestedNumber
        config = $configData
    }
}

function Find-OfflinePostedVoucher {
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][bool]$IsSqlServer,
        [Parameter(Mandatory = $true)][int]$VchType,
        [string]$VchNo = '',
        [string]$SeriesName = '',
        [int]$VchCode = 0
    )

    # Strongest identity: Create-Voucher returns the BUSY VchCode after a
    # successful create/approval-state normalization. Prefer it whenever we
    # have it because it avoids ambiguity across voucher series.
    if ($VchCode -gt 0) {
        return Invoke-OfflineSyncQuerySingle `
            -Connection $Connection `
            -IsSqlServer $IsSqlServer `
            -Sql @"
SELECT TOP 1
    [VchCode],
    [VchNo],
    [Date],
    [VchSeriesCode],
    [MasterCode1],
    [VchAmtBaseCur]
FROM [Tran1]
WHERE [VchCode] = @VchCode
  AND [VchType] = @VchType
"@ `
            -Parameters @{
                VchCode = $VchCode
                VchType = $VchType
            }
    }

    if ([string]::IsNullOrWhiteSpace($VchNo)) {
        return $null
    }

    $seriesFilter = ''
    $parameters = @{
        VchType = $VchType
        VchNo = $VchNo.Trim()
    }

    if (-not [string]::IsNullOrWhiteSpace($SeriesName)) {
        $seriesFilter = @"
  AND [VchSeriesCode] IN (
      SELECT [Code]
      FROM [Master1]
      WHERE [MasterType] = 21
        AND ([Name] = @SeriesName OR [Name] = @PrefixedSeriesName)
  )
"@

        $parameters['SeriesName'] = $SeriesName.Trim()
        $parameters['PrefixedSeriesName'] = ('{0:D2}{1}' -f $VchType, $SeriesName.Trim())
    }

    $sql = @"
SELECT TOP 1
    [VchCode],
    [VchNo],
    [Date],
    [VchSeriesCode],
    [MasterCode1],
    [VchAmtBaseCur]
FROM [Tran1]
WHERE [VchType] = @VchType
  AND [VchNo] = @VchNo
$seriesFilter
ORDER BY [VchCode] DESC
"@

    return Invoke-OfflineSyncQuerySingle `
        -Connection $Connection `
        -IsSqlServer $IsSqlServer `
        -Sql $sql `
        -Parameters $parameters
}

function Get-OfflineCreateResultInfo {
    param($Result)

    $success = ConvertTo-OfflineSyncBool (
        Get-OfflineSyncPropertyValue -Object $Result -Names @('success') -DefaultValue $false
    )

    $created = ConvertTo-OfflineSyncBool (
        Get-OfflineSyncPropertyValue -Object $Result -Names @('created') -DefaultValue $false
    )

    $data = Get-OfflineSyncPropertyValue -Object $Result -Names @('data') -DefaultValue $null

    $vchNo = ([string](
        Get-OfflineSyncPropertyValue `
            -Object $Result `
            -Names @('finalVchNo', 'vchNo', 'voucherNo') `
            -DefaultValue ''
    )).Trim()

    if ([string]::IsNullOrWhiteSpace($vchNo) -and $null -ne $data) {
        $vchNo = ([string](
            Get-OfflineSyncPropertyValue `
                -Object $data `
                -Names @('finalVchNo', 'vchNo', 'voucherNo') `
                -DefaultValue ''
        )).Trim()
    }

    $vchCode = 0
    $rawCode = Get-OfflineSyncPropertyValue `
        -Object $Result `
        -Names @('finalVchCode', 'vchCode', 'voucherCode') `
        -DefaultValue 0

    [void][int]::TryParse([string]$rawCode, [ref]$vchCode)

    if ($vchCode -le 0 -and $null -ne $data) {
        $rawCode = Get-OfflineSyncPropertyValue `
            -Object $data `
            -Names @('finalVchCode', 'vchCode', 'voucherCode') `
            -DefaultValue 0

        [void][int]::TryParse([string]$rawCode, [ref]$vchCode)
    }

    $message = ([string](
        Get-OfflineSyncPropertyValue -Object $Result -Names @('message') -DefaultValue ''
    )).Trim()

    # IMPORTANT:
    # "Sale created successfully" is a normal success message, NOT an error.
    # Only error/warning fields are considered an error condition here.
    $error = ([string](
        Get-OfflineSyncPropertyValue -Object $Result -Names @('error', 'warning') -DefaultValue ''
    )).Trim()

    return @{
        success = $success
        created = $created
        vchNo = $vchNo
        vchCode = $vchCode
        message = $message
        error = $error
    }
}


function Get-OfflineSyncMasterName {
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][bool]$IsSqlServer,
        [int]$Code = 0
    )

    if ($Code -le 0) {
        return ''
    }

    $row = Invoke-OfflineSyncQuerySingle `
        -Connection $Connection `
        -IsSqlServer $IsSqlServer `
        -Sql "SELECT TOP 1 [Name] FROM [Master1] WHERE [Code] = @Code" `
        -Parameters @{ Code = $Code }

    if ($null -eq $row) {
        return ''
    }

    return ([string]$row.Name).Trim()
}

function Normalize-OfflineSyncSeriesName {
    param(
        [string]$Value,
        [int]$VchType
    )

    $name = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($name)) {
        return ''
    }

    $prefix = ('{0:D2}' -f $VchType)
    if (
        $name.StartsWith(
            $prefix,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        return $name.Substring($prefix.Length).Trim()
    }

    return $name
}


function Normalize-OfflineVoucherPostingPayload {
    param(
        [Parameter(Mandatory = $true)]$Payload
    )

    $copy = Copy-OfflineSyncObject -Object $Payload

    $normalizedItems = @()
    $itemIndex = 1

    try {
        foreach ($item in @($copy.items)) {
            if ($null -eq $item) { continue }

            $itemName = ''
            try { $itemName = ([string]$item.itemName).Trim() } catch {}
            if ([string]::IsNullOrWhiteSpace($itemName)) { continue }

            $item | Add-Member `
                -MemberType NoteProperty `
                -Name 'srNo' `
                -Value $itemIndex `
                -Force

            $normalizedItems += $item
            $itemIndex++
        }
    }
    catch {}

    $copy | Add-Member `
        -MemberType NoteProperty `
        -Name 'items' `
        -Value @($normalizedItems) `
        -Force

    $normalizedSundries = @()
    $sundryIndex = 1

    try {
        foreach ($sundry in @($copy.billSundries)) {
            if ($null -eq $sundry) { continue }

            $name = ''
            try { $name = ([string]$sundry.name).Trim() } catch {}
            if ([string]::IsNullOrWhiteSpace($name)) { continue }

            $sundry | Add-Member `
                -MemberType NoteProperty `
                -Name 'srNo' `
                -Value $sundryIndex `
                -Force

            $normalizedSundries += $sundry
            $sundryIndex++
        }
    }
    catch {}

    $copy | Add-Member `
        -MemberType NoteProperty `
        -Name 'billSundries' `
        -Value @($normalizedSundries) `
        -Force

    # BBA/reference rows should also be contiguous when present.
    $normalizedRefs = @()
    $refIndex = 1

    try {
        foreach ($ref in @($copy.refEntries)) {
            if ($null -eq $ref) { continue }

            $ref | Add-Member `
                -MemberType NoteProperty `
                -Name 'srNo' `
                -Value $refIndex `
                -Force

            $normalizedRefs += $ref
            $refIndex++
        }
    }
    catch {}

    if ($null -ne $copy.PSObject.Properties['refEntries']) {
        $copy | Add-Member `
            -MemberType NoteProperty `
            -Name 'refEntries' `
            -Value @($normalizedRefs) `
            -Force
    }

    return $copy
}

function Get-OfflineSyncDateKey {
    param($Value)

    if ($null -eq $Value -or $Value -eq [System.DBNull]::Value) {
        return ''
    }

    try {
        if ($Value -is [datetime]) {
            return ([datetime]$Value).ToString('yyyy-MM-dd')
        }
    }
    catch {}

    $raw = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return ''
    }

    try {
        return ConvertTo-OfflineSyncNumberingDate -Value $raw
    }
    catch {
        return ''
    }
}

function Get-OfflineSyncExpectedTotal {
    param($Payload)

    $total = 0.0

    try {
        foreach ($item in @($Payload.items)) {
            if ($null -eq $item) { continue }
            try { $total += [double]$item.amount } catch {}
        }
    }
    catch {}

    try {
        foreach ($bs in @($Payload.billSundries)) {
            if ($null -eq $bs) { continue }
            try { $total += [double]$bs.amount } catch {}
        }
    }
    catch {}

    return [Math]::Round([Math]::Abs($total), 2)
}

function Get-OfflineSyncItemSignature {
    param($Item)

    $name = ''
    $qty = 0.0

    try { $name = ([string]$Item.itemName).Trim().ToLowerInvariant() } catch {}
    try { $qty = [Math]::Round([double]$Item.qty, 4) } catch {}

    return ('{0}|{1}' -f $name, $qty.ToString('0.####', [System.Globalization.CultureInfo]::InvariantCulture))
}

function Test-OfflineVoucherIntegrity {
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][bool]$IsSqlServer,
        [Parameter(Mandatory = $true)]$Header,
        [Parameter(Mandatory = $true)]$ExpectedPayload,
        [Parameter(Mandatory = $true)][string]$VoucherDate,
        [Parameter(Mandatory = $true)][string]$ExpectedSeries,
        [Parameter(Mandatory = $true)][string]$InstanceId,
        [Parameter(Mandatory = $true)][string]$CompanyCode
    )

    $vchCode = 0
    $vchType = 9
    $vchNo = ''
    $masterCode1 = 0
    $seriesCode = 0
    $headerAmount = 0.0

    try { $vchCode = [int]$Header.VchCode } catch {}
    try { $vchNo = ([string]$Header.VchNo).Trim() } catch {}
    try { $masterCode1 = [int]$Header.MasterCode1 } catch {}
    try { $seriesCode = [int]$Header.VchSeriesCode } catch {}
    try { $headerAmount = [Math]::Abs([double]$Header.VchAmtBaseCur) } catch {}

    $actualParty = Get-OfflineSyncMasterName `
        -Connection $Connection `
        -IsSqlServer $IsSqlServer `
        -Code $masterCode1

    $storedSeries = Get-OfflineSyncMasterName `
        -Connection $Connection `
        -IsSqlServer $IsSqlServer `
        -Code $seriesCode

    $actualSeries = Normalize-OfflineSyncSeriesName `
        -Value $storedSeries `
        -VchType $vchType

    $expectedParty = ''
    try { $expectedParty = ([string]$ExpectedPayload.party).Trim() } catch {}

    $partyMatches = (
        -not [string]::IsNullOrWhiteSpace($expectedParty) -and
        $actualParty.Trim().Equals(
            $expectedParty,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    )

    $expectedDateKey = Get-OfflineSyncDateKey -Value $VoucherDate
    $actualDateKey = Get-OfflineSyncDateKey -Value $Header.Date
    $dateMatches = (
        -not [string]::IsNullOrWhiteSpace($expectedDateKey) -and
        $expectedDateKey -eq $actualDateKey
    )

    $seriesMatches = $true
    $normalizedExpectedSeries = Normalize-OfflineSyncSeriesName `
        -Value $ExpectedSeries `
        -VchType $vchType

    if (
        -not [string]::IsNullOrWhiteSpace($normalizedExpectedSeries) -and
        -not [string]::IsNullOrWhiteSpace($actualSeries)
    ) {
        $seriesMatches = $actualSeries.Equals(
            $normalizedExpectedSeries,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    }

    $detail = $null

    if (Get-Command Get-VoucherDetail -ErrorAction SilentlyContinue) {
        try {
            $detail = Get-VoucherDetail `
                -VchType $vchType `
                -VchNo $vchNo `
                -VchSeries $actualSeries `
                -VchDate $VoucherDate `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode
        }
        catch {
            $detail = @{
                success = $false
                error = $_.Exception.Message
            }
        }
    }
    else {
        $detail = @{
            success = $false
            error = 'Get-VoucherDetail is not loaded.'
        }
    }

    $detailReadable = (
        $null -ne $detail -and
        $detail.success -eq $true -and
        $null -ne $detail.data
    )

    $expectedItems = @()
    $actualItems = @()

    try {
        $expectedItems = @(
            @($ExpectedPayload.items) |
            Where-Object {
                $null -ne $_ -and
                -not [string]::IsNullOrWhiteSpace([string]$_.itemName)
            }
        )
    }
    catch {}

    if ($detailReadable) {
        try {
            $actualItems = @(
                @($detail.data.items) |
                Where-Object {
                    $null -ne $_ -and
                    -not [string]::IsNullOrWhiteSpace([string]$_.itemName)
                }
            )
        }
        catch {}
    }

    $expectedSignatures = @(
        $expectedItems |
        ForEach-Object { Get-OfflineSyncItemSignature -Item $_ } |
        Sort-Object
    )

    $actualSignatures = @(
        $actualItems |
        ForEach-Object { Get-OfflineSyncItemSignature -Item $_ } |
        Sort-Object
    )

    $itemsMatch = $false
    if ($detailReadable -and $expectedSignatures.Count -eq $actualSignatures.Count) {
        $itemsMatch = (($expectedSignatures -join '||') -eq ($actualSignatures -join '||'))
    }

    $expectedTotal = Get-OfflineSyncExpectedTotal -Payload $ExpectedPayload

    $detailTotal = 0.0
    if ($detailReadable) {
        try {
            foreach ($item in $actualItems) {
                try { $detailTotal += [double]$item.amount } catch {}
            }

            foreach ($bs in @($detail.data.billSundries)) {
                try { $detailTotal += [double]$bs.amount } catch {}
            }
        }
        catch {}

        $detailTotal = [Math]::Round([Math]::Abs($detailTotal), 2)
    }

    $headerAmountRounded = [Math]::Round($headerAmount, 2)

    $headerAmountMatches = $true
    $detailAmountMatches = $true

    if ($expectedTotal -gt 0) {
        $headerAmountMatches = ([Math]::Abs($headerAmountRounded - $expectedTotal) -le 0.02)
        $detailAmountMatches = ($detailReadable -and [Math]::Abs($detailTotal - $expectedTotal) -le 0.02)
    }

    $identityMatches = ($partyMatches -and $dateMatches -and $seriesMatches)

    $complete = (
        $identityMatches -and
        $detailReadable -and
        $itemsMatch -and
        $headerAmountMatches -and
        $detailAmountMatches
    )

    $reasons = @()

    if (-not $partyMatches) {
        $reasons += "Party does not match the local transaction."
    }

    if (-not $dateMatches) {
        $reasons += "Voucher date does not match the local transaction."
    }

    if (-not $seriesMatches) {
        $reasons += "Voucher series does not match the local transaction."
    }

    if (-not $detailReadable) {
        $detailError = ''
        try { $detailError = ([string]$detail.error).Trim() } catch {}
        if ([string]::IsNullOrWhiteSpace($detailError)) {
            $detailError = 'BUSY voucher detail could not be loaded.'
        }

        $reasons += $detailError
    }
    elseif (-not $itemsMatch) {
        $reasons += "BUSY voucher item lines do not match the local transaction."
    }

    if (-not $headerAmountMatches) {
        $reasons += "BUSY voucher header amount does not match the local transaction."
    }

    if ($detailReadable -and -not $detailAmountMatches) {
        $reasons += "BUSY voucher detail amount does not match the local transaction."
    }

    return @{
        success = $true
        complete = $complete
        vchCode = $vchCode
        vchNo = $vchNo
        actualParty = $actualParty
        actualSeries = $actualSeries
        voucherDate = $actualDateKey
        expectedItemCount = $expectedItems.Count
        actualItemCount = $actualItems.Count
        expectedItemSignatures = @($expectedSignatures)
        actualItemSignatures = @($actualSignatures)
        itemsMatch = $itemsMatch
        expectedTotal = $expectedTotal
        headerAmount = $headerAmountRounded
        detailTotal = $detailTotal
        headerAmountMatches = $headerAmountMatches
        detailAmountMatches = $detailAmountMatches
        detailReadable = $detailReadable
        headerFound = $true
        reasons = @($reasons)
    }
}

function Set-OfflineSyncLedgerFailedFromReview {
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][bool]$IsSqlServer,
        [Parameter(Mandatory = $true)][string]$LocalId,
        [Parameter(Mandatory = $true)][string]$InstanceId,
        [Parameter(Mandatory = $true)][string]$CompanyCode,
        [Parameter(Mandatory = $true)][string]$ErrorMessage
    )

    $table = $script:OfflineSyncLedgerTable

    return Invoke-OfflineSyncNonQuery `
        -Connection $Connection `
        -IsSqlServer $IsSqlServer `
        -Sql @"
UPDATE [$table]
SET [Status] = 'FAILED',
    [Phase] = 'FAILED',
    [ErrorMessage] = @ErrorMessage,
    [CanRetry] = 1,
    [ProcessingToken] = NULL,
    [UpdatedAt] = @UpdatedAt
WHERE [LocalId] = @LocalId
  AND [InstanceId] = @InstanceId
  AND [CompanyCode] = @CompanyCode
  AND [Status] = 'REVIEW_REQUIRED'
"@ `
        -Parameters @{
            ErrorMessage = $ErrorMessage
            UpdatedAt = Get-Date
            LocalId = $LocalId
            InstanceId = $InstanceId
            CompanyCode = $CompanyCode
        }
}

function Invoke-OfflineVoucherReconciliation {
    param(
        [Parameter(Mandatory = $true)][string]$LocalId,
        [Parameter(Mandatory = $true)]$AuthResult,
        [Parameter(Mandatory = $true)][string]$InstanceId,
        [Parameter(Mandatory = $true)][string]$CompanyCode
    )

    if ([string]::IsNullOrWhiteSpace($LocalId)) {
        return @{
            success = $false
            httpStatus = 400
            errorCode = 'LOCAL_ID_REQUIRED'
            error = 'localId is required.'
        }
    }

    if ($null -eq $AuthResult -or $null -eq $AuthResult.user) {
        return @{
            success = $false
            httpStatus = 401
            errorCode = 'AUTH_REQUIRED'
            error = 'Authentication is required.'
        }
    }

    $ctx = $null
    $connection = $null

    try {
        $ctx = Get-BusyCloudFastConfigDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if ($null -eq $ctx -or $null -eq $ctx.connection) {
            return @{
                success = $false
                httpStatus = 500
                errorCode = 'BUSY_DATABASE_UNAVAILABLE'
                error = 'Could not open the current BUSY company database.'
            }
        }

        $connection = $ctx.connection
        $isSqlServer = ([int]$ctx.dbType -eq 1)

        Ensure-OfflineSyncLedger `
            -Connection $connection `
            -IsSqlServer $isSqlServer

        $record = Get-OfflineSyncLedgerRecord `
            -Connection $connection `
            -IsSqlServer $isSqlServer `
            -LocalId $LocalId.Trim() `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if ($null -eq $record) {
            return @{
                success = $false
                httpStatus = 404
                errorCode = 'OFFLINE_SYNC_RECORD_NOT_FOUND'
                error = 'No server-side synchronization record exists for this localId.'
            }
        }

        $status = ([string]$record.Status).Trim().ToUpperInvariant()

        if ($status -eq 'SYNCED') {
            return @{
                success = $true
                reconciliationStatus = 'SYNCED'
                localId = $LocalId.Trim()
                finalVchNo = ([string]$record.FinalVchNo).Trim()
                finalVchCode = if ($null -ne $record.FinalVchCode) { [int]$record.FinalVchCode } else { $null }
                message = 'This transaction was already confirmed as synchronized.'
            }
        }

        if ($status -ne 'REVIEW_REQUIRED') {
            return @{
                success = $false
                httpStatus = 409
                errorCode = 'OFFLINE_SYNC_NOT_IN_REVIEW'
                error = "This transaction is in server state '$status', not REVIEW_REQUIRED."
            }
        }

        $storedData = $null
        try {
            $storedData = ([string]$record.RequestPayload) | ConvertFrom-Json
        }
        catch {
            return @{
                success = $false
                reviewRequired = $true
                reconciliationStatus = 'REVIEW_REQUIRED'
                httpStatus = 409
                errorCode = 'OFFLINE_SYNC_REVIEW_PAYLOAD_INVALID'
                error = 'The server-side review record cannot read its original local payload.'
            }
        }

        $payload = $storedData.payload

        if ($null -eq $payload) {
            return @{
                success = $false
                reviewRequired = $true
                reconciliationStatus = 'REVIEW_REQUIRED'
                httpStatus = 409
                errorCode = 'OFFLINE_SYNC_REVIEW_PAYLOAD_MISSING'
                error = 'The original local voucher payload is missing from the review record.'
            }
        }

        $candidateNo = ([string]$record.PostingVchNo).Trim()

        if ([string]::IsNullOrWhiteSpace($candidateNo)) {
            $candidateNo = ([string]$record.RequestedVchNo).Trim()
        }

        $expectedSeries = ([string]$record.VchSeries).Trim()

        $voucherDate = ''
        try { $voucherDate = ([string]$storedData.voucherDate).Trim() } catch {}
        if ([string]::IsNullOrWhiteSpace($voucherDate)) {
            try { $voucherDate = ([string]$payload.date).Trim() } catch {}
        }

        # First try the original strict series lookup.
        $header = Find-OfflinePostedVoucher `
            -Connection $connection `
            -IsSqlServer $isSqlServer `
            -VchType ([int]$record.VchType) `
            -VchNo $candidateNo `
            -SeriesName $expectedSeries

        # Some BUSY companies store/return series names differently. If the
        # strict series lookup misses, search the exact voucher number without
        # the series filter, then validate date/party/series/items/amount below.
        if ($null -eq $header) {
            $header = Find-OfflinePostedVoucher `
                -Connection $connection `
                -IsSqlServer $isSqlServer `
                -VchType ([int]$record.VchType) `
                -VchNo $candidateNo
        }

        if ($null -eq $header) {
            # We still do NOT blindly retry here. A prior Create-Voucher call
            # had already reached the dangerous posting phase. The safest
            # outcome is to keep REVIEW_REQUIRED until a human confirms that
            # BUSY truly contains no corresponding voucher.
            $message = "No BUSY header could be found for candidate voucher '$candidateNo'. Automatic retry remains blocked because the previous create reached the BUSY posting phase."

            [void](Set-OfflineSyncLedgerReviewRequired `
                -Connection $connection `
                -IsSqlServer $isSqlServer `
                -Data @{ localId = $LocalId.Trim() } `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode `
                -ErrorMessage $message)

            return @{
                success = $false
                reviewRequired = $true
                reconciliationStatus = 'REVIEW_REQUIRED'
                httpStatus = 409
                errorCode = 'OFFLINE_SYNC_REVIEW_NOT_FOUND'
                localId = $LocalId.Trim()
                error = $message
                diagnostics = @{
                    headerFound = $false
                    candidateVchNo = $candidateNo
                }
            }
        }

        $integrity = Test-OfflineVoucherIntegrity `
            -Connection $connection `
            -IsSqlServer $isSqlServer `
            -Header $header `
            -ExpectedPayload $payload `
            -VoucherDate $voucherDate `
            -ExpectedSeries $expectedSeries `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if ($integrity.complete -eq $true) {
            $finalCode = 0
            try { $finalCode = [int]$integrity.vchCode } catch {}

            [void](Set-OfflineSyncLedgerSynced `
                -Connection $connection `
                -IsSqlServer $isSqlServer `
                -Data @{ localId = $LocalId.Trim() } `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode `
                -FinalVchNo ([string]$integrity.vchNo).Trim() `
                -FinalVchCode $finalCode `
                -WarningMessage 'Confirmed by manual reconciliation check.')

            return @{
                success = $true
                reconciliationStatus = 'SYNCED'
                localId = $LocalId.Trim()
                finalVchNo = ([string]$integrity.vchNo).Trim()
                finalVchCode = if ($finalCode -gt 0) { $finalCode } else { $null }
                message = 'BUSY voucher header, item lines and amount were confirmed.'
                diagnostics = $integrity
            }
        }

        $reasonText = (@($integrity.reasons) -join ' ')
        if ([string]::IsNullOrWhiteSpace($reasonText)) {
            $reasonText = 'The BUSY voucher exists, but its contents could not be confirmed as complete.'
        }

        $reviewMessage = "BUSY voucher '$candidateNo' exists, but reconciliation found an integrity problem. $reasonText"

        [void](Set-OfflineSyncLedgerReviewRequired `
            -Connection $connection `
            -IsSqlServer $isSqlServer `
            -Data @{ localId = $LocalId.Trim() } `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode `
            -ErrorMessage $reviewMessage)

        return @{
            success = $false
            reviewRequired = $true
            reconciliationStatus = 'REVIEW_REQUIRED'
            httpStatus = 409
            errorCode = 'OFFLINE_SYNC_REVIEW_INTEGRITY_MISMATCH'
            localId = $LocalId.Trim()
            error = $reviewMessage
            diagnostics = $integrity
        }
    }
    catch {
        return @{
            success = $false
            reviewRequired = $true
            reconciliationStatus = 'REVIEW_REQUIRED'
            httpStatus = 500
            errorCode = 'OFFLINE_SYNC_RECONCILIATION_FAILED'
            localId = $LocalId.Trim()
            error = $_.Exception.Message
        }
    }
    finally {
        if ($null -ne $connection) {
            try { $connection.Close() } catch {}
            try { $connection.Dispose() } catch {}
        }
    }
}


function Invoke-OfflineVoucherRepair {
    param(
        [Parameter(Mandatory = $true)][string]$LocalId,
        [Parameter(Mandatory = $true)]$AuthResult,
        [Parameter(Mandatory = $true)][string]$InstanceId,
        [Parameter(Mandatory = $true)][string]$CompanyCode
    )

    if ([string]::IsNullOrWhiteSpace($LocalId)) {
        return @{ success=$false; httpStatus=400; errorCode='LOCAL_ID_REQUIRED'; error='localId is required.' }
    }

    if ($null -eq $AuthResult -or $null -eq $AuthResult.user) {
        return @{ success=$false; httpStatus=401; errorCode='AUTH_REQUIRED'; error='Authentication is required.' }
    }

    if (-not (Get-Command Modify-Voucher -ErrorAction SilentlyContinue)) {
        return @{ success=$false; httpStatus=500; errorCode='MODIFY_VOUCHER_NOT_LOADED'; error='Modify-Voucher is not loaded.' }
    }

    $ctx = $null
    $connection = $null

    try {
        $ctx = Get-BusyCloudFastConfigDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if ($null -eq $ctx -or $null -eq $ctx.connection) {
            return @{ success=$false; httpStatus=500; errorCode='BUSY_DATABASE_UNAVAILABLE'; error='Could not open the current BUSY company database.' }
        }

        $connection = $ctx.connection
        $isSqlServer = ([int]$ctx.dbType -eq 1)

        Ensure-OfflineSyncLedger -Connection $connection -IsSqlServer $isSqlServer

        $record = Get-OfflineSyncLedgerRecord `
            -Connection $connection `
            -IsSqlServer $isSqlServer `
            -LocalId $LocalId.Trim() `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if ($null -eq $record) {
            return @{ success=$false; httpStatus=404; errorCode='OFFLINE_SYNC_RECORD_NOT_FOUND'; error='No server-side synchronization record exists for this localId.' }
        }

        $status = ([string]$record.Status).Trim().ToUpperInvariant()
        if ($status -eq 'SYNCED') {
            return @{
                success=$true
                repaired=$false
                alreadyProcessed=$true
                reconciliationStatus='SYNCED'
                localId=$LocalId.Trim()
                finalVchNo=([string]$record.FinalVchNo).Trim()
                finalVchCode=if ($null -ne $record.FinalVchCode) { [int]$record.FinalVchCode } else { $null }
                message='This transaction is already synchronized.'
            }
        }

        if ($status -ne 'REVIEW_REQUIRED') {
            return @{ success=$false; httpStatus=409; errorCode='OFFLINE_SYNC_NOT_IN_REVIEW'; error="This transaction is in server state '$status', not REVIEW_REQUIRED." }
        }

        # A normal user may repair only his/her own local transaction.
        $isAdmin = $false
        if (Get-Command Test-IsPermissionAdminUser -ErrorAction SilentlyContinue) {
            $isAdmin = Test-IsPermissionAdminUser -User $AuthResult.user
        }

        $activeUser = ([string]$AuthResult.user.name).Trim()
        $recordUser = ([string]$record.UserName).Trim()

        if (
            -not $isAdmin -and
            -not [string]::IsNullOrWhiteSpace($recordUser) -and
            -not $recordUser.Equals($activeUser, [System.StringComparison]::OrdinalIgnoreCase)
        ) {
            return @{ success=$false; httpStatus=403; errorCode='OFFLINE_REPAIR_USER_MISMATCH'; error='This local transaction belongs to another BUSY user.' }
        }

        $storedData = $null
        try { $storedData = ([string]$record.RequestPayload) | ConvertFrom-Json }
        catch {
            return @{ success=$false; reviewRequired=$true; httpStatus=409; errorCode='OFFLINE_REPAIR_PAYLOAD_INVALID'; error='The server cannot read the original local transaction payload.' }
        }

        if ($null -eq $storedData -or $null -eq $storedData.payload) {
            return @{ success=$false; reviewRequired=$true; httpStatus=409; errorCode='OFFLINE_REPAIR_PAYLOAD_MISSING'; error='The original local transaction payload is missing.' }
        }

        $candidateNo = ([string]$record.PostingVchNo).Trim()
        if ([string]::IsNullOrWhiteSpace($candidateNo)) {
            $candidateNo = ([string]$record.RequestedVchNo).Trim()
        }

        $expectedSeries = ([string]$record.VchSeries).Trim()
        $voucherDate = ''
        try { $voucherDate = ([string]$storedData.voucherDate).Trim() } catch {}
        if ([string]::IsNullOrWhiteSpace($voucherDate)) {
            try { $voucherDate = ([string]$storedData.payload.date).Trim() } catch {}
        }

        $header = Find-OfflinePostedVoucher `
            -Connection $connection `
            -IsSqlServer $isSqlServer `
            -VchType ([int]$record.VchType) `
            -VchNo $candidateNo `
            -SeriesName $expectedSeries

        if ($null -eq $header) {
            $header = Find-OfflinePostedVoucher `
                -Connection $connection `
                -IsSqlServer $isSqlServer `
                -VchType ([int]$record.VchType) `
                -VchNo $candidateNo
        }

        if ($null -eq $header) {
            return @{
                success=$false
                reviewRequired=$true
                httpStatus=409
                errorCode='OFFLINE_REPAIR_HEADER_NOT_FOUND'
                error="The BUSY voucher '$candidateNo' cannot be found. Repair was not attempted."
            }
        }

        # Before modifying anything, prove that this is the same voucher identity.
        $before = Test-OfflineVoucherIntegrity `
            -Connection $connection `
            -IsSqlServer $isSqlServer `
            -Header $header `
            -ExpectedPayload $storedData.payload `
            -VoucherDate $voucherDate `
            -ExpectedSeries $expectedSeries `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        $identitySafe = $true
        if (-not ([string]$before.actualParty).Trim().Equals(([string]$storedData.payload.party).Trim(), [System.StringComparison]::OrdinalIgnoreCase)) {
            $identitySafe = $false
        }
        if ((Get-OfflineSyncDateKey -Value $voucherDate) -ne ([string]$before.voucherDate).Trim()) {
            $identitySafe = $false
        }

        if (-not $identitySafe) {
            return @{
                success=$false
                reviewRequired=$true
                httpStatus=409
                errorCode='OFFLINE_REPAIR_IDENTITY_MISMATCH'
                error='Repair was blocked because the BUSY voucher party/date does not match the original local transaction.'
                diagnostics=$before
            }
        }

        $payload = Normalize-OfflineVoucherPostingPayload -Payload $storedData.payload
        $payload | Add-Member -MemberType NoteProperty -Name 'vchType' -Value 9 -Force
        $payload | Add-Member -MemberType NoteProperty -Name 'vchNo' -Value ([string]$header.VchNo).Trim() -Force
        $payload | Add-Member -MemberType NoteProperty -Name 'vchSeries' -Value $(if (-not [string]::IsNullOrWhiteSpace([string]$before.actualSeries)) { [string]$before.actualSeries } else { $expectedSeries }) -Force
        $payload | Add-Member -MemberType NoteProperty -Name 'date' -Value $voucherDate -Force
        $payload | Add-Member -MemberType NoteProperty -Name 'bridgeUserName' -Value $activeUser -Force

        # Re-run current permissions/salesman rules. Admins retain the explicit bypass.
        $accessResult = Invoke-OfflineVoucherCurrentAccessChecks `
            -AuthResult $AuthResult `
            -Payload $payload `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (-not $accessResult.success) {
            return $accessResult
        }

        $payload = Normalize-OfflineVoucherPostingPayload -Payload $accessResult.data

        # Preserve the current BUSY approval state while repairing the line data.
        if (Get-Command Get-VoucherApprovalStateByIdentity -ErrorAction SilentlyContinue) {
            try {
                $approvalState = Get-VoucherApprovalStateByIdentity `
                    -VchType 9 `
                    -VchNo ([string]$payload.vchNo) `
                    -VchDate ([string]$payload.date) `
                    -VchSeries ([string]$payload.vchSeries) `
                    -InstanceId $InstanceId `
                    -CompanyCode $CompanyCode

                if ($approvalState.success -and $approvalState.found) {
                    $payload | Add-Member -MemberType NoteProperty -Name 'preserveApprovalStatus' -Value ([int]$approvalState.approvalStatus) -Force
                    $payload | Add-Member -MemberType NoteProperty -Name 'preserveApprovedState' -Value ([int]$approvalState.approvalStatus -eq 1) -Force
                }
            }
            catch {}
        }

        Write-Host (
            "  [OFFLINE-REPAIR] localId={0} vchNo='{1}' series='{2}' normalizedItems={3}" -f `
            $LocalId.Trim(),
            [string]$payload.vchNo,
            [string]$payload.vchSeries,
            @($payload.items).Count
        ) -ForegroundColor Yellow

        $modifyResult = Modify-Voucher `
            -Data $payload `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if ($null -eq $modifyResult -or $modifyResult.success -ne $true) {
            $message = if ($null -ne $modifyResult -and $modifyResult.error) { [string]$modifyResult.error } else { 'BUSY did not confirm the repair.' }

            [void](Set-OfflineSyncLedgerReviewRequired `
                -Connection $connection `
                -IsSqlServer $isSqlServer `
                -Data @{ localId=$LocalId.Trim() } `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode `
                -ErrorMessage "Repair attempt failed. $message")

            return @{ success=$false; reviewRequired=$true; httpStatus=409; errorCode='OFFLINE_REPAIR_MODIFY_FAILED'; error="Repair attempt failed. $message" }
        }

        # Re-read by exact voucher identity after Modify-Voucher.
        $afterHeader = Find-OfflinePostedVoucher `
            -Connection $connection `
            -IsSqlServer $isSqlServer `
            -VchType 9 `
            -VchNo ([string]$payload.vchNo) `
            -SeriesName ([string]$payload.vchSeries)

        if ($null -eq $afterHeader) {
            $afterHeader = Find-OfflinePostedVoucher `
                -Connection $connection `
                -IsSqlServer $isSqlServer `
                -VchType 9 `
                -VchNo ([string]$payload.vchNo)
        }

        if ($null -eq $afterHeader) {
            return @{ success=$false; reviewRequired=$true; httpStatus=409; errorCode='OFFLINE_REPAIR_RECHECK_NOT_FOUND'; error='BUSY accepted the repair call, but the voucher could not be re-read for verification.' }
        }

        $after = Test-OfflineVoucherIntegrity `
            -Connection $connection `
            -IsSqlServer $isSqlServer `
            -Header $afterHeader `
            -ExpectedPayload $payload `
            -VoucherDate $voucherDate `
            -ExpectedSeries ([string]$payload.vchSeries) `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (-not $after.complete) {
            $reasonText = (@($after.reasons) -join ' ')
            if ([string]::IsNullOrWhiteSpace($reasonText)) { $reasonText = 'The repaired voucher still failed integrity verification.' }

            [void](Set-OfflineSyncLedgerReviewRequired `
                -Connection $connection `
                -IsSqlServer $isSqlServer `
                -Data @{ localId=$LocalId.Trim() } `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode `
                -ErrorMessage "Repair completed but verification still failed. $reasonText")

            return @{
                success=$false
                reviewRequired=$true
                httpStatus=409
                errorCode='OFFLINE_REPAIR_VERIFY_FAILED'
                error="Repair completed but verification still failed. $reasonText"
                diagnostics=$after
            }
        }

        $finalCode = 0
        try { $finalCode = [int]$after.vchCode } catch {}

        [void](Set-OfflineSyncLedgerSynced `
            -Connection $connection `
            -IsSqlServer $isSqlServer `
            -Data @{ localId=$LocalId.Trim() } `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode `
            -FinalVchNo ([string]$after.vchNo).Trim() `
            -FinalVchCode $finalCode `
            -WarningMessage 'Existing BUSY voucher was repaired from the original local payload and verified.')

        return @{
            success=$true
            repaired=$true
            reconciliationStatus='SYNCED'
            localId=$LocalId.Trim()
            finalVchNo=([string]$after.vchNo).Trim()
            finalVchCode=if ($finalCode -gt 0) { $finalCode } else { $null }
            message='Existing BUSY voucher was repaired and verified successfully.'
            diagnostics=$after
        }
    }
    catch {
        return @{ success=$false; reviewRequired=$true; httpStatus=500; errorCode='OFFLINE_REPAIR_FAILED'; error=$_.Exception.Message }
    }
    finally {
        if ($null -ne $connection) {
            try { $connection.Close() } catch {}
            try { $connection.Dispose() } catch {}
        }
    }
}

function Invoke-OfflineVoucherCurrentAccessChecks {
    param(
        [Parameter(Mandatory = $true)]$AuthResult,
        [Parameter(Mandatory = $true)]$Payload,
        [Parameter(Mandatory = $true)][string]$InstanceId,
        [Parameter(Mandatory = $true)][string]$CompanyCode
    )

    $isAdmin = $false
    if (Get-Command Test-IsPermissionAdminUser -ErrorAction SilentlyContinue) {
        $isAdmin = Test-IsPermissionAdminUser -User $AuthResult.user
    }

    # Keep the existing Salesman business helper in the same flow. For admins
    # its current implementation already bypasses assignment restrictions.
    if (Get-Command Apply-SalesmanAssignmentToVoucherData -ErrorAction SilentlyContinue) {
        $salesmanCheck = Apply-SalesmanAssignmentToVoucherData `
            -AuthResult $AuthResult `
            -Data $Payload `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode `
            -RequireAuth $true

        if (-not $salesmanCheck.success) {
            return @{
                success = $false
                httpStatus = 500
                errorCode = 'SALESMAN_VALIDATION_FAILED'
                error = if ($salesmanCheck.error) { [string]$salesmanCheck.error } else { 'Could not validate Salesman assignment.' }
            }
        }

        if (-not $salesmanCheck.allowed) {
            return @{
                success = $false
                httpStatus = 403
                errorCode = if ($salesmanCheck.errorCode) { [string]$salesmanCheck.errorCode } else { 'SALESMAN_ACCESS_DENIED' }
                error = if ($salesmanCheck.error) { [string]$salesmanCheck.error } else { 'Salesman access denied.' }
            }
        }

        if ($null -ne $salesmanCheck.data) {
            $Payload = $salesmanCheck.data
        }
    }

    # The user asked not to apply normal permission restrictions to
    # superuser/admin accounts. Existing LIVE behavior already gives admins a
    # bypass; we make that explicit here and still let Create-Voucher perform
    # normal BUSY/business validation.
    if ($isAdmin) {
        return @{ success = $true; data = $Payload; adminBypass = $true }
    }

    if (Get-Command Test-VoucherPartyAccountAccess -ErrorAction SilentlyContinue) {
        $partyCheck = Test-VoucherPartyAccountAccess `
            -AuthResult $AuthResult `
            -Data $Payload `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode `
            -RequireAuth $true

        if (-not $partyCheck.success) {
            return @{ success = $false; httpStatus = 500; errorCode = 'PARTY_ACCOUNT_VALIDATION_FAILED'; error = if ($partyCheck.error) { [string]$partyCheck.error } else { 'Could not validate Party Account Access.' } }
        }

        if (-not $partyCheck.allowed) {
            return @{ success = $false; httpStatus = 403; errorCode = 'PARTY_ACCOUNT_ACCESS_DENIED'; error = 'The selected Party/account is not allowed for this user and voucher type.'; deniedAccounts = @($partyCheck.deniedAccounts) }
        }
    }

    if (Get-Command Test-VoucherDebitCreditAccountAccess -ErrorAction SilentlyContinue) {
        $dcCheck = Test-VoucherDebitCreditAccountAccess `
            -AuthResult $AuthResult `
            -Data $Payload `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode `
            -RequireAuth $true

        if (-not $dcCheck.success) {
            return @{ success = $false; httpStatus = 500; errorCode = 'ACCOUNT_SIDE_ACCESS_VALIDATION_FAILED'; error = if ($dcCheck.error) { [string]$dcCheck.error } else { 'Could not validate Debit/Credit Account Access.' } }
        }

        if (-not $dcCheck.allowed) {
            return @{ success = $false; httpStatus = 403; errorCode = 'ACCOUNT_SIDE_ACCESS_DENIED'; error = 'One or more accounts are not allowed for their Debit/Credit side.'; deniedAccounts = @($dcCheck.deniedAccounts) }
        }
    }

    if (Get-Command Test-VoucherItemGroupAccess -ErrorAction SilentlyContinue) {
        $itemCheck = Test-VoucherItemGroupAccess `
            -AuthResult $AuthResult `
            -Data $Payload `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode `
            -RequireAuth $true

        if (-not $itemCheck.success) {
            return @{ success = $false; httpStatus = 500; errorCode = 'ITEM_GROUP_ACCESS_VALIDATION_FAILED'; error = if ($itemCheck.error) { [string]$itemCheck.error } else { 'Could not validate Item Group Access.' } }
        }

        if (-not $itemCheck.allowed) {
            return @{ success = $false; httpStatus = 403; errorCode = 'ITEM_GROUP_ACCESS_DENIED'; error = 'One or more Items are outside the Item Group branches allowed for this user and voucher type.'; deniedItems = @($itemCheck.deniedItems) }
        }
    }

    return @{ success = $true; data = $Payload; adminBypass = $false }
}

# -----------------------------------------------------------------------------
# Main orchestration function called by routes.ps1
# -----------------------------------------------------------------------------

function Invoke-OfflineVoucherSync {
    param(
        [Parameter(Mandatory = $true)]$Data,
        [Parameter(Mandatory = $true)]$AuthResult,
        [Parameter(Mandatory = $true)][string]$InstanceId,
        [Parameter(Mandatory = $true)][string]$CompanyCode
    )

    $validation = Test-OfflineVoucherSyncRequest -Data $Data
    if (-not $validation.success) {
        return $validation
    }

    if (
        $null -eq $AuthResult -or
        $null -eq $AuthResult.user -or
        [string]::IsNullOrWhiteSpace([string]$AuthResult.user.name)
    ) {
        return @{
            success = $false
            httpStatus = 401
            errorCode = 'AUTH_REQUIRED'
            error = 'An authenticated BUSY user is required to synchronize a local voucher.'
        }
    }

    if ([string]::IsNullOrWhiteSpace($InstanceId) -or [string]::IsNullOrWhiteSpace($CompanyCode)) {
        return @{
            success = $false
            httpStatus = 400
            errorCode = 'BUSY_CONTEXT_REQUIRED'
            error = 'Current BUSY instance and company are required.'
        }
    }

    $userName = ([string]$AuthResult.user.name).Trim()
    $requestHash = Get-OfflineSyncRequestHash -Data $Data
    $processingToken = [guid]::NewGuid().ToString('N')

    $ctx = $null
    $connection = $null
    $isSqlServer = $false
    $ledgerClaimed = $false

    try {
        if (-not (Get-Command Get-BusyCloudFastConfigDbContext -ErrorAction SilentlyContinue)) {
            return @{
                success = $false
                httpStatus = 500
                errorCode = 'OFFLINE_SYNC_DB_RESOLVER_NOT_LOADED'
                error = 'Direct BUSY company database resolver is not loaded.'
            }
        }

        $ctx = Get-BusyCloudFastConfigDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if ($null -eq $ctx -or $null -eq $ctx.connection) {
            return @{
                success = $false
                httpStatus = 500
                errorCode = 'BUSY_DATABASE_UNAVAILABLE'
                error = 'Could not open the current BUSY company database.'
            }
        }

        $connection = $ctx.connection
        $isSqlServer = ([int]$ctx.dbType -eq 1)

        Ensure-OfflineSyncLedger `
            -Connection $connection `
            -IsSqlServer $isSqlServer

        $record = Get-OfflineSyncLedgerRecord `
            -Connection $connection `
            -IsSqlServer $isSqlServer `
            -LocalId ([string]$Data.localId).Trim() `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if ($null -eq $record) {
            try {
                [void](Add-OfflineSyncLedgerRecord `
                    -Connection $connection `
                    -IsSqlServer $isSqlServer `
                    -Data $Data `
                    -InstanceId $InstanceId `
                    -CompanyCode $CompanyCode `
                    -UserName $userName `
                    -RequestHash $requestHash `
                    -ProcessingToken $processingToken)

                $ledgerClaimed = $true
            }
            catch {
                # Most commonly this is a unique-key race. Re-read and treat it
                # exactly like any other existing idempotency record.
                $record = Get-OfflineSyncLedgerRecord `
                    -Connection $connection `
                    -IsSqlServer $isSqlServer `
                    -LocalId ([string]$Data.localId).Trim() `
                    -InstanceId $InstanceId `
                    -CompanyCode $CompanyCode

                if ($null -eq $record) {
                    throw
                }
            }
        }

        if (-not $ledgerClaimed) {
            $claim = Claim-OfflineSyncExistingRecord `
                -Connection $connection `
                -IsSqlServer $isSqlServer `
                -Record $record `
                -Data $Data `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode `
                -RequestHash $requestHash `
                -ProcessingToken $processingToken

            if ($claim.success -and $claim.alreadyProcessed) {
                return $claim
            }

            if (-not $claim.success -and $claim.stalePosting) {
                # Safe crash recovery: if the previous attempt had already
                # entered POSTING_TO_BUSY, first look for the exact candidate
                # voucher before refusing a retry.
                $postingNo = ([string]$record.PostingVchNo).Trim()
                $series = ([string]$record.VchSeries).Trim()

                if (-not [string]::IsNullOrWhiteSpace($postingNo)) {
                    $found = Find-OfflinePostedVoucher `
                        -Connection $connection `
                        -IsSqlServer $isSqlServer `
                        -VchType ([int]$record.VchType) `
                        -VchNo $postingNo `
                        -SeriesName $series

                    if ($null -ne $found) {
                        $foundCode = 0
                        try { $foundCode = [int]$found.VchCode } catch {}

                        # A Tran1 header alone is not enough. A partially-created
                        # BUSY voucher can have a header while its item lines / amount
                        # are missing. Verify the full voucher before marking SYNCED.
                        $integrity = Test-OfflineVoucherIntegrity `
                            -Connection $connection `
                            -IsSqlServer $isSqlServer `
                            -Header $found `
                            -ExpectedPayload $Data.payload `
                            -VoucherDate ([string]$validation.voucherDate) `
                            -ExpectedSeries $series `
                            -InstanceId $InstanceId `
                            -CompanyCode $CompanyCode

                        if ($integrity.complete -eq $true) {
                            [void](Set-OfflineSyncLedgerSynced `
                                -Connection $connection `
                                -IsSqlServer $isSqlServer `
                                -Data $Data `
                                -InstanceId $InstanceId `
                                -CompanyCode $CompanyCode `
                                -FinalVchNo ([string]$found.VchNo).Trim() `
                                -FinalVchCode $foundCode `
                                -WarningMessage 'Recovered after an interrupted synchronization. BUSY voucher contents were verified before returning success.')

                            return @{
                                success = $true
                                alreadyProcessed = $true
                                recovered = $true
                                localId = ([string]$Data.localId).Trim()
                                finalVchNo = ([string]$found.VchNo).Trim()
                                finalVchCode = if ($foundCode -gt 0) { $foundCode } else { $null }
                                syncedAt = (Get-Date).ToString('o')
                                warning = 'Recovered a previously created BUSY voucher after an interrupted synchronization.'
                                diagnostics = $integrity
                            }
                        }

                        $reasonText = (@($integrity.reasons) -join ' ')
                        $message = "A previous synchronization created BUSY voucher '$postingNo', but its contents are incomplete or do not match the local transaction. $reasonText"

                        [void](Set-OfflineSyncLedgerReviewRequired `
                            -Connection $connection `
                            -IsSqlServer $isSqlServer `
                            -Data $Data `
                            -InstanceId $InstanceId `
                            -CompanyCode $CompanyCode `
                            -ErrorMessage $message)

                        return @{
                            success = $false
                            conflict = $true
                            reviewRequired = $true
                            httpStatus = 409
                            errorCode = 'OFFLINE_SYNC_REVIEW_INTEGRITY_MISMATCH'
                            error = $message
                            diagnostics = $integrity
                        }
                    }
                }

                $message = 'A previous synchronization stopped while BUSY creation may have been in progress. No automatic second create will be attempted. Reconciliation is required.'

                [void](Set-OfflineSyncLedgerReviewRequired `
                    -Connection $connection `
                    -IsSqlServer $isSqlServer `
                    -Data $Data `
                    -InstanceId $InstanceId `
                    -CompanyCode $CompanyCode `
                    -ErrorMessage $message)

                return @{
                    success = $false
                    conflict = $true
                    reviewRequired = $true
                    httpStatus = 409
                    errorCode = 'OFFLINE_SYNC_REVIEW_REQUIRED'
                    error = $message
                }
            }

            if (-not $claim.success) {
                return $claim
            }

            $ledgerClaimed = $true
        }

        # -----------------------------------------------------------------
        # Build a fresh posting payload from the immutable local snapshot.
        # -----------------------------------------------------------------
        $payload = Copy-OfflineSyncObject -Object $Data.payload

        $payload | Add-Member `
            -MemberType NoteProperty `
            -Name 'vchType' `
            -Value 9 `
            -Force

        $payload | Add-Member `
            -MemberType NoteProperty `
            -Name 'vchSeries' `
            -Value ([string]$validation.series) `
            -Force

        $payload | Add-Member `
            -MemberType NoteProperty `
            -Name 'date' `
            -Value ([string]$validation.voucherDate) `
            -Force

        $payload | Add-Member `
            -MemberType NoteProperty `
            -Name 'bridgeUserName' `
            -Value $userName `
            -Force

        # Current access checks. Admin/superuser permission restrictions are
        # intentionally bypassed; normal users are revalidated against current
        # server-side permissions at sync time.
        $accessResult = Invoke-OfflineVoucherCurrentAccessChecks `
            -AuthResult $AuthResult `
            -Payload $payload `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (-not $accessResult.success) {
            [void](Set-OfflineSyncLedgerFailed `
                -Connection $connection `
                -IsSqlServer $isSqlServer `
                -Data $Data `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode `
                -ProcessingToken $processingToken `
                -ErrorMessage ([string]$accessResult.error) `
                -CanRetry $true)

            return $accessResult
        }

        $payload = Normalize-OfflineVoucherPostingPayload `
            -Payload $accessResult.data

        Write-Host (
            "  [OFFLINE-SYNC ROWS] normalized items={0} sundries={1}" -f `
            @($payload.items).Count,
            @($payload.billSundries).Count
        ) -ForegroundColor DarkCyan

        [void](Set-OfflineSyncLedgerPhase `
            -Connection $connection `
            -IsSqlServer $isSqlServer `
            -Data $Data `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode `
            -ProcessingToken $processingToken `
            -Phase 'NUMBERING')

        # Resolve numbering immediately before posting. For auto-number series,
        # the stale number stored in IndexedDB is deliberately replaced by the
        # current next number from the same numbering engine used by the app.
        $numberResult = Resolve-OfflineVoucherPostingNumber `
            -Data $Data `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode `
            -SeriesName ([string]$validation.series) `
            -VoucherDate ([string]$validation.voucherDate)

        if (-not $numberResult.success) {
            [void](Set-OfflineSyncLedgerFailed `
                -Connection $connection `
                -IsSqlServer $isSqlServer `
                -Data $Data `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode `
                -ProcessingToken $processingToken `
                -ErrorMessage ([string]$numberResult.error) `
                -CanRetry $true)

            return $numberResult
        }

        $postingVchNo = ([string]$numberResult.vchNo).Trim()

        $payload | Add-Member `
            -MemberType NoteProperty `
            -Name 'vchNo' `
            -Value $postingVchNo `
            -Force

        # IMPORTANT CRASH BOUNDARY:
        # Persist the exact candidate number and POSTING_TO_BUSY state BEFORE
        # calling Create-Voucher. A retry after a process crash will reconcile
        # this exact number rather than blindly creating another voucher.
        $phaseRows = Set-OfflineSyncLedgerPhase `
            -Connection $connection `
            -IsSqlServer $isSqlServer `
            -Data $Data `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode `
            -ProcessingToken $processingToken `
            -Phase 'POSTING_TO_BUSY' `
            -PostingVchNo $postingVchNo

        if ($phaseRows -ne 1) {
            return @{
                success = $false
                conflict = $true
                httpStatus = 409
                errorCode = 'OFFLINE_SYNC_LOST_OWNERSHIP'
                error = 'This synchronization no longer owns the server-side processing record.'
            }
        }

        if (-not (Get-Command Create-Voucher -ErrorAction SilentlyContinue)) {
            [void](Set-OfflineSyncLedgerFailed `
                -Connection $connection `
                -IsSqlServer $isSqlServer `
                -Data $Data `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode `
                -ProcessingToken $processingToken `
                -ErrorMessage 'Create-Voucher is not loaded.' `
                -CanRetry $true)

            return @{ success = $false; httpStatus = 500; errorCode = 'CREATE_VOUCHER_NOT_LOADED'; error = 'Create-Voucher is not loaded.' }
        }

        $createResult = $null

        try {
            $createResult = Create-Voucher `
                -Data $payload `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode
        }
        catch {
            # At this point BUSY creation may have happened before the exception.
            # Never downgrade this directly to retryable FAILED. First reconcile
            # the exact candidate voucher number.
            $foundAfterException = Find-OfflinePostedVoucher `
                -Connection $connection `
                -IsSqlServer $isSqlServer `
                -VchType 9 `
                -VchNo $postingVchNo `
                -SeriesName ([string]$validation.series)

            if ($null -ne $foundAfterException) {
                $foundCode = 0
                try { $foundCode = [int]$foundAfterException.VchCode } catch {}

                $integrityAfterException = Test-OfflineVoucherIntegrity `
                    -Connection $connection `
                    -IsSqlServer $isSqlServer `
                    -Header $foundAfterException `
                    -ExpectedPayload $payload `
                    -VoucherDate ([string]$validation.voucherDate) `
                    -ExpectedSeries ([string]$validation.series) `
                    -InstanceId $InstanceId `
                    -CompanyCode $CompanyCode

                if ($integrityAfterException.complete -eq $true) {
                    [void](Set-OfflineSyncLedgerSynced `
                        -Connection $connection `
                        -IsSqlServer $isSqlServer `
                        -Data $Data `
                        -InstanceId $InstanceId `
                        -CompanyCode $CompanyCode `
                        -ProcessingToken $processingToken `
                        -FinalVchNo ([string]$foundAfterException.VchNo).Trim() `
                        -FinalVchCode $foundCode `
                        -WarningMessage "Create-Voucher raised an error after BUSY creation, but the complete voucher was confirmed: $($_.Exception.Message)")

                    return @{
                        success = $true
                        created = $true
                        recovered = $true
                        localId = ([string]$Data.localId).Trim()
                        finalVchNo = ([string]$foundAfterException.VchNo).Trim()
                        finalVchCode = if ($foundCode -gt 0) { $foundCode } else { $null }
                        syncedAt = (Get-Date).ToString('o')
                        warning = $_.Exception.Message
                        diagnostics = $integrityAfterException
                    }
                }

                $reasonText = (@($integrityAfterException.reasons) -join ' ')
                $reviewMessage = "Create-Voucher raised an exception and BUSY voucher '$postingVchNo' exists, but its contents did not pass integrity verification. $reasonText Error: $($_.Exception.Message)"

                [void](Set-OfflineSyncLedgerReviewRequired `
                    -Connection $connection `
                    -IsSqlServer $isSqlServer `
                    -Data $Data `
                    -InstanceId $InstanceId `
                    -CompanyCode $CompanyCode `
                    -ProcessingToken $processingToken `
                    -ErrorMessage $reviewMessage)

                return @{
                    success = $false
                    conflict = $true
                    reviewRequired = $true
                    httpStatus = 409
                    errorCode = 'OFFLINE_SYNC_REVIEW_INTEGRITY_MISMATCH'
                    error = $reviewMessage
                    diagnostics = $integrityAfterException
                }
            }

            $reviewMessage = "Create-Voucher raised an exception after the server entered the BUSY posting phase. A second create is blocked until reconciliation. Error: $($_.Exception.Message)"

            [void](Set-OfflineSyncLedgerReviewRequired `
                -Connection $connection `
                -IsSqlServer $isSqlServer `
                -Data $Data `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode `
                -ProcessingToken $processingToken `
                -ErrorMessage $reviewMessage)

            return @{
                success = $false
                conflict = $true
                reviewRequired = $true
                httpStatus = 409
                errorCode = 'OFFLINE_SYNC_REVIEW_REQUIRED'
                error = $reviewMessage
            }
        }

        [void](Set-OfflineSyncLedgerPhase `
            -Connection $connection `
            -IsSqlServer $isSqlServer `
            -Data $Data `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode `
            -ProcessingToken $processingToken `
            -Phase 'COMPLETING' `
            -PostingVchNo $postingVchNo)

        $resultInfo = Get-OfflineCreateResultInfo -Result $createResult
        $candidateToCheck = if (-not [string]::IsNullOrWhiteSpace([string]$resultInfo.vchNo)) {
            [string]$resultInfo.vchNo
        }
        else {
            $postingVchNo
        }

        # Prefer the VchCode returned by Create-Voucher. It is a much stronger
        # identity than voucher number + series text.
        $confirmedVoucher = Find-OfflinePostedVoucher `
            -Connection $connection `
            -IsSqlServer $isSqlServer `
            -VchType 9 `
            -VchNo $candidateToCheck `
            -SeriesName ([string]$validation.series) `
            -VchCode ([int]$resultInfo.vchCode)

        # If the strict series lookup did not find the voucher and Create-Voucher
        # did not return a usable VchCode, fall back to the exact VchNo and let
        # the integrity checker verify party/date/series/items/amount.
        if ($null -eq $confirmedVoucher) {
            $confirmedVoucher = Find-OfflinePostedVoucher `
                -Connection $connection `
                -IsSqlServer $isSqlServer `
                -VchType 9 `
                -VchNo $candidateToCheck
        }

        if ($null -ne $confirmedVoucher) {
            $confirmedCode = 0
            try { $confirmedCode = [int]$confirmedVoucher.VchCode } catch {}

            $integrity = Test-OfflineVoucherIntegrity `
                -Connection $connection `
                -IsSqlServer $isSqlServer `
                -Header $confirmedVoucher `
                -ExpectedPayload $payload `
                -VoucherDate ([string]$validation.voucherDate) `
                -ExpectedSeries ([string]$validation.series) `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode

            if (-not $integrity.complete) {
                $reasonText = (@($integrity.reasons) -join ' ')
                if ([string]::IsNullOrWhiteSpace($reasonText)) {
                    $reasonText = 'The created BUSY voucher could not be verified as complete.'
                }

                $reviewMessage = "BUSY voucher '$candidateToCheck' was found after creation, but its saved contents did not pass integrity verification. $reasonText"

                [void](Set-OfflineSyncLedgerReviewRequired `
                    -Connection $connection `
                    -IsSqlServer $isSqlServer `
                    -Data $Data `
                    -InstanceId $InstanceId `
                    -CompanyCode $CompanyCode `
                    -ProcessingToken $processingToken `
                    -ErrorMessage $reviewMessage)

                return @{
                    success = $false
                    conflict = $true
                    reviewRequired = $true
                    created = $true
                    httpStatus = 409
                    errorCode = 'OFFLINE_SYNC_REVIEW_INTEGRITY_MISMATCH'
                    error = $reviewMessage
                    diagnostics = $integrity
                }
            }

            $warning = ''
            if (-not $resultInfo.success) {
                $warning = if (-not [string]::IsNullOrWhiteSpace([string]$resultInfo.error)) {
                    [string]$resultInfo.error
                }
                else {
                    'BUSY voucher was created, but post-create processing returned a warning/error.'
                }
            }

            [void](Set-OfflineSyncLedgerSynced `
                -Connection $connection `
                -IsSqlServer $isSqlServer `
                -Data $Data `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode `
                -ProcessingToken $processingToken `
                -FinalVchNo ([string]$confirmedVoucher.VchNo).Trim() `
                -FinalVchCode $confirmedCode `
                -WarningMessage $warning)

            return @{
                success = $true
                created = $true
                alreadyProcessed = $false
                localId = ([string]$Data.localId).Trim()
                finalVchNo = ([string]$confirmedVoucher.VchNo).Trim()
                finalVchCode = if ($confirmedCode -gt 0) { $confirmedCode } else { $null }
                syncedAt = (Get-Date).ToString('o')
                warning = if ([string]::IsNullOrWhiteSpace($warning)) { $null } else { $warning }
                diagnostics = $integrity
            }
        }

        # If Create-Voucher explicitly says the voucher was created, but direct
        # reconciliation cannot find it, automatic retry is still unsafe.
        if ($resultInfo.created -or $resultInfo.success) {
            $reviewMessage = if (-not [string]::IsNullOrWhiteSpace([string]$resultInfo.error)) {
                "BUSY creation may have succeeded but could not be reconciled automatically. $($resultInfo.error)"
            }
            else {
                'BUSY creation reported success/created, but the voucher could not be reconciled automatically. A second create is blocked.'
            }

            [void](Set-OfflineSyncLedgerReviewRequired `
                -Connection $connection `
                -IsSqlServer $isSqlServer `
                -Data $Data `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode `
                -ProcessingToken $processingToken `
                -ErrorMessage $reviewMessage)

            return @{
                success = $false
                conflict = $true
                reviewRequired = $true
                created = [bool]$resultInfo.created
                httpStatus = 409
                errorCode = 'OFFLINE_SYNC_REVIEW_REQUIRED'
                error = $reviewMessage
            }
        }

        # Normal pre-creation rejection: safe to retry later. The next attempt
        # will resolve numbering again from current server state.
        $createError = if (-not [string]::IsNullOrWhiteSpace([string]$resultInfo.error)) {
            [string]$resultInfo.error
        }
        else {
            'BUSY rejected the voucher before creation.'
        }

        [void](Set-OfflineSyncLedgerFailed `
            -Connection $connection `
            -IsSqlServer $isSqlServer `
            -Data $Data `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode `
            -ProcessingToken $processingToken `
            -ErrorMessage $createError `
            -CanRetry $true)

        return @{
            success = $false
            httpStatus = 400
            errorCode = 'BUSY_VOUCHER_CREATE_FAILED'
            error = $createError
            retryable = $true
        }
    }
    catch {
        # If we own a ledger row and an unexpected error happened BEFORE
        # POSTING_TO_BUSY, make it retryable. If the row already reached posting,
        # protect against duplicates by moving it to REVIEW_REQUIRED.
        if ($ledgerClaimed -and $null -ne $connection) {
            try {
                $current = Get-OfflineSyncLedgerRecord `
                    -Connection $connection `
                    -IsSqlServer $isSqlServer `
                    -LocalId ([string]$Data.localId).Trim() `
                    -InstanceId $InstanceId `
                    -CompanyCode $CompanyCode

                $phase = if ($null -ne $current) { ([string]$current.Phase).Trim().ToUpperInvariant() } else { '' }

                if ($phase -in @('POSTING_TO_BUSY', 'COMPLETING')) {
                    [void](Set-OfflineSyncLedgerReviewRequired `
                        -Connection $connection `
                        -IsSqlServer $isSqlServer `
                        -Data $Data `
                        -InstanceId $InstanceId `
                        -CompanyCode $CompanyCode `
                        -ProcessingToken $processingToken `
                        -ErrorMessage "Unexpected server error after entering BUSY posting phase. Automatic retry blocked: $($_.Exception.Message)")
                }
                else {
                    [void](Set-OfflineSyncLedgerFailed `
                        -Connection $connection `
                        -IsSqlServer $isSqlServer `
                        -Data $Data `
                        -InstanceId $InstanceId `
                        -CompanyCode $CompanyCode `
                        -ProcessingToken $processingToken `
                        -ErrorMessage $_.Exception.Message `
                        -CanRetry $true)
                }
            }
            catch {
            }
        }

        return @{
            success = $false
            httpStatus = 500
            errorCode = 'OFFLINE_SYNC_SERVER_ERROR'
            error = $_.Exception.Message
        }
    }
    finally {
        if ($null -ne $connection) {
            try { $connection.Close() } catch {}
            try { $connection.Dispose() } catch {}
        }
    }
}
