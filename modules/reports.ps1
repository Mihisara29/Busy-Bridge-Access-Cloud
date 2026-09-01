
# ─────────────────────────────────────────────────────────────
# REPORT ENGINE DIAGNOSTICS AND CACHE
# ─────────────────────────────────────────────────────────────
if ($null -eq $script:OutstandingReportCache) {
    $script:OutstandingReportCache = @{}
}

if ($null -eq $script:StockStatusReportCache) {
    $script:StockStatusReportCache = @{}
}

$script:OutstandingReportCacheSeconds = 20

# Stock Status is much more expensive than Outstanding because BUSY method-5
# valuation must replay inventory history chronologically. Keep a short-lived
# in-process result cache so revisiting the page or repeating the same filter
# does not replay thousands of rows again.
$script:StockStatusReportCacheSeconds = 1800
$script:StockStatusReportCacheMaxEntries = 12

$script:ReportDebugEnabled = $true

function Write-ReportEngineLog {
    param(
        [string]$Stage,
        [string]$Message = "",
        [string]$Color = "Cyan"
    )

    if (-not $script:ReportDebugEnabled) {
        return
    }

    try {
        $timestamp = (Get-Date).ToString("HH:mm:ss.fff")
        Write-Host (
            "[REPORT-ENGINE][$timestamp][$Stage] $Message"
        ) -ForegroundColor $Color
    }
    catch {}
}

function New-ReportEngineWatch {
    try {
        return [System.Diagnostics.Stopwatch]::StartNew()
    }
    catch {
        return $null
    }
}

function Stop-ReportEngineWatch {
    param(
        $Watch,
        [string]$Stage,
        [string]$Message = ""
    )

    try {
        if ($null -ne $Watch) {
            $Watch.Stop()
            $elapsed = [Math]::Round(
                $Watch.Elapsed.TotalMilliseconds,
                1
            )

            Write-ReportEngineLog `
                -Stage $Stage `
                -Message "$Message (${elapsed} ms)" `
                -Color "Green"
        }
    }
    catch {}
}

function Get-ReportEngineMemory {
    try {
        $process =
            [System.Diagnostics.Process]::GetCurrentProcess()

        return @{
            workingSetMB = [Math]::Round(
                $process.WorkingSet64 / 1MB,
                1
            )
            privateMemoryMB = [Math]::Round(
                $process.PrivateMemorySize64 / 1MB,
                1
            )
        }
    }
    catch {
        return @{
            workingSetMB = 0
            privateMemoryMB = 0
        }
    }
}

function Clear-OutstandingReportCache {
    param(
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ($null -eq $script:OutstandingReportCache) {
        return
    }

    $prefix = "$InstanceId|$CompanyCode|"

    foreach (
        $key in @(
            $script:OutstandingReportCache.Keys
        )
    ) {
        if (
            [string]::IsNullOrWhiteSpace($InstanceId) -or
            ([string]$key).StartsWith($prefix)
        ) {
            $script:OutstandingReportCache.Remove($key)
        }
    }
}

function Clear-StockStatusReportCache {
    param(
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ($null -eq $script:StockStatusReportCache) {
        return
    }

    foreach ($key in @($script:StockStatusReportCache.Keys)) {
        if ([string]::IsNullOrWhiteSpace($InstanceId)) {
            $script:StockStatusReportCache.Remove($key)
            continue
        }

        $entry = $script:StockStatusReportCache[$key]
        $entryInstance = ""
        $entryCompany = ""

        try { $entryInstance = [string]$entry.instanceId } catch {}
        try { $entryCompany = [string]$entry.companyCode } catch {}

        if (
            $entryInstance -eq $InstanceId -and
            (
                [string]::IsNullOrWhiteSpace($CompanyCode) -or
                $entryCompany -eq $CompanyCode
            )
        ) {
            $script:StockStatusReportCache.Remove($key)
        }
    }
}

function Remove-ExpiredStockStatusReportCache {
    if ($null -eq $script:StockStatusReportCache) {
        return
    }

    $now = Get-Date

    foreach ($key in @($script:StockStatusReportCache.Keys)) {
        $entry = $script:StockStatusReportCache[$key]

        $createdAt = $null
        try { $createdAt = [datetime]$entry.createdAt } catch {}

        if (
            $null -eq $createdAt -or
            ($now - $createdAt).TotalSeconds -gt
            $script:StockStatusReportCacheSeconds
        ) {
            $script:StockStatusReportCache.Remove($key)
        }
    }
}


# Keep the expensive Stock Status cache correct when the existing voucher
# engine invalidates its normal stock/item caches after create/modify/delete.
# reports.ps1 is loaded after vouchers.ps1, so we can safely wrap the already
# loaded Clear-StockCaches function without changing vouchers.ps1.
if (
    $null -eq $script:BaseClearStockCachesForReports -and
    $null -ne (Get-Command Clear-StockCaches -ErrorAction SilentlyContinue)
) {
    $script:BaseClearStockCachesForReports = ${function:Clear-StockCaches}
}

if ($null -ne $script:BaseClearStockCachesForReports) {
    function Clear-StockCaches {
        param(
            [string]$InstanceId = "",
            [string]$CompanyCode = ""
        )

        try {
            & $script:BaseClearStockCachesForReports `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode
        }
        catch {}

        try {
            Clear-StockStatusReportCache `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode
        }
        catch {}
    }
}


# modules/reports.ps1
# Professional Outstanding and Stock Status Reports
# Compatible with BUSY Access and SQL Server database modes.

. "$PSScriptRoot\connection.ps1"
. "$PSScriptRoot\utils.ps1"

function ConvertTo-ReportDate {
    param(
        [string]$Value,
        [datetime]$Fallback = (Get-Date)
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return $Fallback.Date }

    $formats = @(
        "yyyy-MM-dd", "dd-MM-yyyy", "dd/MM/yyyy",
        "MM/dd/yyyy", "yyyy/MM/dd"
    )

    foreach ($format in $formats) {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParseExact(
            $Value.Trim(),
            $format,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None,
            [ref]$parsed
        )) {
            return $parsed.Date
        }
    }

    $fallbackParsed = [datetime]::MinValue
    if ([datetime]::TryParse($Value, [ref]$fallbackParsed)) {
        return $fallbackParsed.Date
    }

    return $Fallback.Date
}

function Get-ReportSqlDateLiteral {
    param(
        [datetime]$Date,
        [bool]$IsSql
    )

    $value = $Date.ToString("yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
    if ($IsSql) { return "'$value'" }
    return "#$($Date.ToString('MM/dd/yyyy', [System.Globalization.CultureInfo]::InvariantCulture))#"
}

function Get-ReportDatabaseMode {
    param([string]$InstanceId)

    $instance = Get-InstanceConfig -InstanceId $InstanceId
    $dbType = 0
    if ($null -ne $instance -and $null -ne $instance.dbType) {
        $dbType = [int]$instance.dbType
    }

    return @{
        dbType = $dbType
        isSql  = ($dbType -eq 1)
    }
}

function Get-ReportFieldValue {
    param($Recordset, [string]$Name, $Default = $null)

    try {
        $value = $Recordset.Fields.Item($Name).Value
        if ($value -eq [System.DBNull]::Value -or $null -eq $value) { return $Default }
        return $value
    } catch {
        return $Default
    }
}

function ConvertTo-ReportDouble {
    param($Value)

    if ($null -eq $Value -or $Value -eq [System.DBNull]::Value) { return 0.0 }

    $number = 0.0
    [double]::TryParse(
        $Value.ToString(),
        [System.Globalization.NumberStyles]::Any,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$number
    ) | Out-Null

    return $number
}

function ConvertTo-ReportInt {
    param($Value)

    if ($null -eq $Value -or $Value -eq [System.DBNull]::Value) { return 0 }

    $number = 0
    [int]::TryParse($Value.ToString(), [ref]$number) | Out-Null
    return $number
}

function Close-ReportRecordset {
    param($Recordset)
    if ($null -ne $Recordset) {
        try { $Recordset.Close() } catch {}
    }
}

function Get-OutstandingAgeBucket {
    param([int]$DaysOverdue, [bool]$IsNotDue)

    if ($IsNotDue) { return "notDue" }
    if ($DaysOverdue -le 30)  { return "days0To30" }
    if ($DaysOverdue -le 60)  { return "days31To60" }
    if ($DaysOverdue -le 90)  { return "days61To90" }
    if ($DaysOverdue -le 180) { return "days91To180" }
    return "above180"
}


function Get-ReportPagedMasterCodes {
    param(
        $Fi,
        [int]$MasterType,
        [string]$Search = "",
        [string]$Group = "",
        [int]$Page = 1,
        [int]$PageSize = 50
    )

    if ($Page -lt 1) { $Page = 1 }
    if ($PageSize -lt 1) { $PageSize = 50 }
    if ($PageSize -gt 50) { $PageSize = 50 }

    $where = "M.MasterType = $MasterType"

    if (-not [string]::IsNullOrWhiteSpace($Search)) {
        $safe = $Search.Trim() -replace "'", "''"
        $where += " AND (M.Name LIKE '*$safe*' OR M.Alias LIKE '*$safe*')"
    }

    if (-not [string]::IsNullOrWhiteSpace($Group) -and $Group -ne "all") {
        $safeGroup = $Group.Trim() -replace "'", "''"
        $where += " AND G.Name = '$safeGroup'"
    }

    $offset = ($Page - 1) * $PageSize

    $countQry = @"
SELECT COUNT(*) AS TotalRows
FROM Master1 M
LEFT JOIN Master1 G ON G.Code = M.ParentGrp
WHERE $where
"@

    $countRst = $Fi.GetRecordset($countQry)
    $total = 0
    if ($countRst -and -not $countRst.EOF) {
        $total = ConvertTo-ReportInt (
            Get-ReportFieldValue $countRst "TotalRows" 0
        )
    }
    Close-ReportRecordset $countRst

    if ($offset -eq 0) {
        $pageQry = @"
SELECT TOP $PageSize
    M.Code,
    M.Name,
    M.Alias,
    M.ParentGrp,
    G.Name AS GroupName
FROM Master1 M
LEFT JOIN Master1 G ON G.Code = M.ParentGrp
WHERE $where
ORDER BY M.Name, M.Code
"@
    }
    else {
        $pageQry = @"
SELECT TOP $PageSize
    M.Code,
    M.Name,
    M.Alias,
    M.ParentGrp,
    G.Name AS GroupName
FROM Master1 M
LEFT JOIN Master1 G ON G.Code = M.ParentGrp
WHERE
    $where
    AND M.Code NOT IN (
        SELECT TOP $offset M2.Code
        FROM Master1 M2
        LEFT JOIN Master1 G2 ON G2.Code = M2.ParentGrp
        WHERE $($where.Replace("M.", "M2.").Replace("G.", "G2."))
        ORDER BY M2.Name, M2.Code
    )
ORDER BY M.Name, M.Code
"@
    }

    $rst = $Fi.GetRecordset($pageQry)
    $rows = [System.Collections.Generic.List[object]]::new()

    if ($rst -and -not $rst.EOF) {
        $rst.MoveFirst()
        while (-not $rst.EOF) {
            $rows.Add([pscustomobject]@{
                code      = ConvertTo-ReportInt (
                    Get-ReportFieldValue $rst "Code" 0
                )
                name      = [string](
                    Get-ReportFieldValue $rst "Name" ""
                )
                alias     = [string](
                    Get-ReportFieldValue $rst "Alias" ""
                )
                groupCode = ConvertTo-ReportInt (
                    Get-ReportFieldValue $rst "ParentGrp" 0
                )
                groupName = [string](
                    Get-ReportFieldValue $rst "GroupName" ""
                )
            })
            $rst.MoveNext()
        }
    }

    Close-ReportRecordset $rst

    return @{
        total = $total
        rows  = @($rows)
    }
}

function Get-OutstandingReport {
    param(
        [string]$From = "",
        [string]$To = "",
        [string]$AsOf = "",

        [string]$Type = "all",
        [string]$Account = "",
        [string]$Search = "",
        [string]$Group = "",
        [string]$Status = "all",
        [string]$Aging = "all",
        [int]$VoucherType = 0,

        [double]$MinAmount = 0,
        [double]$MaxAmount = 0,
        [bool]$IncludeZero = $false,

        [int]$Page = 1,
        [int]$PageSize = 50,

        [string]$SortBy = "dueDate",
        [string]$SortDirection = "asc",

        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $overallWatch = New-ReportEngineWatch
    $stage = "initialisation"

    if ($Page -lt 1) {
        $Page = 1
    }

    # PageSize=0 means "return all matching rows".
    if ($PageSize -lt 0) {
        $PageSize = 500
    }

    if ($PageSize -gt 2500) {
        $PageSize = 2500
    }

    $asOfDate = ConvertTo-ReportDate `
        -Value $AsOf `
        -Fallback (Get-Date)

    $fromDate = if (
        [string]::IsNullOrWhiteSpace($From)
    ) {
        [datetime]::new(
            $asOfDate.Year,
            4,
            1
        )
    }
    else {
        ConvertTo-ReportDate `
            -Value $From `
            -Fallback $asOfDate
    }

    $toDate = if (
        [string]::IsNullOrWhiteSpace($To)
    ) {
        $asOfDate
    }
    else {
        ConvertTo-ReportDate `
            -Value $To `
            -Fallback $asOfDate
    }

    if ($fromDate.Date -gt $toDate.Date) {
        $temporaryDate = $fromDate
        $fromDate = $toDate
        $toDate = $temporaryDate
    }

    if ($toDate.Date -gt $asOfDate.Date) {
        $toDate = $asOfDate
    }

    $instance = Get-InstanceConfig `
        -InstanceId $InstanceId

    $dbType = 0

    if (
        $null -ne $instance -and
        $null -ne $instance.dbType
    ) {
        $dbType = [int]$instance.dbType
    }

    $isSql = ($dbType -eq 1)
    $engineName = if ($isSql) {
        "SQL Server direct"
    }
    else {
        "Access through BUSY"
    }

    Write-ReportEngineLog `
        -Stage "START" `
        -Message (
            "Engine=$engineName; " +
            "Instance=$InstanceId; " +
            "Company=$CompanyCode; " +
            "From=$($fromDate.ToString('yyyy-MM-dd')); " +
            "To=$($toDate.ToString('yyyy-MM-dd')); " +
            "AsOf=$($asOfDate.ToString('yyyy-MM-dd')); " +
            "Type='$Type'; Status='$Status'; Aging='$Aging'; " +
            "Account='$Account'; Search='$Search'; Group='$Group'; " +
            "VoucherType=$VoucherType; MinAmount=$MinAmount; " +
            "MaxAmount=$MaxAmount; SortBy='$SortBy'; " +
            "SortDirection='$SortDirection'; " +
            "Page=$Page; PageSize=$PageSize"
        ) `
        -Color "Yellow"

    $sqlConnection = $null
    $busyConnection = $null

    function Convert-ReportDatabaseValue {
        param($Value)

        if (
            $null -eq $Value -or
            $Value -eq [System.DBNull]::Value
        ) {
            return $null
        }

        return $Value
    }

    function Invoke-OutstandingDatabaseRows {
        param(
            [string]$Query,
            [string]$QueryName,
            [int]$Timeout = 180
        )

        $queryWatch = New-ReportEngineWatch

        Write-ReportEngineLog `
            -Stage "$QueryName-SQL" `
            -Message $Query `
            -Color "DarkGray"

        $rows =
            [System.Collections.Generic.List[object]]::new()

        if ($isSql) {
            $command = $null
            $reader = $null

            try {
                $command = $sqlConnection.CreateCommand()
                $command.CommandText = $Query
                $command.CommandTimeout = $Timeout
                $reader = $command.ExecuteReader()

                while ($reader.Read()) {
                    $output = [ordered]@{}

                    for (
                        $fieldIndex = 0;
                        $fieldIndex -lt $reader.FieldCount;
                        $fieldIndex++
                    ) {
                        $output[
                            $reader.GetName($fieldIndex)
                        ] = Convert-ReportDatabaseValue (
                            $reader.GetValue($fieldIndex)
                        )
                    }

                    $rows.Add(
                        [pscustomobject]$output
                    )
                }
            }
            finally {
                if ($null -ne $reader) {
                    try { $reader.Close() } catch {}
                    try { $reader.Dispose() } catch {}
                }

                if ($null -ne $command) {
                    try { $command.Dispose() } catch {}
                }
            }
        }
        else {
            $recordset = $null

            try {
                $recordset =
                    $busyConnection.GetRecordset($Query)

                if (
                    $recordset -and
                    -not $recordset.EOF
                ) {
                    $recordset.MoveFirst()

                    while (-not $recordset.EOF) {
                        $output = [ordered]@{}

                        for (
                            $fieldIndex = 0;
                            $fieldIndex -lt $recordset.Fields.Count;
                            $fieldIndex++
                        ) {
                            $field =
                                $recordset.Fields.Item(
                                    $fieldIndex
                                )

                            $output[[string]$field.Name] =
                                Convert-ReportDatabaseValue (
                                    $field.Value
                                )
                        }

                        $rows.Add(
                            [pscustomobject]$output
                        )

                        $recordset.MoveNext()
                    }
                }
            }
            finally {
                if ($recordset) {
                    try {
                        $recordset.Close()
                    }
                    catch {}
                }
            }
        }

        Stop-ReportEngineWatch `
            -Watch $queryWatch `
            -Stage $QueryName `
            -Message "Rows=$($rows.Count)"

        return @($rows)
    }

    function Get-OutstandingDateLiteral {
        param([datetime]$Value)

        if ($isSql) {
            return (
                "'" +
                $Value.ToString("yyyy-MM-dd") +
                "'"
            )
        }

        return (
            "#" +
            $Value.ToString("MM/dd/yyyy") +
            "#"
        )
    }

    function Resolve-OutstandingNature {
        param(
            [string]$GroupName,
            [double]$OriginalSigned
        )

        $normalisedGroup =
            ([string]$GroupName).Trim().ToLowerInvariant()

        foreach ($token in @(
            "customer",
            "sundry debtor",
            "debtor",
            "trade debtor",
            "receivable"
        )) {
            if (
                $normalisedGroup.Contains($token)
            ) {
                return "receivable"
            }
        }

        foreach ($token in @(
            "supplier",
            "sundry creditor",
            "creditor",
            "trade creditor",
            "payable"
        )) {
            if (
                $normalisedGroup.Contains($token)
            ) {
                return "payable"
            }
        }

        # Confirmed by the exported BUSY data:
        # sale bill references are negative;
        # purchase bill references are positive.
        if ($OriginalSigned -lt 0) {
            return "receivable"
        }

        return "payable"
    }

    function Get-OutstandingVoucherTypeName {
        param([int]$Value)

        switch ($Value) {
            1  { return "Opening Balance" }
            2  { return "Purchase" }
            3  { return "Sale Return" }
            4  { return "Material Receipt" }
            5  { return "Stock Transfer" }
            8  { return "Stock Journal" }
            9  { return "Sale" }
            10 { return "Purchase Return" }
            11 { return "Material Issue" }
            12 { return "Sale Order" }
            13 { return "Purchase Order" }
            14 { return "Receipt" }
            15 { return "Contra" }
            16 { return "Journal" }
            19 { return "Payment" }
            default { return "Voucher $Value" }
        }
    }

    function Get-OutstandingVoucherTypeShortName {
        param([int]$Value)

        switch ($Value) {
            1  { return "OpBl" }
            2  { return "Purc" }
            3  { return "SlRt" }
            4  { return "MtRc" }
            5  { return "StTr" }
            8  { return "StJr" }
            9  { return "Sale" }
            10 { return "PrRt" }
            11 { return "MtIs" }
            12 { return "SlOd" }
            13 { return "PrOd" }
            14 { return "Rcpt" }
            15 { return "Cntr" }
            16 { return "Jrnl" }
            19 { return "Pymt" }
            default { return "Vch$Value" }
        }
    }

    function Test-OutstandingStatus {
        param(
            [string]$RequestedStatus,
            [string]$CurrentStatus,
            [bool]$PartiallyAdjusted,
            [bool]$Unadjusted
        )

        switch (
            $RequestedStatus.Trim().ToLowerInvariant()
        ) {
            "" {
                return $true
            }

            "all" {
                return $true
            }

            "overdue" {
                return (
                    $CurrentStatus -eq "overdue"
                )
            }

            "due" {
                return (
                    $CurrentStatus -eq "due"
                )
            }

            "not-due" {
                return (
                    $CurrentStatus -eq "not-due"
                )
            }

            "partially-adjusted" {
                return $PartiallyAdjusted
            }

            "unadjusted" {
                return $Unadjusted
            }

            default {
                return $true
            }
        }
    }

    try {
        $stage = "connection"
        $connectionWatch = New-ReportEngineWatch

        if ($isSql) {
            $directConnection = Get-DirectConnection `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode

            if (
                $null -eq $directConnection -or
                $null -eq $directConnection.connection
            ) {
                throw (
                    "Could not build the SQL Server " +
                    "connection for the selected company."
                )
            }

            $sqlConnection =
                $directConnection.connection

            if (
                $sqlConnection.State.ToString() -ne "Open"
            ) {
                $sqlConnection.Open()
            }
        }
        else {
            # The BUSY connection resolves the Access database
            # location dynamically for the selected cloud server
            # and company. No .bds path is hardcoded.
            $busyConnection = Connect-BUSY `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode

            if (-not $busyConnection) {
                throw (
                    "Could not connect to the selected " +
                    "BUSY Access company."
                )
            }
        }

        Stop-ReportEngineWatch `
            -Watch $connectionWatch `
            -Stage "CONNECTION" `
            -Message "Connected using $engineName"

        $asOfLiteral =
            Get-OutstandingDateLiteral $asOfDate

        # Cache raw, database-aggregated rows by company and
        # as-of date. Filters and paging are still applied on
        # every request, but repeated database scans are avoided.
        $cacheKey = (
            "$InstanceId|$CompanyCode|" +
            $asOfDate.ToString("yyyy-MM-dd") +
            "|$dbType"
        )

        $cached = $null

        if (
            $script:OutstandingReportCache.ContainsKey(
                $cacheKey
            )
        ) {
            $candidate =
                $script:OutstandingReportCache[$cacheKey]

            if (
                $null -ne $candidate -and
                $null -ne $candidate.createdAt -and
                (
                    (Get-Date) -
                    [datetime]$candidate.createdAt
                ).TotalSeconds -le
                $script:OutstandingReportCacheSeconds
            ) {
                $cached = $candidate
            }
            else {
                $script:OutstandingReportCache.Remove(
                    $cacheKey
                )
            }
        }

        $cacheHit = ($null -ne $cached)

        if ($cacheHit) {
            Write-ReportEngineLog `
                -Stage "CACHE" `
                -Message "Using cached aggregated rows." `
                -Color "Green"

            $accountRows = @($cached.accountRows)
            $referenceRows = @($cached.referenceRows)
            $onAccountRows = @($cached.onAccountRows)
        }
        else {
            $stage = "account query"

            $accountQuery = @"
SELECT
    A.Code AS AccountCode,
    A.Name AS AccountName,
    A.Alias AS AccountAlias,
    G.Name AS AccountGroup
FROM Master1 A
LEFT JOIN Master1 G
    ON G.Code = A.ParentGrp
WHERE A.MasterType = 2
"@

            $accountRows =
                Invoke-OutstandingDatabaseRows `
                    -Query $accountQuery `
                    -QueryName "ACCOUNT-MASTERS" `
                    -Timeout 60

            $stage = "reference query"

            if ($isSql) {
                $referenceQuery = @"
SELECT
    T.MasterCode1 AS AccountCode,
    T.RefCode,

    MAX(
        CASE
            WHEN T.Method = 1
                THEN T.[No]
            ELSE NULL
        END
    ) AS RefNo,

    MIN(
        CASE
            WHEN T.Method = 1
                THEN T.[Date]
            ELSE NULL
        END
    ) AS RefDate,

    MIN(
        CASE
            WHEN T.Method = 1
                THEN T.DueDate
            ELSE NULL
        END
    ) AS DueDate,

    MAX(
        CASE
            WHEN T.Method = 1
                THEN T.VchCode
            ELSE 0
        END
    ) AS VoucherCode,

    MAX(
        CASE
            WHEN T.Method = 1
                THEN T.VchType
            ELSE 0
        END
    ) AS VoucherType,

    SUM(
        CASE
            WHEN T.Method = 1
                THEN ISNULL(T.Value1, 0)
            ELSE 0
        END
    ) AS OriginalSigned,

    SUM(
        CASE
            WHEN T.Method = 2
                THEN ISNULL(T.Value1, 0)
            ELSE 0
        END
    ) AS AdjustmentSigned,

    SUM(
        CASE
            WHEN T.Method IN (1, 2)
                THEN ISNULL(T.Value1, 0)
            ELSE 0
        END
    ) AS PendingSigned,

    SUM(
        CASE
            WHEN T.Method = 2
                THEN 1
            ELSE 0
        END
    ) AS AdjustmentCount

FROM Tran3 T
INNER JOIN Tran1 V
    ON V.VchCode = T.VchCode

WHERE
    T.RecType = 1
    AND T.Method IN (1, 2)
    AND T.VchType IN (2, 3, 9, 10)
    AND T.RefCode <> 0
    AND ISNULL(T.[No], '') <> ''
    AND T.[Date] <= $asOfLiteral
    AND ISNULL(V.Cancelled, 0) = 0
    AND ISNULL(V.VchCancelled, 0) = 0

GROUP BY
    T.MasterCode1,
    T.RefCode

HAVING
    SUM(
        CASE
            WHEN T.Method IN (1, 2)
                THEN ISNULL(T.Value1, 0)
            ELSE 0
        END
    ) <> 0
"@
            }
            else {
                # Access/Jet SQL version.
                $referenceQuery = @"
SELECT
    T.MasterCode1 AS AccountCode,
    T.RefCode,

    MAX(
        IIF(
            T.Method = 1,
            T.[No],
            NULL
        )
    ) AS RefNo,

    MIN(
        IIF(
            T.Method = 1,
            T.[Date],
            NULL
        )
    ) AS RefDate,

    MIN(
        IIF(
            T.Method = 1,
            T.DueDate,
            NULL
        )
    ) AS DueDate,

    MAX(
        IIF(
            T.Method = 1,
            T.VchCode,
            0
        )
    ) AS VoucherCode,

    MAX(
        IIF(
            T.Method = 1,
            T.VchType,
            0
        )
    ) AS VoucherType,

    SUM(
        IIF(
            T.Method = 1,
            IIF(
                T.Value1 IS NULL,
                0,
                T.Value1
            ),
            0
        )
    ) AS OriginalSigned,

    SUM(
        IIF(
            T.Method = 2,
            IIF(
                T.Value1 IS NULL,
                0,
                T.Value1
            ),
            0
        )
    ) AS AdjustmentSigned,

    SUM(
        IIF(
            T.Method = 1 OR T.Method = 2,
            IIF(
                T.Value1 IS NULL,
                0,
                T.Value1
            ),
            0
        )
    ) AS PendingSigned,

    SUM(
        IIF(
            T.Method = 2,
            1,
            0
        )
    ) AS AdjustmentCount

FROM Tran3 T
INNER JOIN Tran1 V
    ON V.VchCode = T.VchCode

WHERE
    T.RecType = 1
    AND T.Method IN (1, 2)
    AND T.VchType IN (2, 3, 9, 10)
    AND T.RefCode <> 0
    AND T.[No] IS NOT NULL
    AND T.[No] <> ''
    AND T.[Date] <= $asOfLiteral
    AND (
        V.Cancelled = False
        OR V.Cancelled IS NULL
    )
    AND (
        V.VchCancelled = False
        OR V.VchCancelled IS NULL
    )

GROUP BY
    T.MasterCode1,
    T.RefCode

HAVING
    SUM(
        IIF(
            T.Method = 1 OR T.Method = 2,
            IIF(
                T.Value1 IS NULL,
                0,
                T.Value1
            ),
            0
        )
    ) <> 0
"@
            }

            $referenceRows =
                Invoke-OutstandingDatabaseRows `
                    -Query $referenceQuery `
                    -QueryName "BILL-REFERENCES" `
                    -Timeout 180

            $stage = "on-account query"

            if ($isSql) {
                $onAccountQuery = @"
SELECT
    T.MasterCode1 AS AccountCode,
    SUM(ISNULL(T.Value1, 0)) AS OnAccountSigned
FROM Tran3 T
INNER JOIN Tran1 V
    ON V.VchCode = T.VchCode
WHERE
    T.RecType = 1
    AND T.VchType IN (2, 3, 9, 10)
    AND T.[Date] <= $asOfLiteral
    AND (
        T.Method NOT IN (1, 2)
        OR ISNULL(T.RefCode, 0) = 0
        OR ISNULL(T.[No], '') = ''
    )
    AND ISNULL(V.Cancelled, 0) = 0
    AND ISNULL(V.VchCancelled, 0) = 0
GROUP BY
    T.MasterCode1
"@
            }
            else {
                $onAccountQuery = @"
SELECT
    T.MasterCode1 AS AccountCode,
    SUM(
        IIF(
            T.Value1 IS NULL,
            0,
            T.Value1
        )
    ) AS OnAccountSigned
FROM Tran3 T
INNER JOIN Tran1 V
    ON V.VchCode = T.VchCode
WHERE
    T.RecType = 1
    AND T.VchType IN (2, 3, 9, 10)
    AND T.[Date] <= $asOfLiteral
    AND (
        T.Method NOT IN (1, 2)
        OR T.RefCode IS NULL
        OR T.RefCode = 0
        OR T.[No] IS NULL
        OR T.[No] = ''
    )
    AND (
        V.Cancelled = False
        OR V.Cancelled IS NULL
    )
    AND (
        V.VchCancelled = False
        OR V.VchCancelled IS NULL
    )
GROUP BY
    T.MasterCode1
"@
            }

            $onAccountRows =
                Invoke-OutstandingDatabaseRows `
                    -Query $onAccountQuery `
                    -QueryName "ON-ACCOUNT" `
                    -Timeout 120

            $script:OutstandingReportCache[$cacheKey] = @{
                createdAt = Get-Date
                accountRows = @($accountRows)
                referenceRows = @($referenceRows)
                onAccountRows = @($onAccountRows)
            }

            Write-ReportEngineLog `
                -Stage "CACHE" `
                -Message (
                    "Cached account=$($accountRows.Count), " +
                    "references=$($referenceRows.Count), " +
                    "onAccount=$($onAccountRows.Count)."
                ) `
                -Color "Green"
        }

        $stage = "lookup maps"
        $mapWatch = New-ReportEngineWatch

        $accountMap = @{}

        foreach ($row in $accountRows) {
            $accountCode =
                ConvertTo-ReportInt $row.AccountCode

            if ($accountCode -gt 0) {
                $accountMap[$accountCode] = @{
                    code = $accountCode
                    name = [string]$row.AccountName
                    alias = [string]$row.AccountAlias
                    groupName = [string]$row.AccountGroup
                }
            }
        }

        $onAccountMap = @{}

        foreach ($row in $onAccountRows) {
            $accountCode =
                ConvertTo-ReportInt $row.AccountCode

            if ($accountCode -gt 0) {
                $signedValue = [Math]::Round(
                    (
                        ConvertTo-ReportDouble `
                            $row.OnAccountSigned
                    ),
                    2
                )

                if (-not $onAccountMap.ContainsKey($accountCode)) {
                    $onAccountMap[$accountCode] = 0.0
                }

                $onAccountMap[$accountCode] = [Math]::Round(
                    [double]$onAccountMap[$accountCode] +
                    $signedValue,
                    2
                )
            }
        }

        Stop-ReportEngineWatch `
            -Watch $mapWatch `
            -Stage "LOOKUP-MAPS" `
            -Message (
                "Accounts=$($accountMap.Count); " +
                "OnAccountAccounts=$($onAccountMap.Count)"
            )

        $stage = "filter and transform"
        $filterWatch = New-ReportEngineWatch

        $transactions =
            [System.Collections.Generic.List[object]]::new()

        $accountTotals = @{}
        $accountPendingSignedMap = @{}

        $filterCounters = [ordered]@{
            totalReferences = 0
            missingAccount = 0
            missingOriginalReference = 0
            outsideFromDate = 0
            outsideToDate = 0
            zeroPending = 0
            belowMinimum = 0
            aboveMaximum = 0
            typeFilter = 0
            accountFilter = 0
            groupFilter = 0
            statusFilter = 0
            agingFilter = 0
            voucherTypeFilter = 0
            searchFilter = 0
            included = 0
        }

        $accountFilterText =
            $Account.Trim().ToLowerInvariant()

        $groupFilterText =
            $Group.Trim().ToLowerInvariant()

        $searchFilterText =
            $Search.Trim().ToLowerInvariant()

        foreach ($referenceRow in $referenceRows) {
            $filterCounters.totalReferences++

            $accountCode =
                ConvertTo-ReportInt `
                    $referenceRow.AccountCode

            if (
                $accountCode -le 0 -or
                -not $accountMap.ContainsKey(
                    $accountCode
                )
            ) {
                $filterCounters.missingAccount++
                continue
            }

            $originalSigned = [Math]::Round(
                (
                    ConvertTo-ReportDouble `
                        $referenceRow.OriginalSigned
                ),
                2
            )

            # A valid bill reference requires a Method=1 row.
            # Method=2-only groups are not independent bills.
            if ([Math]::Abs($originalSigned) -lt 0.005) {
                $filterCounters.missingOriginalReference++
                continue
            }

            $adjustmentSigned = [Math]::Round(
                (
                    ConvertTo-ReportDouble `
                        $referenceRow.AdjustmentSigned
                ),
                2
            )

            $pendingSigned = [Math]::Round(
                (
                    ConvertTo-ReportDouble `
                        $referenceRow.PendingSigned
                ),
                2
            )

            $pendingAmount =
                [Math]::Abs($pendingSigned)

            if (
                -not $IncludeZero -and
                $pendingAmount -lt 0.005
            ) {
                $filterCounters.zeroPending++
                continue
            }

            if (
                $MinAmount -gt 0 -and
                $pendingAmount -lt $MinAmount
            ) {
                $filterCounters.belowMinimum++
                continue
            }

            if (
                $MaxAmount -gt 0 -and
                $pendingAmount -gt $MaxAmount
            ) {
                $filterCounters.aboveMaximum++
                continue
            }

            $refDate = $null

            try {
                if ($null -ne $referenceRow.RefDate) {
                    $refDate =
                        ([datetime]$referenceRow.RefDate).Date
                }
            }
            catch {}

            if ($null -eq $refDate) {
                $filterCounters.missingOriginalReference++
                continue
            }

            $dueDate = $refDate

            try {
                if ($null -ne $referenceRow.DueDate) {
                    $dueDate =
                        ([datetime]$referenceRow.DueDate).Date
                }
            }
            catch {}

            # From/To apply to the original Method=1 reference date.
            # Method=2 adjustments are included according to AsOf.
            if ($refDate.Date -lt $fromDate.Date) {
                $filterCounters.outsideFromDate++
                continue
            }

            if ($refDate.Date -gt $toDate.Date) {
                $filterCounters.outsideToDate++
                continue
            }

            $accountData = $accountMap[$accountCode]

            $nature = Resolve-OutstandingNature `
                -GroupName $accountData.groupName `
                -OriginalSigned $originalSigned

            if (
                -not [string]::IsNullOrWhiteSpace($Type) -and
                $Type -ne "all" -and
                $nature -ne $Type
            ) {
                $filterCounters.typeFilter++
                continue
            }

            if (
                -not [string]::IsNullOrWhiteSpace(
                    $accountFilterText
                )
            ) {
                $accountMatch =
                    ([string]$accountData.name).
                        ToLowerInvariant().
                        Contains($accountFilterText) -or
                    ([string]$accountData.alias).
                        ToLowerInvariant().
                        Contains($accountFilterText)

                if (-not $accountMatch) {
                    $filterCounters.accountFilter++
                    continue
                }
            }

            if (
                -not [string]::IsNullOrWhiteSpace(
                    $groupFilterText
                ) -and
                $groupFilterText -ne "all" -and
                -not (
                    ([string]$accountData.groupName).
                        ToLowerInvariant().
                        Contains($groupFilterText)
                )
            ) {
                $filterCounters.groupFilter++
                continue
            }

            $voucherCode =
                ConvertTo-ReportInt `
                    $referenceRow.VoucherCode

            $voucherTypeValue =
                ConvertTo-ReportInt `
                    $referenceRow.VoucherType

            if (
                $VoucherType -gt 0 -and
                $voucherTypeValue -ne $VoucherType
            ) {
                $filterCounters.voucherTypeFilter++
                continue
            }

            $refNo =
                ([string]$referenceRow.RefNo).Trim()

            $voucherTypeName =
                Get-OutstandingVoucherTypeName `
                    $voucherTypeValue

            if (
                -not [string]::IsNullOrWhiteSpace(
                    $searchFilterText
                )
            ) {
                $searchable = (
                    "$($accountData.name) " +
                    "$($accountData.alias) " +
                    "$($accountData.groupName) " +
                    "$refNo $voucherTypeName"
                ).ToLowerInvariant()

                if (
                    -not $searchable.Contains(
                        $searchFilterText
                    )
                ) {
                    $filterCounters.searchFilter++
                    continue
                }
            }

            $isNotDue =
                $dueDate.Date -gt $asOfDate.Date

            $daysOverdue = if ($isNotDue) {
                0
            }
            else {
                [Math]::Max(
                    0,
                    [int](
                        $asOfDate.Date -
                        $dueDate.Date
                    ).TotalDays
                )
            }

            $currentStatus = if ($isNotDue) {
                "not-due"
            }
            elseif ($daysOverdue -gt 0) {
                "overdue"
            }
            else {
                "due"
            }

            $agingBucket =
                Get-OutstandingAgeBucket `
                    -DaysOverdue $daysOverdue `
                    -IsNotDue $isNotDue

            $originalAmount =
                [Math]::Abs($originalSigned)

            $adjustedAmount =
                [Math]::Abs($adjustmentSigned)

            $partiallyAdjusted =
                $adjustedAmount -gt 0.005 -and
                $pendingAmount -gt 0.005

            $unadjusted =
                $adjustedAmount -lt 0.005 -and
                $pendingAmount -gt 0.005

            if (
                -not (
                    Test-OutstandingStatus `
                        -RequestedStatus $Status `
                        -CurrentStatus $currentStatus `
                        -PartiallyAdjusted $partiallyAdjusted `
                        -Unadjusted $unadjusted
                )
            ) {
                $filterCounters.statusFilter++
                continue
            }

            if (
                -not [string]::IsNullOrWhiteSpace($Aging) -and
                $Aging -ne "all" -and
                $agingBucket -ne $Aging
            ) {
                $filterCounters.agingFilter++
                continue
            }

            $onAccountSigned = if (
                $onAccountMap.ContainsKey($accountCode)
            ) {
                [double]$onAccountMap[$accountCode]
            }
            else {
                0.0
            }

            $accountKey = [string]$accountCode

            if (
                -not $accountTotals.ContainsKey(
                    $accountKey
                )
            ) {
                $accountTotals[$accountKey] = @{
                    accountCode = $accountCode
                    accountName = [string]$accountData.name
                    accountGroup = [string]$accountData.groupName
                    nature = $nature

                    billCount = 0
                    originalTotal = 0.0
                    adjustedTotal = 0.0
                    pendingTotal = 0.0

                    overdueTotal = 0.0
                    notDueTotal = 0.0
                    dueTodayTotal = 0.0

                    onAccountSigned = $onAccountSigned
                    onAccountAmount =
                        [Math]::Abs($onAccountSigned)

                    ledgerBalanceSigned = 0.0
                    ledgerBalanceAmount = 0.0
                }

                $accountPendingSignedMap[$accountKey] =
                    0.0
            }

            $accountTotal =
                $accountTotals[$accountKey]

            $accountTotal.billCount++
            $accountTotal.originalTotal +=
                $originalAmount
            $accountTotal.adjustedTotal +=
                $adjustedAmount
            $accountTotal.pendingTotal +=
                $pendingAmount

            if ($currentStatus -eq "overdue") {
                $accountTotal.overdueTotal +=
                    $pendingAmount
            }
            elseif ($currentStatus -eq "not-due") {
                $accountTotal.notDueTotal +=
                    $pendingAmount
            }
            else {
                $accountTotal.dueTodayTotal +=
                    $pendingAmount
            }

            $accountPendingSignedMap[$accountKey] +=
                $pendingSigned

            $transactions.Add([pscustomobject]@{
                rowId = (
                    "$accountCode-" +
                    (ConvertTo-ReportInt $referenceRow.RefCode)
                )

                accountCode = $accountCode
                accountName = [string]$accountData.name
                accountAlias = [string]$accountData.alias
                accountGroup = [string]$accountData.groupName

                nature = $nature
                type = $nature

                refCode =
                    ConvertTo-ReportInt `
                        $referenceRow.RefCode

                refNo = $refNo

                voucherCode = $voucherCode
                voucherNo = $refNo
                voucherType = $voucherTypeValue
                voucherTypeName = $voucherTypeName
                voucherTypeShortName =
                    Get-OutstandingVoucherTypeShortName `
                        $voucherTypeValue

                voucherSeriesCode = 0

                refDate =
                    $refDate.ToString("yyyy-MM-dd")

                dueDate =
                    $dueDate.ToString("yyyy-MM-dd")

                originalAmount =
                    [Math]::Round($originalAmount, 2)

                adjustedAmount =
                    [Math]::Round($adjustedAmount, 2)

                pendingAmount =
                    [Math]::Round($pendingAmount, 2)

                outstandingAmount =
                    [Math]::Round($pendingAmount, 2)

                originalSigned = $originalSigned
                adjustmentSigned = $adjustmentSigned
                pendingSigned = $pendingSigned

                due = (-not $isNotDue)
                status = $currentStatus
                daysOverdue = $daysOverdue
                agingBucket = $agingBucket

                partiallyAdjusted = $partiallyAdjusted
                unadjusted = $unadjusted

                adjustmentCount =
                    ConvertTo-ReportInt `
                        $referenceRow.AdjustmentCount

                onAccountSigned = $onAccountSigned
                onAccountAmount =
                    [Math]::Abs($onAccountSigned)

                ledgerBalanceSigned = 0.0
                ledgerBalanceAmount = 0.0
            })

            $filterCounters.included++
        }

        Stop-ReportEngineWatch `
            -Watch $filterWatch `
            -Stage "FILTER" `
            -Message (
                "Included=$($transactions.Count); " +
                "Counters=" +
                (
                    $filterCounters |
                    ConvertTo-Json -Compress
                )
            )

        $stage = "account finalisation"
        $finaliseWatch = New-ReportEngineWatch

        $ledgerMap = @{}

        foreach ($accountKey in @($accountTotals.Keys)) {
            $accountTotal =
                $accountTotals[$accountKey]

            $accountTotal.ledgerBalanceSigned =
                [Math]::Round(
                    (
                        [double]$accountPendingSignedMap[
                            $accountKey
                        ] +
                        [double]$accountTotal.onAccountSigned
                    ),
                    2
                )

            $accountTotal.ledgerBalanceAmount =
                [Math]::Abs(
                    [double]$accountTotal.ledgerBalanceSigned
                )

            foreach ($fieldName in @(
                "originalTotal",
                "adjustedTotal",
                "pendingTotal",
                "overdueTotal",
                "notDueTotal",
                "dueTodayTotal"
            )) {
                $accountTotal[$fieldName] =
                    [Math]::Round(
                        [double]$accountTotal[$fieldName],
                        2
                    )
            }

            $ledgerMap[$accountKey] = @{
                signed =
                    $accountTotal.ledgerBalanceSigned
                amount =
                    $accountTotal.ledgerBalanceAmount
            }
        }

        foreach ($transaction in $transactions) {
            $accountKey =
                [string]$transaction.accountCode

            if ($ledgerMap.ContainsKey($accountKey)) {
                $transaction.ledgerBalanceSigned =
                    $ledgerMap[$accountKey].signed

                $transaction.ledgerBalanceAmount =
                    $ledgerMap[$accountKey].amount
            }
        }

        Stop-ReportEngineWatch `
            -Watch $finaliseWatch `
            -Stage "ACCOUNT-TOTALS" `
            -Message "Accounts=$($accountTotals.Count)"

        $stage = "sorting"
        $sortWatch = New-ReportEngineWatch

        $allRows = @($transactions)
        $descending =
            $SortDirection.Trim().
                ToLowerInvariant() -eq "desc"

        switch (
            $SortBy.Trim().ToLowerInvariant()
        ) {
            "account" {
                $allRows = @(
                    $allRows |
                    Sort-Object `
                        accountName,
                        dueDate,
                        refNo `
                        -Descending:$descending
                )
            }

            "refdate" {
                $allRows = @(
                    $allRows |
                    Sort-Object `
                        refDate,
                        accountName,
                        refNo `
                        -Descending:$descending
                )
            }

            "pending" {
                $allRows = @(
                    $allRows |
                    Sort-Object `
                        pendingAmount,
                        accountName,
                        refNo `
                        -Descending:$descending
                )
            }

            "daysoverdue" {
                $allRows = @(
                    $allRows |
                    Sort-Object `
                        daysOverdue,
                        accountName,
                        refNo `
                        -Descending:$descending
                )
            }

            default {
                $allRows = @(
                    $allRows |
                    Sort-Object `
                        dueDate,
                        accountName,
                        refNo `
                        -Descending:$descending
                )
            }
        }

        Stop-ReportEngineWatch `
            -Watch $sortWatch `
            -Stage "SORT" `
            -Message "Rows=$($allRows.Count)"

        $stage = "pagination"

        $totalTransactions = $allRows.Count

        if ($PageSize -eq 0) {
            # "All" mode.
            $totalPages = 1
            $Page = 1
            $skip = 0
            $pageRows = @($allRows)
        }
        else {
            $totalPages = [Math]::Max(
                1,
                [Math]::Ceiling(
                    $totalTransactions / $PageSize
                )
            )

            if ($Page -gt $totalPages) {
                $Page = $totalPages
            }

            $skip = ($Page - 1) * $PageSize

            $pageRows = @(
                $allRows |
                Select-Object `
                    -Skip $skip `
                    -First $PageSize
            )
        }

        $stage = "summary"
        $summaryWatch = New-ReportEngineWatch

        $totalReceivable = [Math]::Round(
            (@(
                $allRows |
                Where-Object {
                    $_.nature -eq "receivable"
                } |
                Measure-Object `
                    pendingAmount `
                    -Sum
            ).Sum),
            2
        )

        $totalPayable = [Math]::Round(
            (@(
                $allRows |
                Where-Object {
                    $_.nature -eq "payable"
                } |
                Measure-Object `
                    pendingAmount `
                    -Sum
            ).Sum),
            2
        )

        $overdueReceivable = [Math]::Round(
            (@(
                $allRows |
                Where-Object {
                    $_.nature -eq "receivable" -and
                    $_.status -eq "overdue"
                } |
                Measure-Object `
                    pendingAmount `
                    -Sum
            ).Sum),
            2
        )

        $overduePayable = [Math]::Round(
            (@(
                $allRows |
                Where-Object {
                    $_.nature -eq "payable" -and
                    $_.status -eq "overdue"
                } |
                Measure-Object `
                    pendingAmount `
                    -Sum
            ).Sum),
            2
        )

        $notDueAmount = [Math]::Round(
            (@(
                $allRows |
                Where-Object {
                    $_.status -eq "not-due"
                } |
                Measure-Object `
                    pendingAmount `
                    -Sum
            ).Sum),
            2
        )

        $dueTodayAmount = [Math]::Round(
            (@(
                $allRows |
                Where-Object {
                    $_.status -eq "due"
                } |
                Measure-Object `
                    pendingAmount `
                    -Sum
            ).Sum),
            2
        )

        $agingSummary = @{
            notDue = 0.0
            days0To30 = 0.0
            days31To60 = 0.0
            days61To90 = 0.0
            days91To180 = 0.0
            above180 = 0.0
        }

        foreach ($row in $allRows) {
            $bucket = [string]$row.agingBucket

            if ($agingSummary.ContainsKey($bucket)) {
                $agingSummary[$bucket] +=
                    [double]$row.pendingAmount
            }
        }

        foreach ($bucket in @($agingSummary.Keys)) {
            $agingSummary[$bucket] =
                [Math]::Round(
                    [double]$agingSummary[$bucket],
                    2
                )
        }

        $uniqueAccounts = @(
            $allRows |
            Select-Object `
                -ExpandProperty accountCode `
                -Unique
        )

        $receivableAccounts = @(
            $allRows |
            Where-Object {
                $_.nature -eq "receivable"
            } |
            Select-Object `
                -ExpandProperty accountCode `
                -Unique
        ).Count

        $payableAccounts = @(
            $allRows |
            Where-Object {
                $_.nature -eq "payable"
            } |
            Select-Object `
                -ExpandProperty accountCode `
                -Unique
        ).Count

        # accountTotals.Values contains hashtables, not PSCustomObjects.
        # Measure-Object -Property cannot reliably read hashtable keys.
        $totalOnAccountSignedRaw = 0.0

        foreach ($accountTotalValue in $accountTotals.Values) {
            if (
                $null -ne $accountTotalValue -and
                $accountTotalValue.ContainsKey("onAccountSigned")
            ) {
                $totalOnAccountSignedRaw +=
                    [double]$accountTotalValue["onAccountSigned"]
            }
        }

        $totalOnAccountSigned = [Math]::Round(
            $totalOnAccountSignedRaw,
            2
        )

        Stop-ReportEngineWatch `
            -Watch $summaryWatch `
            -Stage "SUMMARY" `
            -Message (
                "Receivable=$totalReceivable; " +
                "Payable=$totalPayable; " +
                "Bills=$totalTransactions; " +
                "Accounts=$($uniqueAccounts.Count); " +
                "OnAccount=$totalOnAccountSigned"
            )

        Stop-ReportEngineWatch `
            -Watch $overallWatch `
            -Stage "COMPLETE" `
            -Message (
                "PageRows=$($pageRows.Count); " +
                "CacheHit=$cacheHit"
            )

        return @{
            success = $true

            databaseMode = if ($isSql) {
                "sql-server-direct"
            }
            else {
                "access-busy-sql"
            }

            fromDate =
                $fromDate.ToString("yyyy-MM-dd")

            toDate =
                $toDate.ToString("yyyy-MM-dd")

            asOfDate =
                $asOfDate.ToString("yyyy-MM-dd")

            generatedAt =
                (Get-Date).ToString(
                    "yyyy-MM-ddTHH:mm:ss"
                )

            page = $Page
            pageSize = $PageSize
            total = $totalTransactions
            totalTransactions = $totalTransactions
            totalPages = $totalPages
            summaryScope = "all"

            summary = @{
                totalReceivable = $totalReceivable
                totalPayable = $totalPayable

                netPosition = [Math]::Round(
                    $totalReceivable -
                    $totalPayable,
                    2
                )

                overdueReceivable =
                    $overdueReceivable

                overduePayable =
                    $overduePayable

                notDueAmount =
                    $notDueAmount

                dueTodayAmount =
                    $dueTodayAmount

                totalOutstanding = [Math]::Round(
                    $totalReceivable +
                    $totalPayable,
                    2
                )

                totalBills =
                    $totalTransactions

                totalAccounts =
                    $uniqueAccounts.Count

                receivableAccounts =
                    $receivableAccounts

                payableAccounts =
                    $payableAccounts

                receivableParties =
                    $receivableAccounts

                payableParties =
                    $payableAccounts

                partiallyAdjustedBills = @(
                    $allRows |
                    Where-Object {
                        $_.partiallyAdjusted
                    }
                ).Count

                unadjustedBills = @(
                    $allRows |
                    Where-Object {
                        $_.unadjusted
                    }
                ).Count

                totalOriginalAmount =
                    [Math]::Round(
                        (@(
                            $allRows |
                            Measure-Object `
                                originalAmount `
                                -Sum
                        ).Sum),
                        2
                    )

                totalAdjustedAmount =
                    [Math]::Round(
                        (@(
                            $allRows |
                            Measure-Object `
                                adjustedAmount `
                                -Sum
                        ).Sum),
                        2
                    )

                totalOnAccountSigned =
                    $totalOnAccountSigned

                totalOnAccountAmount =
                    [Math]::Abs(
                        $totalOnAccountSigned
                    )

                aging = $agingSummary
            }

            accountTotals = $accountTotals
            data = $pageRows

            debug = @{
                stage = "complete"
                engine = $engineName
                dbType = $dbType
                cacheHit = $cacheHit
                cacheSeconds =
                    $script:OutstandingReportCacheSeconds

                accountRows = $accountRows.Count
                accountMap = $accountMap.Count

                aggregatedReferenceRows =
                    $referenceRows.Count

                onAccountRows =
                    $onAccountRows.Count

                onAccountAccounts =
                    $onAccountMap.Count

                totalOnAccountSigned =
                    $totalOnAccountSigned

                filteredTransactions =
                    $totalTransactions

                pageRows =
                    $pageRows.Count

                filterCounters =
                    $filterCounters

                memory =
                    Get-ReportEngineMemory
            }
        }
    }
    catch {
        Write-ReportEngineLog `
            -Stage "ERROR" `
            -Message (
                "Stage=$stage; " +
                $_.Exception.ToString()
            ) `
            -Color "Red"

        Stop-ReportEngineWatch `
            -Watch $overallWatch `
            -Stage "FAILED" `
            -Message "Stage=$stage"

        return @{
            success = $false
            error = $_.Exception.Message
            exception = $_.Exception.ToString()
            stage = $stage

            databaseMode = if ($isSql) {
                "sql-server-direct"
            }
            else {
                "access-busy-sql"
            }

            data = @()

            debug = @{
                engine = $engineName
                dbType = $dbType
                instanceId = $InstanceId
                companyCode = $CompanyCode
                memory =
                    Get-ReportEngineMemory
            }
        }
    }
    finally {
        if ($null -ne $sqlConnection) {
            try {
                if (
                    $sqlConnection.State.ToString() -ne
                    "Closed"
                ) {
                    $sqlConnection.Close()
                }
            }
            catch {}

            try {
                $sqlConnection.Dispose()
            }
            catch {}
        }

        if ($null -ne $busyConnection) {
            Disconnect-BUSY $busyConnection
        }
    }
}


function Get-StockStatusReport {
    param(
        [string]$From                  = "",
        [string]$To                    = "",
        [string]$AsOf                  = "",
        [string]$View                  = "balances",
        [string]$MaterialCentre        = "",
        [string]$MaterialCentres       = "",
        [string]$ItemGroup             = "",
        [string]$Search                = "",
        [string]$Status                = "all",
        [bool]$IncludeZero             = $true,
        [double]$LowStockLevel         = 5,
        [string]$ValueBy               = "busy",
        [string]$UnitMode              = "both",
        [bool]$ShowValue               = $true,
        [bool]$IncludeStockTransfers   = $true,
        [bool]$ShowSalePurchaseSeparately = $true,
        [string]$MastersMode           = "moved-closing",
        [bool]$ShowParentGroup         = $true,
        [int]$Page                     = 1,
        [int]$PageSize                 = 5000,
        [string]$InstanceId            = "",
        [string]$CompanyCode           = ""
    )

    # ============================================================
    # BUSY-LIKE SINGLE-PAGE STOCK STATUS ENGINE
    #
    # One response powers the web app's Balances, Detailed,
    # Columnar (MC-wise), Grouped and Hierarchical views.
    #
    # Quantity source of truth:
    #   Opening stock : Tran4.D1
    #   Main movement : Tran2.Value1
    #   Alt movement  : Tran2.Value2
    #
    # BUSY valuation source of truth:
    #   Opening quantity/value : Tran4.D1 / Tran4.D3
    #   Inventory movements    : Tran2.Value1 / Value3 / D2 / D5
    #
    # IMPORTANT:
    # BUSY StockValMethod=5 is NOT "sum Value3 / closing qty".
    # Outward rows (Sale / Issue / Consumption / Transfer Out) must
    # reduce the valuation pool at the CURRENT moving-average rate.
    # Zero-valued Production / Stock Journal inward rows are preserved
    # as zero unless the same item's same voucher has an outward source.
    #
    # The batch engine below reproduces the same MC-specific method-5
    # logic already proven against BUSY voucher costing, but does it in
    # one transaction scan so the Stock Status page stays fast.
    # ============================================================

    $eps = 0.000001

    if ($Page -lt 1) { $Page = 1 }
    if ($PageSize -lt 1) { $PageSize = 250 }
    if ($PageSize -gt 5000) { $PageSize = 5000 }

    $viewMode = ([string]$View).Trim().ToLowerInvariant()
    if ($viewMode -notin @("balances", "detailed", "columnar", "grouped", "hierarchical")) {
        $viewMode = "balances"
    }

    $valueMode = ([string]$ValueBy).Trim().ToLowerInvariant()
    if ($valueMode -notin @("busy", "purchase", "sale", "mrp")) {
        $valueMode = "busy"
    }

    $unitDisplay = ([string]$UnitMode).Trim().ToLowerInvariant()
    if ($unitDisplay -notin @("main", "alt", "both")) {
        $unitDisplay = "both"
    }

    $mastersFilter = ([string]$MastersMode).Trim().ToLowerInvariant()
    if ($mastersFilter -notin @("all", "moved-closing", "moved", "closing")) {
        $mastersFilter = "moved-closing"
    }

    $statusFilter = ([string]$Status).Trim().ToLowerInvariant()
    if ($statusFilter -notin @("all", "in-stock", "low-stock", "out-of-stock", "negative")) {
        $statusFilter = "all"
    }

    $endDate = if (-not [string]::IsNullOrWhiteSpace($To)) {
        ConvertTo-ReportDate -Value $To
    }
    else {
        ConvertTo-ReportDate -Value $AsOf
    }

    $asOfDate = if (-not [string]::IsNullOrWhiteSpace($AsOf)) {
        ConvertTo-ReportDate -Value $AsOf -Fallback $endDate
    }
    else {
        $endDate
    }

    # BUSY financial years in this project start on 1 April.  This is
    # only the default movement start. A supplied From date always wins.
    $fyYear = if ($endDate.Month -ge 4) { $endDate.Year } else { $endDate.Year - 1 }
    $defaultFrom = [datetime]::new($fyYear, 4, 1)

    $fromDate = if (-not [string]::IsNullOrWhiteSpace($From)) {
        ConvertTo-ReportDate -Value $From -Fallback $defaultFrom
    }
    else {
        $defaultFrom
    }

    if ($viewMode -ne "detailed") {
        # Balance/Columnar/Grouped/Hierarchical are point-in-time views.
        $endDate = $asOfDate
    }

    if ($fromDate.Date -gt $endDate.Date) {
        return @{
            success = $false
            error = "From date cannot be after the report end date."
            data = @()
        }
    }

    # ------------------------------------------------------------
    # FAST RESULT CACHE
    # ------------------------------------------------------------
    # This cache is intentionally short-lived. It protects the bridge from
    # repeating the full chronological method-5 replay when the user revisits
    # the report, switches away and back, or repeats the same filters.
    #
    # A unit-separator character is used between key parts so item/search text
    # containing ordinary pipes cannot accidentally collide with another key.
    # ------------------------------------------------------------
    Remove-ExpiredStockStatusReportCache

    $cacheSeparator = [char]31
    $stockStatusCacheKey = (
        @(
            [string]$InstanceId,
            [string]$CompanyCode,
            $fromDate.ToString("yyyy-MM-dd"),
            $endDate.ToString("yyyy-MM-dd"),
            $asOfDate.ToString("yyyy-MM-dd"),
            [string]$viewMode,
            ([string]$MaterialCentre).Trim().ToLowerInvariant(),
            ([string]$MaterialCentres).Trim().ToLowerInvariant(),
            ([string]$ItemGroup).Trim().ToLowerInvariant(),
            ([string]$Search).Trim().ToLowerInvariant(),
            [string]$statusFilter,
            [string]([bool]$IncludeZero),
            [string]([double]$LowStockLevel),
            [string]$valueMode,
            [string]$unitDisplay,
            [string]([bool]$ShowValue),
            [string]([bool]$IncludeStockTransfers),
            [string]([bool]$ShowSalePurchaseSeparately),
            [string]$mastersFilter,
            [string]([bool]$ShowParentGroup),
            [string]([int]$Page),
            [string]([int]$PageSize)
        ) -join $cacheSeparator
    )

    if ($script:StockStatusReportCache.ContainsKey($stockStatusCacheKey)) {
        $cachedEntry =
            $script:StockStatusReportCache[$stockStatusCacheKey]

        $cacheAgeSeconds = 0.0

        try {
            $cacheAgeSeconds = (
                (Get-Date) -
                [datetime]$cachedEntry.createdAt
            ).TotalSeconds
        }
        catch {
            $cacheAgeSeconds =
                $script:StockStatusReportCacheSeconds + 1
        }

        if (
            $cacheAgeSeconds -le
            $script:StockStatusReportCacheSeconds -and
            $null -ne $cachedEntry.result
        ) {
            Write-ReportEngineLog `
                -Stage "STOCK-CACHE" `
                -Message (
                    "Cache hit; age=" +
                    [Math]::Round($cacheAgeSeconds, 1) +
                    "s"
                ) `
                -Color "Green"

            return $cachedEntry.result
        }

        $script:StockStatusReportCache.Remove(
            $stockStatusCacheKey
        )
    }

    $stockTotalWatch = New-ReportEngineWatch
    Write-ReportEngineLog `
        -Stage "STOCK-START" `
        -Message ("Cold calculation; view=" + $viewMode + "; asOf=" + $asOfDate.ToString("yyyy-MM-dd"))

    $mode = Get-ReportDatabaseMode -InstanceId $InstanceId
    $fromLiteral = Get-ReportSqlDateLiteral -Date $fromDate -IsSql $mode.isSql
    $endLiteral = Get-ReportSqlDateLiteral -Date $endDate -IsSql $mode.isSql

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) {
        return @{ success = $false; error = "BUSY connection failed"; data = @() }
    }

    try {
        # --------------------------------------------------------
        # Master lookups: item groups and material centres.
        # --------------------------------------------------------
        $groupMap = @{}
        $groupOptions = [System.Collections.Generic.List[object]]::new()

        $groupRst = $fi.GetRecordset(
            "SELECT Code, Name, Alias, ParentGrp FROM Master1 WHERE MasterType=5 ORDER BY Name"
        )

        if ($groupRst -and -not $groupRst.EOF) {
            $groupRst.MoveFirst()
            while (-not $groupRst.EOF) {
                $code = ConvertTo-ReportInt (Get-ReportFieldValue $groupRst "Code" 0)
                $parentCode = ConvertTo-ReportInt (Get-ReportFieldValue $groupRst "ParentGrp" 0)
                $name = [string](Get-ReportFieldValue $groupRst "Name" "")
                $alias = [string](Get-ReportFieldValue $groupRst "Alias" "")

                if ($code -gt 0) {
                    $groupMap[$code] = [pscustomobject]@{
                        code = $code
                        name = $name
                        alias = $alias
                        parentCode = $parentCode
                    }
                }

                $groupRst.MoveNext()
            }
        }
        Close-ReportRecordset $groupRst

        function Get-StockGroupPath {
            param([int]$GroupCode)

            $path = [System.Collections.ArrayList]::new()
            $visited = @{}
            $current = $GroupCode
            $guard = 0

            while (
                $current -gt 0 -and
                $groupMap.ContainsKey($current) -and
                $guard -lt 25
            ) {
                if ($visited.ContainsKey([string]$current)) { break }
                $visited[[string]$current] = $true

                $group = $groupMap[$current]
                if (-not [string]::IsNullOrWhiteSpace([string]$group.name)) {
                    [void]$path.Insert(0, [string]$group.name)
                }

                $current = [int]$group.parentCode
                $guard++
            }

            return @($path)
        }

        foreach ($groupCode in @($groupMap.Keys)) {
            $group = $groupMap[$groupCode]
            $path = @(Get-StockGroupPath -GroupCode ([int]$groupCode))
            $groupOptions.Add([pscustomobject]@{
                code = [int]$group.code
                name = [string]$group.name
                alias = [string]$group.alias
                parentCode = [int]$group.parentCode
                path = @($path)
                pathLabel = ($path -join " > ")
            })
        }

        $mcMap = @{}
        $mcNameToCode = @{}
        $mcOptions = [System.Collections.Generic.List[object]]::new()

        $mcRst = $fi.GetRecordset(
            "SELECT Code, Name FROM Master1 WHERE MasterType=11 ORDER BY Name"
        )

        if ($mcRst -and -not $mcRst.EOF) {
            $mcRst.MoveFirst()
            while (-not $mcRst.EOF) {
                $mcCode = ConvertTo-ReportInt (Get-ReportFieldValue $mcRst "Code" 0)
                $mcName = [string](Get-ReportFieldValue $mcRst "Name" "")

                if ($mcCode -gt 0) {
                    $mcMap[$mcCode] = $mcName
                    $mcNameToCode[$mcName.Trim().ToLowerInvariant()] = $mcCode
                    $mcOptions.Add([pscustomobject]@{
                        code = $mcCode
                        name = $mcName
                    })
                }

                $mcRst.MoveNext()
            }
        }
        Close-ReportRecordset $mcRst

        # Material-centre selection supports the legacy single value plus
        # the new pipe-separated multi-selection used by the single page.
        $requestedMcNames = [System.Collections.ArrayList]::new()

        if (-not [string]::IsNullOrWhiteSpace($MaterialCentres)) {
            foreach ($part in ($MaterialCentres -split '\|')) {
                $candidate = ([string]$part).Trim()
                if (
                    -not [string]::IsNullOrWhiteSpace($candidate) -and
                    $candidate.ToLowerInvariant() -ne "all"
                ) {
                    [void]$requestedMcNames.Add($candidate)
                }
            }
        }
        elseif (
            -not [string]::IsNullOrWhiteSpace($MaterialCentre) -and
            $MaterialCentre.Trim().ToLowerInvariant() -ne "all"
        ) {
            [void]$requestedMcNames.Add($MaterialCentre.Trim())
        }

        $selectedMcCodes = @{}
        foreach ($mcName in @($requestedMcNames)) {
            $key = ([string]$mcName).Trim().ToLowerInvariant()
            if ($mcNameToCode.ContainsKey($key)) {
                $selectedMcCodes[[int]$mcNameToCode[$key]] = $true
            }
        }

        $useAllMc = ($selectedMcCodes.Count -eq 0)

        function Test-StockMcSelected {
            param([int]$McCode)
            if ($useAllMc) { return $true }
            return $selectedMcCodes.ContainsKey($McCode)
        }

        # --------------------------------------------------------
        # FAST item-master load.
        #
        # IMPORTANT:
        # The old implementation used two correlated subqueries for every
        # item row to resolve MainUnit / AltUnit names. On large BUSY
        # companies Access may execute those lookups thousands of times.
        #
        # Load the Code -> Name map once, then read item masters with a
        # simple indexed query. This removes another N+1-style bottleneck.
        # --------------------------------------------------------
        $masterNameMap = @{}
        $masterNameRst = $fi.GetRecordset(
            "SELECT Code, Name FROM Master1"
        )

        if ($masterNameRst -and -not $masterNameRst.EOF) {
            $masterNameRst.MoveFirst()
            while (-not $masterNameRst.EOF) {
                $masterCode = ConvertTo-ReportInt (Get-ReportFieldValue $masterNameRst "Code" 0)
                if ($masterCode -gt 0) {
                    $masterNameMap[$masterCode] = [string](Get-ReportFieldValue $masterNameRst "Name" "")
                }
                $masterNameRst.MoveNext()
            }
        }
        Close-ReportRecordset $masterNameRst

        $itemQry = @"
SELECT
    M.Code,
    M.Name,
    M.Alias,
    M.ParentGrp,
    M.CM1 AS MainUnitCode,
    M.CM2 AS AltUnitCode,
    M.D2 AS MRP,
    M.D3 AS SalePrice,
    M.D4 AS PurchasePrice
FROM Master1 M
WHERE M.MasterType=6
ORDER BY M.Name, M.Code
"@

        $items = @{}
        $itemOrder = [System.Collections.ArrayList]::new()
        $itemRst = $fi.GetRecordset($itemQry)

        if ($itemRst -and -not $itemRst.EOF) {
            $itemRst.MoveFirst()
            while (-not $itemRst.EOF) {
                $itemCode = ConvertTo-ReportInt (Get-ReportFieldValue $itemRst "Code" 0)
                $itemName = [string](Get-ReportFieldValue $itemRst "Name" "")
                $alias = [string](Get-ReportFieldValue $itemRst "Alias" "")
                $groupCode = ConvertTo-ReportInt (Get-ReportFieldValue $itemRst "ParentGrp" 0)
                $groupPath = @(Get-StockGroupPath -GroupCode $groupCode)
                $groupName = if ($groupPath.Count -gt 0) { [string]$groupPath[$groupPath.Count - 1] } else { "General" }

                $matchesSearch = $true
                if (-not [string]::IsNullOrWhiteSpace($Search)) {
                    $needle = $Search.Trim().ToLowerInvariant()
                    $matchesSearch = (
                        $itemName.ToLowerInvariant().Contains($needle) -or
                        $alias.ToLowerInvariant().Contains($needle)
                    )
                }

                $matchesGroup = $true
                if (
                    -not [string]::IsNullOrWhiteSpace($ItemGroup) -and
                    $ItemGroup.Trim().ToLowerInvariant() -ne "all"
                ) {
                    $groupNeedle = $ItemGroup.Trim().ToLowerInvariant()
                    $pathLabel = ($groupPath -join " > ").Trim().ToLowerInvariant()
                    $matchesGroup = ($pathLabel -eq $groupNeedle)

                    if (-not $matchesGroup -and -not $groupNeedle.Contains(" > ")) {
                        foreach ($segment in $groupPath) {
                            if (([string]$segment).Trim().ToLowerInvariant() -eq $groupNeedle) {
                                $matchesGroup = $true
                                break
                            }
                        }
                    }
                }

                if ($itemCode -gt 0 -and $matchesSearch -and $matchesGroup) {
                    $mainUnitCode = ConvertTo-ReportInt (Get-ReportFieldValue $itemRst "MainUnitCode" 0)
                    $altUnitCode = ConvertTo-ReportInt (Get-ReportFieldValue $itemRst "AltUnitCode" 0)

                    $mainUnitName = if ($mainUnitCode -gt 0 -and $masterNameMap.ContainsKey($mainUnitCode)) {
                        [string]$masterNameMap[$mainUnitCode]
                    }
                    else { "" }

                    $altUnitName = if ($altUnitCode -gt 0 -and $masterNameMap.ContainsKey($altUnitCode)) {
                        [string]$masterNameMap[$altUnitCode]
                    }
                    else { "" }

                    $items[$itemCode] = [pscustomobject]@{
                        code = $itemCode
                        name = $itemName
                        alias = $alias
                        groupCode = $groupCode
                        group = $groupName
                        groupPath = @($groupPath)
                        mainUnitCode = $mainUnitCode
                        altUnitCode = $altUnitCode
                        mainUnit = $mainUnitName
                        altUnit = $altUnitName
                        conversionFactor = 1.0
                        conversionType = 1
                        mrp = ConvertTo-ReportDouble (Get-ReportFieldValue $itemRst "MRP" 0)
                        salePrice = ConvertTo-ReportDouble (Get-ReportFieldValue $itemRst "SalePrice" 0)
                        purchasePrice = ConvertTo-ReportDouble (Get-ReportFieldValue $itemRst "PurchasePrice" 0)
                        mc = @{}
                    }
                    [void]$itemOrder.Add($itemCode)
                }

                $itemRst.MoveNext()
            }
        }
        Close-ReportRecordset $itemRst

        # --------------------------------------------------------
        # SQL item filter for narrow searches / groups.
        #
        # The expensive valuation scan used to read the whole Tran2 table
        # even when the user searched for one item. When the current master
        # filter contains a reasonably small number of item codes, push those
        # codes into Access so unrelated transaction rows never cross COM.
        #
        # For the normal "All Items" report we deliberately leave the query
        # unfiltered; thousands of values in an Access IN (...) list can be
        # slower than the original scan and can exceed parser limits.
        # --------------------------------------------------------
        $tran2ItemPredicate = ""
        $tran4ItemPredicate = ""

        if ($itemOrder.Count -eq 0) {
            $tran2ItemPredicate = "AND 1=0"
            $tran4ItemPredicate = "AND 1=0"
        }
        elseif ($itemOrder.Count -le 400) {
            $itemSqlCodes = (
                @($itemOrder) |
                ForEach-Object { [int]$_ }
            ) -join ","

            $tran2ItemPredicate =
                "AND T.MasterCode1 IN ($itemSqlCodes)"

            $tran4ItemPredicate =
                "AND MasterCode1 IN ($itemSqlCodes)"
        }

        function New-StockMcState {
            return [pscustomobject]@{
                baseOpeningMain = 0.0
                baseOpeningAlt = 0.0
                baseOpeningValue = 0.0
                priorMain = 0.0
                priorAlt = 0.0
                priorValue = 0.0

                inwardMain = 0.0
                inwardAlt = 0.0
                inwardValue = 0.0
                outwardMain = 0.0
                outwardAlt = 0.0
                outwardValue = 0.0

                purchaseQty = 0.0
                saleQty = 0.0
                saleReturnQty = 0.0
                purchaseReturnQty = 0.0
                transferInQty = 0.0
                transferInAltQty = 0.0
                transferInValue = 0.0
                transferOutQty = 0.0
                transferOutAltQty = 0.0
                transferOutValue = 0.0
                productionGeneratedQty = 0.0
                productionConsumedQty = 0.0
                stockJournalGeneratedQty = 0.0
                stockJournalConsumedQty = 0.0
                materialReceiptQty = 0.0
                materialIssueQty = 0.0
                otherInQty = 0.0
                otherOutQty = 0.0
            }
        }

        function Get-StockMcState {
            param($Item, [int]$McCode)
            $key = [string]$McCode
            if (-not $Item.mc.ContainsKey($key)) {
                $Item.mc[$key] = New-StockMcState
            }
            return $Item.mc[$key]
        }

        # --------------------------------------------------------
        # Base opening stock (financial-year opening rows).
        # --------------------------------------------------------
        $opQry = @"
SELECT
    MasterCode1 AS ItemCode,
    MasterCode2 AS MCCode,
    SUM(D1) AS MainQty,
    SUM(D2) AS AltQty,
    SUM(D3) AS StockValue
FROM Tran4
WHERE RecType=0
    $tran4ItemPredicate
GROUP BY MasterCode1, MasterCode2
"@

        $opRst = $fi.GetRecordset($opQry)
        if ($opRst -and -not $opRst.EOF) {
            $opRst.MoveFirst()
            while (-not $opRst.EOF) {
                $itemCode = ConvertTo-ReportInt (Get-ReportFieldValue $opRst "ItemCode" 0)
                $mcCode = ConvertTo-ReportInt (Get-ReportFieldValue $opRst "MCCode" 0)

                if ($items.ContainsKey($itemCode)) {
                    $state = Get-StockMcState -Item $items[$itemCode] -McCode $mcCode
                    $state.baseOpeningMain += ConvertTo-ReportDouble (Get-ReportFieldValue $opRst "MainQty" 0)
                    $state.baseOpeningAlt += ConvertTo-ReportDouble (Get-ReportFieldValue $opRst "AltQty" 0)
                    $state.baseOpeningValue += ConvertTo-ReportDouble (Get-ReportFieldValue $opRst "StockValue" 0)
                }

                $opRst.MoveNext()
            }
        }
        Close-ReportRecordset $opRst

        # --------------------------------------------------------
        # Movement totals are intentionally NOT queried separately here.
        #
        # The previous implementation executed three additional joined Tran2
        # scans (prior / inward / outward) and then scanned Tran2 again for
        # BUSY method-5 valuation. On large Access companies those repeated
        # scans dominate the report time.
        #
        # The chronological method-5 scan below now also fills prior/opening
        # movement, inward/outward totals and voucher-type breakdowns in the
        # same pass. This keeps the result identical while reducing Tran2 work
        # from four passes to one.
        # --------------------------------------------------------

        # ========================================================
        # FAST BUSY STOCK VALUATION METHOD-5 ENGINE
        # ========================================================
        #
        # Why this exists:
        #
        # The earlier Stock Status implementation calculated:
        #
        #   (Opening D3 + Inward Value3 - Outward Value3) / Closing Qty
        #
        # That is NOT BUSY method-5 valuation.  A Sale row's Value3 can
        # differ from the inventory cost that BUSY removes from stock.
        #
        # Proven example supplied from BUSY:
        #   Item: 20 SMD RESISTOR WE043
        #
        #   Main Store sale row:
        #       Qty     = -10
        #       Value3  = -400
        #       D2      = 50
        #       D5      = 500
        #
        # BUSY keeps the stock cost at 50.  Therefore subtracting Value3
        # (400) directly makes the report price wrong.
        #
        # This engine replays ALL item inventory movements ONCE and keeps
        # a moving weighted-average pool for every Item + Material Centre.
        # No per-item GetMasterXML loop is used, so it remains fast.
        #
        # The supplied company/item diagnostic confirmed StockValMethod=5.
        # If a future item uses a different per-item valuation method, that
        # method can be added separately without changing this fast engine.
        # ========================================================

        $busyWaStates = @{}
        $busyReturnParentMap = @{}

        function Get-BusyWaKey {
            param(
                [int]$ItemCode,
                [int]$McCode
            )

            return "$ItemCode|$McCode"
        }

        function Get-BusyWaState {
            param(
                [int]$ItemCode,
                [int]$McCode
            )

            $key = Get-BusyWaKey `
                -ItemCode $ItemCode `
                -McCode $McCode

            if (-not $busyWaStates.ContainsKey($key)) {
                $busyWaStates[$key] = [pscustomobject]@{
                    itemCode  = $ItemCode
                    mcCode    = $McCode
                    ledgerQty = 0.0
                    poolQty   = 0.0
                    poolValue = 0.0

                    # Keep the last meaningful valuation rate even when
                    # physical stock goes negative and the positive pool
                    # has already been exhausted.
                    lastRate  = 0.0
                }
            }

            return $busyWaStates[$key]
        }

        function Get-BusyWaRate {
            param(
                [int]$ItemCode,
                [int]$McCode
            )

            $state = Get-BusyWaState `
                -ItemCode $ItemCode `
                -McCode $McCode

            # BUSY method-5 keeps a signed valuation balance.  Once stock
            # becomes negative, both poolQty and poolValue are negative.
            # The report Price is still a positive unit rate.
            if ([Math]::Abs([double]$state.poolQty) -gt $eps) {
                return [Math]::Abs(
                    [double]$state.poolValue /
                    [double]$state.poolQty
                )
            }

            return [Math]::Abs([double]$state.lastRate)
        }

        function Add-BusyWaPool {
            param(
                [int]$ItemCode,
                [int]$McCode,
                [double]$Qty,
                [double]$Value
            )

            if ($Qty -le $eps) {
                return
            }

            $state = Get-BusyWaState `
                -ItemCode $ItemCode `
                -McCode $McCode

            $incomingQty = [double]$Qty
            $incomingValue = [Math]::Abs([double]$Value)
            $incomingRate = if ($incomingQty -gt $eps) {
                $incomingValue / $incomingQty
            }
            else {
                0.0
            }

            # ----------------------------------------------------
            # BUSY NEGATIVE-STOCK RULE FOR INWARD MOVEMENT
            # ----------------------------------------------------
            # If this MC is already negative, an inward row first fills
            # that negative balance at the CURRENT negative-stock average.
            # Therefore the negative average does not jump merely because
            # a Sale Return / Purchase / Transfer In has a different posted
            # rate.  Only any quantity that crosses above zero is valued at
            # the inward row's own rate/value.
            #
            # Proven example: GF002 Main Store ends at 1,045.227... even
            # after 3 Sale-Return units posted at 1,050 because those units
            # only reduce an existing negative balance.
            # ----------------------------------------------------
            if ([double]$state.poolQty -lt -$eps) {
                $negativeQty = [Math]::Abs([double]$state.poolQty)
                $negativeRate = Get-BusyWaRate `
                    -ItemCode $ItemCode `
                    -McCode $McCode

                $offsetQty = [Math]::Min($incomingQty, $negativeQty)

                if ($offsetQty -gt $eps) {
                    $state.poolQty = [double]$state.poolQty + $offsetQty
                    $state.poolValue = [double]$state.poolValue + ($offsetQty * $negativeRate)
                    $incomingQty -= $offsetQty
                }

                if ([Math]::Abs([double]$state.poolQty) -le $eps) {
                    $state.poolQty = 0.0
                    $state.poolValue = 0.0
                }
            }

            # Any quantity left after clearing negative stock becomes
            # ordinary positive inward stock and uses the row's own value.
            if ($incomingQty -gt $eps) {
                $remainingValue = $incomingQty * $incomingRate

                $state.poolQty = [double]$state.poolQty + $incomingQty
                $state.poolValue = [double]$state.poolValue + $remainingValue
            }

            if ([Math]::Abs([double]$state.poolQty) -gt $eps) {
                $state.lastRate = [Math]::Abs(
                    [double]$state.poolValue /
                    [double]$state.poolQty
                )
            }
            elseif ($incomingRate -gt $eps) {
                $state.lastRate = $incomingRate
            }
        }

        function Remove-BusyWaPool {
            param(
                [int]$ItemCode,
                [int]$McCode,
                [double]$Qty,
                [double]$UnitRate
            )

            if ($Qty -le $eps) {
                return
            }

            $state = Get-BusyWaState `
                -ItemCode $ItemCode `
                -McCode $McCode

            $outQty = [double]$Qty
            $currentQty = [double]$state.poolQty
            $currentRate = Get-BusyWaRate `
                -ItemCode $ItemCode `
                -McCode $McCode

            $postedRate = [Math]::Max(0.0, [double]$UnitRate)

            # When stock is positive before the outward transaction, BUSY
            # issues the WHOLE row at the current moving-average rate even if
            # that row crosses the balance below zero.  AIR CLEANER GS024 is
            # the proven case: 1 @ 170 then Sale 4 uses 170 for all 4 units.
            if ($currentQty -gt $eps) {
                $useRate = $currentRate
            }
            else {
                # Once stock is already zero/negative, BUSY uses the row's
                # posted inventory valuation rate.  Each additional negative
                # issue therefore participates in a signed weighted average.
                $useRate = if ($postedRate -gt $eps) {
                    $postedRate
                }
                else {
                    $currentRate
                }
            }

            $state.poolQty = [double]$state.poolQty - $outQty
            $state.poolValue = [double]$state.poolValue - ($outQty * $useRate)

            if ([Math]::Abs([double]$state.poolQty) -le $eps) {
                $state.poolQty = 0.0
                $state.poolValue = 0.0
            }
            else {
                $state.lastRate = [Math]::Abs(
                    [double]$state.poolValue /
                    [double]$state.poolQty
                )
            }

            if ($useRate -gt $eps -and [Math]::Abs([double]$state.poolQty) -le $eps) {
                $state.lastRate = $useRate
            }
        }

        function Get-BusyWaNegativeOutRate {
            param(
                $Row,
                [double]$Qty
            )

            if ($Qty -le $eps) {
                return 0.0
            }

            # IMPORTANT: while stock is already negative, BUSY's signed
            # valuation follows the transaction's posted stock value.  In the
            # supplied GF002 Sale row Value3/Qty = 1,039.50 while D2/D5 show
            # 1,050.  BUSY's final All-MC price proves Value3 is the rate that
            # participates in the negative-stock weighted balance.
            $value3 = [Math]::Abs([double]$Row.value3)
            if ($value3 -gt $eps) {
                return $value3 / $Qty
            }

            $d5 = [Math]::Abs([double]$Row.d5)
            if ($d5 -gt $eps) {
                return $d5 / $Qty
            }

            $d2 = [Math]::Abs([double]$Row.d2)
            if ($d2 -gt $eps) {
                return $d2
            }

            return 0.0
        }

        function Get-BusyWaExplicitValue {
            param(
                $Row,
                [double]$Qty
            )

            # BUSY inventory-cost fields:
            #
            # D5 is the strongest observed inventory valuation amount.
            # D2 behaves as BUSY's inventory unit cost/rate.
            #
            # IMPORTANT: Value3 is NOT always the stock-cost amount.
            # A proven BUSY Sale row contained:
            #
            #   Qty    = 10
            #   Value3 = 400
            #   D2     = 50
            #   D5     = 500
            #
            # BUSY stock valuation removes 500 (10 x 50), not 400.
            # Therefore D5/D2 must win over Value3.
            $d5 = [Math]::Abs(
                [double]$Row.d5
            )

            if ($d5 -gt $eps) {
                return $d5
            }

            $d2 = [Math]::Abs(
                [double]$Row.d2
            )

            if ($d2 -gt $eps) {
                return (
                    [double]$Qty *
                    $d2
                )
            }

            # Fallback only when BUSY did not post a D5 amount or D2 rate.
            $value3 = [Math]::Abs(
                [double]$Row.value3
            )

            if ($value3 -gt $eps) {
                return $value3
            }

            return 0.0
        }

        function Get-BusyWaExplicitRate {
            param(
                $Row,
                [double]$Qty
            )

            if ($Qty -le $eps) {
                return 0.0
            }

            $value = Get-BusyWaExplicitValue `
                -Row $Row `
                -Qty $Qty

            if ($value -le $eps) {
                return 0.0
            }

            return (
                [double]$value /
                [double]$Qty
            )
        }

        # --------------------------------------------------------
        # Parent Purchase/Sale -> Return links.
        # --------------------------------------------------------
        #
        # Method=2 is the standard linked-return relation seen in BUSY.
        # Some datasets also expose Method=3 reference rows.  Method 3 is
        # included only as a fallback relation and DISTINCT prevents
        # duplicate parent/return pairs.
        # --------------------------------------------------------
        try {
            $returnLinkQry = @"
SELECT DISTINCT
    P.VchCode AS ParentVchCode,
    R.VchCode AS ReturnVchCode
FROM Tran3 P, Tran3 R, Tran1 H
WHERE
    P.RefCode = R.RefCode
    AND P.VchCode <> R.VchCode
    AND P.Method = 1
    AND (R.Method = 2 OR R.Method = 3)
    AND R.VchCode = H.VchCode
    AND (H.VchType = 3 OR H.VchType = 10)
"@

            $returnLinkRst =
                $fi.GetRecordset($returnLinkQry)

            if (
                $returnLinkRst -and
                -not $returnLinkRst.EOF
            ) {
                $returnLinkRst.MoveFirst()

                while (-not $returnLinkRst.EOF) {
                    $parentCode = ConvertTo-ReportInt (
                        Get-ReportFieldValue `
                            $returnLinkRst `
                            "ParentVchCode" `
                            0
                    )

                    $returnCode = ConvertTo-ReportInt (
                        Get-ReportFieldValue `
                            $returnLinkRst `
                            "ReturnVchCode" `
                            0
                    )

                    if (
                        $parentCode -gt 0 -and
                        $returnCode -gt 0 -and
                        -not $busyReturnParentMap.ContainsKey(
                            [string]$returnCode
                        )
                    ) {
                        $busyReturnParentMap[
                            [string]$returnCode
                        ] = $parentCode
                    }

                    $returnLinkRst.MoveNext()
                }
            }

            Close-ReportRecordset $returnLinkRst
        }
        catch {
            # Return linking improves exact parent-return valuation.
            # If an older BUSY schema does not expose the expected Tran3
            # relation, the engine safely falls back to current/explicit
            # valuation for that return.
        }

        # --------------------------------------------------------
        # Initialize method-5 pools from financial-year opening stock.
        # --------------------------------------------------------
        foreach ($itemCodeRaw in @($itemOrder)) {
            $itemCode = [int]$itemCodeRaw

            if (-not $items.ContainsKey($itemCode)) {
                continue
            }

            $item = $items[$itemCode]

            foreach ($mcKey in @($item.mc.Keys)) {
                $mcCode = 0

                [void][int]::TryParse(
                    [string]$mcKey,
                    [ref]$mcCode
                )

                $mcState =
                    $item.mc[[string]$mcCode]

                $openingQty =
                    [double]$mcState.baseOpeningMain

                $openingValue =
                    [double]$mcState.baseOpeningValue

                $waState = Get-BusyWaState `
                    -ItemCode $itemCode `
                    -McCode $mcCode

                $waState.ledgerQty =
                    [double]$waState.ledgerQty +
                    $openingQty

                if ([Math]::Abs($openingQty) -gt $eps) {
                    # Opening stock is already a signed BUSY valuation
                    # balance.  Preserve negative opening quantity/value too.
                    $waState.poolQty = [double]$waState.poolQty + $openingQty
                    $waState.poolValue = [double]$waState.poolValue + $openingValue

                    if ([Math]::Abs([double]$waState.poolQty) -gt $eps) {
                        $waState.lastRate = [Math]::Abs(
                            [double]$waState.poolValue /
                            [double]$waState.poolQty
                        )
                    }
                }
            }
        }

        function Process-BusyWaItemRows {
            param(
                [int]$ItemCode,
                [System.Collections.ArrayList]$Rows
            )

            if (
                $ItemCode -le 0 -or
                $null -eq $Rows -or
                $Rows.Count -eq 0
            ) {
                return
            }

            # Physical closing quantity follows every signed inventory row.
            foreach ($row in $Rows) {
                $state = Get-BusyWaState `
                    -ItemCode $ItemCode `
                    -McCode ([int]$row.mcCode)

                $state.ledgerQty =
                    [double]$state.ledgerQty +
                    [double]$row.qty
            }

            $purchaseSourceByVch = @{}
            $saleRateByVch = @{}

            $dateGroups = @(
                $Rows |
                Group-Object voucherDate |
                Sort-Object Name
            )

            foreach ($dateGroup in $dateGroups) {
                $dateRows = @($dateGroup.Group)

                # ====================================================
                # PHASE 1
                # Direct valuation-bearing inward rows +
                # Purchase Returns.
                # ====================================================
                foreach ($row in $dateRows) {
                    $qty = [double]$row.qty
                    $vchType = [int]$row.vchType
                    $mcCode = [int]$row.mcCode
                    $vchCode = [int]$row.vchCode

                    if (
                        $qty -gt $eps -and
                        $vchType -ne 3
                    ) {
                        $explicitValue =
                            Get-BusyWaExplicitValue `
                                -Row $row `
                                -Qty $qty

                        if ($explicitValue -gt $eps) {
                            Add-BusyWaPool `
                                -ItemCode $ItemCode `
                                -McCode $mcCode `
                                -Qty $qty `
                                -Value $explicitValue

                            if ($vchType -eq 2) {
                                $purchaseRate =
                                    $explicitValue /
                                    $qty

                                $purchaseKey =
                                    [string]$vchCode

                                if (
                                    -not $purchaseSourceByVch.ContainsKey(
                                        $purchaseKey
                                    )
                                ) {
                                    $purchaseSourceByVch[
                                        $purchaseKey
                                    ] = [System.Collections.ArrayList]::new()
                                }

                                [void]$purchaseSourceByVch[
                                    $purchaseKey
                                ].Add(
                                    [pscustomobject]@{
                                        mcCode = $mcCode
                                        originalQty = $qty
                                        rate = [double]$purchaseRate
                                        returnedQty = 0.0
                                    }
                                )
                            }
                        }
                    }

                    # --------------------------------------------
                    # Purchase Return
                    # --------------------------------------------
                    if (
                        $qty -lt -$eps -and
                        $vchType -eq 10
                    ) {
                        $returnQty =
                            [Math]::Abs($qty)

                        $parentCode = 0

                        if (
                            $busyReturnParentMap.ContainsKey(
                                [string]$vchCode
                            )
                        ) {
                            $parentCode =
                                [int]$busyReturnParentMap[
                                    [string]$vchCode
                                ]
                        }

                        $remainingReturnQty =
                            $returnQty

                        if (
                            $parentCode -gt 0 -and
                            $purchaseSourceByVch.ContainsKey(
                                [string]$parentCode
                            )
                        ) {
                            foreach (
                                $source in
                                $purchaseSourceByVch[
                                    [string]$parentCode
                                ]
                            ) {
                                if (
                                    $remainingReturnQty -le
                                    $eps
                                ) {
                                    break
                                }

                                if (
                                    [int]$source.mcCode -ne
                                    $mcCode
                                ) {
                                    continue
                                }

                                $available =
                                    [double]$source.originalQty -
                                    [double]$source.returnedQty

                                if ($available -le $eps) {
                                    continue
                                }

                                $take = [Math]::Min(
                                    $available,
                                    $remainingReturnQty
                                )

                                Remove-BusyWaPool `
                                    -ItemCode $ItemCode `
                                    -McCode $mcCode `
                                    -Qty $take `
                                    -UnitRate ([double]$source.rate)

                                $source.returnedQty =
                                    [double]$source.returnedQty +
                                    $take

                                $remainingReturnQty -=
                                    $take
                            }
                        }

                        if (
                            $remainingReturnQty -gt
                            $eps
                        ) {
                            $returnRate =
                                Get-BusyWaExplicitRate `
                                    -Row $row `
                                    -Qty $remainingReturnQty

                            if ($returnRate -le $eps) {
                                $returnRate =
                                    Get-BusyWaRate `
                                        -ItemCode $ItemCode `
                                        -McCode $mcCode
                            }

                            Remove-BusyWaPool `
                                -ItemCode $ItemCode `
                                -McCode $mcCode `
                                -Qty $remainingReturnQty `
                                -UnitRate $returnRate
                        }
                    }
                }

                # ====================================================
                # PHASE 2
                # Zero-valued Transfer / Production / Stock-Journal
                # inward rows.
                # ====================================================
                $voucherGroups = @(
                    $dateRows |
                    Group-Object vchCode
                )

                foreach ($voucherGroup in $voucherGroups) {
                    $voucherRows =
                        @($voucherGroup.Group)

                    $zeroValueInRows = @(
                        $voucherRows |
                        Where-Object {
                            $candidateQty =
                                [double]$_.qty

                            $candidateValue =
                                Get-BusyWaExplicitValue `
                                    -Row $_ `
                                    -Qty (
                                        [Math]::Abs(
                                            $candidateQty
                                        )
                                    )

                            (
                                $candidateQty -gt $eps -and
                                [int]$_.vchType -ne 3 -and
                                $candidateValue -le $eps
                            )
                        }
                    )

                    if ($zeroValueInRows.Count -eq 0) {
                        continue
                    }

                    $sourceValue = 0.0

                    foreach ($outRow in @(
                        $voucherRows |
                        Where-Object {
                            [double]$_.qty -lt -$eps
                        }
                    )) {
                        $outQty =
                            [Math]::Abs(
                                [double]$outRow.qty
                            )

                        $sourceRate =
                            Get-BusyWaRate `
                                -ItemCode $ItemCode `
                                -McCode ([int]$outRow.mcCode)

                        # Negative/empty source MC:
                        # use the row's own posted cost hint when BUSY has
                        # no positive valuation pool there.
                        if ($sourceRate -le $eps) {
                            $sourceRate =
                                Get-BusyWaExplicitRate `
                                    -Row $outRow `
                                    -Qty $outQty
                        }

                        $sourceValue +=
                            $outQty *
                            $sourceRate
                    }

                    $inQtyTotal = 0.0

                    foreach ($inRow in $zeroValueInRows) {
                        $inQtyTotal +=
                            [double]$inRow.qty
                    }

                    foreach ($inRow in $zeroValueInRows) {
                        $qty =
                            [double]$inRow.qty

                        $mcCode =
                            [int]$inRow.mcCode

                        $value = 0.0

                        if (
                            $sourceValue -gt $eps -and
                            $inQtyTotal -gt $eps
                        ) {
                            $value =
                                $sourceValue *
                                (
                                    $qty /
                                    $inQtyTotal
                                )
                        }
                        else {
                            # BUSY parity: a standalone zero-valued inward
                            # remains zero. Do NOT manufacture current-rate
                            # value for Production/Stock Journal generated
                            # quantity.
                            $value = 0.0
                        }

                        Add-BusyWaPool `
                            -ItemCode $ItemCode `
                            -McCode $mcCode `
                            -Qty $qty `
                            -Value $value
                    }
                }

                # ====================================================
                # PHASE 3
                # Normal outward movements.
                # ====================================================
                foreach ($row in $dateRows) {
                    $rowQty = [double]$row.qty
                    $vchType = [int]$row.vchType

                    if (
                        $rowQty -lt -$eps -and
                        $vchType -ne 10
                    ) {
                        $vchCode =
                            [int]$row.vchCode

                        $mcCode =
                            [int]$row.mcCode

                        $outQty =
                            [Math]::Abs($rowQty)

                        $stateBeforeOut = Get-BusyWaState `
                            -ItemCode $ItemCode `
                            -McCode $mcCode

                        if ([double]$stateBeforeOut.poolQty -gt $eps) {
                            # Positive before issue: current moving average
                            # applies to the complete outward row, even if the
                            # transaction crosses stock below zero.
                            $outRate =
                                Get-BusyWaRate `
                                    -ItemCode $ItemCode `
                                    -McCode $mcCode
                        }
                        else {
                            # Already zero/negative: use BUSY's posted stock
                            # value for this row (Value3 first), so successive
                            # negative issues form a signed weighted average.
                            $outRate =
                                Get-BusyWaNegativeOutRate `
                                    -Row $row `
                                    -Qty $outQty

                            if ($outRate -le $eps) {
                                $outRate =
                                    Get-BusyWaRate `
                                        -ItemCode $ItemCode `
                                        -McCode $mcCode
                            }
                        }

                        if ($vchType -eq 9) {
                            $saleKey =
                                [string]$vchCode

                            if (
                                -not $saleRateByVch.ContainsKey(
                                    $saleKey
                                )
                            ) {
                                $saleRateByVch[
                                    $saleKey
                                ] = [System.Collections.ArrayList]::new()
                            }

                            [void]$saleRateByVch[
                                $saleKey
                            ].Add(
                                [pscustomobject]@{
                                    mcCode = $mcCode
                                    qty = $outQty
                                    rate = [double]$outRate
                                }
                            )
                        }

                        Remove-BusyWaPool `
                            -ItemCode $ItemCode `
                            -McCode $mcCode `
                            -Qty $outQty `
                            -UnitRate $outRate
                    }
                }

                # ====================================================
                # PHASE 4
                # Sale Returns.
                # ====================================================
                foreach ($row in $dateRows) {
                    if (
                        [double]$row.qty -gt $eps -and
                        [int]$row.vchType -eq 3
                    ) {
                        $qty =
                            [double]$row.qty

                        $mcCode =
                            [int]$row.mcCode

                        $vchCode =
                            [int]$row.vchCode

                        $parentCode = 0

                        if (
                            $busyReturnParentMap.ContainsKey(
                                [string]$vchCode
                            )
                        ) {
                            $parentCode =
                                [int]$busyReturnParentMap[
                                    [string]$vchCode
                                ]
                        }

                        $returnRate = 0.0

                        if (
                            $parentCode -gt 0 -and
                            $saleRateByVch.ContainsKey(
                                [string]$parentCode
                            )
                        ) {
                            $matchedQty = 0.0
                            $matchedValue = 0.0

                            foreach (
                                $salePart in
                                $saleRateByVch[
                                    [string]$parentCode
                                ]
                            ) {
                                if (
                                    [int]$salePart.mcCode -ne
                                    $mcCode
                                ) {
                                    continue
                                }

                                $remaining =
                                    [Math]::Max(
                                        0.0,
                                        $qty - $matchedQty
                                    )

                                $take = [Math]::Min(
                                    [double]$salePart.qty,
                                    $remaining
                                )

                                if ($take -le $eps) {
                                    continue
                                }

                                $matchedQty += $take

                                $matchedValue +=
                                    $take *
                                    [double]$salePart.rate

                                if (
                                    $matchedQty -ge
                                    $qty - $eps
                                ) {
                                    break
                                }
                            }

                            if ($matchedQty -gt $eps) {
                                $returnRate =
                                    $matchedValue /
                                    $matchedQty
                            }
                        }

                        if ($returnRate -le $eps) {
                            $returnRate =
                                Get-BusyWaRate `
                                    -ItemCode $ItemCode `
                                    -McCode $mcCode
                        }

                        if ($returnRate -le $eps) {
                            $returnRate =
                                Get-BusyWaExplicitRate `
                                    -Row $row `
                                    -Qty $qty
                        }

                        Add-BusyWaPool `
                            -ItemCode $ItemCode `
                            -McCode $mcCode `
                            -Qty $qty `
                            -Value (
                                $qty *
                                $returnRate
                            )
                    }
                }
            }
        }

        # --------------------------------------------------------
        # One chronological transaction scan for every filtered item.
        # --------------------------------------------------------
        $stockWaWatch = New-ReportEngineWatch
        Write-ReportEngineLog -Stage "STOCK-WA" -Message "Starting single Tran2 valuation + movement scan"

        $busyWaTxnQry = @"
SELECT
    T.MasterCode1 AS ItemCode,
    T.VchCode,
    T.MasterCode2 AS MCCode,
    T.SrNo,
    T.Value1,
    T.Value2,
    T.Value3,
    T.D2,
    T.D5,
    V.VchType,
    V.[Date] AS VoucherDate
FROM Tran2 T
LEFT JOIN Tran1 V ON V.VchCode = T.VchCode
WHERE
    T.RecType = 2
    $tran2ItemPredicate
    AND (V.[Date] IS NULL OR V.[Date] <= $endLiteral)
ORDER BY
    T.MasterCode1,
    T.VchCode,
    T.SrNo
"@

        $busyWaTxnRst =
            $fi.GetRecordset($busyWaTxnQry)

        $currentBusyWaItemCode = 0
        $currentBusyWaRows =
            [System.Collections.ArrayList]::new()

        if (
            $busyWaTxnRst -and
            -not $busyWaTxnRst.EOF
        ) {
            $busyWaTxnRst.MoveFirst()

            # Cache COM Field objects once. Repeated Fields.Item("...") lookups
            # inside a large Access recordset are surprisingly expensive.
            # Each Field.Value follows the recordset cursor automatically.
            $fieldItemCode = $busyWaTxnRst.Fields.Item("ItemCode")
            $fieldVchCode = $busyWaTxnRst.Fields.Item("VchCode")
            $fieldMcCode = $busyWaTxnRst.Fields.Item("MCCode")
            $fieldSrNo = $busyWaTxnRst.Fields.Item("SrNo")
            $fieldValue1 = $busyWaTxnRst.Fields.Item("Value1")
            $fieldValue2 = $busyWaTxnRst.Fields.Item("Value2")
            $fieldValue3 = $busyWaTxnRst.Fields.Item("Value3")
            $fieldD2 = $busyWaTxnRst.Fields.Item("D2")
            $fieldD5 = $busyWaTxnRst.Fields.Item("D5")
            $fieldVchType = $busyWaTxnRst.Fields.Item("VchType")
            $fieldVoucherDate = $busyWaTxnRst.Fields.Item("VoucherDate")

            while (-not $busyWaTxnRst.EOF) {
                $itemCode =
                    ConvertTo-ReportInt $fieldItemCode.Value

                if (-not $items.ContainsKey($itemCode)) {
                    $busyWaTxnRst.MoveNext()
                    continue
                }

                if (
                    $currentBusyWaItemCode -ne 0 -and
                    $itemCode -ne $currentBusyWaItemCode
                ) {
                    Process-BusyWaItemRows `
                        -ItemCode $currentBusyWaItemCode `
                        -Rows $currentBusyWaRows

                    $currentBusyWaRows =
                        [System.Collections.ArrayList]::new()
                }

                $currentBusyWaItemCode =
                    $itemCode

                # Read each COM field once. Besides feeding method-5 valuation,
                # these same scalar values now power opening/prior and movement
                # summaries, removing three extra Tran2 queries.
                $rowVchCode = ConvertTo-ReportInt $fieldVchCode.Value
                $rowMcCode = ConvertTo-ReportInt $fieldMcCode.Value
                $rowSrNo = ConvertTo-ReportInt $fieldSrNo.Value
                $rowQty = ConvertTo-ReportDouble $fieldValue1.Value
                $rowAltQty = ConvertTo-ReportDouble $fieldValue2.Value
                $rowValue3 = ConvertTo-ReportDouble $fieldValue3.Value
                $rowD2 = ConvertTo-ReportDouble $fieldD2.Value
                $rowD5 = ConvertTo-ReportDouble $fieldD5.Value
                $rowVchType = ConvertTo-ReportInt $fieldVchType.Value

                $voucherDateText = ""
                $voucherDateValue = $null

                try {
                    $voucherDateRaw = $fieldVoucherDate.Value

                    if (
                        $null -ne $voucherDateRaw -and
                        $voucherDateRaw -ne [System.DBNull]::Value
                    ) {
                        $voucherDateValue = [datetime]$voucherDateRaw
                        $voucherDateText = $voucherDateValue.ToString("yyyy-MM-dd")
                    }
                }
                catch {
                    $voucherDateValue = $null
                    $voucherDateText = ""
                }

                $movementState = Get-StockMcState `
                    -Item $items[$itemCode] `
                    -McCode $rowMcCode

                # Match the old prior query exactly: NULL-date rows were part
                # of opening/prior; dated rows before From are prior movement.
                if (
                    $null -eq $voucherDateValue -or
                    $voucherDateValue.Date -lt $fromDate.Date
                ) {
                    $movementState.priorMain += $rowQty
                    $movementState.priorAlt += $rowAltQty
                    $movementState.priorValue += $rowValue3
                }
                elseif (
                    $voucherDateValue.Date -le $endDate.Date
                ) {
                    $absMainQty = [Math]::Abs([double]$rowQty)
                    $absAltQty = [Math]::Abs([double]$rowAltQty)
                    $absStockValue = [Math]::Abs([double]$rowValue3)

                    if ($rowQty -gt $eps) {
                        $movementState.inwardMain += $absMainQty
                        $movementState.inwardAlt += $absAltQty
                        $movementState.inwardValue += $absStockValue

                        switch ($rowVchType) {
                            2 { $movementState.purchaseQty += $absMainQty }
                            3 { $movementState.saleReturnQty += $absMainQty }
                            4 { $movementState.materialReceiptQty += $absMainQty }
                            5 {
                                $movementState.transferInQty += $absMainQty
                                $movementState.transferInAltQty += $absAltQty
                                $movementState.transferInValue += $absStockValue
                            }
                            6 { $movementState.productionGeneratedQty += $absMainQty }
                            8 { $movementState.stockJournalGeneratedQty += $absMainQty }
                            default { $movementState.otherInQty += $absMainQty }
                        }
                    }
                    elseif ($rowQty -lt -$eps) {
                        $movementState.outwardMain += $absMainQty
                        $movementState.outwardAlt += $absAltQty
                        $movementState.outwardValue += $absStockValue

                        switch ($rowVchType) {
                            5 {
                                $movementState.transferOutQty += $absMainQty
                                $movementState.transferOutAltQty += $absAltQty
                                $movementState.transferOutValue += $absStockValue
                            }
                            6 { $movementState.productionConsumedQty += $absMainQty }
                            8 { $movementState.stockJournalConsumedQty += $absMainQty }
                            9 { $movementState.saleQty += $absMainQty }
                            10 { $movementState.purchaseReturnQty += $absMainQty }
                            11 { $movementState.materialIssueQty += $absMainQty }
                            default { $movementState.otherOutQty += $absMainQty }
                        }
                    }
                }

                [void]$currentBusyWaRows.Add(
                    [pscustomobject]@{
                        vchCode = $rowVchCode
                        mcCode = $rowMcCode
                        srNo = $rowSrNo
                        qty = $rowQty
                        value3 = $rowValue3
                        d2 = $rowD2
                        d5 = $rowD5
                        vchType = $rowVchType
                        voucherDate = $voucherDateText
                    }
                )

                $busyWaTxnRst.MoveNext()
            }

            if (
                $currentBusyWaItemCode -gt 0 -and
                $currentBusyWaRows.Count -gt 0
            ) {
                Process-BusyWaItemRows `
                    -ItemCode $currentBusyWaItemCode `
                    -Rows $currentBusyWaRows
            }
        }

        Close-ReportRecordset $busyWaTxnRst

        Stop-ReportEngineWatch `
            -Watch $stockWaWatch `
            -Stage "STOCK-WA" `
            -Message "Single Tran2 valuation + movement scan completed"

        # --------------------------------------------------------
        # Alternate-unit quantities.
        #
        # Do NOT call GetMasterXML once per item here.
        #
        # BUSY already posts the alternate quantity we need:
        #   opening stock : Tran4.D2
        #   transactions  : Tran2.Value2
        #
        # Reading thousands of item XML documents through COM was the
        # dominant reason this report took several minutes to load.
        # Using the posted quantities is also more faithful to BUSY because
        # the report displays the quantities BUSY actually stored.
        # --------------------------------------------------------

        # --------------------------------------------------------
        # Build final item rows.
        # --------------------------------------------------------
        $allRows = [System.Collections.Generic.List[object]]::new()

        foreach ($itemCode in @($itemOrder)) {
            $item = $items[[int]$itemCode]
            $centres = [System.Collections.Generic.List[object]]::new()

            $openingMain = 0.0
            $openingAlt = 0.0
            $openingValue = 0.0
            $inwardMain = 0.0
            $inwardAlt = 0.0
            $inwardValue = 0.0
            $outwardMain = 0.0
            $outwardAlt = 0.0
            $outwardValue = 0.0

            # Correct BUSY method-5 closing valuation is accumulated
            # from the MC-specific moving-average snapshots below.
            $busyClosingValue = 0.0

            $purchaseQty = 0.0
            $saleQty = 0.0
            $saleReturnQty = 0.0
            $purchaseReturnQty = 0.0
            $transferInQty = 0.0
            $transferOutQty = 0.0
            $productionGeneratedQty = 0.0
            $productionConsumedQty = 0.0
            $stockJournalGeneratedQty = 0.0
            $stockJournalConsumedQty = 0.0
            $materialReceiptQty = 0.0
            $materialIssueQty = 0.0
            $otherInQty = 0.0
            $otherOutQty = 0.0

            $candidateMcCodes = [System.Collections.ArrayList]::new()
            foreach ($mcKey in @($item.mc.Keys)) {
                $mcCode = 0
                [int]::TryParse([string]$mcKey, [ref]$mcCode) | Out-Null
                if (Test-StockMcSelected -McCode $mcCode) {
                    [void]$candidateMcCodes.Add($mcCode)
                }
            }

            # If one/some MCs were explicitly selected, preserve a zero row
            # for those centres even when the item has no movements there.
            if (-not $useAllMc) {
                foreach ($selectedCode in @($selectedMcCodes.Keys)) {
                    if (-not $candidateMcCodes.Contains([int]$selectedCode)) {
                        [void]$candidateMcCodes.Add([int]$selectedCode)
                    }
                }
            }

            foreach ($mcCodeRaw in @($candidateMcCodes | Sort-Object -Unique)) {
                $mcCode = [int]$mcCodeRaw
                $state = if ($item.mc.ContainsKey([string]$mcCode)) {
                    $item.mc[[string]$mcCode]
                }
                else {
                    New-StockMcState
                }

                $mcOpeningMain = [double]$state.baseOpeningMain + [double]$state.priorMain
                $mcOpeningAlt = [double]$state.baseOpeningAlt + [double]$state.priorAlt
                $mcOpeningValue = [double]$state.baseOpeningValue + [double]$state.priorValue

                $mcInwardMain = [double]$state.inwardMain
                $mcInwardAlt = [double]$state.inwardAlt
                $mcInwardValue = [double]$state.inwardValue
                $mcOutwardMain = [double]$state.outwardMain
                $mcOutwardAlt = [double]$state.outwardAlt
                $mcOutwardValue = [double]$state.outwardValue

                if (-not $IncludeStockTransfers) {
                    $mcInwardMain = [Math]::Max(0.0, $mcInwardMain - [double]$state.transferInQty)
                    $mcInwardAlt = [Math]::Max(0.0, $mcInwardAlt - [double]$state.transferInAltQty)
                    $mcInwardValue = [Math]::Max(0.0, $mcInwardValue - [double]$state.transferInValue)
                    $mcOutwardMain = [Math]::Max(0.0, $mcOutwardMain - [double]$state.transferOutQty)
                    $mcOutwardAlt = [Math]::Max(0.0, $mcOutwardAlt - [double]$state.transferOutAltQty)
                    $mcOutwardValue = [Math]::Max(0.0, $mcOutwardValue - [double]$state.transferOutValue)
                }

                $mcClosingMain = $mcOpeningMain + $mcInwardMain - $mcOutwardMain
                $mcClosingAlt = $mcOpeningAlt + $mcInwardAlt - $mcOutwardAlt

                # Posted ledger value is retained for detailed movement
                # diagnostics, but BUSY's report Price/Value comes from
                # the method-5 moving valuation snapshot.
                $mcPostedClosingValue =
                    $mcOpeningValue +
                    $mcInwardValue -
                    $mcOutwardValue

                $busyWaKey = Get-BusyWaKey `
                    -ItemCode ([int]$item.code) `
                    -McCode $mcCode

                $mcRate = 0.0

                if ($busyWaStates.ContainsKey($busyWaKey)) {
                    $mcRate =
                        Get-BusyWaRate `
                            -ItemCode ([int]$item.code) `
                            -McCode $mcCode
                }

                # With stock transfers included (the BUSY default), the
                # signed WA state is the exact valuation balance.  Using it
                # directly preserves negative-stock weighted values instead
                # of reconstructing them from a clamped positive pool.
                if (
                    $IncludeStockTransfers -and
                    $busyWaStates.ContainsKey($busyWaKey)
                ) {
                    $mcClosingValue =
                        [double]$busyWaStates[$busyWaKey].poolValue
                }
                else {
                    $mcClosingValue =
                        $mcClosingMain *
                        $mcRate
                }

                $busyClosingValue +=
                    $mcClosingValue

                $mcName = if ($mcMap.ContainsKey($mcCode)) {
                    [string]$mcMap[$mcCode]
                }
                else {
                    "Default"
                }

                $centres.Add([pscustomobject]@{
                    code = $mcCode
                    name = $mcName
                    openingQuantity = [Math]::Round($mcOpeningMain, 3)
                    openingAltQuantity = [Math]::Round($mcOpeningAlt, 3)
                    inwardQuantity = [Math]::Round($mcInwardMain, 3)
                    inwardAltQuantity = [Math]::Round($mcInwardAlt, 3)
                    outwardQuantity = [Math]::Round($mcOutwardMain, 3)
                    outwardAltQuantity = [Math]::Round($mcOutwardAlt, 3)
                    closingQuantity = [Math]::Round($mcClosingMain, 3)
                    closingAltQuantity = [Math]::Round($mcClosingAlt, 3)
                    quantity = [Math]::Round($mcClosingMain, 3)
                    stockValue = [Math]::Round($mcClosingValue, 2)
                    valuationRate = [Math]::Round($mcRate, 6)
                    postedLedgerValue = [Math]::Round($mcPostedClosingValue, 2)
                })

                $openingMain += $mcOpeningMain
                $openingAlt += $mcOpeningAlt
                $openingValue += $mcOpeningValue
                $inwardMain += $mcInwardMain
                $inwardAlt += $mcInwardAlt
                $inwardValue += $mcInwardValue
                $outwardMain += $mcOutwardMain
                $outwardAlt += $mcOutwardAlt
                $outwardValue += $mcOutwardValue

                $purchaseQty += [double]$state.purchaseQty
                $saleQty += [double]$state.saleQty
                $saleReturnQty += [double]$state.saleReturnQty
                $purchaseReturnQty += [double]$state.purchaseReturnQty
                $transferInQty += [double]$state.transferInQty
                $transferOutQty += [double]$state.transferOutQty
                $productionGeneratedQty += [double]$state.productionGeneratedQty
                $productionConsumedQty += [double]$state.productionConsumedQty
                $stockJournalGeneratedQty += [double]$state.stockJournalGeneratedQty
                $stockJournalConsumedQty += [double]$state.stockJournalConsumedQty
                $materialReceiptQty += [double]$state.materialReceiptQty
                $materialIssueQty += [double]$state.materialIssueQty
                $otherInQty += [double]$state.otherInQty
                $otherOutQty += [double]$state.otherOutQty
            }

            $closingMain = $openingMain + $inwardMain - $outwardMain
            $closingAlt = $openingAlt + $inwardAlt - $outwardAlt

            # Raw posted ledger value is useful for diagnostics only.
            # BUSY method-5 Stock Status value is $busyClosingValue.
            $closingLedgerValue =
                $openingValue +
                $inwardValue -
                $outwardValue

            $movementQty = $inwardMain + $outwardMain

            $stockStatus =
                if ($closingMain -lt -$eps) { "negative" }
                elseif ([Math]::Abs($closingMain) -le $eps) { "out-of-stock" }
                elseif ($closingMain -le $LowStockLevel) { "low-stock" }
                else { "in-stock" }

            if ($statusFilter -ne "all" -and $stockStatus -ne $statusFilter) {
                continue
            }

            $passesMasters = switch ($mastersFilter) {
                "moved" { $movementQty -gt $eps }
                "closing" { [Math]::Abs($closingMain) -gt $eps }
                "moved-closing" { ($movementQty -gt $eps -or [Math]::Abs($closingMain) -gt $eps) }
                default { $true }
            }

            if (-not $passesMasters) { continue }
            if (-not $IncludeZero -and [Math]::Abs($closingMain) -le $eps) { continue }

            $busyRate = if ([Math]::Abs($closingMain) -gt $eps) {
                $busyClosingValue / $closingMain
            }
            else {
                0.0
            }

            $valuationRate = switch ($valueMode) {
                "purchase" { [double]$item.purchasePrice }
                "sale" { [double]$item.salePrice }
                "mrp" { [double]$item.mrp }
                default { [double]$busyRate }
            }

            $stockValue = switch ($valueMode) {
                "busy" {
                    [double]$busyClosingValue
                }

                default {
                    [double]$closingMain *
                    [double]$valuationRate
                }
            }

            $parentGroup = if ($item.groupPath.Count -gt 1) {
                [string]$item.groupPath[0]
            }
            elseif ($item.groupPath.Count -eq 1) {
                [string]$item.groupPath[0]
            }
            else { "General" }

            $allRows.Add([pscustomobject]@{
                itemCode = [int]$item.code
                itemName = [string]$item.name
                alias = [string]$item.alias
                group = [string]$item.group
                parentGroup = $parentGroup
                groupPath = @($item.groupPath)

                unit = if ([string]::IsNullOrWhiteSpace([string]$item.mainUnit)) { "Units" } else { [string]$item.mainUnit }
                mainUnit = if ([string]::IsNullOrWhiteSpace([string]$item.mainUnit)) { "Units" } else { [string]$item.mainUnit }
                altUnit = [string]$item.altUnit
                conversionFactor = [Math]::Round([double]$item.conversionFactor, 6)
                conversionType = [int]$item.conversionType

                openingQuantity = [Math]::Round($openingMain, 3)
                openingAltQuantity = [Math]::Round($openingAlt, 3)
                inwardQuantity = [Math]::Round($inwardMain, 3)
                inwardAltQuantity = [Math]::Round($inwardAlt, 3)
                outwardQuantity = [Math]::Round($outwardMain, 3)
                outwardAltQuantity = [Math]::Round($outwardAlt, 3)
                closingQuantity = [Math]::Round($closingMain, 3)
                closingAltQuantity = [Math]::Round($closingAlt, 3)

                quantity = [Math]::Round($closingMain, 3)
                availableQuantity = [Math]::Round($closingMain, 3)
                movementQuantity = [Math]::Round($movementQty, 3)
                lowStockLevel = $LowStockLevel
                status = $stockStatus

                purchaseQuantity = [Math]::Round($purchaseQty, 3)
                saleQuantity = [Math]::Round($saleQty, 3)
                saleReturnQuantity = [Math]::Round($saleReturnQty, 3)
                purchaseReturnQuantity = [Math]::Round($purchaseReturnQty, 3)
                transferInQuantity = [Math]::Round($transferInQty, 3)
                transferOutQuantity = [Math]::Round($transferOutQty, 3)
                productionGeneratedQuantity = [Math]::Round($productionGeneratedQty, 3)
                productionConsumedQuantity = [Math]::Round($productionConsumedQty, 3)
                stockJournalGeneratedQuantity = [Math]::Round($stockJournalGeneratedQty, 3)
                stockJournalConsumedQuantity = [Math]::Round($stockJournalConsumedQty, 3)
                materialReceiptQuantity = [Math]::Round($materialReceiptQty, 3)
                materialIssueQuantity = [Math]::Round($materialIssueQty, 3)
                otherInQuantity = [Math]::Round($otherInQty, 3)
                otherOutQuantity = [Math]::Round($otherOutQty, 3)

                mrp = [Math]::Round([double]$item.mrp, 2)
                salePrice = [Math]::Round([double]$item.salePrice, 2)
                purchasePrice = [Math]::Round([double]$item.purchasePrice, 2)
                busyValuationRate = [Math]::Round($busyRate, 6)
                busyStockValue = [Math]::Round($busyClosingValue, 2)

                # Keep old posted-net value visible for diagnostics so
                # future mismatches can be traced without changing UI.
                postedLedgerClosingValue = [Math]::Round($closingLedgerValue, 2)

                valuationRate = [Math]::Round($valuationRate, 6)
                stockValue = [Math]::Round($stockValue, 2)
                openingValue = [Math]::Round($openingValue, 2)
                inwardValue = [Math]::Round($inwardValue, 2)
                outwardValue = [Math]::Round($outwardValue, 2)

                materialCentres = @($centres | Sort-Object name)
            })
        }

        $sortedRows = @($allRows | Sort-Object itemName, itemCode)
        $totalRows = $sortedRows.Count
        $skip = ($Page - 1) * $PageSize
        $pageRows = @($sortedRows | Select-Object -Skip $skip -First $PageSize)

        # --------------------------------------------------------
        # Group summary powers both BUSY-like Grouped and Hierarchical
        # views without forcing the browser to recalculate aggregate totals.
        # --------------------------------------------------------
        $groupTotalsMap = @{}

        foreach ($row in $sortedRows) {
            $groupKey = if ($row.groupPath.Count -gt 0) {
                (@($row.groupPath) -join " > ")
            }
            elseif ([string]::IsNullOrWhiteSpace([string]$row.group)) {
                "General"
            }
            else {
                [string]$row.group
            }

            if (-not $groupTotalsMap.ContainsKey($groupKey)) {
                $groupTotalsMap[$groupKey] = [pscustomobject]@{
                    group = $groupKey
                    parentGroup = [string]$row.parentGroup
                    groupPath = @($row.groupPath)
                    itemCount = 0
                    openingQuantity = 0.0
                    inwardQuantity = 0.0
                    outwardQuantity = 0.0
                    closingQuantity = 0.0
                    stockValue = 0.0
                }
            }

            $g = $groupTotalsMap[$groupKey]
            $g.itemCount++
            $g.openingQuantity += [double]$row.openingQuantity
            $g.inwardQuantity += [double]$row.inwardQuantity
            $g.outwardQuantity += [double]$row.outwardQuantity
            $g.closingQuantity += [double]$row.closingQuantity
            $g.stockValue += [double]$row.stockValue
        }

        $groupTotals = foreach ($groupName in @($groupTotalsMap.Keys | Sort-Object)) {
            $g = $groupTotalsMap[$groupName]
            [pscustomobject]@{
                group = [string]$g.group
                parentGroup = [string]$g.parentGroup
                groupPath = @($g.groupPath)
                itemCount = [int]$g.itemCount
                openingQuantity = [Math]::Round([double]$g.openingQuantity, 3)
                inwardQuantity = [Math]::Round([double]$g.inwardQuantity, 3)
                outwardQuantity = [Math]::Round([double]$g.outwardQuantity, 3)
                closingQuantity = [Math]::Round([double]$g.closingQuantity, 3)
                stockValue = [Math]::Round([double]$g.stockValue, 2)
            }
        }

        $sum = {
            param([string]$Property)
            $result = @($sortedRows | Measure-Object -Property $Property -Sum).Sum
            if ($null -eq $result) { return 0.0 }
            return [double]$result
        }

        $summary = @{
            totalItems = $sortedRows.Count
            inStockItems = @($sortedRows | Where-Object status -eq "in-stock").Count
            lowStockItems = @($sortedRows | Where-Object status -eq "low-stock").Count
            outOfStockItems = @($sortedRows | Where-Object status -eq "out-of-stock").Count
            negativeStockItems = @($sortedRows | Where-Object status -eq "negative").Count
            totalQuantity = [Math]::Round((& $sum "closingQuantity"), 3)
            openingQuantity = [Math]::Round((& $sum "openingQuantity"), 3)
            inwardQuantity = [Math]::Round((& $sum "inwardQuantity"), 3)
            outwardQuantity = [Math]::Round((& $sum "outwardQuantity"), 3)
            purchaseQuantity = [Math]::Round((& $sum "purchaseQuantity"), 3)
            saleQuantity = [Math]::Round((& $sum "saleQuantity"), 3)
            stockValue = [Math]::Round((& $sum "stockValue"), 2)
            busyStockValue = [Math]::Round((& $sum "busyStockValue"), 2)
        }

        $result = @{
            success = $true
            view = $viewMode
            fromDate = $fromDate.ToString("yyyy-MM-dd")
            toDate = $endDate.ToString("yyyy-MM-dd")
            asOfDate = $asOfDate.ToString("yyyy-MM-dd")
            generatedAt = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
            page = $Page
            pageSize = $PageSize
            total = $totalRows
            totalPages = [Math]::Max(1, [Math]::Ceiling($totalRows / [double]$PageSize))
            summaryScope = "all"
            valuationSource = if ($valueMode -eq "busy") { "busy-stockvalmethod-5-batch-moving-weighted-average" } else { "item-master-$valueMode-price" }
            stockValuationMethod = 5
            stockValuationMode = "moving-weighted-average-by-material-centre"
            options = @{
                unitMode = $unitDisplay
                showValue = $ShowValue
                includeStockTransfers = $IncludeStockTransfers
                showSalePurchaseSeparately = $ShowSalePurchaseSeparately
                mastersMode = $mastersFilter
                showParentGroup = $ShowParentGroup
                valueBy = $valueMode
            }
            filterOptions = @{
                materialCentres = @($mcOptions)
                itemGroups = @($groupOptions | Sort-Object pathLabel, name)
            }
            summary = $summary
            groupTotals = @($groupTotals)
            data = $pageRows
        }

        # Save only successful responses. Failed/partial calculations are never
        # cached. Keep the cache bounded because each full report can contain
        # thousands of material-centre rows.
        $script:StockStatusReportCache[$stockStatusCacheKey] = @{
            createdAt = Get-Date
            instanceId = $InstanceId
            companyCode = $CompanyCode
            result = $result
        }

        if (
            $script:StockStatusReportCache.Count -gt
            $script:StockStatusReportCacheMaxEntries
        ) {
            $oldestEntries = @(
                $script:StockStatusReportCache.GetEnumerator() |
                Sort-Object {
                    try { [datetime]$_.Value.createdAt }
                    catch { [datetime]::MinValue }
                }
            )

            while (
                $script:StockStatusReportCache.Count -gt
                $script:StockStatusReportCacheMaxEntries -and
                $oldestEntries.Count -gt 0
            ) {
                $oldestKey = [string]$oldestEntries[0].Key
                $script:StockStatusReportCache.Remove($oldestKey)
                $oldestEntries = @($oldestEntries | Select-Object -Skip 1)
            }
        }

        Write-ReportEngineLog `
            -Stage "STOCK-CACHE" `
            -Message (
                "Stored Stock Status result; rows=" +
                $totalRows +
                "; cacheEntries=" +
                $script:StockStatusReportCache.Count
            ) `
            -Color "Green"

        Stop-ReportEngineWatch `
            -Watch $stockTotalWatch `
            -Stage "STOCK-TOTAL" `
            -Message ("Cold Stock Status completed; rows=" + $totalRows)

        return $result
    }
    catch {
        $errorLine = 0
        try { $errorLine = [int]$_.InvocationInfo.ScriptLineNumber } catch {}

        return @{
            success = $false
            error = $_.Exception.Message
            exception = $_.Exception.GetType().FullName
            line = $errorLine
            data = @()
        }
    }
    finally {
        Disconnect-BUSY $fi
    }
}

