
# ─────────────────────────────────────────────────────────────
# REPORT ENGINE DIAGNOSTICS AND CACHE
# ─────────────────────────────────────────────────────────────
if ($null -eq $script:OutstandingReportCache) {
    $script:OutstandingReportCache = @{}
}

$script:OutstandingReportCacheSeconds = 20
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
        [string]$AsOf           = "",
        [string]$MaterialCentre = "",
        [string]$ItemGroup      = "",
        [string]$Search         = "",
        [string]$Status         = "all",
        [bool]$IncludeZero      = $true,
        [double]$LowStockLevel  = 5,
        [string]$ValueBy        = "purchase",
        [int]$Page              = 1,
        [int]$PageSize          = 50,
        [string]$InstanceId     = "",
        [string]$CompanyCode    = ""
    )

    if ($Page -lt 1) { $Page = 1 }
    if ($PageSize -lt 1) { $PageSize = 50 }
    if ($PageSize -gt 50) { $PageSize = 50 }

    $asOfDate = ConvertTo-ReportDate -Value $AsOf
    $mode = Get-ReportDatabaseMode -InstanceId $InstanceId
    $dateLiteral = Get-ReportSqlDateLiteral -Date $asOfDate -IsSql $mode.isSql

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) {
        return @{ success = $false; error = "BUSY connection failed" }
    }

    try {
        $itemPage = Get-ReportPagedMasterCodes `
            -Fi $fi `
            -MasterType 6 `
            -Search $Search `
            -Group $ItemGroup `
            -Page $Page `
            -PageSize $PageSize

        $pageItems = @($itemPage.rows)

        if ($pageItems.Count -eq 0) {
            return @{
                success = $true
                asOfDate = $asOfDate.ToString("yyyy-MM-dd")
                generatedAt = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
                page = $Page
                pageSize = $PageSize
                total = $itemPage.total
                totalPages = [Math]::Max(
                    1, [Math]::Ceiling($itemPage.total / $PageSize)
                )
                summaryScope = "page"
                summary = @{
                    totalItems = 0
                    inStockItems = 0
                    lowStockItems = 0
                    outOfStockItems = 0
                    negativeStockItems = 0
                    totalQuantity = 0
                    stockValue = 0
                }
                data = @()
            }
        }

        $items = @{}
        foreach ($entry in $pageItems) {
            $items[[int]$entry.code] = @{
                code = [int]$entry.code
                name = [string]$entry.name
                alias = [string]$entry.alias
                group = [string]$entry.groupName
                unit = ""
                mrp = 0.0
                salePrice = 0.0
                purchasePrice = 0.0
                mc = @{}
            }
        }

        $itemCodes = ($items.Keys | Sort-Object) -join ","

        $mcMap = @{}
        $mcRst = $fi.GetRecordset(
            "SELECT Code, Name FROM Master1 WHERE MasterType = 11"
        )
        if ($mcRst -and -not $mcRst.EOF) {
            $mcRst.MoveFirst()
            while (-not $mcRst.EOF) {
                $code = ConvertTo-ReportInt (
                    Get-ReportFieldValue $mcRst "Code" 0
                )
                if ($code -gt 0) {
                    $mcMap[$code] = [string](
                        Get-ReportFieldValue $mcRst "Name" ""
                    )
                }
                $mcRst.MoveNext()
            }
        }
        Close-ReportRecordset $mcRst

        # Only load XML for the 50 items in the requested page.
        foreach ($itemCode in @($items.Keys)) {
            try {
                $cacheKey =
                    "$InstanceId|$CompanyCode|report-item-master|$itemCode"
                $cached = Get-Cache $cacheKey

                if ($cached) {
                    $items[$itemCode].unit = $cached.unit
                    $items[$itemCode].mrp = $cached.mrp
                    $items[$itemCode].salePrice = $cached.salePrice
                    $items[$itemCode].purchasePrice = $cached.purchasePrice
                    continue
                }

                $xmlString = $fi.GetMasterXML([int]$itemCode)
                if (-not $xmlString) { continue }

                $xml = [xml]$xmlString
                $master = @{
                    unit = [string]$xml.Item.MainUnit
                    mrp = ConvertTo-ReportDouble ([string]$xml.Item.MRP)
                    salePrice =
                        ConvertTo-ReportDouble ([string]$xml.Item.SalePrice)
                    purchasePrice =
                        ConvertTo-ReportDouble ([string]$xml.Item.PurchasePrice)
                }

                $items[$itemCode].unit = $master.unit
                $items[$itemCode].mrp = $master.mrp
                $items[$itemCode].salePrice = $master.salePrice
                $items[$itemCode].purchasePrice = $master.purchasePrice

                Set-Cache -Key $cacheKey -Value $master
            }
            catch {}
        }

        $opQry = @"
SELECT
    MasterCode1 AS ItemCode,
    MasterCode2 AS MCCode,
    SUM(D1) AS Qty
FROM Tran4
WHERE
    RecType = 0
    AND MasterCode1 IN ($itemCodes)
GROUP BY MasterCode1, MasterCode2
"@

        $opRst = $fi.GetRecordset($opQry)
        if ($opRst -and -not $opRst.EOF) {
            $opRst.MoveFirst()
            while (-not $opRst.EOF) {
                $itemCode = ConvertTo-ReportInt (
                    Get-ReportFieldValue $opRst "ItemCode" 0
                )
                $mcCode = ConvertTo-ReportInt (
                    Get-ReportFieldValue $opRst "MCCode" 0
                )
                $qty = ConvertTo-ReportDouble (
                    Get-ReportFieldValue $opRst "Qty" 0
                )

                if ($items.ContainsKey($itemCode)) {
                    if (-not $items[$itemCode].mc.ContainsKey($mcCode)) {
                        $items[$itemCode].mc[$mcCode] = 0.0
                    }
                    $items[$itemCode].mc[$mcCode] += $qty
                }

                $opRst.MoveNext()
            }
        }
        Close-ReportRecordset $opRst

        $txnQry = @"
SELECT
    T.MasterCode1 AS ItemCode,
    T.MasterCode2 AS MCCode,
    SUM(T.Value1) AS Qty
FROM Tran2 T
LEFT JOIN Tran1 V ON V.VchCode = T.VchCode
WHERE
    T.RecType = 2
    AND T.MasterCode1 IN ($itemCodes)
    AND (V.[Date] IS NULL OR V.[Date] <= $dateLiteral)
GROUP BY T.MasterCode1, T.MasterCode2
"@

        $txnRst = $fi.GetRecordset($txnQry)
        if ($txnRst -and -not $txnRst.EOF) {
            $txnRst.MoveFirst()
            while (-not $txnRst.EOF) {
                $itemCode = ConvertTo-ReportInt (
                    Get-ReportFieldValue $txnRst "ItemCode" 0
                )
                $mcCode = ConvertTo-ReportInt (
                    Get-ReportFieldValue $txnRst "MCCode" 0
                )
                $qty = ConvertTo-ReportDouble (
                    Get-ReportFieldValue $txnRst "Qty" 0
                )

                if ($items.ContainsKey($itemCode)) {
                    if (-not $items[$itemCode].mc.ContainsKey($mcCode)) {
                        $items[$itemCode].mc[$mcCode] = 0.0
                    }
                    $items[$itemCode].mc[$mcCode] += $qty
                }

                $txnRst.MoveNext()
            }
        }
        Close-ReportRecordset $txnRst

        $rows = [System.Collections.Generic.List[object]]::new()

        foreach ($entry in $pageItems) {
            $item = $items[[int]$entry.code]
            $centres = [System.Collections.Generic.List[object]]::new()
            $qtyTotal = 0.0

            foreach ($mcCode in $item.mc.Keys) {
                $mcName =
                    if ($mcMap.ContainsKey([int]$mcCode)) {
                        $mcMap[[int]$mcCode]
                    } else { "Default" }

                if (
                    -not [string]::IsNullOrWhiteSpace($MaterialCentre) -and
                    $MaterialCentre -ne "all" -and
                    $mcName.Trim().ToLowerInvariant() -ne
                        $MaterialCentre.Trim().ToLowerInvariant()
                ) {
                    continue
                }

                $qty = [Math]::Round(
                    [double]$item.mc[$mcCode], 3
                )
                $qtyTotal += $qty

                $centres.Add([pscustomobject]@{
                    code = [int]$mcCode
                    name = $mcName
                    quantity = $qty
                })
            }

            if (-not $IncludeZero -and [Math]::Abs($qtyTotal) -lt 0.0005) {
                continue
            }

            $stockStatus =
                if ($qtyTotal -lt 0) { "negative" }
                elseif ([Math]::Abs($qtyTotal) -lt 0.0005) { "out-of-stock" }
                elseif ($qtyTotal -le $LowStockLevel) { "low-stock" }
                else { "in-stock" }

            if ($Status -ne "all" -and $Status -ne $stockStatus) {
                continue
            }

            $rate =
                switch ($ValueBy.ToLowerInvariant()) {
                    "sale" { [double]$item.salePrice }
                    "mrp" { [double]$item.mrp }
                    default { [double]$item.purchasePrice }
                }

            $rows.Add([pscustomobject]@{
                itemCode = $item.code
                itemName = $item.name
                alias = $item.alias
                group = $item.group
                unit = $item.unit
                quantity = [Math]::Round($qtyTotal, 3)
                availableQuantity = [Math]::Round($qtyTotal, 3)
                lowStockLevel = $LowStockLevel
                status = $stockStatus
                mrp = [Math]::Round($item.mrp, 2)
                salePrice = [Math]::Round($item.salePrice, 2)
                purchasePrice = [Math]::Round($item.purchasePrice, 2)
                valuationRate = [Math]::Round($rate, 2)
                stockValue = [Math]::Round($qtyTotal * $rate, 2)
                materialCentres = @($centres | Sort-Object name)
            })
        }

        $pageRows = @($rows | Sort-Object itemName)

        return @{
            success = $true
            asOfDate = $asOfDate.ToString("yyyy-MM-dd")
            generatedAt = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
            page = $Page
            pageSize = $PageSize
            total = $itemPage.total
            totalPages = [Math]::Max(
                1, [Math]::Ceiling($itemPage.total / $PageSize)
            )
            summaryScope = "page"
            summary = @{
                totalItems = $pageRows.Count
                inStockItems =
                    @($pageRows | Where-Object status -eq "in-stock").Count
                lowStockItems =
                    @($pageRows | Where-Object status -eq "low-stock").Count
                outOfStockItems =
                    @($pageRows | Where-Object status -eq "out-of-stock").Count
                negativeStockItems =
                    @($pageRows | Where-Object status -eq "negative").Count
                totalQuantity = [Math]::Round(
                    (@($pageRows | Measure-Object quantity -Sum).Sum), 3
                )
                stockValue = [Math]::Round(
                    (@($pageRows | Measure-Object stockValue -Sum).Sum), 2
                )
            }
            data = $pageRows
        }
    }
    catch {
        return @{
            success = $false
            error = $_.Exception.Message
            data = @()
        }
    }
    finally {
        Disconnect-BUSY $fi
    }
}