
# ─────────────────────────────────────────────────────────────
# REPORT ENGINE DIAGNOSTICS AND CACHE
# ─────────────────────────────────────────────────────────────
if ($null -eq $script:OutstandingReportCache) {
    $script:OutstandingReportCache = @{}
}

if ($null -eq $script:StockStatusReportCache) {
    $script:StockStatusReportCache = @{}
}

# Fast-page caches are separate from the complete-result cache.
# They let the default Balances view return the requested page without
# valuing every stock item in the company first.
if ($null -eq $script:StockStatusFastIndexCache) {
    $script:StockStatusFastIndexCache = @{}
}

if ($null -eq $script:StockStatusFastPageCache) {
    $script:StockStatusFastPageCache = @{}
}

$script:OutstandingReportCacheSeconds = 20

# Stock Status is much more expensive than Outstanding because BUSY method-5
# valuation must replay inventory history chronologically. Keep a short-lived
# in-process result cache so revisiting the page or repeating the same filter
# does not replay thousands of rows again.
$script:StockStatusReportCacheSeconds = 1800
$script:StockStatusReportCacheMaxEntries = 4

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

    foreach ($cacheName in @(
        "StockStatusReportCache",
        "StockStatusFastIndexCache",
        "StockStatusFastPageCache"
    )) {
        $cache = Get-Variable `
            -Name $cacheName `
            -Scope Script `
            -ValueOnly `
            -ErrorAction SilentlyContinue

        if ($null -eq $cache) {
            continue
        }

        foreach ($key in @($cache.Keys)) {
            if ([string]::IsNullOrWhiteSpace($InstanceId)) {
                $cache.Remove($key)
                continue
            }

            $entry = $cache[$key]
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
                $cache.Remove($key)
            }
        }
    }

    return

    # Legacy block retained below only for source compatibility; it is
    # unreachable because the unified cache clearing above already handled
    # every Stock Status cache.
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
    $now = Get-Date

    foreach ($cacheName in @(
        "StockStatusReportCache",
        "StockStatusFastIndexCache",
        "StockStatusFastPageCache"
    )) {
        $cache = Get-Variable `
            -Name $cacheName `
            -Scope Script `
            -ValueOnly `
            -ErrorAction SilentlyContinue

        if ($null -eq $cache) {
            continue
        }

        foreach ($key in @($cache.Keys)) {
            $entry = $cache[$key]

            $createdAt = $null
            try { $createdAt = [datetime]$entry.createdAt } catch {}

            if (
                $null -eq $createdAt -or
                ($now - $createdAt).TotalSeconds -gt
                $script:StockStatusReportCacheSeconds
            ) {
                $cache.Remove($key)
            }
        }
    }

    return

    # Legacy complete-cache expiry code below is intentionally unreachable;
    # the unified loop above expires all Stock Status cache tiers.
    if ($null -eq $script:StockStatusReportCache) {
        return
    }

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


# Build Stock Status rows-per-page choices from the ACTUAL matching report row
# count, not from the raw number of stock masters in Master1. This means a
# report with 2,573 matching rows exposes sizes up to 2,500, then All.
# There is deliberately no hard-coded maximum. After the common sizes, the
# list grows in 5,000-row steps for as long as the matching result requires.
# A value of 0 is the API representation of "All".
function Get-StockStatusPageSizeOptions {
    param(
        [int]$ItemCount
    )

    $options = [System.Collections.Generic.List[int]]::new()

    $baseSizes = @(
        50,
        100,
        250,
        500,
        750,
        1000,
        1250,
        1500,
        1750,
        2000,
        2500,
        3000,
        4000,
        5000,
        7000,
        10000
    )

    foreach ($size in $baseSizes) {
        if ($size -lt $ItemCount) {
            $options.Add([int]$size)
        }
    }

    $dynamicSize = 15000
    while ($dynamicSize -lt $ItemCount) {
        $options.Add([int]$dynamicSize)
        $dynamicSize += 5000
    }

    # 0 means "All".  The frontend renders this as the All option.
    $options.Add(0)

    return @($options)
}


# Apply page/pageSize to a previously calculated complete Stock Status result.
# The expensive BUSY method-5 calculation is cached without page information,
# so moving between pages or changing rows-per-page only slices the cached rows.
function New-StockStatusPagedResponse {
    param(
        $BaseResult,
        $AllRows,
        [int]$Page = 1,
        [int]$PageSize = 100
    )

    $rows = @($AllRows)
    $totalRows = $rows.Count

    if ($Page -lt 1) {
        $Page = 1
    }

    # Negative values are invalid.  Zero is intentionally reserved for All.
    if ($PageSize -lt 0) {
        $PageSize = 100
    }

    # If the requested page size is greater than or equal to the actual
    # matching report row count, it is equivalent to All. This is intentionally
    # based on the report result (for example 2,573 rows), not the raw company
    # stock-master count (for example 4,198 masters).
    if (
        $PageSize -gt 0 -and
        $totalRows -gt 0 -and
        $PageSize -ge $totalRows
    ) {
        $PageSize = 0
    }

    $totalPages = 1
    $pageRows = @()

    if ($PageSize -eq 0) {
        $Page = 1
        $pageRows = @($rows)
    }
    else {
        $totalPages = [Math]::Max(
            1,
            [int][Math]::Ceiling(
                $totalRows / [double]$PageSize
            )
        )

        if ($Page -gt $totalPages) {
            $Page = $totalPages
        }

        $skip = ($Page - 1) * $PageSize
        $pageRows = @(
            $rows |
            Select-Object -Skip $skip -First $PageSize
        )
    }

    $response = @{}
    foreach ($key in @($BaseResult.Keys)) {
        $response[$key] = $BaseResult[$key]
    }

    $response.page = $Page
    $response.pageSize = $PageSize
    $response.total = $totalRows
    $response.totalPages = $totalPages
    $response.data = @($pageRows)

    return $response
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



# ============================================================
# STOCK STATUS DIRECT ACCESS FAST PATH
# ============================================================
# BUSY COM recordsets are reliable but expensive for company-wide
# aggregate/index work.  auth_native.ps1 is loaded before reports.ps1
# and already exposes Open-BdsConnection for Access companies.
#
# Use direct OLE DB only for READ-ONLY report queries. Voucher writes
# remain on the existing BUSY APIs.
# ============================================================

function Invoke-StockStatusDirectTable {
    param(
        [Parameter(Mandatory)]
        [System.Data.OleDb.OleDbConnection]$Connection,

        [Parameter(Mandatory)]
        [string]$Sql
    )

    $command = $Connection.CreateCommand()
    $command.CommandText = $Sql
    $adapter = New-Object System.Data.OleDb.OleDbDataAdapter($command)
    $table = New-Object System.Data.DataTable

    try {
        [void]$adapter.Fill($table)
        return ,$table
    }
    finally {
        try { $adapter.Dispose() } catch {}
        try { $command.Dispose() } catch {}
    }
}

function Get-StockStatusDirectAccessContext {
    param(
        [string]$InstanceId,
        [string]$CompanyCode,
        [int]$FinancialYearStart
    )

    try {
        if (
            -not (Get-Command Get-InstanceConfig -ErrorAction SilentlyContinue) -or
            -not (Get-Command Open-BdsConnection -ErrorAction SilentlyContinue)
        ) {
            return $null
        }

        $instance = Get-InstanceConfig -InstanceId $InstanceId
        if ($null -eq $instance) {
            return $null
        }

        $dbType = 0
        try {
            if ($null -ne $instance.dbType) {
                $dbType = [int]$instance.dbType
            }
        }
        catch {}

        # Direct OLE DB optimization is for Access/BDS only.
        if ($dbType -ne 0) {
            return $null
        }

        $dataPath = [string]$instance.dataPath
        if ([string]::IsNullOrWhiteSpace($dataPath)) {
            return $null
        }

        $companyFolder =
            Join-Path $dataPath $CompanyCode

        $candidates = @(
            (Join-Path $companyFolder ("db1{0}.bds" -f $FinancialYearStart)),
            (Join-Path $companyFolder "db.bds")
        )

        foreach ($dbFile in $candidates) {
            if (-not (Test-Path -LiteralPath $dbFile -PathType Leaf)) {
                continue
            }

            try {
                $connection = Open-BdsConnection -DbFile $dbFile

                if ($null -ne $connection) {
                    $fileInfo = Get-Item -LiteralPath $dbFile

                    return [pscustomobject]@{
                        connection = $connection
                        path = $dbFile
                        writeTicks = [int64]$fileInfo.LastWriteTimeUtc.Ticks
                        length = [int64]$fileInfo.Length
                    }
                }
            }
            catch {
                # COM remains the safe fallback.
            }
        }
    }
    catch {}

    return $null
}

function Get-StockStatusStableHash {
    param([string]$Text)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        $hash = $sha.ComputeHash($bytes)
        return (
            ($hash | ForEach-Object { $_.ToString("x2") }) -join ""
        )
    }
    finally {
        $sha.Dispose()
    }
}


function Get-StockStatusMasterSignature {
    param(
        [Parameter(Mandatory)]
        [System.Data.OleDb.OleDbConnection]$Connection
    )

    try {
        $table =
            Invoke-StockStatusDirectTable `
                -Connection $Connection `
                -Sql @"
SELECT
    COUNT(*) AS RowCount,
    MAX(Code) AS MaxCode,
    MAX(ModificationTime) AS MaxModified
FROM Master1
WHERE MasterType IN (5,6,8,11)
"@

        if ($table.Rows.Count -le 0) {
            return ""
        }

        $row = $table.Rows[0]
        $count = ConvertTo-ReportInt $row["RowCount"]
        $maxCode = ConvertTo-ReportInt $row["MaxCode"]
        $maxModified = ""

        try {
            if (
                $null -ne $row["MaxModified"] -and
                $row["MaxModified"] -ne [System.DBNull]::Value
            ) {
                $maxModified =
                    ([datetime]$row["MaxModified"]).
                        ToUniversalTime().
                        Ticks
            }
        }
        catch {}

        return "$count|$maxCode|$maxModified"
    }
    catch {
        return ""
    }
}

function Get-StockStatusMasterSnapshotPath {
    param(
        [string]$InstanceId,
        [string]$CompanyCode
    )

    try {
        $cacheDir =
            Join-Path `
                (Join-Path $PSScriptRoot "..\cache") `
                "stock-status"

        [void][System.IO.Directory]::CreateDirectory($cacheDir)

        $hash =
            Get-StockStatusStableHash `
                -Text (
                    "master-snapshot-v8|" +
                    $InstanceId +
                    "|" +
                    $CompanyCode
                )

        return Join-Path $cacheDir ($hash + "-masters.json")
    }
    catch {
        return ""
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
        [int]$PageSize                 = 100,
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

    # pageSize=0 means All.  Do not cap the maximum here; the real company
    # item count is queried below and drives the available page-size choices.
    if ($PageSize -lt 0) { $PageSize = 100 }

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
    # FAST RESULT CACHE + DYNAMIC PAGINATION METADATA
    # ------------------------------------------------------------
    # Page and pageSize are deliberately NOT part of the expensive calculation
    # cache key.  One complete calculation is cached and every later page is a
    # cheap server-side slice of that cached result.
    # ------------------------------------------------------------
    Remove-ExpiredStockStatusReportCache

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) {
        return @{ success = $false; error = "BUSY connection failed"; data = @() }
    }

    $directStockContext = $null
    $directStockConn = $null

    try {
        # For Access companies, use the current financial-year BDS directly
        # for read-only aggregate/master queries. This bypasses the slow COM
        # recordset bridge for thousands of rows.
        $directStockContext =
            Get-StockStatusDirectAccessContext `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode `
                -FinancialYearStart $fyYear

        if ($null -ne $directStockContext) {
            $directStockConn = $directStockContext.connection

            Write-ReportEngineLog `
                -Stage "STOCK-DIRECT" `
                -Message (
                    "Direct Access read path active; db=" +
                    [System.IO.Path]::GetFileName(
                        [string]$directStockContext.path
                    )
                ) `
                -Color "Green"
        }

        # Keep the raw number of stock-item masters for diagnostics/UI only.
        # IMPORTANT: this is NOT used to generate rows-per-page choices because
        # the Stock Status report can exclude masters according to the current
        # report filters / masters mode. Page-size choices are based on the
        # final matching report-row count instead.
        $companyItemCount = 0

        if ($null -ne $directStockConn) {
            $itemCountTable =
                Invoke-StockStatusDirectTable `
                    -Connection $directStockConn `
                    -Sql (
                        "SELECT COUNT(*) AS ItemCount " +
                        "FROM Master1 WHERE MasterType=6"
                    )

            if ($itemCountTable.Rows.Count -gt 0) {
                $companyItemCount =
                    ConvertTo-ReportInt (
                        $itemCountTable.Rows[0]["ItemCount"]
                    )
            }
        }
        else {
            $itemCountRst = $null

            try {
                $itemCountRst = $fi.GetRecordset(
                    "SELECT COUNT(*) AS ItemCount FROM Master1 WHERE MasterType=6"
                )

                if ($itemCountRst -and -not $itemCountRst.EOF) {
                    $companyItemCount = ConvertTo-ReportInt (
                        Get-ReportFieldValue $itemCountRst "ItemCount" 0
                    )
                }
            }
            finally {
                Close-ReportRecordset $itemCountRst
            }
        }

        $cacheSeparator = [char]31
        $stockStatusCacheKey = (
            @(
                "stock-wa-v8-busy-parity-fast-cache",
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
                [string]([bool]$ShowParentGroup)
            ) -join $cacheSeparator
        )

        if ($script:StockStatusReportCache.ContainsKey($stockStatusCacheKey)) {
            $cachedEntry = $script:StockStatusReportCache[$stockStatusCacheKey]
            $cacheAgeSeconds = 0.0

            try {
                $cacheAgeSeconds = (
                    (Get-Date) - [datetime]$cachedEntry.createdAt
                ).TotalSeconds
            }
            catch {
                $cacheAgeSeconds =
                    $script:StockStatusReportCacheSeconds + 1
            }

            if (
                $cacheAgeSeconds -le $script:StockStatusReportCacheSeconds -and
                $null -ne $cachedEntry.baseResult -and
                $null -ne $cachedEntry.allRows
            ) {
                # Refresh cheap metadata even when the heavy stock calculation
                # is served from cache. The page-size menu must follow the ACTUAL
                # matching report rows, not the raw Master1 item count.
                $cachedTotalRows = @($cachedEntry.allRows).Count
                $cachedPageSizeOptions = @(
                    Get-StockStatusPageSizeOptions -ItemCount $cachedTotalRows
                )

                $cachedEntry.baseResult.companyItemCount = $companyItemCount
                $cachedEntry.baseResult.pageSizeOptions = @($cachedPageSizeOptions)

                Write-ReportEngineLog `
                    -Stage "STOCK-CACHE" `
                    -Message (
                        "Cache hit; age=" +
                        [Math]::Round($cacheAgeSeconds, 1) +
                        "s; page=" + $Page +
                        "; pageSize=" + $PageSize
                    ) `
                    -Color "Green"

                return New-StockStatusPagedResponse `
                    -BaseResult $cachedEntry.baseResult `
                    -AllRows $cachedEntry.allRows `
                    -Page $Page `
                    -PageSize $PageSize
            }

            # Old-format or expired entries are discarded safely.
            $script:StockStatusReportCache.Remove($stockStatusCacheKey)
        }

        # --------------------------------------------------------
        # FAST PAGE CACHE
        # --------------------------------------------------------
        # The default Balances view is page-oriented. Cache each exact
        # already-valued page independently so revisiting a page is immediate.
        $fastPageModeRequested = (
            $viewMode -eq "balances" -and
            $PageSize -gt 0 -and
            $PageSize -le 250
        )

        $fastPageCacheKey = (
            @(
                [string]$stockStatusCacheKey,
                "fast-page",
                [string]$Page,
                [string]$PageSize
            ) -join $cacheSeparator
        )

        if (
            $fastPageModeRequested -and
            $script:StockStatusFastPageCache.ContainsKey($fastPageCacheKey)
        ) {
            $fastPageEntry = $script:StockStatusFastPageCache[$fastPageCacheKey]
            $fastPageAge = 0.0

            try {
                $fastPageAge = (
                    (Get-Date) - [datetime]$fastPageEntry.createdAt
                ).TotalSeconds
            }
            catch {
                $fastPageAge =
                    $script:StockStatusReportCacheSeconds + 1
            }

            if (
                $fastPageAge -le $script:StockStatusReportCacheSeconds -and
                $null -ne $fastPageEntry.response
            ) {
                $fastCachedResponse = @{}
                foreach ($key in @($fastPageEntry.response.Keys)) {
                    $fastCachedResponse[$key] = $fastPageEntry.response[$key]
                }

                $fastCachedResponse.companyItemCount = $companyItemCount

                Write-ReportEngineLog `
                    -Stage "STOCK-FAST-CACHE" `
                    -Message (
                        "Fast page cache hit; age=" +
                        [Math]::Round($fastPageAge, 1) +
                        "s; page=" + $Page +
                        "; pageSize=" + $PageSize
                    ) `
                    -Color "Green"

                return $fastCachedResponse
            }

            $script:StockStatusFastPageCache.Remove($fastPageCacheKey)
        }

        $stockTotalWatch = New-ReportEngineWatch
        Write-ReportEngineLog `
            -Stage "STOCK-START" `
            -Message (
                "Cold calculation; view=" + $viewMode +
                "; asOf=" + $asOfDate.ToString("yyyy-MM-dd") +
                "; companyItems=" + $companyItemCount
            )

        $mode = Get-ReportDatabaseMode -InstanceId $InstanceId
        $fromLiteral = Get-ReportSqlDateLiteral -Date $fromDate -IsSql $mode.isSql
        $endLiteral = Get-ReportSqlDateLiteral -Date $endDate -IsSql $mode.isSql

        # --------------------------------------------------------
        # Master lookups: item groups and material centres.
        # --------------------------------------------------------
        # --------------------------------------------------------
        # PERSISTENT MASTER METADATA SNAPSHOT
        # --------------------------------------------------------
        # Voucher writes change the BDS file timestamp, so validating master
        # metadata by the database file timestamp invalidated the cache far too
        # often.  Instead use a compact Master1 signature
        # (row-count/max-code/max-modification-time). This lets master/group/
        # unit/MC metadata survive normal voucher activity and bridge restarts.
        # --------------------------------------------------------
        $stockMasterSnapshot = $null
        $stockMasterSnapshotHit = $false
        $stockMasterSnapshotPath = ""
        $stockMasterSignature = ""
        $stockMasterSnapshotBuild = @{
            groups = @()
            materialCentres = @()
            units = @()
            items = @()
        }

        if ($null -ne $directStockConn) {
            $stockMasterSnapshotPath =
                Get-StockStatusMasterSnapshotPath `
                    -InstanceId $InstanceId `
                    -CompanyCode $CompanyCode

            $stockMasterSignature =
                Get-StockStatusMasterSignature `
                    -Connection $directStockConn

            if (
                -not [string]::IsNullOrWhiteSpace(
                    $stockMasterSnapshotPath
                ) -and
                -not [string]::IsNullOrWhiteSpace(
                    $stockMasterSignature
                ) -and
                (Test-Path -LiteralPath $stockMasterSnapshotPath)
            ) {
                try {
                    $candidateSnapshot =
                        Get-Content `
                            -LiteralPath $stockMasterSnapshotPath `
                            -Raw |
                        ConvertFrom-Json

                    if (
                        [string]$candidateSnapshot.engine -eq
                            "stock-master-snapshot-v8" -and
                        [string]$candidateSnapshot.signature -eq
                            $stockMasterSignature -and
                        $null -ne $candidateSnapshot.groups -and
                        $null -ne $candidateSnapshot.materialCentres -and
                        $null -ne $candidateSnapshot.units -and
                        $null -ne $candidateSnapshot.items
                    ) {
                        $stockMasterSnapshot = $candidateSnapshot
                        $stockMasterSnapshotHit = $true

                        Write-ReportEngineLog `
                            -Stage "STOCK-MASTERS" `
                            -Message (
                                "Persistent master snapshot hit; items=" +
                                @($candidateSnapshot.items).Count
                            ) `
                            -Color "Green"
                    }
                }
                catch {}
            }
        }

        $groupMap = @{}
        $groupOptions = [System.Collections.Generic.List[object]]::new()

        if ($null -ne $directStockConn) {
            $groupRowsForLoad = @()

            if ($stockMasterSnapshotHit) {
                $groupRowsForLoad =
                    @($stockMasterSnapshot.groups)
            }
            else {
                $groupTable =
                    Invoke-StockStatusDirectTable `
                        -Connection $directStockConn `
                        -Sql (
                            "SELECT Code, Name, Alias, ParentGrp " +
                            "FROM Master1 WHERE MasterType=5 ORDER BY Name"
                        )

                foreach ($groupRow in $groupTable.Rows) {
                    $groupObj = [pscustomobject]@{
                        code =
                            ConvertTo-ReportInt $groupRow["Code"]
                        name = [string]$groupRow["Name"]
                        alias = [string]$groupRow["Alias"]
                        parentCode =
                            ConvertTo-ReportInt $groupRow["ParentGrp"]
                    }

                    $groupRowsForLoad += $groupObj
                    $stockMasterSnapshotBuild.groups += $groupObj
                }
            }

            foreach ($groupObj in @($groupRowsForLoad)) {
                $code = ConvertTo-ReportInt $groupObj.code
                $parentCode =
                    ConvertTo-ReportInt $groupObj.parentCode
                $name = [string]$groupObj.name
                $alias = [string]$groupObj.alias

                if ($code -gt 0) {
                    $groupMap[$code] = [pscustomobject]@{
                        code = $code
                        name = $name
                        alias = $alias
                        parentCode = $parentCode
                    }
                }
            }
        }
        else {
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
        }

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

        if ($null -ne $directStockConn) {
            $mcRowsForLoad = @()

            if ($stockMasterSnapshotHit) {
                $mcRowsForLoad =
                    @($stockMasterSnapshot.materialCentres)
            }
            else {
                $mcTable =
                    Invoke-StockStatusDirectTable `
                        -Connection $directStockConn `
                        -Sql (
                            "SELECT Code, Name FROM Master1 " +
                            "WHERE MasterType=11 ORDER BY Name"
                        )

                foreach ($mcRow in $mcTable.Rows) {
                    $mcObj = [pscustomobject]@{
                        code =
                            ConvertTo-ReportInt $mcRow["Code"]
                        name = [string]$mcRow["Name"]
                    }

                    $mcRowsForLoad += $mcObj
                    $stockMasterSnapshotBuild.materialCentres +=
                        $mcObj
                }
            }

            foreach ($mcObj in @($mcRowsForLoad)) {
                $mcCode = ConvertTo-ReportInt $mcObj.code
                $mcName = [string]$mcObj.name

                if ($mcCode -gt 0) {
                    $mcMap[$mcCode] = $mcName
                    $mcNameToCode[
                        $mcName.Trim().ToLowerInvariant()
                    ] = $mcCode

                    $mcOptions.Add([pscustomobject]@{
                        code = $mcCode
                        name = $mcName
                    })
                }
            }
        }
        else {
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
        }

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

        if ($null -ne $directStockConn) {
            $unitRowsForLoad = @()

            if ($stockMasterSnapshotHit) {
                $unitRowsForLoad =
                    @($stockMasterSnapshot.units)
            }
            else {
                $unitTable =
                    Invoke-StockStatusDirectTable `
                        -Connection $directStockConn `
                        -Sql (
                            "SELECT Code, Name FROM Master1 " +
                            "WHERE MasterType=8"
                        )

                foreach ($unitRow in $unitTable.Rows) {
                    $unitObj = [pscustomobject]@{
                        code =
                            ConvertTo-ReportInt $unitRow["Code"]
                        name = [string]$unitRow["Name"]
                    }

                    $unitRowsForLoad += $unitObj
                    $stockMasterSnapshotBuild.units += $unitObj
                }
            }

            foreach ($unitObj in @($unitRowsForLoad)) {
                $masterCode =
                    ConvertTo-ReportInt $unitObj.code

                if ($masterCode -gt 0) {
                    $masterNameMap[$masterCode] =
                        [string]$unitObj.name
                }
            }
        }
        else {
            $masterNameRst = $fi.GetRecordset(
                "SELECT Code, Name FROM Master1 WHERE MasterType=8"
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
        }

                # BUSY parity:
        # Item master XML emits <DoNotMaintainStkBal>True</...> for items
        # whose Master1.B21 is NULL. Native BUSY excludes those items from
        # Stock Status entirely. KH567 (code 7849) proved this exact rule.
        # Do not hard-code item codes; apply the BUSY master flag globally.
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
WHERE
    M.MasterType=6
    AND M.B21 IS NOT NULL
ORDER BY M.Name, M.Code
"@

        $items = @{}
        $itemOrder = [System.Collections.ArrayList]::new()

        $addStockItem = {
            param(
                [int]$itemCode,
                [string]$itemName,
                [string]$alias,
                [int]$groupCode,
                [int]$mainUnitCode,
                [int]$altUnitCode,
                [double]$mrp,
                [double]$salePrice,
                [double]$purchasePrice
            )

            $groupPath =
                @(Get-StockGroupPath -GroupCode $groupCode)

            $groupName = if ($groupPath.Count -gt 0) {
                [string]$groupPath[$groupPath.Count - 1]
            }
            else {
                "General"
            }

            $matchesSearch = $true
            if (-not [string]::IsNullOrWhiteSpace($Search)) {
                $searchNeedle =
                    $Search.Trim().ToLowerInvariant()

                $matchesSearch = (
                    $itemName.ToLowerInvariant().Contains(
                        $searchNeedle
                    ) -or
                    $alias.ToLowerInvariant().Contains(
                        $searchNeedle
                    )
                )
            }

            $matchesGroup = $true
            if (
                -not [string]::IsNullOrWhiteSpace($ItemGroup) -and
                $ItemGroup.Trim().ToLowerInvariant() -ne "all"
            ) {
                $groupNeedle =
                    $ItemGroup.Trim().ToLowerInvariant()

                $pathLabel =
                    ($groupPath -join " > ").
                        Trim().
                        ToLowerInvariant()

                $matchesGroup =
                    ($pathLabel -eq $groupNeedle)

                if (
                    -not $matchesGroup -and
                    -not $groupNeedle.Contains(" > ")
                ) {
                    foreach ($segment in $groupPath) {
                        if (
                            ([string]$segment).
                                Trim().
                                ToLowerInvariant() -eq
                            $groupNeedle
                        ) {
                            $matchesGroup = $true
                            break
                        }
                    }
                }
            }

            if (
                $itemCode -le 0 -or
                -not $matchesSearch -or
                -not $matchesGroup
            ) {
                return
            }

            $mainUnitName = if (
                $mainUnitCode -gt 0 -and
                $masterNameMap.ContainsKey($mainUnitCode)
            ) {
                [string]$masterNameMap[$mainUnitCode]
            }
            else {
                ""
            }

            $altUnitName = if (
                $altUnitCode -gt 0 -and
                $masterNameMap.ContainsKey($altUnitCode)
            ) {
                [string]$masterNameMap[$altUnitCode]
            }
            else {
                ""
            }

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
                mrp = $mrp
                salePrice = $salePrice
                purchasePrice = $purchasePrice
                mc = @{}
            }

            [void]$itemOrder.Add($itemCode)
        }

        if ($null -ne $directStockConn) {
            $itemRowsForLoad = @()

            if ($stockMasterSnapshotHit) {
                $itemRowsForLoad =
                    @($stockMasterSnapshot.items)
            }
            else {
                $itemTable =
                    Invoke-StockStatusDirectTable `
                        -Connection $directStockConn `
                        -Sql $itemQry

                foreach ($itemRow in $itemTable.Rows) {
                    $itemObj = [pscustomobject]@{
                        code =
                            ConvertTo-ReportInt $itemRow["Code"]
                        name = [string]$itemRow["Name"]
                        alias = [string]$itemRow["Alias"]
                        parentGrp =
                            ConvertTo-ReportInt $itemRow["ParentGrp"]
                        mainUnitCode =
                            ConvertTo-ReportInt $itemRow["MainUnitCode"]
                        altUnitCode =
                            ConvertTo-ReportInt $itemRow["AltUnitCode"]
                        mrp =
                            ConvertTo-ReportDouble $itemRow["MRP"]
                        salePrice =
                            ConvertTo-ReportDouble $itemRow["SalePrice"]
                        purchasePrice =
                            ConvertTo-ReportDouble $itemRow["PurchasePrice"]
                    }

                    $itemRowsForLoad += $itemObj
                    $stockMasterSnapshotBuild.items += $itemObj
                }
            }

            foreach ($itemObj in @($itemRowsForLoad)) {
                & $addStockItem `
                    -itemCode (
                        ConvertTo-ReportInt $itemObj.code
                    ) `
                    -itemName ([string]$itemObj.name) `
                    -alias ([string]$itemObj.alias) `
                    -groupCode (
                        ConvertTo-ReportInt $itemObj.parentGrp
                    ) `
                    -mainUnitCode (
                        ConvertTo-ReportInt $itemObj.mainUnitCode
                    ) `
                    -altUnitCode (
                        ConvertTo-ReportInt $itemObj.altUnitCode
                    ) `
                    -mrp (
                        ConvertTo-ReportDouble $itemObj.mrp
                    ) `
                    -salePrice (
                        ConvertTo-ReportDouble $itemObj.salePrice
                    ) `
                    -purchasePrice (
                        ConvertTo-ReportDouble $itemObj.purchasePrice
                    )
            }
        }
        else {
            $itemRst = $fi.GetRecordset($itemQry)

            if ($itemRst -and -not $itemRst.EOF) {
                $itemRst.MoveFirst()

                while (-not $itemRst.EOF) {
                    & $addStockItem `
                        -itemCode (
                            ConvertTo-ReportInt (
                                Get-ReportFieldValue $itemRst "Code" 0
                            )
                        ) `
                        -itemName (
                            [string](
                                Get-ReportFieldValue $itemRst "Name" ""
                            )
                        ) `
                        -alias (
                            [string](
                                Get-ReportFieldValue $itemRst "Alias" ""
                            )
                        ) `
                        -groupCode (
                            ConvertTo-ReportInt (
                                Get-ReportFieldValue $itemRst "ParentGrp" 0
                            )
                        ) `
                        -mainUnitCode (
                            ConvertTo-ReportInt (
                                Get-ReportFieldValue $itemRst "MainUnitCode" 0
                            )
                        ) `
                        -altUnitCode (
                            ConvertTo-ReportInt (
                                Get-ReportFieldValue $itemRst "AltUnitCode" 0
                            )
                        ) `
                        -mrp (
                            ConvertTo-ReportDouble (
                                Get-ReportFieldValue $itemRst "MRP" 0
                            )
                        ) `
                        -salePrice (
                            ConvertTo-ReportDouble (
                                Get-ReportFieldValue $itemRst "SalePrice" 0
                            )
                        ) `
                        -purchasePrice (
                            ConvertTo-ReportDouble (
                                Get-ReportFieldValue $itemRst "PurchasePrice" 0
                            )
                        )

                    $itemRst.MoveNext()
                }
            }

            Close-ReportRecordset $itemRst
        }

        if (
            $null -ne $directStockConn -and
            -not $stockMasterSnapshotHit -and
            -not [string]::IsNullOrWhiteSpace(
                $stockMasterSnapshotPath
            ) -and
            -not [string]::IsNullOrWhiteSpace(
                $stockMasterSignature
            )
        ) {
            try {
                $snapshotPayload = [ordered]@{
                    engine = "stock-master-snapshot-v8"
                    signature = $stockMasterSignature
                    groups = @($stockMasterSnapshotBuild.groups)
                    materialCentres =
                        @($stockMasterSnapshotBuild.materialCentres)
                    units = @($stockMasterSnapshotBuild.units)
                    items = @($stockMasterSnapshotBuild.items)
                }

                $snapshotPayload |
                    ConvertTo-Json -Depth 8 -Compress |
                    Set-Content `
                        -LiteralPath $stockMasterSnapshotPath `
                        -Encoding UTF8

                Write-ReportEngineLog `
                    -Stage "STOCK-MASTERS" `
                    -Message (
                        "Persistent master snapshot stored; items=" +
                        $stockMasterSnapshotBuild.items.Count
                    ) `
                    -Color "Green"
            }
            catch {}
        }

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

        # ========================================================
        # TRUE FAST PAGINATION FOR THE DEFAULT BALANCES VIEW
        # ========================================================
        #
        # Old cold path:
        #   every stock master -> every transaction -> method-5 valuation
        #   -> filter/sort -> finally return first 100 rows.
        #
        # Fast cold path:
        #   1) Access performs cheap GROUP BY quantity aggregation.
        #   2) Apply status / masters / zero filters globally.
        #   3) Sort the matching item masters.
        #   4) Select only the requested page item codes.
        #   5) The existing exact BUSY method-5 engine below replays
        #      transactions ONLY for those page items.
        #
        # This preserves exact row Price/Amount while removing thousands of
        # unrelated transaction rows from the user-facing request.
        #
        # PageSize=0 ("All"), Grouped/Hierarchical, and very large pages keep
        # the complete calculation path because they genuinely need all rows.
        # ========================================================
        $fastPageModeActive = (
            $viewMode -eq "balances" -and
            $PageSize -gt 0 -and
            $PageSize -le 250
        )

        $fastTotalRows = 0
        $fastEffectivePage = $Page
        $fastEffectivePageSize = $PageSize
        $fastTotalPages = 1
        $fastGlobalSummary = $null

        if ($fastPageModeActive) {
            $fastIndexWatch = New-ReportEngineWatch
            $fastIndexKey = (
                @(
                    [string]$stockStatusCacheKey,
                    "quantity-index-v2-busy-maintain-stock"
                ) -join $cacheSeparator
            )

            $fastCandidateCodes = @()
            $fastGlobalGroupTotals = @()
            $fastIndexHit = $false

            # Persistent index survives bridge restarts. It is accepted only
            # while the underlying BDS file's write timestamp/size are
            # unchanged, so normal voucher modifications invalidate it.
            $fastDiskCachePath = ""
            $fastDiskDbStamp = ""

            if ($null -ne $directStockContext) {
                $fastDiskDbStamp = (
                    [string]$directStockContext.writeTicks +
                    ":" +
                    [string]$directStockContext.length
                )

                try {
                    $fastDiskCacheDir =
                        Join-Path `
                            (Join-Path $PSScriptRoot "..\cache") `
                            "stock-status"

                    [void][System.IO.Directory]::CreateDirectory(
                        $fastDiskCacheDir
                    )

                    $fastDiskHash =
                        Get-StockStatusStableHash `
                            -Text (
                                $fastIndexKey +
                                "|" +
                                [string]$directStockContext.path
                            )

                    $fastDiskCachePath =
                        Join-Path `
                            $fastDiskCacheDir `
                            ($fastDiskHash + ".json")
                }
                catch {
                    $fastDiskCachePath = ""
                }
            }

            if ($script:StockStatusFastIndexCache.ContainsKey($fastIndexKey)) {
                $fastIndexEntry =
                    $script:StockStatusFastIndexCache[$fastIndexKey]

                $fastIndexAge = 0.0
                try {
                    $fastIndexAge = (
                        (Get-Date) -
                        [datetime]$fastIndexEntry.createdAt
                    ).TotalSeconds
                }
                catch {
                    $fastIndexAge =
                        $script:StockStatusReportCacheSeconds + 1
                }

                if (
                    $fastIndexAge -le
                    $script:StockStatusReportCacheSeconds -and
                    $null -ne $fastIndexEntry.candidateCodes -and
                    $null -ne $fastIndexEntry.summary
                ) {
                    $fastCandidateCodes =
                        @($fastIndexEntry.candidateCodes)

                    $fastGlobalGroupTotals =
                        @($fastIndexEntry.groupTotals)

                    $fastGlobalSummary = @{}
                    foreach ($key in @($fastIndexEntry.summary.Keys)) {
                        $fastGlobalSummary[$key] =
                            $fastIndexEntry.summary[$key]
                    }

                    $fastIndexHit = $true

                    Write-ReportEngineLog `
                        -Stage "STOCK-FAST-INDEX" `
                        -Message (
                            "Quantity index cache hit; age=" +
                            [Math]::Round($fastIndexAge, 1) +
                            "s; matching=" +
                            $fastCandidateCodes.Count
                        ) `
                        -Color "Green"
                }
                else {
                    $script:StockStatusFastIndexCache.Remove(
                        $fastIndexKey
                    )
                }
            }

            if (
                -not $fastIndexHit -and
                -not [string]::IsNullOrWhiteSpace(
                    $fastDiskCachePath
                ) -and
                (Test-Path -LiteralPath $fastDiskCachePath)
            ) {
                try {
                    $fastDiskEntry =
                        Get-Content `
                            -LiteralPath $fastDiskCachePath `
                            -Raw |
                        ConvertFrom-Json

                    if (
                        [string]$fastDiskEntry.engine -eq
                            "stock-fast-index-v8" -and
                        [string]$fastDiskEntry.dbStamp -eq
                            $fastDiskDbStamp -and
                        $null -ne $fastDiskEntry.candidateCodes -and
                        $null -ne $fastDiskEntry.summary
                    ) {
                        $fastCandidateCodes =
                            @(
                                $fastDiskEntry.candidateCodes |
                                ForEach-Object { [int]$_ }
                            )

                        $fastGlobalSummary = @{}
                        foreach (
                            $property in
                            $fastDiskEntry.summary.PSObject.Properties
                        ) {
                            $fastGlobalSummary[
                                [string]$property.Name
                            ] = $property.Value
                        }

                        $fastGlobalGroupTotals =
                            @($fastDiskEntry.groupTotals)

                        $fastIndexHit = $true

                        $script:StockStatusFastIndexCache[
                            $fastIndexKey
                        ] = @{
                            createdAt = Get-Date
                            instanceId = $InstanceId
                            companyCode = $CompanyCode
                            candidateCodes =
                                @($fastCandidateCodes)
                            summary = $fastGlobalSummary
                            groupTotals =
                                @($fastGlobalGroupTotals)
                        }

                        Write-ReportEngineLog `
                            -Stage "STOCK-FAST-INDEX" `
                            -Message (
                                "Persistent quantity index hit; matching=" +
                                $fastCandidateCodes.Count
                            ) `
                            -Color "Green"
                    }
                }
                catch {
                    # Ignore malformed/stale disk cache and rebuild.
                }
            }

            if (-not $fastIndexHit) {
                # Keep the quantity index deliberately small: only fields
                # required to determine global matching rows and top quantity
                # metrics. Expensive BUSY valuation fields are page-scoped.
                $fastStatesByItem = @{}

                function Get-FastStockState {
                    param(
                        [int]$ItemCode,
                        [int]$McCode
                    )

                    $itemKey = [string]$ItemCode
                    $mcKey = [string]$McCode

                    if (-not $fastStatesByItem.ContainsKey($itemKey)) {
                        $fastStatesByItem[$itemKey] = @{}
                    }

                    $itemStates = $fastStatesByItem[$itemKey]

                    if (-not $itemStates.ContainsKey($mcKey)) {
                        $itemStates[$mcKey] = [pscustomobject]@{
                            opening = 0.0
                            prior = 0.0
                            inward = 0.0
                            outward = 0.0
                            transferIn = 0.0
                            transferOut = 0.0
                            purchase = 0.0
                            sale = 0.0
                        }
                    }

                    return $itemStates[$mcKey]
                }

                # Financial-year opening quantities.
                # Direct OLE DB is dramatically faster than BUSY COM for this
                # company-wide grouped read.
                $fastOpenQry = @"
SELECT
    MasterCode1 AS ItemCode,
    MasterCode2 AS MCCode,
    SUM(D1) AS OpeningQty
FROM Tran4
WHERE
    RecType=0
    $tran4ItemPredicate
GROUP BY
    MasterCode1,
    MasterCode2
"@

                if ($null -ne $directStockConn) {
                    $fastOpenTable =
                        Invoke-StockStatusDirectTable `
                            -Connection $directStockConn `
                            -Sql $fastOpenQry

                    foreach ($fastOpenRow in $fastOpenTable.Rows) {
                        $fastItemCode =
                            ConvertTo-ReportInt (
                                $fastOpenRow["ItemCode"]
                            )

                        $fastMcCode =
                            ConvertTo-ReportInt (
                                $fastOpenRow["MCCode"]
                            )

                        if ($items.ContainsKey($fastItemCode)) {
                            $fastState =
                                Get-FastStockState `
                                    -ItemCode $fastItemCode `
                                    -McCode $fastMcCode

                            $fastState.opening +=
                                ConvertTo-ReportDouble (
                                    $fastOpenRow["OpeningQty"]
                                )
                        }
                    }
                }
                else {
                    $fastOpenRst =
                        $fi.GetRecordset($fastOpenQry)

                    if (
                        $fastOpenRst -and
                        -not $fastOpenRst.EOF
                    ) {
                        $fastOpenRst.MoveFirst()

                        while (-not $fastOpenRst.EOF) {
                            $fastItemCode =
                                ConvertTo-ReportInt (
                                    Get-ReportFieldValue `
                                        $fastOpenRst `
                                        "ItemCode" `
                                        0
                                )

                            $fastMcCode =
                                ConvertTo-ReportInt (
                                    Get-ReportFieldValue `
                                        $fastOpenRst `
                                        "MCCode" `
                                        0
                                )

                            if ($items.ContainsKey($fastItemCode)) {
                                $fastState =
                                    Get-FastStockState `
                                        -ItemCode $fastItemCode `
                                        -McCode $fastMcCode

                                $fastState.opening +=
                                    ConvertTo-ReportDouble (
                                        Get-ReportFieldValue `
                                            $fastOpenRst `
                                            "OpeningQty" `
                                            0
                                    )
                            }

                            $fastOpenRst.MoveNext()
                        }
                    }

                    Close-ReportRecordset $fastOpenRst
                }

                # --------------------------------------------------------
                # Can the current FY movement query skip Tran1 entirely?
                #
                # For the normal current-period report, if:
                #   * From = FY start,
                #   * no active voucher is future-dated past AsOf,
                #   * no voucher is cancelled,
                # then every RecType=2 row in this FY database is in scope.
                #
                # Avoiding the Tran2 -> Tran1 join is substantially faster in
                # large Access/BDS companies.
                # --------------------------------------------------------
                $fastUseTran2OnlyAggregate = $false

                if (
                    $null -ne $directStockConn -and
                    $fromDate.Date -eq $defaultFrom.Date
                ) {
                    try {
                        $scopeTable =
                            Invoke-StockStatusDirectTable `
                                -Connection $directStockConn `
                                -Sql @"
SELECT
    MIN([Date]) AS MinVoucherDate,
    MAX([Date]) AS MaxVoucherDate,
    SUM(
        IIF(
            Cancelled<>0 OR VchCancelled<>0,
            1,
            0
        )
    ) AS CancelledCount
FROM Tran1
"@

                        if ($scopeTable.Rows.Count -gt 0) {
                            $minVoucherDateOk = $true
                            $maxVoucherDateOk = $true
                            $cancelledCount = 0

                            try {
                                $minRaw =
                                    $scopeTable.Rows[0]["MinVoucherDate"]

                                if (
                                    $null -ne $minRaw -and
                                    $minRaw -ne [System.DBNull]::Value
                                ) {
                                    $minVoucherDateOk = (
                                        ([datetime]$minRaw).Date -ge
                                        $defaultFrom.Date
                                    )
                                }
                            }
                            catch {
                                $minVoucherDateOk = $false
                            }

                            try {
                                $maxRaw =
                                    $scopeTable.Rows[0]["MaxVoucherDate"]

                                if (
                                    $null -ne $maxRaw -and
                                    $maxRaw -ne [System.DBNull]::Value
                                ) {
                                    $maxVoucherDateOk = (
                                        ([datetime]$maxRaw).Date -le
                                        $endDate.Date
                                    )
                                }
                            }
                            catch {
                                $maxVoucherDateOk = $false
                            }

                            $cancelledCount =
                                ConvertTo-ReportInt (
                                    $scopeTable.Rows[0][
                                        "CancelledCount"
                                    ]
                                )

                            $fastUseTran2OnlyAggregate = (
                                $minVoucherDateOk -and
                                $maxVoucherDateOk -and
                                $cancelledCount -eq 0
                            )
                        }
                    }
                    catch {}
                }

                # --------------------------------------------------------
                # FAST MOVEMENT AGGREGATION
                # --------------------------------------------------------
                # The previous query performed seven SUM(IIF(...)) expressions
                # across the whole Tran2/Tran1 join and took ~30 seconds in
                # Access. For the default Balances view we only need current
                # period Qty In/Out to decide status/pagination.
                #
                # Financial-year opening is already in Tran4, so when From is
                # the FY start there is no "prior movement" query at all.
                # --------------------------------------------------------
                if ($fastUseTran2OnlyAggregate) {
                    $fastMovementQry = @"
SELECT
    T.MasterCode1 AS ItemCode,
    T.MasterCode2 AS MCCode,
    SUM(IIF(T.Value1 > 0, T.Value1, 0)) AS InwardQty,
    SUM(IIF(T.Value1 < 0, -T.Value1, 0)) AS OutwardQty
FROM Tran2 T
WHERE
    T.RecType = 2
    $tran2ItemPredicate
GROUP BY
    T.MasterCode1,
    T.MasterCode2
"@

                    Write-ReportEngineLog `
                        -Stage "STOCK-FAST-MOVEMENT" `
                        -Message (
                            "Using direct FY Tran2 aggregate " +
                            "(Tran1 join safely skipped)"
                        ) `
                        -Color "Green"
                }
                else {
                    $fastMovementQry = @"
SELECT
    T.MasterCode1 AS ItemCode,
    T.MasterCode2 AS MCCode,
    SUM(IIF(T.Value1 > 0, T.Value1, 0)) AS InwardQty,
    SUM(IIF(T.Value1 < 0, -T.Value1, 0)) AS OutwardQty
FROM
    Tran2 T
    INNER JOIN Tran1 V
        ON V.VchCode = T.VchCode
WHERE
    T.RecType = 2
    $tran2ItemPredicate
    AND V.[Date] >= $fromLiteral
    AND V.[Date] <= $endLiteral
    AND (V.Cancelled = 0 OR V.Cancelled IS NULL)
    AND (V.VchCancelled = 0 OR V.VchCancelled IS NULL)
GROUP BY
    T.MasterCode1,
    T.MasterCode2
"@
                }

                $consumeFastMovementRow = {
                    param(
                        [int]$fastItemCode,
                        [int]$fastMcCode,
                        [double]$fastInQty,
                        [double]$fastOutQty
                    )

                    if (-not $items.ContainsKey($fastItemCode)) {
                        return
                    }

                    $fastState =
                        Get-FastStockState `
                            -ItemCode $fastItemCode `
                            -McCode $fastMcCode

                    $fastState.inward += $fastInQty
                    $fastState.outward += $fastOutQty
                }

                if ($null -ne $directStockConn) {
                    $fastMoveTable =
                        Invoke-StockStatusDirectTable `
                            -Connection $directStockConn `
                            -Sql $fastMovementQry

                    foreach ($fastMoveRow in $fastMoveTable.Rows) {
                        & $consumeFastMovementRow `
                            -fastItemCode (
                                ConvertTo-ReportInt (
                                    $fastMoveRow["ItemCode"]
                                )
                            ) `
                            -fastMcCode (
                                ConvertTo-ReportInt (
                                    $fastMoveRow["MCCode"]
                                )
                            ) `
                            -fastInQty (
                                ConvertTo-ReportDouble (
                                    $fastMoveRow["InwardQty"]
                                )
                            ) `
                            -fastOutQty (
                                ConvertTo-ReportDouble (
                                    $fastMoveRow["OutwardQty"]
                                )
                            )
                    }
                }
                else {
                    $fastMoveRst =
                        $fi.GetRecordset($fastMovementQry)

                    if (
                        $fastMoveRst -and
                        -not $fastMoveRst.EOF
                    ) {
                        $fastMoveRst.MoveFirst()

                        while (-not $fastMoveRst.EOF) {
                            & $consumeFastMovementRow `
                                -fastItemCode (
                                    ConvertTo-ReportInt (
                                        Get-ReportFieldValue `
                                            $fastMoveRst `
                                            "ItemCode" `
                                            0
                                    )
                                ) `
                                -fastMcCode (
                                    ConvertTo-ReportInt (
                                        Get-ReportFieldValue `
                                            $fastMoveRst `
                                            "MCCode" `
                                            0
                                    )
                                ) `
                                -fastInQty (
                                    ConvertTo-ReportDouble (
                                        Get-ReportFieldValue `
                                            $fastMoveRst `
                                            "InwardQty" `
                                            0
                                    )
                                ) `
                                -fastOutQty (
                                    ConvertTo-ReportDouble (
                                        Get-ReportFieldValue `
                                            $fastMoveRst `
                                            "OutwardQty" `
                                            0
                                    )
                                )

                            $fastMoveRst.MoveNext()
                        }
                    }

                    Close-ReportRecordset $fastMoveRst
                }

                # If a custom From is after the FY start, add movement between
                # FY opening and From into opening. The normal Stock Status UI
                # uses the FY start, so this query is skipped on first load.
                if ($fromDate.Date -gt $defaultFrom.Date) {
                    $defaultFromLiteral =
                        Get-ReportSqlDateLiteral `
                            -Date $defaultFrom `
                            -IsSql $mode.isSql

                    $fastPriorQry = @"
SELECT
    T.MasterCode1 AS ItemCode,
    T.MasterCode2 AS MCCode,
    SUM(T.Value1) AS PriorQty
FROM
    Tran2 T
    INNER JOIN Tran1 V
        ON V.VchCode = T.VchCode
WHERE
    T.RecType = 2
    $tran2ItemPredicate
    AND V.[Date] >= $defaultFromLiteral
    AND V.[Date] < $fromLiteral
    AND (V.Cancelled = 0 OR V.Cancelled IS NULL)
    AND (V.VchCancelled = 0 OR V.VchCancelled IS NULL)
GROUP BY
    T.MasterCode1,
    T.MasterCode2
"@

                    if ($null -ne $directStockConn) {
                        $fastPriorTable =
                            Invoke-StockStatusDirectTable `
                                -Connection $directStockConn `
                                -Sql $fastPriorQry

                        foreach (
                            $fastPriorRow in
                            $fastPriorTable.Rows
                        ) {
                            $fastItemCode =
                                ConvertTo-ReportInt (
                                    $fastPriorRow["ItemCode"]
                                )

                            $fastMcCode =
                                ConvertTo-ReportInt (
                                    $fastPriorRow["MCCode"]
                                )

                            if ($items.ContainsKey($fastItemCode)) {
                                $fastState =
                                    Get-FastStockState `
                                        -ItemCode $fastItemCode `
                                        -McCode $fastMcCode

                                $fastState.prior +=
                                    ConvertTo-ReportDouble (
                                        $fastPriorRow["PriorQty"]
                                    )
                            }
                        }
                    }
                    else {
                        $fastPriorRst =
                            $fi.GetRecordset($fastPriorQry)

                        if (
                            $fastPriorRst -and
                            -not $fastPriorRst.EOF
                        ) {
                            $fastPriorRst.MoveFirst()

                            while (-not $fastPriorRst.EOF) {
                                $fastItemCode =
                                    ConvertTo-ReportInt (
                                        Get-ReportFieldValue `
                                            $fastPriorRst `
                                            "ItemCode" `
                                            0
                                    )

                                $fastMcCode =
                                    ConvertTo-ReportInt (
                                        Get-ReportFieldValue `
                                            $fastPriorRst `
                                            "MCCode" `
                                            0
                                    )

                                if ($items.ContainsKey($fastItemCode)) {
                                    $fastState =
                                        Get-FastStockState `
                                            -ItemCode $fastItemCode `
                                            -McCode $fastMcCode

                                    $fastState.prior +=
                                        ConvertTo-ReportDouble (
                                            Get-ReportFieldValue `
                                                $fastPriorRst `
                                                "PriorQty" `
                                                0
                                        )
                                }

                                $fastPriorRst.MoveNext()
                            }
                        }

                        Close-ReportRecordset $fastPriorRst
                    }
                }

                # Stock-transfer exclusion is an uncommon optional mode.
                # Run one small extra aggregate only when explicitly needed.
                if (-not $IncludeStockTransfers) {
                    $fastTransferQry = @"
SELECT
    T.MasterCode1 AS ItemCode,
    T.MasterCode2 AS MCCode,
    SUM(IIF(T.Value1 > 0, T.Value1, 0)) AS TransferInQty,
    SUM(IIF(T.Value1 < 0, -T.Value1, 0)) AS TransferOutQty
FROM
    Tran2 T
    INNER JOIN Tran1 V
        ON V.VchCode = T.VchCode
WHERE
    T.RecType = 2
    $tran2ItemPredicate
    AND V.[Date] >= $fromLiteral
    AND V.[Date] <= $endLiteral
    AND V.VchType = 5
    AND (V.Cancelled = 0 OR V.Cancelled IS NULL)
    AND (V.VchCancelled = 0 OR V.VchCancelled IS NULL)
GROUP BY
    T.MasterCode1,
    T.MasterCode2
"@

                    if ($null -ne $directStockConn) {
                        $fastTransferTable =
                            Invoke-StockStatusDirectTable `
                                -Connection $directStockConn `
                                -Sql $fastTransferQry

                        foreach (
                            $fastTransferRow in
                            $fastTransferTable.Rows
                        ) {
                            $fastItemCode =
                                ConvertTo-ReportInt (
                                    $fastTransferRow["ItemCode"]
                                )

                            $fastMcCode =
                                ConvertTo-ReportInt (
                                    $fastTransferRow["MCCode"]
                                )

                            if ($items.ContainsKey($fastItemCode)) {
                                $fastState =
                                    Get-FastStockState `
                                        -ItemCode $fastItemCode `
                                        -McCode $fastMcCode

                                $fastState.transferIn +=
                                    ConvertTo-ReportDouble (
                                        $fastTransferRow[
                                            "TransferInQty"
                                        ]
                                    )

                                $fastState.transferOut +=
                                    ConvertTo-ReportDouble (
                                        $fastTransferRow[
                                            "TransferOutQty"
                                        ]
                                    )
                            }
                        }
                    }
                    else {
                        $fastTransferRst =
                            $fi.GetRecordset($fastTransferQry)

                        if (
                            $fastTransferRst -and
                            -not $fastTransferRst.EOF
                        ) {
                            $fastTransferRst.MoveFirst()

                            while (-not $fastTransferRst.EOF) {
                                $fastItemCode =
                                    ConvertTo-ReportInt (
                                        Get-ReportFieldValue `
                                            $fastTransferRst `
                                            "ItemCode" `
                                            0
                                    )

                                $fastMcCode =
                                    ConvertTo-ReportInt (
                                        Get-ReportFieldValue `
                                            $fastTransferRst `
                                            "MCCode" `
                                            0
                                    )

                                if ($items.ContainsKey($fastItemCode)) {
                                    $fastState =
                                        Get-FastStockState `
                                            -ItemCode $fastItemCode `
                                            -McCode $fastMcCode

                                    $fastState.transferIn +=
                                        ConvertTo-ReportDouble (
                                            Get-ReportFieldValue `
                                                $fastTransferRst `
                                                "TransferInQty" `
                                                0
                                        )

                                    $fastState.transferOut +=
                                        ConvertTo-ReportDouble (
                                            Get-ReportFieldValue `
                                                $fastTransferRst `
                                                "TransferOutQty" `
                                                0
                                        )
                                }

                                $fastTransferRst.MoveNext()
                            }
                        }

                        Close-ReportRecordset $fastTransferRst
                    }
                }

                $fastCandidates =
                    [System.Collections.Generic.List[object]]::new()

                # Exact GLOBAL quantity totals by item group. These do not
                # require expensive valuation and therefore remain correct
                # even though method-5 Price/Amount is page-scoped.
                $fastGlobalGroupTotalsMap = @{}

                $fastSummaryOpening = 0.0
                $fastSummaryInward = 0.0
                $fastSummaryOutward = 0.0
                $fastSummaryClosing = 0.0
                $fastSummaryPurchase = 0.0
                $fastSummarySale = 0.0
                $fastInStock = 0
                $fastLowStock = 0
                $fastOutOfStock = 0
                $fastNegative = 0

                foreach ($fastItemCodeRaw in @($itemOrder)) {
                    $fastItemCode = [int]$fastItemCodeRaw
                    $fastItem = $items[$fastItemCode]

                    $fastOpening = 0.0
                    $fastInward = 0.0
                    $fastOutward = 0.0
                    $fastPurchase = 0.0
                    $fastSale = 0.0

                    $fastItemKey = [string]$fastItemCode
                    $fastItemStates = if (
                        $fastStatesByItem.ContainsKey($fastItemKey)
                    ) {
                        $fastStatesByItem[$fastItemKey]
                    }
                    else {
                        @{}
                    }

                    $fastMcCodes = [System.Collections.ArrayList]::new()

                    if ($useAllMc) {
                        foreach ($fastMcKey in @($fastItemStates.Keys)) {
                            $fastParsedMc = 0
                            [void][int]::TryParse(
                                [string]$fastMcKey,
                                [ref]$fastParsedMc
                            )
                            [void]$fastMcCodes.Add($fastParsedMc)
                        }
                    }
                    else {
                        foreach ($fastSelectedMc in @($selectedMcCodes.Keys)) {
                            [void]$fastMcCodes.Add([int]$fastSelectedMc)
                        }
                    }

                    foreach (
                        $fastMcCodeRaw in
                        @($fastMcCodes | Sort-Object -Unique)
                    ) {
                        $fastMcCode = [int]$fastMcCodeRaw
                        $fastMcKey = [string]$fastMcCode

                        if ($fastItemStates.ContainsKey($fastMcKey)) {
                            $fastState = $fastItemStates[$fastMcKey]
                        }
                        else {
                            $fastState = [pscustomobject]@{
                                opening = 0.0
                                prior = 0.0
                                inward = 0.0
                                outward = 0.0
                                transferIn = 0.0
                                transferOut = 0.0
                                purchase = 0.0
                                sale = 0.0
                            }
                        }

                        $fastOpening +=
                            [double]$fastState.opening +
                            [double]$fastState.prior

                        $fastMcInward = [double]$fastState.inward
                        $fastMcOutward = [double]$fastState.outward

                        if (-not $IncludeStockTransfers) {
                            $fastMcInward = [Math]::Max(
                                0.0,
                                $fastMcInward -
                                [double]$fastState.transferIn
                            )

                            $fastMcOutward = [Math]::Max(
                                0.0,
                                $fastMcOutward -
                                [double]$fastState.transferOut
                            )
                        }

                        $fastInward += $fastMcInward
                        $fastOutward += $fastMcOutward
                        $fastPurchase += [double]$fastState.purchase
                        $fastSale += [double]$fastState.sale
                    }

                    $fastClosing =
                        $fastOpening +
                        $fastInward -
                        $fastOutward

                    $fastMovement =
                        $fastInward +
                        $fastOutward

                    $fastStatus =
                        if ($fastClosing -lt -$eps) {
                            "negative"
                        }
                        elseif ([Math]::Abs($fastClosing) -le $eps) {
                            "out-of-stock"
                        }
                        elseif ($fastClosing -le $LowStockLevel) {
                            "low-stock"
                        }
                        else {
                            "in-stock"
                        }

                    if (
                        $statusFilter -ne "all" -and
                        $fastStatus -ne $statusFilter
                    ) {
                        continue
                    }

                    $fastPassesMasters = switch ($mastersFilter) {
                        "moved" {
                            $fastMovement -gt $eps
                        }
                        "closing" {
                            [Math]::Abs($fastClosing) -gt $eps
                        }
                        "moved-closing" {
                            (
                                $fastMovement -gt $eps -or
                                [Math]::Abs($fastClosing) -gt $eps
                            )
                        }
                        default {
                            $true
                        }
                    }

                    if (-not $fastPassesMasters) {
                        continue
                    }

                    if (
                        -not $IncludeZero -and
                        [Math]::Abs($fastClosing) -le $eps
                    ) {
                        continue
                    }

                    $fastCandidates.Add(
                        [pscustomobject]@{
                            itemCode = $fastItemCode
                            itemName = [string]$fastItem.name
                            opening = $fastOpening
                            inward = $fastInward
                            outward = $fastOutward
                            closing = $fastClosing
                            purchase = $fastPurchase
                            sale = $fastSale
                            status = $fastStatus
                        }
                    )

                    $fastGroupKey = if (
                        $fastItem.groupPath.Count -gt 0
                    ) {
                        (@($fastItem.groupPath) -join " > ")
                    }
                    elseif (
                        [string]::IsNullOrWhiteSpace(
                            [string]$fastItem.group
                        )
                    ) {
                        "General"
                    }
                    else {
                        [string]$fastItem.group
                    }

                    if (
                        -not $fastGlobalGroupTotalsMap.ContainsKey(
                            $fastGroupKey
                        )
                    ) {
                        $fastParentGroup = if (
                            $fastItem.groupPath.Count -gt 1
                        ) {
                            [string]$fastItem.groupPath[
                                $fastItem.groupPath.Count - 2
                            ]
                        }
                        else {
                            ""
                        }

                        $fastGlobalGroupTotalsMap[
                            $fastGroupKey
                        ] = [pscustomobject]@{
                            group = $fastGroupKey
                            parentGroup = $fastParentGroup
                            groupPath = @($fastItem.groupPath)
                            itemCount = 0
                            openingQuantity = 0.0
                            inwardQuantity = 0.0
                            outwardQuantity = 0.0
                            closingQuantity = 0.0
                            stockValue = 0.0
                        }
                    }

                    $fastGlobalGroupTotal =
                        $fastGlobalGroupTotalsMap[
                            $fastGroupKey
                        ]

                    $fastGlobalGroupTotal.itemCount++
                    $fastGlobalGroupTotal.openingQuantity +=
                        $fastOpening
                    $fastGlobalGroupTotal.inwardQuantity +=
                        $fastInward
                    $fastGlobalGroupTotal.outwardQuantity +=
                        $fastOutward
                    $fastGlobalGroupTotal.closingQuantity +=
                        $fastClosing

                    $fastSummaryOpening += $fastOpening
                    $fastSummaryInward += $fastInward
                    $fastSummaryOutward += $fastOutward
                    $fastSummaryClosing += $fastClosing
                    $fastSummaryPurchase += $fastPurchase
                    $fastSummarySale += $fastSale

                    switch ($fastStatus) {
                        "in-stock" {
                            $fastInStock++
                        }
                        "low-stock" {
                            $fastLowStock++
                        }
                        "out-of-stock" {
                            $fastOutOfStock++
                        }
                        "negative" {
                            $fastNegative++
                        }
                    }
                }

                $fastSortedCandidates = @(
                    $fastCandidates |
                    Sort-Object itemName, itemCode
                )

                $fastCandidateCodes = @(
                    $fastSortedCandidates |
                    ForEach-Object {
                        [int]$_.itemCode
                    }
                )

                $fastGlobalSummary = @{
                    totalItems = $fastCandidateCodes.Count
                    inStockItems = $fastInStock
                    lowStockItems = $fastLowStock
                    outOfStockItems = $fastOutOfStock
                    negativeStockItems = $fastNegative
                    totalQuantity = [Math]::Round(
                        $fastSummaryClosing,
                        3
                    )
                    openingQuantity = [Math]::Round(
                        $fastSummaryOpening,
                        3
                    )
                    inwardQuantity = [Math]::Round(
                        $fastSummaryInward,
                        3
                    )
                    outwardQuantity = [Math]::Round(
                        $fastSummaryOutward,
                        3
                    )
                    purchaseQuantity = [Math]::Round(
                        $fastSummaryPurchase,
                        3
                    )
                    saleQuantity = [Math]::Round(
                        $fastSummarySale,
                        3
                    )

                    # The expensive value is filled from the exact page
                    # valuation after method-5 replay below.
                    stockValue = 0.0
                    busyStockValue = 0.0
                }

                $fastGlobalGroupTotals = @(
                    foreach (
                        $fastGroupName in
                        @(
                            $fastGlobalGroupTotalsMap.Keys |
                            Sort-Object
                        )
                    ) {
                        $fastGroup =
                            $fastGlobalGroupTotalsMap[
                                $fastGroupName
                            ]

                        [pscustomobject]@{
                            group = [string]$fastGroup.group
                            parentGroup =
                                [string]$fastGroup.parentGroup
                            groupPath =
                                @($fastGroup.groupPath)
                            itemCount =
                                [int]$fastGroup.itemCount
                            openingQuantity =
                                [Math]::Round(
                                    [double]$fastGroup.openingQuantity,
                                    3
                                )
                            inwardQuantity =
                                [Math]::Round(
                                    [double]$fastGroup.inwardQuantity,
                                    3
                                )
                            outwardQuantity =
                                [Math]::Round(
                                    [double]$fastGroup.outwardQuantity,
                                    3
                                )
                            closingQuantity =
                                [Math]::Round(
                                    [double]$fastGroup.closingQuantity,
                                    3
                                )
                            stockValue = 0.0
                        }
                    }
                )

                $script:StockStatusFastIndexCache[$fastIndexKey] = @{
                    createdAt = Get-Date
                    instanceId = $InstanceId
                    companyCode = $CompanyCode
                    candidateCodes = @($fastCandidateCodes)
                    summary = $fastGlobalSummary
                    groupTotals = @($fastGlobalGroupTotals)
                }

                if (
                    -not [string]::IsNullOrWhiteSpace(
                        $fastDiskCachePath
                    )
                ) {
                    try {
                        $diskPayload = [ordered]@{
                            engine = "stock-fast-index-v8"
                            dbStamp = $fastDiskDbStamp
                            candidateCodes =
                                @($fastCandidateCodes)
                            summary = $fastGlobalSummary
                            groupTotals =
                                @($fastGlobalGroupTotals)
                        }

                        $diskPayload |
                            ConvertTo-Json -Depth 8 -Compress |
                            Set-Content `
                                -LiteralPath $fastDiskCachePath `
                                -Encoding UTF8
                    }
                    catch {}
                }

                Stop-ReportEngineWatch `
                    -Watch $fastIndexWatch `
                    -Stage "STOCK-FAST-INDEX" `
                    -Message (
                        "Built global quantity index; matching=" +
                        $fastCandidateCodes.Count
                    )
            }

            $fastTotalRows = $fastCandidateCodes.Count

            if (
                $fastEffectivePageSize -gt 0 -and
                $fastTotalRows -gt 0 -and
                $fastEffectivePageSize -ge $fastTotalRows
            ) {
                $fastEffectivePage = 1
                $fastEffectivePageSize = 0
                $fastTotalPages = 1
            }
            else {
                $fastTotalPages = [Math]::Max(
                    1,
                    [int][Math]::Ceiling(
                        $fastTotalRows /
                        [double]$fastEffectivePageSize
                    )
                )

                if ($fastEffectivePage -gt $fastTotalPages) {
                    $fastEffectivePage = $fastTotalPages
                }
            }

            if ($fastEffectivePageSize -eq 0) {
                $fastPageItemCodes = @($fastCandidateCodes)
            }
            else {
                $fastSkip =
                    ($fastEffectivePage - 1) *
                    $fastEffectivePageSize

                $fastPageItemCodes = @(
                    $fastCandidateCodes |
                    Select-Object `
                        -Skip $fastSkip `
                        -First $fastEffectivePageSize
                )
            }

            # The exact method-5 engine below now sees ONLY page items.
            $itemOrder =
                [System.Collections.ArrayList]::new()

            foreach ($fastCode in @($fastPageItemCodes)) {
                [void]$itemOrder.Add([int]$fastCode)
            }

            # Rebuild SQL predicates after page selection. Fast-page mode is
            # capped at 250 rows specifically so Access receives a small,
            # safe IN(...) list and never scans unrelated item transactions.
            if ($itemOrder.Count -eq 0) {
                $tran2ItemPredicate = "AND 1=0"
                $tran4ItemPredicate = "AND 1=0"
            }
            else {
                $fastPageSqlCodes = (
                    @($itemOrder) |
                    ForEach-Object {
                        [int]$_
                    }
                ) -join ","

                $tran2ItemPredicate =
                    "AND T.MasterCode1 IN ($fastPageSqlCodes)"

                $tran4ItemPredicate =
                    "AND MasterCode1 IN ($fastPageSqlCodes)"
            }

            Write-ReportEngineLog `
                -Stage "STOCK-FAST-PAGE" `
                -Message (
                    "Selected exact valuation page; page=" +
                    $fastEffectivePage +
                    "; pageSize=" +
                    $fastEffectivePageSize +
                    "; pageItems=" +
                    $itemOrder.Count +
                    "; matching=" +
                    $fastTotalRows
                ) `
                -Color "Green"
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
            # BUSY METHOD-5: ZERO-VALUE INWARD INTO NEGATIVE STOCK
            # ----------------------------------------------------
            #
            # Native BUSY proves that a ZERO-valued inward transaction into
            # an already-negative Material Centre does NOT offset the negative
            # valuation at the old negative-stock rate.
            #
            # Proven WE043 / LP-1696 sequence:
            #
            #   01-05-2026 Stock Journal:
            #       Qty Out = 15
            #       Amount  = 750
            #       Balance = -15
            #
            #   06-07-2026 Stock Journal:
            #       Qty In  = 1
            #       Amount  = 0
            #       Balance = -14
            #
            # Native BUSY Stock Status then reports:
            #       Qty     = -14
            #       Price   = 0
            #       Amount  = 0
            #
            # Therefore:
            #   * physical/signed valuation quantity still moves by +Qty;
            #   * the negative valuation amount is cleared to ZERO;
            #   * the resulting BUSY rate is ZERO until a later valued
            #     transaction establishes a new valuation.
            #
            # Positive-stock behaviour remains different and unchanged:
            # LP-1437 had 10 qty / 480 value, then Production +10 / 0 value,
            # and BUSY correctly reports 20 qty / 480 value = 24 each.
            # ----------------------------------------------------
            if (
                [double]$state.poolQty -lt -$eps -and
                $incomingValue -le $eps
            ) {
                $state.poolQty =
                    [double]$state.poolQty +
                    $incomingQty

                if ([Math]::Abs([double]$state.poolQty) -le $eps) {
                    $state.poolQty = 0.0
                }

                $state.poolValue = 0.0
                $state.lastRate = 0.0
                return
            }

            # ----------------------------------------------------
            # BUSY NEGATIVE-STOCK RULE FOR VALUED INWARD MOVEMENT
            # ----------------------------------------------------
            # For a valued inward row, first offset an existing negative
            # balance at the CURRENT negative-stock average. Only quantity
            # that crosses above zero uses the inward row's own valuation.
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

            # The caller supplies the BUSY voucher-specific issue rate.
            # Most ordinary outward rows use posted D2/D5 when available,
            # while Stock Transfer OUT (VchType 5) uses the source MC's
            # current method-5 valuation rate.  This function applies the
            # selected rate to the signed valuation balance.
            #
            # Proven WE043 case:
            #   Qty   = -10
            #   D2    = 50
            #   D5    = 500
            #   Value3= -400
            #
            # Using the reconstructed current average here keeps WE043 at
            # 44.53, while BUSY removes the row at its posted inventory cost
            # and reports 51.89 at 02-09-2026.
            #
            # UnitRate is therefore authoritative when supplied.  The
            # reconstructed current rate is only a fallback for rows where BUSY
            # did not persist D2/D5 (or another explicit inventory-cost hint).
            $useRate = if ($postedRate -gt $eps) {
                $postedRate
            }
            else {
                $currentRate
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
        # PERFORMANCE:
        # The old implementation self-joined ALL Tran3 rows before the
        # valuation scan. On Access this was one of the largest cold-load
        # costs. When the current request is page-scoped, resolve only return
        # references belonging to the selected page items:
        #
        #   page item -> return voucher/ref -> parent voucher
        #
        # The complete All-rows path keeps the legacy fallback.
        # --------------------------------------------------------
        try {
            if (
                $itemOrder.Count -gt 0 -and
                $itemOrder.Count -le 400
            ) {
                $returnItemCodes = (
                    @($itemOrder) |
                    ForEach-Object {
                        [int]$_
                    }
                ) -join ","

                $returnRefsByCode = @{}
                $returnRefCodes =
                    [System.Collections.ArrayList]::new()

                $returnRefQry = @"
SELECT DISTINCT
    R.VchCode AS ReturnVchCode,
    R.RefCode
FROM
    (Tran3 R
     INNER JOIN Tran1 H
        ON R.VchCode = H.VchCode)
    INNER JOIN Tran2 T
        ON R.VchCode = T.VchCode
WHERE
    T.RecType = 2
    AND T.MasterCode1 IN ($returnItemCodes)
    AND (H.VchType = 3 OR H.VchType = 10)
    AND (R.Method = 2 OR R.Method = 3)
"@

                if ($null -ne $directStockConn) {
                    $returnRefTable =
                        Invoke-StockStatusDirectTable `
                            -Connection $directStockConn `
                            -Sql $returnRefQry

                    foreach ($returnRefRow in $returnRefTable.Rows) {
                        $returnCode =
                            ConvertTo-ReportInt (
                                $returnRefRow["ReturnVchCode"]
                            )

                        $refCode =
                            ConvertTo-ReportInt (
                                $returnRefRow["RefCode"]
                            )

                        if (
                            $returnCode -gt 0 -and
                            $refCode -gt 0
                        ) {
                            $returnRefsByCode[
                                [string]$returnCode
                            ] = $refCode

                            if (-not $returnRefCodes.Contains($refCode)) {
                                [void]$returnRefCodes.Add($refCode)
                            }
                        }
                    }
                }
                else {
                    $returnRefRst =
                        $fi.GetRecordset($returnRefQry)

                    if (
                        $returnRefRst -and
                        -not $returnRefRst.EOF
                    ) {
                        $returnRefRst.MoveFirst()

                        while (-not $returnRefRst.EOF) {
                            $returnCode = ConvertTo-ReportInt (
                                Get-ReportFieldValue `
                                    $returnRefRst `
                                    "ReturnVchCode" `
                                    0
                            )

                            $refCode = ConvertTo-ReportInt (
                                Get-ReportFieldValue `
                                    $returnRefRst `
                                    "RefCode" `
                                    0
                            )

                            if (
                                $returnCode -gt 0 -and
                                $refCode -gt 0
                            ) {
                                $returnRefsByCode[
                                    [string]$returnCode
                                ] = $refCode

                                if (-not $returnRefCodes.Contains($refCode)) {
                                    [void]$returnRefCodes.Add($refCode)
                                }
                            }

                            $returnRefRst.MoveNext()
                        }
                    }

                    Close-ReportRecordset $returnRefRst
                }

                if ($returnRefCodes.Count -gt 0) {
                    $returnRefSql = (
                        @($returnRefCodes) |
                        ForEach-Object {
                            [int]$_
                        }
                    ) -join ","

                    $parentByRef = @{}

                    $parentQry = @"
SELECT
    VchCode AS ParentVchCode,
    RefCode
FROM Tran3
WHERE
    Method = 1
    AND RefCode IN ($returnRefSql)
"@

                    if ($null -ne $directStockConn) {
                        $parentTable =
                            Invoke-StockStatusDirectTable `
                                -Connection $directStockConn `
                                -Sql $parentQry

                        foreach ($parentRow in $parentTable.Rows) {
                            $parentCode =
                                ConvertTo-ReportInt (
                                    $parentRow["ParentVchCode"]
                                )

                            $refCode =
                                ConvertTo-ReportInt (
                                    $parentRow["RefCode"]
                                )

                            if (
                                $parentCode -gt 0 -and
                                $refCode -gt 0 -and
                                -not $parentByRef.ContainsKey(
                                    [string]$refCode
                                )
                            ) {
                                $parentByRef[
                                    [string]$refCode
                                ] = $parentCode
                            }
                        }
                    }
                    else {
                        $parentRst =
                            $fi.GetRecordset($parentQry)

                        if (
                            $parentRst -and
                            -not $parentRst.EOF
                        ) {
                            $parentRst.MoveFirst()

                            while (-not $parentRst.EOF) {
                                $parentCode = ConvertTo-ReportInt (
                                    Get-ReportFieldValue `
                                        $parentRst `
                                        "ParentVchCode" `
                                        0
                                )

                                $refCode = ConvertTo-ReportInt (
                                    Get-ReportFieldValue `
                                        $parentRst `
                                        "RefCode" `
                                        0
                                )

                                if (
                                    $parentCode -gt 0 -and
                                    $refCode -gt 0 -and
                                    -not $parentByRef.ContainsKey(
                                        [string]$refCode
                                    )
                                ) {
                                    $parentByRef[
                                        [string]$refCode
                                    ] = $parentCode
                                }

                                $parentRst.MoveNext()
                            }
                        }

                        Close-ReportRecordset $parentRst
                    }

                    foreach ($returnKey in @($returnRefsByCode.Keys)) {
                        $refCode =
                            [int]$returnRefsByCode[$returnKey]

                        if (
                            $parentByRef.ContainsKey(
                                [string]$refCode
                            )
                        ) {
                            $busyReturnParentMap[
                                [string]$returnKey
                            ] = [int]$parentByRef[
                                [string]$refCode
                            ]
                        }
                    }
                }

                Write-ReportEngineLog `
                    -Stage "STOCK-RETURN-LINKS" `
                    -Message (
                        "Page-scoped return links; items=" +
                        $itemOrder.Count +
                        "; links=" +
                        $busyReturnParentMap.Count
                    ) `
                    -Color "Green"
            }
            else {
                # Full / All-rows fallback. This path is intentionally not used
                # by the default page=1,pageSize=100 user experience.
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
        }
        catch {
            # Return linking improves exact parent-return valuation.
            # The engine safely falls back to current/explicit valuation if
            # an older BUSY schema does not expose the expected relation.
            Write-ReportEngineLog `
                -Stage "STOCK-RETURN-LINKS" `
                -Message (
                    "Return-link optimization fallback: " +
                    $_.Exception.Message
                ) `
                -Color "Yellow"
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
                # BUSY moving weighted-average valuation is sequence-sensitive.
                # Preserve voucher order inside each date instead of processing
                # every inward row for the whole day before every outward row.
                #
                # The phases below are intentionally kept, but they now run
                # inside ONE voucher at a time. This is required for stock
                # transfer / production / stock-journal rows that may need to
                # value an inward destination from the same voucher's source row.
                $sequencedVoucherGroups = @(
                    $dateGroup.Group |
                    Group-Object vchCode |
                    Sort-Object { [int]$_.Name }
                )

                foreach ($sequencedVoucherGroup in $sequencedVoucherGroups) {
                    $dateRows = @(
                        $sequencedVoucherGroup.Group |
                        Sort-Object srNo
                    )

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
                #
                # IMPORTANT BUSY METHOD-5 RULE:
                #
                # A Tran2 inward row that carries no BUSY inventory value
                # (D5/D2/Value3 all zero) must remain ZERO-valued.
                #
                # Do NOT manufacture value from another row in the same
                # voucher. Native BUSY MC ledgers prove this:
                #
                # SP701 / LP-1437:
                #   Stock Transfer +3  -> Amt.In = 0
                #   Production    +2   -> Amt.In = 0
                #   physical qty = 5, valuation value = 0
                #
                # WE043 / LP-1437:
                #   existing 10 qty / 480 value
                #   Production +10 / 0 value
                #   final 20 qty / 480 value = 24 each
                #
                # WE043 / LP-1696:
                #   existing -15 qty / -750 value
                #   zero-value inward +1
                #   BUSY resets valuation to 0 while qty becomes -14.
                #
                # Add-BusyWaPool owns the positive/negative state behavior;
                # this phase must pass the row's REAL zero value through.
                # ====================================================
                $zeroValueInRows = @(
                    $dateRows |
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

                foreach ($inRow in $zeroValueInRows) {
                    Add-BusyWaPool `
                        -ItemCode $ItemCode `
                        -McCode ([int]$inRow.mcCode) `
                        -Qty ([double]$inRow.qty) `
                        -Value 0.0
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
                            if ([Math]::Abs([double]$stateBeforeOut.poolValue) -le $eps) {
                                # BUSY METHOD-5 ZERO-VALUATION POOL RULE:
                                #
                                # If physical stock is positive but the MC's
                                # current stock valuation is ZERO, an outward
                                # row must NOT manufacture valuation from its
                                # posted D2/D5 fields.
                                #
                                # Proven SP701 / LP-1437:
                                #   +3 Stock Transfer @ value 0
                                #   +2 Production     @ value 0
                                #   => 5 qty / 0 value
                                #
                                # Then two Sales of 3 each carry transaction
                                # Amt.Out 1,440, but BUSY stock valuation stays:
                                #   5 -> 2 -> -1 qty
                                #   value 0 throughout
                                #
                                # Therefore the issue valuation rate is ZERO.
                                $outRate = 0.0
                            }
                            elseif ($vchType -eq 5) {
                                # BUSY Stock Transfer OUT uses the source
                                # material-centre's CURRENT method-5 valuation
                                # rate, not the Tran2 D2/D5 amount.
                                #
                                # Proven WE043 / Main Store:
                                #   opening          148 @ 50 = 7,400
                                #   transfer SJ-690  -10
                                #   Tran2 inward/outward posted value = 480
                                #
                                # Native BUSY MC ledger closes the source
                                # transfer row at a valuation balance of 6,900,
                                # which means BUSY removed 10 @ 50 = 500 from
                                # Main Store, while LP-1437 still received the
                                # posted inward value 480.
                                $outRate =
                                    Get-BusyWaRate `
                                        -ItemCode $ItemCode `
                                        -McCode $mcCode
                            }
                            else {
                                # For ordinary positive-valued stock, BUSY's
                                # posted inventory-cost D5/D2 is authoritative
                                # when present.
                                #
                                # Proven WE043 Sale:
                                #   Qty=-10, Value3=-400, D2=50, D5=500
                                # BUSY removes 500 from stock, not 400.
                                $outRate =
                                    Get-BusyWaExplicitRate `
                                        -Row $row `
                                        -Qty $outQty

                                if ($outRate -le $eps) {
                                    $outRate =
                                        Get-BusyWaRate `
                                            -ItemCode $ItemCode `
                                            -McCode $mcCode
                                }
                            }
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
        }

        # --------------------------------------------------------
        # One chronological transaction scan for every filtered item.
        # --------------------------------------------------------
        $stockWaWatch = New-ReportEngineWatch
        Write-ReportEngineLog -Stage "STOCK-WA" -Message "Starting voucher-sequenced Tran2 valuation + movement scan"

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
INNER JOIN Tran1 V ON V.VchCode = T.VchCode
WHERE
    T.RecType = 2
    $tran2ItemPredicate
    AND V.[Date] <= $endLiteral
    AND (V.Cancelled = 0 OR V.Cancelled IS NULL)
    AND (V.VchCancelled = 0 OR V.VchCancelled IS NULL)
ORDER BY
    T.MasterCode1,
    V.[Date],
    T.VchCode,
    T.SrNo
"@

        $busyWaCursorState = [pscustomobject]@{
            itemCode = 0
            rows = [System.Collections.ArrayList]::new()
        }

        $consumeBusyWaTxnRow = {
            param(
                [int]$itemCode,
                [int]$rowVchCode,
                [int]$rowMcCode,
                [int]$rowSrNo,
                [double]$rowQty,
                [double]$rowAltQty,
                [double]$rowValue3,
                [double]$rowD2,
                [double]$rowD5,
                [int]$rowVchType,
                $voucherDateRaw
            )

            if (-not $items.ContainsKey($itemCode)) {
                return
            }

            if (
                [int]$busyWaCursorState.itemCode -ne 0 -and
                $itemCode -ne [int]$busyWaCursorState.itemCode
            ) {
                Process-BusyWaItemRows `
                    -ItemCode ([int]$busyWaCursorState.itemCode) `
                    -Rows $busyWaCursorState.rows

                $busyWaCursorState.rows =
                    [System.Collections.ArrayList]::new()
            }

            $busyWaCursorState.itemCode = $itemCode

            $voucherDateText = ""
            $voucherDateValue = $null

            try {
                if (
                    $null -ne $voucherDateRaw -and
                    $voucherDateRaw -ne [System.DBNull]::Value
                ) {
                    $voucherDateValue = [datetime]$voucherDateRaw
                    $voucherDateText =
                        $voucherDateValue.ToString("yyyy-MM-dd")
                }
            }
            catch {
                $voucherDateValue = $null
                $voucherDateText = ""
            }

            $movementState = Get-StockMcState `
                -Item $items[$itemCode] `
                -McCode $rowMcCode

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
                $absStockValue =
                    [Math]::Abs([double]$rowValue3)

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

            [void]$busyWaCursorState.rows.Add(
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
        }

        if ($null -ne $directStockConn) {
            $busyWaTxnTable =
                Invoke-StockStatusDirectTable `
                    -Connection $directStockConn `
                    -Sql $busyWaTxnQry

            foreach ($txnRow in $busyWaTxnTable.Rows) {
                & $consumeBusyWaTxnRow `
                    -itemCode (
                        ConvertTo-ReportInt $txnRow["ItemCode"]
                    ) `
                    -rowVchCode (
                        ConvertTo-ReportInt $txnRow["VchCode"]
                    ) `
                    -rowMcCode (
                        ConvertTo-ReportInt $txnRow["MCCode"]
                    ) `
                    -rowSrNo (
                        ConvertTo-ReportInt $txnRow["SrNo"]
                    ) `
                    -rowQty (
                        ConvertTo-ReportDouble $txnRow["Value1"]
                    ) `
                    -rowAltQty (
                        ConvertTo-ReportDouble $txnRow["Value2"]
                    ) `
                    -rowValue3 (
                        ConvertTo-ReportDouble $txnRow["Value3"]
                    ) `
                    -rowD2 (
                        ConvertTo-ReportDouble $txnRow["D2"]
                    ) `
                    -rowD5 (
                        ConvertTo-ReportDouble $txnRow["D5"]
                    ) `
                    -rowVchType (
                        ConvertTo-ReportInt $txnRow["VchType"]
                    ) `
                    -voucherDateRaw $txnRow["VoucherDate"]
            }
        }
        else {
            $busyWaTxnRst =
                $fi.GetRecordset($busyWaTxnQry)

            if (
                $busyWaTxnRst -and
                -not $busyWaTxnRst.EOF
            ) {
                $busyWaTxnRst.MoveFirst()

                $fieldItemCode =
                    $busyWaTxnRst.Fields.Item("ItemCode")
                $fieldVchCode =
                    $busyWaTxnRst.Fields.Item("VchCode")
                $fieldMcCode =
                    $busyWaTxnRst.Fields.Item("MCCode")
                $fieldSrNo =
                    $busyWaTxnRst.Fields.Item("SrNo")
                $fieldValue1 =
                    $busyWaTxnRst.Fields.Item("Value1")
                $fieldValue2 =
                    $busyWaTxnRst.Fields.Item("Value2")
                $fieldValue3 =
                    $busyWaTxnRst.Fields.Item("Value3")
                $fieldD2 =
                    $busyWaTxnRst.Fields.Item("D2")
                $fieldD5 =
                    $busyWaTxnRst.Fields.Item("D5")
                $fieldVchType =
                    $busyWaTxnRst.Fields.Item("VchType")
                $fieldVoucherDate =
                    $busyWaTxnRst.Fields.Item("VoucherDate")

                while (-not $busyWaTxnRst.EOF) {
                    & $consumeBusyWaTxnRow `
                        -itemCode (
                            ConvertTo-ReportInt $fieldItemCode.Value
                        ) `
                        -rowVchCode (
                            ConvertTo-ReportInt $fieldVchCode.Value
                        ) `
                        -rowMcCode (
                            ConvertTo-ReportInt $fieldMcCode.Value
                        ) `
                        -rowSrNo (
                            ConvertTo-ReportInt $fieldSrNo.Value
                        ) `
                        -rowQty (
                            ConvertTo-ReportDouble $fieldValue1.Value
                        ) `
                        -rowAltQty (
                            ConvertTo-ReportDouble $fieldValue2.Value
                        ) `
                        -rowValue3 (
                            ConvertTo-ReportDouble $fieldValue3.Value
                        ) `
                        -rowD2 (
                            ConvertTo-ReportDouble $fieldD2.Value
                        ) `
                        -rowD5 (
                            ConvertTo-ReportDouble $fieldD5.Value
                        ) `
                        -rowVchType (
                            ConvertTo-ReportInt $fieldVchType.Value
                        ) `
                        -voucherDateRaw $fieldVoucherDate.Value

                    $busyWaTxnRst.MoveNext()
                }
            }

            Close-ReportRecordset $busyWaTxnRst
        }

        if (
            [int]$busyWaCursorState.itemCode -gt 0 -and
            $busyWaCursorState.rows.Count -gt 0
        ) {
            Process-BusyWaItemRows `
                -ItemCode ([int]$busyWaCursorState.itemCode) `
                -Rows $busyWaCursorState.rows
        }

        Stop-ReportEngineWatch `
            -Watch $stockWaWatch `
            -Stage "STOCK-WA" `
            -Message "Voucher-sequenced Tran2 valuation + movement scan completed"

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

        # --------------------------------------------------------
        # FAST PAGE RESPONSE
        # --------------------------------------------------------
        # $sortedRows contains only the exact-valued requested page in this
        # mode. Global quantity metrics came from the cheap quantity index.
        # The page's Stock Value is exact; calculating an exact ALL-items value
        # would defeat the purpose and block the first render again.
        if ($fastPageModeActive) {
            $fastPageStockValue = 0.0

            foreach ($fastRow in @($sortedRows)) {
                $fastPageStockValue +=
                    [double]$fastRow.stockValue
            }

            $fastSummary = @{}
            foreach ($key in @($fastGlobalSummary.Keys)) {
                $fastSummary[$key] =
                    $fastGlobalSummary[$key]
            }

            $fastSummary.stockValue =
                [Math]::Round(
                    $fastPageStockValue,
                    2
                )

            $fastSummary.busyStockValue =
                [Math]::Round(
                    $fastPageStockValue,
                    2
                )

            $fastPageSizeOptions = @(
                Get-StockStatusPageSizeOptions `
                    -ItemCount $fastTotalRows
            )

            # GROUPED / HIERARCHICAL quantity totals are GLOBAL and exact.
            # Only the expensive stock-value contribution remains page-scoped.
            $fastPageGroupValueMap = @{}

            foreach ($fastGroupRow in @($sortedRows)) {
                $fastGroupKey = if (
                    $fastGroupRow.groupPath.Count -gt 0
                ) {
                    (@($fastGroupRow.groupPath) -join " > ")
                }
                elseif (
                    [string]::IsNullOrWhiteSpace(
                        [string]$fastGroupRow.group
                    )
                ) {
                    "General"
                }
                else {
                    [string]$fastGroupRow.group
                }

                if (
                    -not $fastPageGroupValueMap.ContainsKey(
                        $fastGroupKey
                    )
                ) {
                    $fastPageGroupValueMap[$fastGroupKey] = 0.0
                }

                $fastPageGroupValueMap[$fastGroupKey] +=
                    [double]$fastGroupRow.stockValue
            }

            $fastGroupTotals = @(
                foreach (
                    $fastGlobalGroupTotal in
                    @($fastGlobalGroupTotals)
                ) {
                    $fastGroupName =
                        [string]$fastGlobalGroupTotal.group

                    $fastPageGroupValue = 0.0
                    if (
                        $fastPageGroupValueMap.ContainsKey(
                            $fastGroupName
                        )
                    ) {
                        $fastPageGroupValue =
                            [double]$fastPageGroupValueMap[
                                $fastGroupName
                            ]
                    }

                    [pscustomobject]@{
                        group =
                            [string]$fastGlobalGroupTotal.group
                        parentGroup =
                            [string]$fastGlobalGroupTotal.parentGroup
                        groupPath =
                            @($fastGlobalGroupTotal.groupPath)
                        itemCount =
                            [int]$fastGlobalGroupTotal.itemCount
                        openingQuantity =
                            [double]$fastGlobalGroupTotal.openingQuantity
                        inwardQuantity =
                            [double]$fastGlobalGroupTotal.inwardQuantity
                        outwardQuantity =
                            [double]$fastGlobalGroupTotal.outwardQuantity
                        closingQuantity =
                            [double]$fastGlobalGroupTotal.closingQuantity

                        # Page contribution only; frontend labels it.
                        stockValue =
                            [Math]::Round(
                                $fastPageGroupValue,
                                2
                            )
                    }
                }
            )

            $fastResponse = @{
                success = $true
                view = $viewMode
                fromDate = $fromDate.ToString("yyyy-MM-dd")
                toDate = $endDate.ToString("yyyy-MM-dd")
                asOfDate = $asOfDate.ToString("yyyy-MM-dd")
                generatedAt = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
                page = $fastEffectivePage
                pageSize = $fastEffectivePageSize
                total = $fastTotalRows
                totalPages = $fastTotalPages
                companyItemCount = $companyItemCount
                pageSizeOptions = @($fastPageSizeOptions)

                # Quantities/counts are global and exact. Valuation is exact
                # for the loaded page, which the frontend labels explicitly.
                summaryScope = "page"
                quantitySummaryScope = "all"
                fastPageValuation = $true

                valuationSource = if ($valueMode -eq "busy") {
                    "busy-stockvalmethod-5-v8-busy-parity-fast-cache"
                }
                else {
                    "item-master-$valueMode-price"
                }

                stockValuationMethod = 5
                stockValuationMode =
                    "busy-method5-v8-busy-parity-fast-cache-by-material-centre"

                options = @{
                    unitMode = $unitDisplay
                    showValue = $ShowValue
                    includeStockTransfers = $IncludeStockTransfers
                    showSalePurchaseSeparately =
                        $ShowSalePurchaseSeparately
                    mastersMode = $mastersFilter
                    showParentGroup = $ShowParentGroup
                    valueBy = $valueMode
                }

                filterOptions = @{
                    materialCentres = @($mcOptions)
                    itemGroups = @(
                        $groupOptions |
                        Sort-Object pathLabel, name
                    )
                }

                summary = $fastSummary
                # Exact GROUP totals for the currently loaded page.
                # The frontend labels these as page totals in fast mode.
                groupTotals = @($fastGroupTotals)
                data = @($sortedRows)
            }

            $script:StockStatusFastPageCache[
                $fastPageCacheKey
            ] = @{
                createdAt = Get-Date
                instanceId = $InstanceId
                companyCode = $CompanyCode
                response = $fastResponse
            }

            Stop-ReportEngineWatch `
                -Watch $stockTotalWatch `
                -Stage "STOCK-FAST-TOTAL" `
                -Message (
                    "Fast Stock Status page completed; matching=" +
                    $fastTotalRows +
                    "; loaded=" +
                    $sortedRows.Count
                )

            return $fastResponse
        }

        $totalRows = $sortedRows.Count

        # Generate rows-per-page choices from the final matching report rows.
        # Example: 2,573 rows => 50..2,500, All. 3,000/4,000 must not appear.
        $pageSizeOptions = @(
            Get-StockStatusPageSizeOptions -ItemCount $totalRows
        )

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

        $baseResult = @{
            success = $true
            view = $viewMode
            fromDate = $fromDate.ToString("yyyy-MM-dd")
            toDate = $endDate.ToString("yyyy-MM-dd")
            asOfDate = $asOfDate.ToString("yyyy-MM-dd")
            generatedAt = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
            page = 1
            pageSize = 0
            total = $totalRows
            totalPages = 1
            companyItemCount = $companyItemCount
            pageSizeOptions = @($pageSizeOptions)
            summaryScope = "all"
            valuationSource = if ($valueMode -eq "busy") { "busy-stockvalmethod-5-v8-busy-parity-fast-cache" } else { "item-master-$valueMode-price" }
            stockValuationMethod = 5
            stockValuationMode = "busy-method5-v8-busy-parity-fast-cache-by-material-centre"
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
            data = @()
        }

        # Cache the complete calculated rows once. Page/pageSize are intentionally
        # excluded from the cache key, so later page requests do not replay BUSY
        # method-5 valuation or transaction movement scans.
        $script:StockStatusReportCache[$stockStatusCacheKey] = @{
            createdAt = Get-Date
            instanceId = $InstanceId
            companyCode = $CompanyCode
            baseResult = $baseResult
            allRows = @($sortedRows)
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

        return New-StockStatusPagedResponse `
            -BaseResult $baseResult `
            -AllRows $sortedRows `
            -Page $Page `
            -PageSize $PageSize
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
        if ($null -ne $directStockConn) {
            try { $directStockConn.Close() } catch {}
            try { $directStockConn.Dispose() } catch {}
        }

        Disconnect-BUSY $fi
    }
}

