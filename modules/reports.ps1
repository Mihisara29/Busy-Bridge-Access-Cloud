
# ─────────────────────────────────────────────────────────────
# REPORT ENGINE DIAGNOSTICS AND CACHE
# ─────────────────────────────────────────────────────────────
if ($null -eq $script:OutstandingReportCache) {
    $script:OutstandingReportCache = @{}
}

# [OUTSTANDING-SPLIT-V1] A bounded, short-lived computed snapshot supports
# independent summary and paged-table HTTP requests without recalculating
# filtered bill amounts and financial cards for every page navigation.
if ($null -eq $script:OutstandingComputedSnapshotCache) {
    $script:OutstandingComputedSnapshotCache = @{}
}
$script:OutstandingComputedSnapshotSeconds = 90

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
Write-Host "  [OUTSTANDING-PAYMENT-SIGN-PARITY-V11] Payment new-reference pending sign parity loaded." -ForegroundColor Green
Write-Host "  [OUTSTANDING-NATIVE-CSV-PARITY-V10] Payment/Receipt direction + Credit Note parity loaded." -ForegroundColor Green
Write-Host "  [OUTSTANDING-VOUCHER-DATE-V9] Voucher-date filtering loaded." -ForegroundColor Green
Write-Host "  [OUTSTANDING-BUSY-PARITY-V8] Closed-reference pending + adjusted parity loaded." -ForegroundColor Green
Write-Host "  [OUTSTANDING-BUSY-PARITY-V7] Method3 + signed pending + accountless BUSY refs loaded." -ForegroundColor Green
Write-Host "  [OUTSTANDING-BUSY-PARITY-V6] Native BUSY voucher scope + RefStatus nature loaded." -ForegroundColor Green
Write-Host "  [OUTSTANDING-NO-CONTRA-V5] Contra excluded from Outstanding processing." -ForegroundColor Green
Write-Host "  [OUTSTANDING-CARDS-V4] Pending/adjusted financial cards loaded." -ForegroundColor Green
Write-Host "  [OUTSTANDING-PARITY-V3] New Ref nature + voucher filter parity loaded." -ForegroundColor Green

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

    # Same invalidation boundary for computed Outstanding snapshots.
    if ($null -ne $script:OutstandingComputedSnapshotCache) {
        foreach ($key in @($script:OutstandingComputedSnapshotCache.Keys)) {
            if ([string]::IsNullOrWhiteSpace($InstanceId) -or
                ([string]$key).StartsWith($prefix)) {
                $script:OutstandingComputedSnapshotCache.Remove($key)
            }
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
Write-Host "  [STOCK-ACCESS-COMPAT-V1] Access/BDS Stock Status compatibility path loaded." -ForegroundColor Green

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
        [int]$AccountCode = 0,
        [string]$Search = "",
        [string]$Group = "",
        [int]$GroupCode = 0,
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
        [string]$CompanyCode = "",
        [int[]]$AccountCodes = @()  # BUSYWEB_SCOPED_OUTSTANDING_V1: optional internal bulk account scope
    )

    # Preserve the requested filter independently from the per-row account code.
    # PowerShell variable names are CASE-INSENSITIVE: $AccountCode and
    # $accountCode refer to the same variable. The latter is reused throughout
    # account-map and reference loops, so comparing against $AccountCode was
    # silently comparing each row with itself (and showing all accounts).
    $requestedAccountMasterCode = [int]$AccountCode

    # BUSYWEB_SCOPED_OUTSTANDING_V1
    # Restrict SQL/Jet read volume for dashboard Top Customers, while keeping
    # the original native BUSY parity and sign/due-date calculations below.
    # Accept only positive numeric master codes; never interpolate arbitrary
    # request text into a SQL fragment. No AccountCodes = legacy all-company.
    $scopedAccountCodes = @(
        $AccountCodes | Where-Object { [int]$_ -gt 0 } |
            ForEach-Object { [int]$_ } | Sort-Object -Unique
    )
    if ($AccountCodes.Count -gt 0 -and $scopedAccountCodes.Count -eq 0) {
        throw 'AccountCodes must contain positive account master codes.'
    }
    if ($scopedAccountCodes.Count -gt 20) {
        throw 'Scoped Outstanding supports a maximum of 20 account codes.'
    }
    # The same integers are used by SQL Server and Jet; no arbitrary SQL
    # values are accepted. Keep the scope inside the aggregate input queries.
    $accountScopeFilter = ''
    $referenceScopeFilter = ''
    $scopeCacheSuffix = ''
    if ($scopedAccountCodes.Count -gt 0) {
        $codeCsv = (($scopedAccountCodes | ForEach-Object { [string]$_ }) -join ',')
        $accountScopeFilter = " AND A.Code IN ($codeCsv)"
        $referenceScopeFilter = " AND T.MasterCode1 IN ($codeCsv)"
        $scopeCacheSuffix = "|accounts=$codeCsv"
        Write-ReportEngineLog -Stage 'SCOPED-ACCOUNTS' -Message (
            'Requested accounts=' + $scopedAccountCodes.Count
        )
    }

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

        # Dashboard enrichment must never hold the worker for minutes on a
        # single small account batch. If DB is congested, preserve sales/item
        # rankings and leave financial enrichment unavailable for this run.
        if ($isSql -and $scopedAccountCodes.Count -gt 0 -and $Timeout -gt 30) {
            $Timeout = 30
        }
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
            [double]$OriginalSigned,
            [int]$VoucherType = 0,
            [int]$RefStatus = 0
        )

        # BUSY native special cases proved by exported Bills reports:
        #   Payment (VchType 19) new references belong to Bills Receivable.
        #   Receipt (VchType 14) new references belong to Bills Payable.
        #
        # Do this BEFORE RefStatus. Some cheque-return Payment references are
        # stored with Status=1 even though native BUSY still lists them under
        # Bills Receivable. Trusting RefStatus first incorrectly moved those
        # rows to Payable and netted their negative pending values into the
        # Payable total.
        if ($VoucherType -eq 19) {
            return "receivable"
        }

        if ($VoucherType -eq 14) {
            return "payable"
        }

        # For the remaining voucher types, BUSY reference status is the best
        # direction signal. This keeps Journal / Debit Note side handling.
        #   Status = 2 -> Bills Receivable
        #   Status = 1 -> Bills Payable
        if ($RefStatus -eq 2) {
            return "receivable"
        }

        if ($RefStatus -eq 1) {
            return "payable"
        }

        # Compatibility fallback for old/unusual records with no reference
        # status. Preserve the previous group/sign behaviour.
        $normalisedGroup =
            ([string]$GroupName).Trim().ToLowerInvariant()

        foreach ($token in @(
            "customer",
            "sundry debtor",
            "debtor",
            "trade debtor",
            "receivable"
        )) {
            if ($normalisedGroup.Contains($token)) {
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
            if ($normalisedGroup.Contains($token)) {
                return "payable"
            }
        }

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
            6  { return "Production" }
            8  { return "Stock Journal" }
            9  { return "Sale" }
            10 { return "Purchase Return" }
            11 { return "Material Issue" }
            12 { return "Sale Order" }
            13 { return "Purchase Order" }
            14 { return "Receipt" }
            15 { return "Contra" }
            16 { return "Journal" }
            17 { return "Debit Note" }
            18 { return "Credit Note" }
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
            6  { return "Prod" }
            8  { return "SJrl" }
            9  { return "Sale" }
            10 { return "PrRt" }
            11 { return "MtIs" }
            12 { return "SlOd" }
            13 { return "PrOd" }
            14 { return "Rcpt" }
            15 { return "Cntr" }
            16 { return "Jrnl" }
            17 { return "DrNt" }
            18 { return "CrNt" }
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
            "|$dbType" +
            $scopeCacheSuffix
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
    A.ParentGrp AS AccountGroupCode,
    G.Name AS AccountGroup
FROM Master1 A
LEFT JOIN Master1 G
    ON G.Code = A.ParentGrp
WHERE A.MasterType = 2$accountScopeFilter
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
                THEN V.[Date]
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

    MAX(
        CASE
            WHEN T.Method = 1
                THEN ISNULL(T.Status, 0)
            ELSE 0
        END
    ) AS RefStatus,

    MAX(
        CASE
            WHEN T.Method = 1
                THEN ISNULL(T.TranType, 0)
            ELSE 0
        END
    ) AS RefTranType,

    SUM(
        CASE
            WHEN T.Method = 1
                THEN ISNULL(T.Value1, 0)
            ELSE 0
        END
    ) AS OriginalSigned,

    SUM(
        CASE
            WHEN T.Method IN (2, 3)
                THEN ISNULL(T.Value1, 0)
            ELSE 0
        END
    ) AS AdjustmentSigned,

    SUM(
        CASE
            WHEN T.Method IN (1, 2, 3)
                THEN ISNULL(T.Value1, 0)
            ELSE 0
        END
    ) AS PendingSigned,

    SUM(
        CASE
            WHEN T.Method IN (2, 3)
                THEN 1
            ELSE 0
        END
    ) AS AdjustmentCount

FROM Tran3 T
INNER JOIN Tran1 V
    ON V.VchCode = T.VchCode

WHERE
    T.RecType = 1
    $referenceScopeFilter
    AND T.Method IN (1, 2, 3)
    AND (
        (
            T.Method = 1
            AND T.VchType IN (2, 3, 6, 8, 9, 10, 14, 16, 17, 18, 19)
        )
        OR (
            T.Method IN (2, 3)
            AND T.VchType <> 15
        )
    )
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
            WHEN T.Method IN (1, 2, 3)
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
            V.[Date],
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

    MAX(
        IIF(
            T.Method = 1,
            IIF(
                T.Status IS NULL,
                0,
                T.Status
            ),
            0
        )
    ) AS RefStatus,

    MAX(
        IIF(
            T.Method = 1,
            IIF(
                T.TranType IS NULL,
                0,
                T.TranType
            ),
            0
        )
    ) AS RefTranType,

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
            (T.Method = 2 OR T.Method = 3) AND T.VchType <> 15,
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
            T.Method = 1 OR (
                (T.Method = 2 OR T.Method = 3)
                AND T.VchType <> 15
            ),
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
            (T.Method = 2 OR T.Method = 3) AND T.VchType <> 15,
            1,
            0
        )
    ) AS AdjustmentCount

FROM Tran3 T
INNER JOIN Tran1 V
    ON V.VchCode = T.VchCode

WHERE
    T.RecType = 1
    $referenceScopeFilter
    AND T.Method IN (1, 2, 3)
    AND (
        (
            T.Method = 1
            AND T.VchType IN (2, 3, 6, 8, 9, 10, 14, 16, 17, 18, 19)
        )
        OR (
            (T.Method = 2 OR T.Method = 3)
            AND T.VchType <> 15
        )
    )
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
            T.Method = 1 OR (
                (T.Method = 2 OR T.Method = 3)
                AND T.VchType <> 15
            ),
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
    $referenceScopeFilter
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
    $referenceScopeFilter
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
                    groupCode = ConvertTo-ReportInt $row.AccountGroupCode
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

        # Native BUSY Ref. Amount totals include zero-pending references.
        # These accumulators preserve their adjustment contribution while the
        # visible web register can continue to hide zero-pending rows.
        $closedReceivableAdjustedTotal = 0.0
        $closedPayableAdjustedTotal = 0.0

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

        # Exact code-based filters for the selected debtor/creditor tree.
        # No fallback to substring matching when an invalid code is supplied.
        $allowedGroupCodes = @{}
        if ($GroupCode -gt 0 -or $requestedAccountMasterCode -gt 0) {
            $partyOptions = Get-OutstandingPartyOptions -InstanceId $InstanceId -CompanyCode $CompanyCode
            if (-not $partyOptions.success) { throw ([string]$partyOptions.error) }
            if ($GroupCode -gt 0) {
                $matchedGroup = @($partyOptions.groups | Where-Object { [int]$_.code -eq $GroupCode })
                if ($matchedGroup.Count -ne 1) { throw 'Selected group is outside the Sundry Debtors/Creditors trees.' }
                foreach ($optionGroup in @($partyOptions.groups)) {
                    if (@($optionGroup.pathCodes) -contains $GroupCode) {
                        $allowedGroupCodes[[string]$optionGroup.code] = $true
                    }
                }
            }
            if ($requestedAccountMasterCode -gt 0) {
                $matchedAccount = @($partyOptions.accounts | Where-Object { [int]$_.code -eq $requestedAccountMasterCode })
                if ($matchedAccount.Count -ne 1) { throw 'Selected account is outside the Sundry Debtors/Creditors trees.' }
                if ($GroupCode -gt 0 -and -not $allowedGroupCodes.ContainsKey(([string]$matchedAccount[0].groupCode))) {
                    throw 'Selected account does not belong to the selected group.'
                }
            }
        }

        foreach ($referenceRow in $referenceRows) {
            $filterCounters.totalReferences++

            $accountCode =
                ConvertTo-ReportInt `
                    $referenceRow.AccountCode

            $voucherTypeValue =
                ConvertTo-ReportInt `
                    $referenceRow.VoucherType

            $isNativeAccountlessReference = (
                $accountCode -le 0 -and
                $voucherTypeValue -in @(6, 8)
            )

            if (
                -not $isNativeAccountlessReference -and
                (
                    $accountCode -le 0 -or
                    -not $accountMap.ContainsKey(
                        $accountCode
                    )
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

            # BUSY displays the pending balance from the reference direction:
            # receivable refs are stored with the opposite sign to payable refs.
            # Keep the sign because BUSY can show an over-adjusted bill as a
            # negative pending amount (for example -960.00).
            $refStatusForPending =
                ConvertTo-ReportInt `
                    $referenceRow.RefStatus

            $refTranType =
                ConvertTo-ReportInt `
                    $referenceRow.RefTranType

            # Native BUSY parity:
            # A Method=1 bill-reference row with TranType=1 is retained by
            # BUSY in the Bills report, but its Pending amount is zero.
            # The supplied native BUSY report proves this for
            # RECLP1696-428 (RefCode 2060, Ref Amt 86,400, Pending 0).
            $pendingAmount = if ($refTranType -eq 1) {
                0.0
            }
            elseif (
                $voucherTypeValue -eq 19 -and
                $refStatusForPending -eq 1 -and
                $pendingSigned -lt 0
            ) {
                # Native BUSY Payment parity:
                # cheque-return Payment new references can carry Status=1 and
                # a negative stored balance, while BUSY still displays them in
                # Bills Receivable as a positive Pending amount.
                #
                # V10 fixed their Receivable/Payable classification but left the
                # stored negative sign untouched. That made Receivable smaller by
                # the same amount instead of larger by it.
                [Math]::Round([Math]::Abs($pendingSigned), 2)
            }
            elseif ($refStatusForPending -eq 2) {
                [Math]::Round((-1.0 * $pendingSigned), 2)
            }
            elseif ($refStatusForPending -eq 1) {
                [Math]::Round($pendingSigned, 2)
            }
            else {
                [Math]::Round([Math]::Abs($pendingSigned), 2)
            }

            $pendingMagnitude =
                [Math]::Abs($pendingAmount)

            # Do not discard zero-pending references yet. Native BUSY includes
            # them in Ref. Amount / adjusted totals even when the transaction
            # register hides them. We defer the visible-row decision until all
            # business filters have been evaluated.
            $isZeroPending =
                ($pendingMagnitude -lt 0.005)

            if (
                $MinAmount -gt 0 -and
                $pendingMagnitude -lt $MinAmount
            ) {
                $filterCounters.belowMinimum++
                continue
            }

            if (
                $MaxAmount -gt 0 -and
                $pendingMagnitude -gt $MaxAmount
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

            # From/To are voucher-date filters. RefDate comes from Tran1.Date
            # for the original Method=1 voucher. Method 2/3 reference movements
            # remain controlled by the AsOf cutoff through Tran3.Date in SQL.
            if ($refDate.Date -lt $fromDate.Date) {
                $filterCounters.outsideFromDate++
                continue
            }

            if ($refDate.Date -gt $toDate.Date) {
                $filterCounters.outsideToDate++
                continue
            }

            $accountData = if ($isNativeAccountlessReference) {
                @{
                    code = 0
                    name = ""
                    alias = ""
                    groupName = ""
                }
            }
            else {
                $accountMap[$accountCode]
            }

            $voucherCode =
                ConvertTo-ReportInt `
                    $referenceRow.VoucherCode

            $refStatus =
                ConvertTo-ReportInt `
                    $referenceRow.RefStatus

            $nature = Resolve-OutstandingNature `
                -GroupName $accountData.groupName `
                -OriginalSigned $originalSigned `
                -VoucherType $voucherTypeValue `
                -RefStatus $refStatus

            if ($requestedAccountMasterCode -gt 0 -and $accountCode -ne $requestedAccountMasterCode) {
                $filterCounters.accountFilter++
                continue
            }
            if ($GroupCode -gt 0 -and
                ($isNativeAccountlessReference -or
                 -not $allowedGroupCodes.ContainsKey([string]$accountData.groupCode))) {
                $filterCounters.groupFilter++
                continue
            }
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

            # Net adjustment in the same arithmetic used by BUSY's totals.
            $adjustedAmount = [Math]::Round(
                $originalAmount - $pendingAmount,
                2
            )

            $partiallyAdjusted =
                [Math]::Abs($adjustedAmount) -gt 0.005 -and
                $pendingMagnitude -gt 0.005

            $unadjusted =
                [Math]::Abs($adjustedAmount) -lt 0.005 -and
                $pendingMagnitude -gt 0.005

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

            if (
                -not $IncludeZero -and
                $isZeroPending
            ) {
                $filterCounters.zeroPending++

                # BUSY still includes a closed reference's Ref. Amount when
                # deriving the Adjusted total:
                #   Adjusted = Ref. Amount - Pending
                if ($nature -eq "receivable") {
                    $closedReceivableAdjustedTotal =
                        [Math]::Round(
                            $closedReceivableAdjustedTotal +
                            $adjustedAmount,
                            2
                        )
                }
                elseif ($nature -eq "payable") {
                    $closedPayableAdjustedTotal =
                        [Math]::Round(
                            $closedPayableAdjustedTotal +
                            $adjustedAmount,
                            2
                        )
                }

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

        # Financial overview card totals.
        # Pending totals are the current open balances; adjusted totals are the
        # amounts already settled/adjusted against the currently matched open
        # references, split by receivable/payable nature.
        $receivableOriginalTotal = [Math]::Round(
            (@(
                $allRows |
                Where-Object {
                    $_.nature -eq "receivable"
                } |
                Measure-Object `
                    originalAmount `
                    -Sum
            ).Sum),
            2
        )

        $payableOriginalTotal = [Math]::Round(
            (@(
                $allRows |
                Where-Object {
                    $_.nature -eq "payable"
                } |
                Measure-Object `
                    originalAmount `
                    -Sum
            ).Sum),
            2
        )

        # Match BUSY's report arithmetic:
        # Adjusted Total = Ref. Amount Total - Pending Amount Total.
        # This remains a NET adjustment, so references whose pending amount
        # increased correctly reduce the adjusted total.
        $receivableAdjustedTotal = [Math]::Round(
            (
                $receivableOriginalTotal -
                $totalReceivable +
                $closedReceivableAdjustedTotal
            ),
            2
        )

        $payableAdjustedTotal = [Math]::Round(
            (
                $payableOriginalTotal -
                $totalPayable +
                $closedPayableAdjustedTotal
            ),
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
            Where-Object {
                (ConvertTo-ReportInt $_.accountCode) -gt 0
            } |
            Select-Object `
                -ExpandProperty accountCode `
                -Unique
        )

        $receivableAccounts = @(
            $allRows |
            Where-Object {
                $_.nature -eq "receivable" -and
                (ConvertTo-ReportInt $_.accountCode) -gt 0
            } |
            Select-Object `
                -ExpandProperty accountCode `
                -Unique
        ).Count

        $payableAccounts = @(
            $allRows |
            Where-Object {
                $_.nature -eq "payable" -and
                (ConvertTo-ReportInt $_.accountCode) -gt 0
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

        try {
            $parityGroups = @(
                $allRows |
                Group-Object nature, voucherTypeShortName
            )

            foreach ($parityGroup in $parityGroups) {
                $parts = [string]$parityGroup.Name -split ", "
                $parityNature = if ($parts.Count -gt 0) { $parts[0] } else { "" }
                $parityType = if ($parts.Count -gt 1) { $parts[1] } else { "" }

                $parityPending = [Math]::Round(
                    (@(
                        $parityGroup.Group |
                        Measure-Object pendingAmount -Sum
                    ).Sum),
                    2
                )

                Write-Host (
                    "  [OUTSTANDING-PARITY] nature={0} type={1} rows={2} pending={3}" -f
                    $parityNature,
                    $parityType,
                    @($parityGroup.Group).Count,
                    $parityPending
                ) -ForegroundColor DarkCyan
            }
        }
        catch {}

        Write-Host (
            "  [OUTSTANDING-CLOSED-REFS] receivableAdjusted={0}; payableAdjusted={1}" -f
            [Math]::Round($closedReceivableAdjustedTotal, 2),
            [Math]::Round($closedPayableAdjustedTotal, 2)
        ) -ForegroundColor DarkCyan

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
                # Explicit card fields. Keep the legacy totals below for
                # compatibility with the rest of the report UI.
                receivablePendingTotal = $totalReceivable
                receivableAdjustedTotal = $receivableAdjustedTotal
                payablePendingTotal = $totalPayable
                payableAdjustedTotal = $payableAdjustedTotal

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
# [OUTSTANDING-PARTY-AGING-V2] Party master selections + aging drill-down
# Only system Sundry Debtors (116) / Sundry Creditors (117) descendants.
# The financial engine above remains the single source of truth.
# ============================================================
if ($null -eq $script:OutstandingPartyOptionsCache) {
    $script:OutstandingPartyOptionsCache = @{}
}
function Get-OutstandingPartyOptions {
    param([string]$InstanceId = '', [string]$CompanyCode = '')
    $key = "$InstanceId|$CompanyCode".ToLowerInvariant()
    if ($script:OutstandingPartyOptionsCache.ContainsKey($key)) {
        $cached = $script:OutstandingPartyOptionsCache[$key]
        if ([datetime]::UtcNow -lt [datetime]$cached.expires) {
            return $cached.response
        }
        $script:OutstandingPartyOptionsCache.Remove($key)
    }
    $conn = $null
    $command = $null
    $reader = $null
    $rst = $null
    $rows = [System.Collections.Generic.List[object]]::new()
    try {
        $instance = Get-InstanceConfig -InstanceId $InstanceId
        if ($null -eq $instance) { throw 'BUSY instance not found.' }
        $qry = 'SELECT Code, Name, Alias, ParentGrp, MasterType FROM Master1 WHERE MasterType IN (1,2) ORDER BY MasterType, Name, Code'
        if ([int]$instance.dbType -eq 1) {
            $direct = Get-DirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
            if ($null -eq $direct -or $null -eq $direct.connection) { throw 'Direct SQL connection failed.' }
            $conn = $direct.connection
            if ($conn.State.ToString() -ne 'Open') { $conn.Open() }
            $command = $conn.CreateCommand()
            $command.CommandText = $qry
            $command.CommandTimeout = 60
            $reader = $command.ExecuteReader()
            while ($reader.Read()) {
                $rows.Add(@{
                    code = ConvertTo-ReportInt $reader['Code']
                    name = if ($reader['Name'] -eq [System.DBNull]::Value) { '' } else { [string]$reader['Name'] }
                    alias = if ($reader['Alias'] -eq [System.DBNull]::Value) { '' } else { [string]$reader['Alias'] }
                    parentCode = ConvertTo-ReportInt $reader['ParentGrp']
                    type = ConvertTo-ReportInt $reader['MasterType']
                })
            }
        }
        else {
            $conn = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
            if ($null -eq $conn) { throw 'Access BUSY connection failed.' }
            $rst = $conn.GetRecordset($qry)
            if ($rst -and -not $rst.EOF) {
                $rst.MoveFirst()
                while (-not $rst.EOF) {
                    $rows.Add(@{
                        code = ConvertTo-ReportInt (Get-ReportFieldValue $rst 'Code' 0)
                        name = [string](Get-ReportFieldValue $rst 'Name' '')
                        alias = [string](Get-ReportFieldValue $rst 'Alias' '')
                        parentCode = ConvertTo-ReportInt (Get-ReportFieldValue $rst 'ParentGrp' 0)
                        type = ConvertTo-ReportInt (Get-ReportFieldValue $rst 'MasterType' 0)
                    })
                    $rst.MoveNext()
                }
            }
        }
        $allGroups = @{}
        foreach ($row in $rows) {
            if ([int]$row.type -eq 1) { $allGroups[[string]$row.code] = $row }
        }
        $rootCodes = @(116,117)
        if (-not $allGroups.ContainsKey('116') -or -not $allGroups.ContainsKey('117')) {
            throw 'Required BUSY Sundry Debtors/Creditors master codes 116/117 are missing.'
        }
        $eligible = @{}
        $groups = [System.Collections.Generic.List[object]]::new()
        foreach ($g in $allGroups.Values) {
            $cursor = [int]$g.code
            $pathCodes = [System.Collections.Generic.List[int]]::new()
            $pathNames = [System.Collections.Generic.List[string]]::new()
            $seen = @{}
            for ($depth=0; $depth -lt 80; $depth++) {
                if ($seen.ContainsKey([string]$cursor) -or -not $allGroups.ContainsKey([string]$cursor)) { break }
                $seen[[string]$cursor] = $true
                $p = $allGroups[[string]$cursor]
                $pathCodes.Insert(0, [int]$p.code)
                $pathNames.Insert(0, [string]$p.name)
                if ($rootCodes -contains $cursor) { break }
                $cursor = [int]$p.parentCode
            }
            if ($pathCodes.Count -gt 0 -and $rootCodes -contains [int]$pathCodes[0]) {
                $eligible[[string]$g.code] = $true
                $groups.Add(@{
                    code = [int]$g.code
                    parentCode = [int]$g.parentCode
                    rootCode = [int]$pathCodes[0]
                    name = [string]$g.name
                    path = @($pathNames.ToArray())
                    pathCodes = @($pathCodes.ToArray())
                })
            }
        }
        $accounts = [System.Collections.Generic.List[object]]::new()
        foreach ($row in $rows) {
            if ([int]$row.type -ne 2 -or -not $eligible.ContainsKey([string]$row.parentCode)) { continue }
            $accounts.Add(@{
                code = [int]$row.code
                name = [string]$row.name
                alias = [string]$row.alias
                groupCode = [int]$row.parentCode
            })
        }
        $result = @{
            success = $true
            groups = @($groups.ToArray() | Sort-Object { $_.path -join ' > ' })
            accounts = @($accounts.ToArray() | Sort-Object name, code)
        }
        $script:OutstandingPartyOptionsCache[$key] = @{
            expires = [datetime]::UtcNow.AddSeconds(120)
            response = $result
        }
        return $result
    }
    catch {
        return @{ success = $false; error = "Party master lookup failed: $($_.Exception.Message)"; groups=@(); accounts=@() }
    }
    finally {
        if ($null -ne $reader) { try { $reader.Close() } catch {}; try { $reader.Dispose() } catch {} }
        if ($null -ne $command) { try { $command.Dispose() } catch {} }
        if ($null -ne $rst) { try { $rst.Close() } catch {} }
        if ($null -ne $conn -and [int](Get-InstanceConfig -InstanceId $InstanceId).dbType -eq 1) {
            try { $conn.Dispose() } catch {}
        }
    }
}

# Aging is an on-demand projection of BUSY's authoritative outstanding snapshot.
# Matrix cells total pending balances; details use the SAME snapshotId and rows.
function Get-OutstandingAgingReport {
    param(
        [ValidateSet('matrix','details')][string]$Mode = 'matrix',
        [string]$From = '', [string]$To = '', [string]$AsOf = '',
        [string]$Search = '', [string]$Group = '', [string]$Account = '',
        [int]$GroupCode = 0, [int]$AccountCode = 0,
        [int]$VoucherType = 0, [double]$MinAmount = 0, [double]$MaxAmount = 0,
        [string]$InstanceId = '', [string]$CompanyCode = '',
        [string]$SnapshotId = '', [int]$CellAccountCode = 0,
        [string]$Nature = '', [string]$Bucket = '',
        [int]$Page = 1, [int]$PageSize = 100
    )
    $entry = $null
    if ($Mode -eq 'matrix') {
        $result = Get-OutstandingSeparatedReport -Mode 'summary' `
            -From $From -To $To -AsOf $AsOf -Type 'all' `
            -Account $Account -AccountCode $AccountCode `
            -Group $Group -GroupCode $GroupCode -Search $Search `
            -Status 'all' -Aging 'all' -VoucherType $VoucherType `
            -MinAmount $MinAmount -MaxAmount $MaxAmount -IncludeZero $false `
            -InstanceId $InstanceId -CompanyCode $CompanyCode
        if (-not $result.success) { return $result }
        $SnapshotId = [string]$result.snapshotId
    }
    else {
        if ([string]::IsNullOrWhiteSpace($SnapshotId)) { return @{success=$false;error='Missing aging snapshot ID.'} }
    }
    $prefix = "$InstanceId|$CompanyCode|"
    foreach ($key in @($script:OutstandingComputedSnapshotCache.Keys)) {
        $candidate = $script:OutstandingComputedSnapshotCache[$key]
        if ($key.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase) -and
            [string]$candidate.snapshotId -eq $SnapshotId -and
            ([datetime]::UtcNow - [datetime]$candidate.createdUtc).TotalSeconds -le 90) {
            $entry = $candidate
            break
        }
    }
    if ($null -eq $entry) { return @{success=$false;code='OUTSTANDING_SNAPSHOT_EXPIRED';error='Aging snapshot expired. Reopen View Aging to refresh.'} }
    $rows = @($entry.report.data)
    $bucketNames = @('notDue','days0To30','days31To60','days61To90','days91To180','above180')
    if ($Mode -eq 'details') {
        if ($Nature -notin @('receivable','payable') -or ($Bucket -and $Bucket -notin $bucketNames)) {
            return @{success=$false;error='Invalid aging cell selection.'}
        }
        if ($Page -lt 1) { $Page = 1 }
        if ($PageSize -lt 1 -or $PageSize -gt 500) { $PageSize = 100 }
        $matched = @($rows | Where-Object {
            $_.nature -eq $Nature -and
            ($CellAccountCode -eq 0 -or [int]$_.accountCode -eq $CellAccountCode) -and
            (-not $Bucket -or [string]$_.agingBucket -eq $Bucket)
        } | Sort-Object dueDate, accountName, refNo)
        # Do not permit arbitrary account codes that are absent from the matrix.
        if ($CellAccountCode -gt 0 -and -not @($rows | Where-Object { [int]$_.accountCode -eq $CellAccountCode }).Count) {
            return @{success=$false;error='Account is not available in this aging snapshot.'}
        }
        $total = $matched.Count
        $pages = [Math]::Max(1,[int][Math]::Ceiling($total / [double]$PageSize))
        if ($Page -gt $pages) { $Page = $pages }
        return @{
            success = $true; mode = 'details'; snapshotId = $SnapshotId
            total = $total; page = $Page; pageSize = $PageSize; totalPages = $pages
            data = @($matched | Select-Object -Skip (($Page-1)*$PageSize) -First $PageSize)
        }
    }
    # Aggregate directly from authoritative bill rows. Group-Object's Group
    # projection in the original implementation could leave a wrapper in the
    # matrix builder, producing accountCode=0 and zero amounts in PS 5.1.
    # Use a compound (nature, accountCode) key so payable and receivable rows
    # for the same account are never merged.
    $matrixByAccount = @{}
    foreach ($bill in $rows) {
        if ($null -eq $bill) { continue }
        $billNature = [string]$bill.nature
        if ($billNature -notin @('receivable','payable')) {
            return @{ success=$false; code='OUTSTANDING_AGING_INVALID_NATURE'; error='An outstanding bill has an unsupported nature.' }
        }
        $billAccountCode = [int]$bill.accountCode
        if ($billAccountCode -le 0) {
            return @{ success=$false; code='OUTSTANDING_AGING_INVALID_ACCOUNT'; error='An outstanding bill has an invalid account code.' }
        }
        $billBucket = [string]$bill.agingBucket
        if ($billBucket -notin $bucketNames) {
            return @{ success=$false; code='OUTSTANDING_AGING_INVALID_BUCKET'; error='An outstanding bill has an unrecognized aging bucket.' }
        }
        try {
            $billAmount = [decimal]$bill.pendingAmount
        }
        catch {
            return @{ success=$false; code='OUTSTANDING_AGING_INVALID_AMOUNT'; error='An outstanding bill has a nonnumeric pending amount.' }
        }
        $matrixKey = '{0}|{1}' -f $billNature, $billAccountCode
        if (-not $matrixByAccount.ContainsKey($matrixKey)) {
            $initialBuckets = @{}
            foreach ($bucketName in $bucketNames) { $initialBuckets[$bucketName] = [decimal]0 }
            $matrixByAccount[$matrixKey] = @{
                accountCode = $billAccountCode
                accountName = [string]$bill.accountName
                groupName = [string]$bill.accountGroup
                nature = $billNature
                count = 0
                buckets = $initialBuckets
                total = [decimal]0
            }
        }
        $accountRow = $matrixByAccount[$matrixKey]
        $accountRow['count'] = [int]$accountRow['count'] + 1
        $accountRow['buckets'][$billBucket] = [decimal]$accountRow['buckets'][$billBucket] + $billAmount
    }
    $matrix = [System.Collections.Generic.List[object]]::new()
    foreach ($matrixKey in @($matrixByAccount.Keys)) {
        $accountRow = $matrixByAccount[$matrixKey]
        $totalAmount = [decimal]0
        foreach ($bucketName in $bucketNames) {
            $roundedAmount = [Math]::Round([decimal]$accountRow['buckets'][$bucketName], 2)
            $accountRow['buckets'][$bucketName] = $roundedAmount
            $totalAmount += $roundedAmount
        }
        $accountRow['total'] = [Math]::Round($totalAmount,2)
        $matrix.Add($accountRow)
    }
    $totals = @{}
    foreach ($natureValue in @('receivable','payable')) {
        $byBucket = @{}
        foreach ($b in $bucketNames) { $byBucket[$b] = [decimal]0 }
        $natureSum = [decimal]0
        foreach ($entryRow in @($matrix | Where-Object { $_.nature -eq $natureValue })) {
            foreach ($b in $bucketNames) { $byBucket[$b] += [decimal]$entryRow.buckets[$b] }
            $natureSum += [decimal]$entryRow.total
        }
        $totals[$natureValue] = @{ buckets=$byBucket; total=[Math]::Round($natureSum,2) }
    }
    return @{
        success = $true; mode = 'matrix'; snapshotId = $SnapshotId
        fromDate = $entry.report.fromDate; toDate=$entry.report.toDate; asOfDate=$entry.report.asOfDate
        rows = @($matrix.ToArray() | Sort-Object nature, accountName, accountCode)
        totals = $totals
        buckets = $bucketNames
        totalBills = $rows.Count
    }
}

# ============================================================
# [OUTSTANDING-SPLIT-V1] SUMMARY AND TABLE RESPONSE CONTRACTS
# ============================================================
# 1. First summary request computes an authoritative full filtered snapshot
#    using the original BUSY parity-verified Get-OutstandingReport engine.
# 2. Page/size/sort requests reuse that computed snapshot, never recomputing
#    financial cards. This is snapshot paging, NOT true SQL OFFSET pagination.
# 3. Separate endpoint payloads prevent transferring all rows for summaries.
# 4. Expiring snapshots ensure old data cannot remain available indefinitely.
#    An explicit Refresh forces re-computation; writes should call
#    Clear-OutstandingReportCache to invalidate both underlying caches.
# The original endpoint and accounting formulas remain completely unchanged.
function Get-OutstandingSeparatedReport {
    param(
        [ValidateSet('summary','table')][string]$Mode = 'table',
        [string]$From = '',
        [string]$To = '',
        [string]$AsOf = '',
        [string]$Type = 'all',
        [string]$Account = '',
        [int]$AccountCode = 0,
        [string]$Search = '',
        [string]$Group = '',
        [int]$GroupCode = 0,
        [string]$Status = 'all',
        [string]$Aging = 'all',
        [int]$VoucherType = 0,
        [double]$MinAmount = 0,
        [double]$MaxAmount = 0,
        [bool]$IncludeZero = $false,
        [int]$Page = 1,
        [int]$PageSize = 50,
        [string]$SortBy = 'dueDate',
        [string]$SortDirection = 'asc',
        [string]$InstanceId = '',
        [string]$CompanyCode = '',
        [bool]$ForceRefresh = $false,
        [string]$ExpectedSnapshotId = ''
    )

    if ($null -eq $script:OutstandingComputedSnapshotCache) {
        $script:OutstandingComputedSnapshotCache = @{}
    }
    $ttl = 90
    $signature = [ordered]@{
        from = $From
        to = $To
        asOf = $AsOf
        type = $Type
        account = $Account
        accountCode = $AccountCode
        search = $Search
        group = $Group
        groupCode = $GroupCode
        status = $Status
        aging = $Aging
        voucherType = $VoucherType
        minAmount = $MinAmount
        maxAmount = $MaxAmount
        includeZero = $IncludeZero
    }
    $key = "$InstanceId|$CompanyCode|" + ($signature | ConvertTo-Json -Compress)
    $now = [datetime]::UtcNow
    $entry = $null

    # Bound memory. Expires all snapshots even if the current request has a
    # different filter signature. Limits retained full-row arrays to 12.
    foreach ($otherKey in @($script:OutstandingComputedSnapshotCache.Keys)) {
        $candidate = $script:OutstandingComputedSnapshotCache[$otherKey]
        if ($null -eq $candidate -or $null -eq $candidate.createdUtc -or
            ($now - [datetime]$candidate.createdUtc).TotalSeconds -gt $ttl) {
            $script:OutstandingComputedSnapshotCache.Remove($otherKey)
        }
    }
    if (-not $ForceRefresh -and
        $script:OutstandingComputedSnapshotCache.ContainsKey($key)) {
        $entry = $script:OutstandingComputedSnapshotCache[$key]
    }
    $hit = ($null -ne $entry)

    # Never build a different full dataset behind an already-displayed card.
    # The frontend refreshes summary first when a snapshot has expired.
    if ($Mode -eq 'table' -and
        (-not $hit -or [string]$entry.snapshotId -ne $ExpectedSnapshotId)) {
        return @{
            success = $false
            code = 'OUTSTANDING_SNAPSHOT_EXPIRED'
            error = 'OUTSTANDING_SNAPSHOT_EXPIRED: The report snapshot changed or expired. Refresh the financial summary.'
        }
    }

    if (-not $hit) {
        if ($ForceRefresh) {
            # Bypass the 20-second raw-source cache as well as computed results.
            Clear-OutstandingReportCache -InstanceId $InstanceId -CompanyCode $CompanyCode
        }
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $full = Get-OutstandingReport `
            -From $From -To $To -AsOf $AsOf `
            -Type $Type -Account $Account -AccountCode $AccountCode `
            -Search $Search -Group $Group -GroupCode $GroupCode `
            -Status $Status -Aging $Aging -VoucherType $VoucherType `
            -MinAmount $MinAmount -MaxAmount $MaxAmount `
            -IncludeZero $IncludeZero -Page 1 -PageSize 0 `
            -SortBy 'dueDate' -SortDirection 'asc' `
            -InstanceId $InstanceId -CompanyCode $CompanyCode
        $watch.Stop()
        if ($null -eq $full -or $full.success -ne $true) {
            return $full
        }
        $entry = @{
            createdUtc = [datetime]::UtcNow
            snapshotId = [guid]::NewGuid().ToString('N')
            report = $full
        }
        $script:OutstandingComputedSnapshotCache[$key] = $entry
        if ($script:OutstandingComputedSnapshotCache.Count -gt 12) {
            $oldest = @($script:OutstandingComputedSnapshotCache.Keys |
                Sort-Object {
                    $script:OutstandingComputedSnapshotCache[$_].createdUtc
                } | Select-Object -First 1)
            foreach ($oldKey in $oldest) {
                if ($oldKey -ne $key) {
                    $script:OutstandingComputedSnapshotCache.Remove($oldKey)
                }
            }
        }
        Write-Host ("[OUTSTANDING-SPLIT] COMPUTE mode={0} elapsedMs={1} rows={2}" -f
            $Mode, $watch.ElapsedMilliseconds, @($full.data).Count) -ForegroundColor DarkCyan
    }
    $snapshot = $entry.report
    $generatedAt = [datetime]$entry.createdUtc
    $snapshotId = [string]$entry.snapshotId
    $allRows = @($snapshot.data)

    if ($Mode -eq 'summary') {
        Write-Host ("[OUTSTANDING-SPLIT] SUMMARY cacheHit={0} rows={1}" -f
            $hit, $allRows.Count) -ForegroundColor DarkCyan
        return @{
            success = $true
            mode = 'summary'
            snapshotId = $snapshotId
            summaryScope = 'all'
            fromDate = $snapshot.fromDate
            toDate = $snapshot.toDate
            asOfDate = $snapshot.asOfDate
            generatedAt = $generatedAt.ToString('o')
            totalTransactions = $snapshot.totalTransactions
            databaseMode = $snapshot.databaseMode
            summary = $snapshot.summary
            debug = @{ snapshotHit = $hit; snapshotSeconds = $ttl }
        }
    }

    $descending = ($SortDirection -eq 'desc')
    switch ($SortBy.ToLowerInvariant()) {
        'account' {
            $orderedRows = @($allRows | Sort-Object accountName, dueDate, refNo -Descending:$descending)
        }
        'refdate' {
            $orderedRows = @($allRows | Sort-Object refDate, accountName, refNo -Descending:$descending)
        }
        'pending' {
            $orderedRows = @($allRows | Sort-Object pendingAmount, accountName, refNo -Descending:$descending)
        }
        'daysoverdue' {
            $orderedRows = @($allRows | Sort-Object daysOverdue, accountName, refNo -Descending:$descending)
        }
        default {
            $orderedRows = @($allRows | Sort-Object dueDate, accountName, refNo -Descending:$descending)
        }
    }
    $total = $orderedRows.Count
    if ($Page -lt 1) { $Page = 1 }
    if ($PageSize -lt 0) { $PageSize = 50 }
    if ($PageSize -gt 2500) { $PageSize = 2500 }
    if ($PageSize -eq 0) {
        $Page = 1
        $totalPages = 1
        $pageRows = $orderedRows
    }
    else {
        $totalPages = [Math]::Max(1, [int][Math]::Ceiling($total / [double]$PageSize))
        if ($Page -gt $totalPages) { $Page = $totalPages }
        $skip = ($Page - 1) * $PageSize
        $pageRows = @($orderedRows | Select-Object -Skip $skip -First $PageSize)
    }
    Write-Host ("[OUTSTANDING-SPLIT] TABLE cacheHit={0} page={1} size={2} returned={3} total={4}" -f
        $hit, $Page, $PageSize, @($pageRows).Count, $total) -ForegroundColor DarkCyan

    return @{
        success = $true
        mode = 'table'
        snapshotId = $snapshotId
        summaryScope = 'none'
        fromDate = $snapshot.fromDate
        toDate = $snapshot.toDate
        asOfDate = $snapshot.asOfDate
        generatedAt = $generatedAt.ToString('o')
        databaseMode = $snapshot.databaseMode
        total = $total
        totalTransactions = $total
        page = $Page
        pageSize = $PageSize
        totalPages = $totalPages
        data = @($pageRows)
        debug = @{ snapshotHit = $hit; snapshotSeconds = $ttl }
    }
}

# ============================================================
# STOCK STATUS DIRECT DATABASE FAST PATH
# [STOCK-ACCESS-COMPAT-V1] Access/BDS Stock Status uses BUSY COM fallback.
# ============================================================
# BUSY COM recordsets are reliable but expensive for company-wide
# aggregate/index work. auth_native.ps1/connection helpers are loaded
# before reports.ps1 and expose direct SQL Server / Access connections.
#
# Use direct DB connections only for READ-ONLY report queries. Voucher
# writes remain on the existing BUSY APIs.
# ============================================================

function Invoke-StockStatusDirectTable {
    param(
        [Parameter(Mandatory)]
        [System.Data.Common.DbConnection]$Connection,

        [Parameter(Mandatory)]
        [string]$Sql,

        [int]$Timeout = 180
    )

    # IMPORTANT:
    # Use a generic DbConnection reader instead of OleDbDataAdapter.
    # Both System.Data.SqlClient.SqlConnection and OleDbConnection derive
    # from DbConnection, so the same fast read path now works for:
    #   * SQL Server companies
    #   * Access/BDS companies
    #
    # This avoids routing large Stock Status GROUP BY queries through
    # BUSY's COM GetRecordset bridge.
    $command = $null
    $reader = $null
    $table = New-Object System.Data.DataTable

    try {
        $command = $Connection.CreateCommand()
        $command.CommandText = $Sql
        $command.CommandTimeout = $Timeout

        $reader = $command.ExecuteReader()
        $table.Load($reader)

        return ,$table
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
        [System.Data.Common.DbConnection]$Connection
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
                    "master-snapshot-v11-single-snapshot|" +
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

# Test-only, item-level method-5 hypothesis derived from native BUSY checkpoints.
# DO NOT apply this routine globally without cross-item regression testing.
function Invoke-StockStatusWaLabReplay {
    param(
        [Parameter(Mandatory=$true)] [AllowEmptyCollection()] [object[]]$Rows,
        [Parameter(Mandatory=$true)] [AllowEmptyCollection()] [object[]]$Openings,
        [Parameter(Mandatory=$true)] [string]$AsOf,
        [switch]$SaleReturnsBeforeOtherInwards,
        [switch]$UsePostedReturnWhenPoolEmpty,
        [switch]$UseRetainedRateForEmptyReturn,
        [switch]$UsePostedIssueWhenPoolEmpty,
        [switch]$PurchaseReceiptsBeforeReturns,
        [switch]$RoundPoolValueAfterEachMovement
    )

    $labEpsilon = 0.00000001
    [double]$totalQty = 0.0
    [double]$totalValue = 0.0
    $mcQty = @{}
    foreach ($opening in $Openings) {
        $key = [string][int]$opening.mcCode
        if (-not $mcQty.ContainsKey($key)) { $mcQty[$key] = 0.0 }
        $mcQty[$key] = [double]$mcQty[$key] + [double]$opening.qty
        $totalQty += [double]$opening.qty
        $totalValue += [double]$opening.value
    }
    [double]$lowestQty = $totalQty
    # [STOCK-RETAINED-RETURN-V3:REPLAY-STATE:BEGIN]
    # A zero quantity clears the stock VALUE, not the item's last known
    # positive weighted-average COST. This rate is consulted ONLY with the
    # explicit opt-in switch, and only for a sales return at exactly zero.
    [double]$retainedRate = 0.0
    $retainedRateAvailable = ($totalQty -gt $labEpsilon)
    if ($retainedRateAvailable) { $retainedRate = $totalValue / $totalQty }
    [int]$retainedEmptyReturnEvents = 0
    # [STOCK-RETAINED-RETURN-V3:REPLAY-STATE:END]
    $rowsByDate = @{}
    foreach ($line in $Rows) {
        $day = [string]$line.voucherDate
        if ($day -gt $AsOf) { continue }
        if (-not $rowsByDate.ContainsKey($day)) {
            $rowsByDate[$day] = [System.Collections.ArrayList]::new()
        }
        [void]$rowsByDate[$day].Add($line)
    }
    # Native checkpoints for this positive-stock item demonstrate a
    # DATE-LEVEL inward-first ordering (not raw voucher-code ordering).
    foreach ($day in @($rowsByDate.Keys | Sort-Object)) {
        $dayRows = @($rowsByDate[$day] | Sort-Object vchCode, srNo, vchType)
        # Experimental, all-items-only sequencing: determine the return cost
        # from the average BEFORE valuation-bearing other inward receipts.
        # Preserve original intra-group voucher order and inward-before-outward.
        # The original isolated item-1198 lab calls this without the switch.
        if ($SaleReturnsBeforeOtherInwards) {
            $returnRows = @($dayRows | Where-Object {
                [int]$_.vchType -eq 3 -and [double]$_.qty -gt $labEpsilon
            })
            if ($PurchaseReceiptsBeforeReturns) {
                # Test-only hypothesis: purchase receipts (type 2) precede
                # sale returns (type 3), which precede other inward rows.
                # Verified against 235 identities at the guarded checkpoint.
                # Distinct from the original item-1198 isolated lab.
                $purchaseRows = @($dayRows | Where-Object {
                    [int]$_.vchType -eq 2 -and [double]$_.qty -gt $labEpsilon
                })
                $otherRows = @($dayRows | Where-Object {
                    [double]$_.qty -le $labEpsilon -or (
                        [int]$_.vchType -ne 2 -and [int]$_.vchType -ne 3
                    )
                })
                $dayRows = @($purchaseRows) + @($returnRows) + @($otherRows)
            }
            else {
                $otherRows = @($dayRows | Where-Object {
                    [int]$_.vchType -ne 3 -or [double]$_.qty -le $labEpsilon
                })
                $dayRows = @($returnRows) + @($otherRows)
            }
        }
        # Phase 1: all inward quantities, including sale returns.
        foreach ($line in $dayRows) {
            [double]$movement = [double]$line.qty
            if ($movement -le $labEpsilon) { continue }
            $mcKey = [string][int]$line.mcCode
            if (-not $mcQty.ContainsKey($mcKey)) { $mcQty[$mcKey] = 0.0 }
            [double]$inValue = 0.0
            if ([int]$line.vchType -eq 3) {
                # The sale return joins were financial references, not
                # sufficient evidence of the actual inventory issue cost.
                # A current-average return matches native BUSY checkpoints.
                if ($totalQty -le $labEpsilon) {
                    if (-not $UsePostedReturnWhenPoolEmpty) {
                        return [pscustomobject]@{ supported=$false; reason='Sale return with nonpositive valuation pool'; quantity=$totalQty }
                    }
                    # Experimental, guarded first-stock cost: use THIS return's
                    # posted Value3 when no current average is available.
                    # The legacy item-1198 lab does not pass this switch.
                    # [STOCK-RETAINED-RETURN-V3:EMPTY-RETURN:BEGIN]
                    if (
                        $UseRetainedRateForEmptyReturn -and
                        [Math]::Abs($totalQty) -le $labEpsilon -and
                        $retainedRateAvailable
                    ) {
                        $inValue = $movement * $retainedRate
                        $retainedEmptyReturnEvents++
                    }
                    else {
                        # No prior positive weighted cost, or negative pool:
                        # keep the established posted-value fallback.
                        $inValue = [double]$line.value3
                    }
                    # [STOCK-RETAINED-RETURN-V3:EMPTY-RETURN:END]
                }
                else {
                    $inValue = $movement * ($totalValue / $totalQty)
                }
            }
            else {
                # Do not use D5 for positive-stock voucher value here;
                # for this fixture native BUSY follows the posted Value3
                # on purchases/other non-sale-return inward movements.
                $inValue = [double]$line.value3
            }
            $totalQty += $movement
            $totalValue += $inValue
            $mcQty[$mcKey] = [double]$mcQty[$mcKey] + $movement
            if ($RoundPoolValueAfterEachMovement) {
                $totalValue = [Math]::Round($totalValue, 2, [MidpointRounding]::ToEven)
            }
            if ($totalQty -gt $labEpsilon) {
                $retainedRate = $totalValue / $totalQty
                $retainedRateAvailable = $true
            }
        }
        # Phase 2: all outward quantities at the current shared average.
        foreach ($line in $dayRows) {
            [double]$movement = [double]$line.qty
            if ($movement -ge -$labEpsilon) { continue }
            if ($totalQty -le $labEpsilon) {
                if (-not $UsePostedIssueWhenPoolEmpty) {
                    return [pscustomobject]@{ supported=$false; reason='Nonpositive or exhausted item-level pool'; quantity=$totalQty }
                }
                # Negative-stock exception: preserve signed issue cost from
                # its posted Value3. Only candidate allowlist and snapshot
                # guards can apply this; not used by original item-1198 lab.
                $totalQty += $movement
                $totalValue += [double]$line.value3
                $mcKey = [string][int]$line.mcCode
                if (-not $mcQty.ContainsKey($mcKey)) { $mcQty[$mcKey] = 0.0 }
                $mcQty[$mcKey] = [double]$mcQty[$mcKey] + $movement
                if ($RoundPoolValueAfterEachMovement) {
                    $totalValue = [Math]::Round($totalValue, 2, [MidpointRounding]::ToEven)
                }
                continue
            }
            [double]$outQty = -$movement
            if (($totalQty - $outQty) -lt -$labEpsilon) {
                return [pscustomobject]@{ supported=$false; reason='Negative stock: not validated by this lab'; quantity=($totalQty - $outQty) }
            }
            [double]$currentRate = $totalValue / $totalQty
            $totalValue -= $outQty * $currentRate
            $totalQty -= $outQty
            if ([Math]::Abs($totalQty) -le $labEpsilon) {
                $totalQty = 0.0
                $totalValue = 0.0
            }
            $mcKey = [string][int]$line.mcCode
            if (-not $mcQty.ContainsKey($mcKey)) { $mcQty[$mcKey] = 0.0 }
            $mcQty[$mcKey] = [double]$mcQty[$mcKey] - $outQty
            if ($RoundPoolValueAfterEachMovement) {
                $totalValue = [Math]::Round($totalValue, 2, [MidpointRounding]::ToEven)
            }
            # Preserve the rate when an issue exhausts the pool; subsequent
            # sale returns can reuse the last known cost in observation mode.
            if ($totalQty -gt $labEpsilon) {
                $retainedRate = $totalValue / $totalQty
                $retainedRateAvailable = $true
            }
            elseif ([Math]::Abs($totalQty) -le $labEpsilon) {
                $retainedRate = $currentRate
                $retainedRateAvailable = $true
            }
        }
        if ($totalQty -lt $lowestQty) { $lowestQty = $totalQty }
    }
    [double]$rate = 0.0
    if ([Math]::Abs($totalQty) -gt $labEpsilon) { $rate = $totalValue / $totalQty }
    $centres = @{}
    [double]$centresQty = 0.0
    foreach ($code in @($mcQty.Keys)) {
        $centres[$code] = [pscustomobject]@{
            quantity = [double]$mcQty[$code]
            amount = [double]$mcQty[$code] * $rate
        }
        $centresQty += [double]$mcQty[$code]
    }
    if ([Math]::Abs($centresQty - $totalQty) -gt 0.000001) {
        return [pscustomobject]@{ supported=$false; reason='Material-centre quantity reconciliation failed'; quantity=$totalQty }
    }
    return [pscustomobject]@{
        supported = $true
        quantity = $totalQty
        amount = $totalValue
        rate = $rate
        lowestQty = $lowestQty
        centres = $centres
        retainedEmptyReturns = $retainedEmptyReturnEvents
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
        [bool]$FastSnapshot            = $false,
        [string]$InstanceId            = "",
        [string]$CompanyCode           = ""
    )

    # ============================================================
    # BUSY-LIKE LAZY VIEW STOCK STATUS ENGINE
    #
    # The frontend now requests only the ACTIVE report view.
    #
    #   balances      -> lazy request
    #   detailed      -> lazy request
    #   columnar      -> lazy request
    #   grouped       -> lazy request
    #   hierarchical  -> lazy request
    #
    # This avoids doing report work for views the user has not opened.
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
        # --------------------------------------------------------
        # DIRECT STOCK READ CONNECTION
        # --------------------------------------------------------
        # Stock Status previously used the direct path only for Access/BDS.
        # SQL Server companies still executed large aggregate queries through
        # BUSY COM GetRecordset, which was the main cold-load bottleneck.
        #
        # Use Get-DirectConnection for SQL Server and keep the existing
        # direct BDS path for Access. Voucher writes still use the existing
        # BUSY APIs; this connection is READ-ONLY report work.
        # --------------------------------------------------------
        $stockInstance =
            Get-InstanceConfig -InstanceId $InstanceId

        $stockDbType = 0
        try {
            if (
                $null -ne $stockInstance -and
                $null -ne $stockInstance.dbType
            ) {
                $stockDbType = [int]$stockInstance.dbType
            }
        }
        catch {}

        if ($stockDbType -eq 1) {
            try {
                if (
                    Get-Command Get-DirectConnection `
                        -ErrorAction SilentlyContinue
                ) {
                    $sqlDirectStockContext =
                        Get-DirectConnection `
                            -InstanceId $InstanceId `
                            -CompanyCode $CompanyCode

                    if (
                        $null -ne $sqlDirectStockContext -and
                        $null -ne $sqlDirectStockContext.connection
                    ) {
                        $directStockConn =
                            $sqlDirectStockContext.connection

                        if (
                            $directStockConn.State.ToString() -ne
                            "Open"
                        ) {
                            $directStockConn.Open()
                        }

                        Write-ReportEngineLog `
                            -Stage "STOCK-DIRECT" `
                            -Message (
                                "Direct SQL Server read path active; " +
                                "large stock aggregates bypass BUSY COM"
                            ) `
                            -Color "Green"
                    }
                }
            }
            catch {
                $directStockConn = $null

                Write-ReportEngineLog `
                    -Stage "STOCK-DIRECT" `
                    -Message (
                        "Direct SQL Server read path unavailable; " +
                        "falling back to BUSY COM. " +
                        $_.Exception.Message
                    ) `
                    -Color "Yellow"
            }
        }
        else {
            # STOCK-ACCESS-COMPAT-V1
            #
            # Access/BDS stock status must prefer the already-proven BUSY COM
            # recordset path. Some of the optimized Stock Status queries use
            # SQL constructs/provider behaviour that is reliable on SQL Server
            # but can fail or return incomplete data through Jet/OLE DB.
            #
            # IMPORTANT:
            # - SQL Server keeps the direct fast path.
            # - Access/BDS keeps exactly the existing report business logic,
            #   but all stock reads are executed through BUSY COM.
            # - No voucher data is written here.
            # - This is a compatibility/reliability switch, not a formula change.
            $directStockContext = $null
            $directStockConn = $null

            Write-ReportEngineLog `
                -Stage "STOCK-ACCESS" `
                -Message (
                    "Access/BDS detected; using BUSY COM stock-report path " +
                    "for compatibility."
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

        # [GREGORY-NATIVE-ITEM-ELIGIBILITY-V1:BEGIN]
        # The Gregory FY2026 native BUSY Stock Status excludes item masters
        # whose B21=1. The old global B21 IS NOT NULL rule included 13 such
        # masters, including three with nonzero closing quantities.
        # Guarded by exact instance, company, SQL fiscal DB and FY.
        # This fixes the proven item eligibility problem ONLY; the Gregory
        # valuation replay is still experimental and is NOT enabled here.
        $gregoryNativeStockEligibility = $false
        if (
            [string]$InstanceId -ceq 'instance_2' -and
            [string]$CompanyCode -ceq 'Comp0004' -and
            $stockDbType -eq 1 -and
            $null -ne $directStockConn -and
            $fromDate.Date -eq [datetime]'2026-04-01' -and
            $endDate.Date -ge [datetime]'2026-04-01' -and
            $endDate.Date -le [datetime]'2027-03-31'
        ) {
            try {
                $gregoryNativeStockEligibility = (
                    [string]$directStockConn.Database -ceq 'BusyComp0004_db12026'
                )
            }
            catch { $gregoryNativeStockEligibility = $false }
        }
        $gregorySnapshotEligibility = if ($gregoryNativeStockEligibility) {
            'gregory-b21-zero-v1'
        }
        else {
            'legacy-b21-not-null'
        }
        # [GREGORY-NATIVE-ITEM-ELIGIBILITY-V1:END]

        # [STOCK-GS-LOCAL-DEFAULT-V1:BEGIN]
        # GS-only local deployment default. All original company, SQL, FY,
        # report-view, transfer, quantity and candidate preflight guards remain.
        # IMPORTANT: this affects ALL callers of this bridge for this company,
        # not only browser tabs running at localhost.
        # Emergency opt-out: set BUSY_STOCK_WA_GENERAL_APPLY=off before restart.
        $stockGeneralApplyRequestedMode = ([string]$env:BUSY_STOCK_WA_GENERAL_APPLY).Trim()
        if (
            [string]::IsNullOrWhiteSpace($stockGeneralApplyRequestedMode) -and
            [string]$InstanceId -ceq 'instance_2' -and
            [string]$CompanyCode -ceq 'Comp0005'
        ) {
            $stockGeneralApplyRequestedMode = 'apply-tested'
        }
        # [STOCK-GS-LOCAL-DEFAULT-V1:END]

        # [STOCK-GENERAL-APPLY-V4:FULLSCAN:BEGIN]
        # Never let a fast-page request apply only 100 independent values
        # while a later export computes 235 values. Reuse the complete
        # calculation cache, then slice using the existing paging helper.
        # This is deliberately limited to the exact test identity/scope.
        $stockGeneralApplyFullScanRequested = $false
        if (
            $stockGeneralApplyRequestedMode -ceq 'apply-tested' -and
            [string]$InstanceId -ceq 'instance_2' -and
            [string]$CompanyCode -ceq 'Comp0005' -and
            $stockDbType -eq 1 -and
            $null -ne $directStockConn -and
            $fromDate.Date -eq [datetime]'2026-04-01' -and
            $endDate.Date -ge [datetime]'2026-04-01' -and
            $endDate.Date -le [datetime]'2027-03-31' -and
            $asOfDate.Date -eq $endDate.Date -and
            $IncludeStockTransfers -and
            $valueMode -eq 'busy' -and
            $viewMode -eq 'balances'
        ) {
            try {
                $stockGeneralApplyFullScanRequested = (
                    [string]$directStockConn.Database -ceq 'BusyComp0005_db12026'
                )
            }
            catch { $stockGeneralApplyFullScanRequested = $false }
        }
        # [STOCK-GENERAL-APPLY-V4:FULLSCAN:END]
        $cacheSeparator = [char]31
        $stockStatusCacheKey = (
            @(
                "stock-wa-v25-gs-local-default-gregory-b21-v1",
                [string]$InstanceId,
                [string]$CompanyCode,
                [string]$gregorySnapshotEligibility,
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
                [string]$env:BUSY_STOCK_WA_LAB_1198,
                [string]$env:BUSY_STOCK_WA_LAB_ALL_ITEMS,
                [string]$env:BUSY_STOCK_WA_LAB_SEP30,
                [string]$env:BUSY_STOCK_WA_LAB_AUG31,
                [string]$env:BUSY_STOCK_WA_GENERAL_OBSERVE,
                [string]$stockGeneralApplyRequestedMode
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
            $viewMode -in @("balances", "detailed", "columnar") -and
            $PageSize -gt 0 -and
            $PageSize -le 250 -and
            -not $stockGeneralApplyFullScanRequested
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
                            "stock-master-snapshot-v11-single-snapshot" -and
                        [string]$candidateSnapshot.signature -eq
                            $stockMasterSignature -and
                        $null -ne $candidateSnapshot.groups -and
                        $null -ne $candidateSnapshot.materialCentres -and
                        $null -ne $candidateSnapshot.units -and
                        $null -ne $candidateSnapshot.items -and
                        (
                            (
                                $gregoryNativeStockEligibility -and
                                [string]$candidateSnapshot.itemEligibility -ceq
                                    'gregory-b21-zero-v1'
                            ) -or (
                                -not $gregoryNativeStockEligibility -and
                                [string]$candidateSnapshot.itemEligibility -cne
                                    'gregory-b21-zero-v1'
                            )
                        )
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
        # whose Master1.B21 is NULL; legacy reports exclude those items.
        # For FY2026 Gregory, native exports additionally exclude B21=1:
        # the exact guarded B21=0 override below applies only to that DB.
        # Neither condition is generalized to unverified companies.
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

        # [GREGORY-NATIVE-ITEM-ELIGIBILITY-V1:QUERY:BEGIN]
        # Scope the native B21=0 rule to Gregory ONLY; the known-good GS
        # source and all other company selection rules remain identical.
        if ($gregoryNativeStockEligibility) {
            $itemQry = $itemQry.Replace(
                'AND M.B21 IS NOT NULL',
                'AND M.B21 = 0'
            )
            Write-ReportEngineLog -Stage 'GREGORY-ITEM-ELIGIBILITY' -Message (
                'Native Gregory FY2026 item filter active: MasterType=6; B21=0. ' +
                'Experimental Gregory valuation NOT applied.'
            ) -Color 'Cyan'
        }
        # [GREGORY-NATIVE-ITEM-ELIGIBILITY-V1:QUERY:END]

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
                # Hierarchical group filtering:
                # selecting a parent/subgroup must include all descendants.
                $groupNeedle =
                    $ItemGroup.
                        Trim().
                        Replace([char]0x203A, ">").
                        ToLowerInvariant()

                $needleSegments = @(
                    $groupNeedle.
                        Split(
                            [string[]]@(">"),
                            [System.StringSplitOptions]::RemoveEmptyEntries
                        ) |
                    ForEach-Object {
                        ([string]$_).Trim()
                    }
                )

                $normalizedGroupNeedle =
                    ($needleSegments -join " > ")

                $normalizedPathSegments = @(
                    $groupPath |
                    ForEach-Object {
                        ([string]$_).
                            Trim().
                            ToLowerInvariant()
                    }
                )

                $pathLabel =
                    ($normalizedPathSegments -join " > ")

                $matchesGroup = (
                    $pathLabel -eq $normalizedGroupNeedle -or
                    $pathLabel.StartsWith(
                        $normalizedGroupNeedle + " > "
                    )
                )

                if (
                    -not $matchesGroup -and
                    $needleSegments.Count -eq 1
                ) {
                    foreach ($segment in $normalizedPathSegments) {
                        if ($segment -eq $normalizedGroupNeedle) {
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
                    engine = "stock-master-snapshot-v11-single-snapshot"
                    signature = $stockMasterSignature
                    itemEligibility = [string]$gregorySnapshotEligibility
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


        # ========================================================
        # V14 FAST SQL SNAPSHOT
        # ========================================================
        #
        # Architecture:
        #
        #      AS ON DATE
        #          |
        #          v
        #   ONE FAST SNAPSHOT REQUEST
        #          |
        #          +--> Direct SQL aggregate (Tran4 + Tran2)
        #          +--> Master data already loaded above
        #          |
        #          v
        #   Complete lightweight frontend snapshot
        #
        # IMPORTANT:
        # This path intentionally DOES NOT run STOCK-WA / chronological
        # method-5 replay across all stock masters. That exact replay is ideal
        # for a small requested page, but it is the wrong algorithm for a full
        # 3,000+ item snapshot.
        #
        # For the snapshot's "BUSY Posted Value" we use BUSY's posted opening
        # value (Tran4.D3) plus posted movement values (Tran2.Value3).
        # Purchase / Sale / MRP alternatives are item-master prices and are
        # recalculated instantly in the frontend.
        # ========================================================
        if ($FastSnapshot -and $null -ne $directStockConn) {
            $snapshotWatch = New-ReportEngineWatch

            Write-ReportEngineLog `
                -Stage "STOCK-SNAPSHOT-FAST" `
                -Message (
                    "Starting set-based snapshot; asOf=" +
                    $asOfDate.ToString("yyyy-MM-dd") +
                    "; masters=" +
                    $itemOrder.Count
                ) `
                -Color "Cyan"

            $snapshotStatesByItem = @{}

            function Get-StockSnapshotState {
                param(
                    [int]$ItemCode,
                    [int]$McCode
                )

                $itemKey = [string]$ItemCode
                $mcKey = [string]$McCode

                if (-not $snapshotStatesByItem.ContainsKey($itemKey)) {
                    $snapshotStatesByItem[$itemKey] = @{}
                }

                $states = $snapshotStatesByItem[$itemKey]

                if (-not $states.ContainsKey($mcKey)) {
                    $states[$mcKey] = [pscustomobject]@{
                        openingMain = 0.0
                        openingAlt = 0.0
                        openingValue = 0.0

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
                        transferOutQty = 0.0
                        transferInAltQty = 0.0
                        transferOutAltQty = 0.0

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

                return $states[$mcKey]
            }

            # Material Centre is a backend filter in V15. Item Group and
            # Search have already reduced $itemOrder/$items above. For MC we
            # additionally restrict Tran4/Tran2 at SQL level so SQL Server does
            # not aggregate centres that the user did not request.
            $snapshotMcSqlClause = ""

            if (-not $useAllMc -and $selectedMcCodes.Count -gt 0) {
                $snapshotMcCodes = @(
                    $selectedMcCodes.Keys |
                    ForEach-Object { [int]$_ } |
                    Sort-Object
                )

                if ($snapshotMcCodes.Count -gt 0) {
                    $snapshotMcSqlClause =
                        " AND MasterCode2 IN (" +
                        (($snapshotMcCodes | ForEach-Object { [string]$_ }) -join ",") +
                        ")"
                }
            }

            # FY opening stock/value is already stored compactly in Tran4.
            $snapshotOpeningSql = @"
SELECT
    MasterCode1 AS ItemCode,
    MasterCode2 AS MCCode,
    SUM(D1) AS OpeningQty,
    SUM(D2) AS OpeningAltQty,
    SUM(D3) AS OpeningValue
FROM Tran4
WHERE RecType = 0$snapshotMcSqlClause
GROUP BY MasterCode1, MasterCode2
"@

            $snapshotOpeningTable =
                Invoke-StockStatusDirectTable `
                    -Connection $directStockConn `
                    -Timeout 180 `
                    -Sql $snapshotOpeningSql

            foreach ($row in $snapshotOpeningTable.Rows) {
                $itemCode =
                    ConvertTo-ReportInt $row["ItemCode"]

                if (-not $items.ContainsKey($itemCode)) {
                    continue
                }

                $mcCode =
                    ConvertTo-ReportInt $row["MCCode"]

                $state =
                    Get-StockSnapshotState `
                        -ItemCode $itemCode `
                        -McCode $mcCode

                $state.openingMain +=
                    ConvertTo-ReportDouble $row["OpeningQty"]

                $state.openingAlt +=
                    ConvertTo-ReportDouble $row["OpeningAltQty"]

                $state.openingValue +=
                    ConvertTo-ReportDouble $row["OpeningValue"]
            }

            $snapshotMovementMcSqlClause = ""
            if (-not $useAllMc -and $selectedMcCodes.Count -gt 0) {
                $snapshotMcCodes = @(
                    $selectedMcCodes.Keys |
                    ForEach-Object { [int]$_ } |
                    Sort-Object
                )
                if ($snapshotMcCodes.Count -gt 0) {
                    $snapshotMovementMcSqlClause =
                        " AND T.MasterCode2 IN (" +
                        (($snapshotMcCodes | ForEach-Object { [string]$_ }) -join ",") +
                        ")"
                }
            }

            # Current-period movement is one set-based SQL GROUP BY. It gives
            # all five frontend views their quantity/movement breakdown without
            # replaying each transaction in PowerShell.
            if ($mode.isSql) {
                $snapshotMovementSql = @"
SELECT
    T.MasterCode1 AS ItemCode,
    T.MasterCode2 AS MCCode,

    SUM(CASE WHEN T.Value1 > 0 THEN T.Value1 ELSE 0 END) AS InwardQty,
    SUM(CASE WHEN T.Value1 < 0 THEN -T.Value1 ELSE 0 END) AS OutwardQty,

    SUM(CASE WHEN T.Value2 > 0 THEN T.Value2 ELSE 0 END) AS InwardAltQty,
    SUM(CASE WHEN T.Value2 < 0 THEN -T.Value2 ELSE 0 END) AS OutwardAltQty,

    SUM(CASE WHEN T.Value1 > 0 THEN ABS(T.Value3) ELSE 0 END) AS InwardValue,
    SUM(CASE WHEN T.Value1 < 0 THEN ABS(T.Value3) ELSE 0 END) AS OutwardValue,

    SUM(CASE WHEN T.Value1 > 0 AND V.VchType = 2 THEN T.Value1 ELSE 0 END) AS PurchaseQty,
    SUM(CASE WHEN T.Value1 < 0 AND V.VchType = 9 THEN -T.Value1 ELSE 0 END) AS SaleQty,
    SUM(CASE WHEN T.Value1 > 0 AND V.VchType = 3 THEN T.Value1 ELSE 0 END) AS SaleReturnQty,
    SUM(CASE WHEN T.Value1 < 0 AND V.VchType = 10 THEN -T.Value1 ELSE 0 END) AS PurchaseReturnQty,

    SUM(CASE WHEN T.Value1 > 0 AND V.VchType = 5 THEN T.Value1 ELSE 0 END) AS TransferInQty,
    SUM(CASE WHEN T.Value1 < 0 AND V.VchType = 5 THEN -T.Value1 ELSE 0 END) AS TransferOutQty,
    SUM(CASE WHEN T.Value2 > 0 AND V.VchType = 5 THEN T.Value2 ELSE 0 END) AS TransferInAltQty,
    SUM(CASE WHEN T.Value2 < 0 AND V.VchType = 5 THEN -T.Value2 ELSE 0 END) AS TransferOutAltQty,

    SUM(CASE WHEN T.Value1 > 0 AND V.VchType = 6 THEN T.Value1 ELSE 0 END) AS ProductionGeneratedQty,
    SUM(CASE WHEN T.Value1 < 0 AND V.VchType = 6 THEN -T.Value1 ELSE 0 END) AS ProductionConsumedQty,

    SUM(CASE WHEN T.Value1 > 0 AND V.VchType = 8 THEN T.Value1 ELSE 0 END) AS StockJournalGeneratedQty,
    SUM(CASE WHEN T.Value1 < 0 AND V.VchType = 8 THEN -T.Value1 ELSE 0 END) AS StockJournalConsumedQty,

    SUM(CASE WHEN T.Value1 > 0 AND V.VchType = 4 THEN T.Value1 ELSE 0 END) AS MaterialReceiptQty,
    SUM(CASE WHEN T.Value1 < 0 AND V.VchType = 11 THEN -T.Value1 ELSE 0 END) AS MaterialIssueQty,

    SUM(
        CASE
            WHEN T.Value1 > 0
             AND V.VchType NOT IN (2, 3, 4, 5, 6, 8)
            THEN T.Value1
            ELSE 0
        END
    ) AS OtherInQty,

    SUM(
        CASE
            WHEN T.Value1 < 0
             AND V.VchType NOT IN (5, 6, 8, 9, 10, 11)
            THEN -T.Value1
            ELSE 0
        END
    ) AS OtherOutQty

FROM Tran2 T
INNER JOIN Tran1 V
    ON V.VchCode = T.VchCode
WHERE
    T.RecType = 2
    AND V.[Date] >= $fromLiteral
    AND V.[Date] <= $endLiteral
    AND (V.Cancelled = 0 OR V.Cancelled IS NULL)
    AND (V.VchCancelled = 0 OR V.VchCancelled IS NULL)$snapshotMovementMcSqlClause
GROUP BY
    T.MasterCode1,
    T.MasterCode2
"@
            }
            else {
                # Access/BDS equivalent. The project is currently SQL Server
                # in the user's main deployment, but keep a compatible direct
                # aggregate instead of falling through to full STOCK-WA.
                $snapshotMovementSql = @"
SELECT
    T.MasterCode1 AS ItemCode,
    T.MasterCode2 AS MCCode,

    SUM(IIF(T.Value1 > 0, T.Value1, 0)) AS InwardQty,
    SUM(IIF(T.Value1 < 0, -T.Value1, 0)) AS OutwardQty,

    SUM(IIF(T.Value2 > 0, T.Value2, 0)) AS InwardAltQty,
    SUM(IIF(T.Value2 < 0, -T.Value2, 0)) AS OutwardAltQty,

    SUM(IIF(T.Value1 > 0, ABS(T.Value3), 0)) AS InwardValue,
    SUM(IIF(T.Value1 < 0, ABS(T.Value3), 0)) AS OutwardValue,

    SUM(IIF(T.Value1 > 0 AND V.VchType = 2, T.Value1, 0)) AS PurchaseQty,
    SUM(IIF(T.Value1 < 0 AND V.VchType = 9, -T.Value1, 0)) AS SaleQty,
    SUM(IIF(T.Value1 > 0 AND V.VchType = 3, T.Value1, 0)) AS SaleReturnQty,
    SUM(IIF(T.Value1 < 0 AND V.VchType = 10, -T.Value1, 0)) AS PurchaseReturnQty,

    SUM(IIF(T.Value1 > 0 AND V.VchType = 5, T.Value1, 0)) AS TransferInQty,
    SUM(IIF(T.Value1 < 0 AND V.VchType = 5, -T.Value1, 0)) AS TransferOutQty,
    SUM(IIF(T.Value2 > 0 AND V.VchType = 5, T.Value2, 0)) AS TransferInAltQty,
    SUM(IIF(T.Value2 < 0 AND V.VchType = 5, -T.Value2, 0)) AS TransferOutAltQty,

    SUM(IIF(T.Value1 > 0 AND V.VchType = 6, T.Value1, 0)) AS ProductionGeneratedQty,
    SUM(IIF(T.Value1 < 0 AND V.VchType = 6, -T.Value1, 0)) AS ProductionConsumedQty,

    SUM(IIF(T.Value1 > 0 AND V.VchType = 8, T.Value1, 0)) AS StockJournalGeneratedQty,
    SUM(IIF(T.Value1 < 0 AND V.VchType = 8, -T.Value1, 0)) AS StockJournalConsumedQty,

    SUM(IIF(T.Value1 > 0 AND V.VchType = 4, T.Value1, 0)) AS MaterialReceiptQty,
    SUM(IIF(T.Value1 < 0 AND V.VchType = 11, -T.Value1, 0)) AS MaterialIssueQty,

    SUM(IIF(T.Value1 > 0 AND V.VchType NOT IN (2,3,4,5,6,8), T.Value1, 0)) AS OtherInQty,
    SUM(IIF(T.Value1 < 0 AND V.VchType NOT IN (5,6,8,9,10,11), -T.Value1, 0)) AS OtherOutQty

FROM Tran2 T
INNER JOIN Tran1 V
    ON V.VchCode = T.VchCode
WHERE
    T.RecType = 2
    AND V.[Date] >= $fromLiteral
    AND V.[Date] <= $endLiteral
    AND (V.Cancelled = 0 OR V.Cancelled IS NULL)
    AND (V.VchCancelled = 0 OR V.VchCancelled IS NULL)$snapshotMovementMcSqlClause
GROUP BY
    T.MasterCode1,
    T.MasterCode2
"@
            }

            $snapshotMovementTable =
                Invoke-StockStatusDirectTable `
                    -Connection $directStockConn `
                    -Timeout 240 `
                    -Sql $snapshotMovementSql

            foreach ($row in $snapshotMovementTable.Rows) {
                $itemCode =
                    ConvertTo-ReportInt $row["ItemCode"]

                if (-not $items.ContainsKey($itemCode)) {
                    continue
                }

                $mcCode =
                    ConvertTo-ReportInt $row["MCCode"]

                $state =
                    Get-StockSnapshotState `
                        -ItemCode $itemCode `
                        -McCode $mcCode

                $state.inwardMain += ConvertTo-ReportDouble $row["InwardQty"]
                $state.outwardMain += ConvertTo-ReportDouble $row["OutwardQty"]
                $state.inwardAlt += ConvertTo-ReportDouble $row["InwardAltQty"]
                $state.outwardAlt += ConvertTo-ReportDouble $row["OutwardAltQty"]
                $state.inwardValue += ConvertTo-ReportDouble $row["InwardValue"]
                $state.outwardValue += ConvertTo-ReportDouble $row["OutwardValue"]

                $state.purchaseQty += ConvertTo-ReportDouble $row["PurchaseQty"]
                $state.saleQty += ConvertTo-ReportDouble $row["SaleQty"]
                $state.saleReturnQty += ConvertTo-ReportDouble $row["SaleReturnQty"]
                $state.purchaseReturnQty += ConvertTo-ReportDouble $row["PurchaseReturnQty"]

                $state.transferInQty += ConvertTo-ReportDouble $row["TransferInQty"]
                $state.transferOutQty += ConvertTo-ReportDouble $row["TransferOutQty"]
                $state.transferInAltQty += ConvertTo-ReportDouble $row["TransferInAltQty"]
                $state.transferOutAltQty += ConvertTo-ReportDouble $row["TransferOutAltQty"]

                $state.productionGeneratedQty += ConvertTo-ReportDouble $row["ProductionGeneratedQty"]
                $state.productionConsumedQty += ConvertTo-ReportDouble $row["ProductionConsumedQty"]
                $state.stockJournalGeneratedQty += ConvertTo-ReportDouble $row["StockJournalGeneratedQty"]
                $state.stockJournalConsumedQty += ConvertTo-ReportDouble $row["StockJournalConsumedQty"]
                $state.materialReceiptQty += ConvertTo-ReportDouble $row["MaterialReceiptQty"]
                $state.materialIssueQty += ConvertTo-ReportDouble $row["MaterialIssueQty"]
                $state.otherInQty += ConvertTo-ReportDouble $row["OtherInQty"]
                $state.otherOutQty += ConvertTo-ReportDouble $row["OtherOutQty"]
            }

            $snapshotRows =
                [System.Collections.Generic.List[object]]::new()

            foreach ($itemCode in @($itemOrder)) {
                $item = $items[[int]$itemCode]
                $itemStates = @{}

                if (
                    $snapshotStatesByItem.ContainsKey(
                        [string]$itemCode
                    )
                ) {
                    $itemStates =
                        $snapshotStatesByItem[
                            [string]$itemCode
                        ]
                }

                # When a Material Centre backend filter is active, an item
                # must have opening or movement state in that selected centre.
                # Search/group can legitimately include masters with zero stock,
                # so this extra test is MC-specific only.
                if (-not $useAllMc -and $itemStates.Count -eq 0) {
                    continue
                }

                $centres =
                    [System.Collections.Generic.List[object]]::new()

                $openingMain = 0.0
                $openingAlt = 0.0
                $openingValue = 0.0
                $inwardMain = 0.0
                $inwardAlt = 0.0
                $inwardValue = 0.0
                $outwardMain = 0.0
                $outwardAlt = 0.0
                $outwardValue = 0.0

                $purchaseQty = 0.0
                $saleQty = 0.0
                $saleReturnQty = 0.0
                $purchaseReturnQty = 0.0
                $transferInQty = 0.0
                $transferOutQty = 0.0
                $transferInAltQty = 0.0
                $transferOutAltQty = 0.0
                $productionGeneratedQty = 0.0
                $productionConsumedQty = 0.0
                $stockJournalGeneratedQty = 0.0
                $stockJournalConsumedQty = 0.0
                $materialReceiptQty = 0.0
                $materialIssueQty = 0.0
                $otherInQty = 0.0
                $otherOutQty = 0.0

                foreach (
                    $mcKey in
                    @($itemStates.Keys | Sort-Object)
                ) {
                    $mcCode = 0

                    if (
                        -not [int]::TryParse(
                            [string]$mcKey,
                            [ref]$mcCode
                        )
                    ) {
                        continue
                    }

                    $state = $itemStates[$mcKey]

                    $mcOpeningMain = [double]$state.openingMain
                    $mcOpeningAlt = [double]$state.openingAlt
                    $mcOpeningValue = [double]$state.openingValue

                    $mcInwardMain = [double]$state.inwardMain
                    $mcInwardAlt = [double]$state.inwardAlt
                    $mcInwardValue = [double]$state.inwardValue

                    $mcOutwardMain = [double]$state.outwardMain
                    $mcOutwardAlt = [double]$state.outwardAlt
                    $mcOutwardValue = [double]$state.outwardValue

                    $mcClosingMain =
                        $mcOpeningMain +
                        $mcInwardMain -
                        $mcOutwardMain

                    $mcClosingAlt =
                        $mcOpeningAlt +
                        $mcInwardAlt -
                        $mcOutwardAlt

                    $mcPostedValue =
                        $mcOpeningValue +
                        $mcInwardValue -
                        $mcOutwardValue

                    $mcRate =
                        if (
                            [Math]::Abs($mcClosingMain) -gt $eps
                        ) {
                            $mcPostedValue / $mcClosingMain
                        }
                        else {
                            0.0
                        }

                    $mcName =
                        if ($mcMap.ContainsKey($mcCode)) {
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

                        stockValue = [Math]::Round($mcPostedValue, 2)
                        postedLedgerValue = [Math]::Round($mcPostedValue, 2)
                        valuationRate = [Math]::Round($mcRate, 6)

                        purchaseQuantity = [Math]::Round([double]$state.purchaseQty, 3)
                        saleQuantity = [Math]::Round([double]$state.saleQty, 3)
                        saleReturnQuantity = [Math]::Round([double]$state.saleReturnQty, 3)
                        purchaseReturnQuantity = [Math]::Round([double]$state.purchaseReturnQty, 3)

                        transferInQuantity = [Math]::Round([double]$state.transferInQty, 3)
                        transferOutQuantity = [Math]::Round([double]$state.transferOutQty, 3)
                        transferInAltQuantity = [Math]::Round([double]$state.transferInAltQty, 3)
                        transferOutAltQuantity = [Math]::Round([double]$state.transferOutAltQty, 3)

                        productionGeneratedQuantity = [Math]::Round([double]$state.productionGeneratedQty, 3)
                        productionConsumedQuantity = [Math]::Round([double]$state.productionConsumedQty, 3)
                        stockJournalGeneratedQuantity = [Math]::Round([double]$state.stockJournalGeneratedQty, 3)
                        stockJournalConsumedQuantity = [Math]::Round([double]$state.stockJournalConsumedQty, 3)
                        materialReceiptQuantity = [Math]::Round([double]$state.materialReceiptQty, 3)
                        materialIssueQuantity = [Math]::Round([double]$state.materialIssueQty, 3)
                        otherInQuantity = [Math]::Round([double]$state.otherInQty, 3)
                        otherOutQuantity = [Math]::Round([double]$state.otherOutQty, 3)
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
                    $transferInAltQty += [double]$state.transferInAltQty
                    $transferOutAltQty += [double]$state.transferOutAltQty
                    $productionGeneratedQty += [double]$state.productionGeneratedQty
                    $productionConsumedQty += [double]$state.productionConsumedQty
                    $stockJournalGeneratedQty += [double]$state.stockJournalGeneratedQty
                    $stockJournalConsumedQty += [double]$state.stockJournalConsumedQty
                    $materialReceiptQty += [double]$state.materialReceiptQty
                    $materialIssueQty += [double]$state.materialIssueQty
                    $otherInQty += [double]$state.otherInQty
                    $otherOutQty += [double]$state.otherOutQty
                }

                $closingMain =
                    $openingMain +
                    $inwardMain -
                    $outwardMain

                $closingAlt =
                    $openingAlt +
                    $inwardAlt -
                    $outwardAlt

                $postedValue =
                    $openingValue +
                    $inwardValue -
                    $outwardValue

                $postedRate =
                    if (
                        [Math]::Abs($closingMain) -gt $eps
                    ) {
                        $postedValue / $closingMain
                    }
                    else {
                        0.0
                    }

                $stockStatus =
                    if ($closingMain -lt -$eps) {
                        "negative"
                    }
                    elseif (
                        [Math]::Abs($closingMain) -le $eps
                    ) {
                        "out-of-stock"
                    }
                    elseif ($closingMain -le 5) {
                        "low-stock"
                    }
                    else {
                        "in-stock"
                    }

                $parentGroup =
                    if ($item.groupPath.Count -gt 0) {
                        [string]$item.groupPath[0]
                    }
                    else {
                        "General"
                    }

                $snapshotRows.Add([pscustomobject]@{
                    itemCode = [int]$item.code
                    itemName = [string]$item.name
                    alias = [string]$item.alias
                    group = [string]$item.group
                    parentGroup = $parentGroup
                    groupPath = @($item.groupPath)

                    unit =
                        if (
                            [string]::IsNullOrWhiteSpace(
                                [string]$item.mainUnit
                            )
                        ) {
                            "Units"
                        }
                        else {
                            [string]$item.mainUnit
                        }

                    mainUnit =
                        if (
                            [string]::IsNullOrWhiteSpace(
                                [string]$item.mainUnit
                            )
                        ) {
                            "Units"
                        }
                        else {
                            [string]$item.mainUnit
                        }

                    altUnit = [string]$item.altUnit
                    conversionFactor = [Math]::Round(
                        [double]$item.conversionFactor,
                        6
                    )
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
                    movementQuantity = [Math]::Round(
                        $inwardMain + $outwardMain,
                        3
                    )

                    lowStockLevel = 5
                    status = $stockStatus

                    purchaseQuantity = [Math]::Round($purchaseQty, 3)
                    saleQuantity = [Math]::Round($saleQty, 3)
                    saleReturnQuantity = [Math]::Round($saleReturnQty, 3)
                    purchaseReturnQuantity = [Math]::Round($purchaseReturnQty, 3)

                    transferInQuantity = [Math]::Round($transferInQty, 3)
                    transferOutQuantity = [Math]::Round($transferOutQty, 3)
                    transferInAltQuantity = [Math]::Round($transferInAltQty, 3)
                    transferOutAltQuantity = [Math]::Round($transferOutAltQty, 3)

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

                    # Snapshot BUSY value = posted ledger value.
                    busyValuationRate = [Math]::Round($postedRate, 6)
                    busyStockValue = [Math]::Round($postedValue, 2)
                    postedLedgerClosingValue = [Math]::Round($postedValue, 2)

                    valuationRate = [Math]::Round($postedRate, 6)
                    stockValue = [Math]::Round($postedValue, 2)

                    openingValue = [Math]::Round($openingValue, 2)
                    inwardValue = [Math]::Round($inwardValue, 2)
                    outwardValue = [Math]::Round($outwardValue, 2)

                    materialCentres =
                        @($centres | Sort-Object name)
                })
            }

            $sortedSnapshotRows =
                @(
                    $snapshotRows |
                    Sort-Object itemName, itemCode
                )

            $snapshotGroupMap = @{}

            foreach ($row in $sortedSnapshotRows) {
                $groupKey =
                    if ($row.groupPath.Count -gt 0) {
                        @($row.groupPath) -join " > "
                    }
                    else {
                        "General"
                    }

                if (-not $snapshotGroupMap.ContainsKey($groupKey)) {
                    $snapshotGroupMap[$groupKey] =
                        [pscustomobject]@{
                            group = [string]$row.group
                            parentGroup =
                                [string]$row.parentGroup
                            groupPath = @($row.groupPath)
                            itemCount = 0
                            openingQuantity = 0.0
                            inwardQuantity = 0.0
                            outwardQuantity = 0.0
                            closingQuantity = 0.0
                            stockValue = 0.0
                        }
                }

                $groupTotal = $snapshotGroupMap[$groupKey]
                $groupTotal.itemCount++
                $groupTotal.openingQuantity += [double]$row.openingQuantity
                $groupTotal.inwardQuantity += [double]$row.inwardQuantity
                $groupTotal.outwardQuantity += [double]$row.outwardQuantity
                $groupTotal.closingQuantity += [double]$row.closingQuantity
                $groupTotal.stockValue += [double]$row.stockValue
            }

            $snapshotGroupTotals = @(
                foreach (
                    $groupKey in
                    @($snapshotGroupMap.Keys | Sort-Object)
                ) {
                    $groupTotal =
                        $snapshotGroupMap[$groupKey]

                    [pscustomobject]@{
                        group = [string]$groupTotal.group
                        parentGroup =
                            [string]$groupTotal.parentGroup
                        groupPath =
                            @($groupTotal.groupPath)
                        itemCount =
                            [int]$groupTotal.itemCount
                        openingQuantity =
                            [Math]::Round(
                                [double]$groupTotal.openingQuantity,
                                3
                            )
                        inwardQuantity =
                            [Math]::Round(
                                [double]$groupTotal.inwardQuantity,
                                3
                            )
                        outwardQuantity =
                            [Math]::Round(
                                [double]$groupTotal.outwardQuantity,
                                3
                            )
                        closingQuantity =
                            [Math]::Round(
                                [double]$groupTotal.closingQuantity,
                                3
                            )
                        stockValue =
                            [Math]::Round(
                                [double]$groupTotal.stockValue,
                                2
                            )
                    }
                }
            )

            $snapshotSummary = @{
                totalItems = $sortedSnapshotRows.Count

                inStockItems =
                    @(
                        $sortedSnapshotRows |
                        Where-Object status -eq "in-stock"
                    ).Count

                lowStockItems =
                    @(
                        $sortedSnapshotRows |
                        Where-Object status -eq "low-stock"
                    ).Count

                outOfStockItems =
                    @(
                        $sortedSnapshotRows |
                        Where-Object status -eq "out-of-stock"
                    ).Count

                negativeStockItems =
                    @(
                        $sortedSnapshotRows |
                        Where-Object status -eq "negative"
                    ).Count

                totalQuantity = [Math]::Round(
                    [double]((
                        $sortedSnapshotRows |
                        Measure-Object -Property closingQuantity -Sum
                    ).Sum),
                    3
                )

                openingQuantity = [Math]::Round(
                    [double]((
                        $sortedSnapshotRows |
                        Measure-Object -Property openingQuantity -Sum
                    ).Sum),
                    3
                )

                inwardQuantity = [Math]::Round(
                    [double]((
                        $sortedSnapshotRows |
                        Measure-Object -Property inwardQuantity -Sum
                    ).Sum),
                    3
                )

                outwardQuantity = [Math]::Round(
                    [double]((
                        $sortedSnapshotRows |
                        Measure-Object -Property outwardQuantity -Sum
                    ).Sum),
                    3
                )

                purchaseQuantity = [Math]::Round(
                    [double]((
                        $sortedSnapshotRows |
                        Measure-Object -Property purchaseQuantity -Sum
                    ).Sum),
                    3
                )

                saleQuantity = [Math]::Round(
                    [double]((
                        $sortedSnapshotRows |
                        Measure-Object -Property saleQuantity -Sum
                    ).Sum),
                    3
                )

                stockValue = [Math]::Round(
                    [double]((
                        $sortedSnapshotRows |
                        Measure-Object -Property stockValue -Sum
                    ).Sum),
                    2
                )

                busyStockValue = [Math]::Round(
                    [double]((
                        $sortedSnapshotRows |
                        Measure-Object -Property busyStockValue -Sum
                    ).Sum),
                    2
                )
            }

            Stop-ReportEngineWatch `
                -Watch $snapshotWatch `
                -Stage "STOCK-SNAPSHOT-FAST" `
                -Message (
                    "Set-based snapshot ready; rows=" +
                    $sortedSnapshotRows.Count +
                    "; mcStates=" +
                    $snapshotStatesByItem.Count +
                    "; NO STOCK-WA replay"
                )

            return @{
                success = $true
                view = "balances"
                fromDate = $fromDate.ToString("yyyy-MM-dd")
                toDate = $endDate.ToString("yyyy-MM-dd")
                asOfDate = $asOfDate.ToString("yyyy-MM-dd")
                generatedAt = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")

                page = 1
                pageSize = 0
                total = $sortedSnapshotRows.Count
                totalPages = 1
                companyItemCount = $companyItemCount
                pageSizeOptions = @(50, 100, 200, 500, 0)

                summaryScope = "all"

                valuationSource =
                    "busy-posted-ledger-fast-snapshot-v17-hierarchical-group-filter"

                stockValuationMethod = 0

                stockValuationMode =
                    "posted-ledger-set-based-snapshot-v17"

                options = @{
                    unitMode = "both"
                    showValue = $true
                    includeStockTransfers = $true
                    showSalePurchaseSeparately = $true
                    mastersMode = "all"
                    showParentGroup = $true
                    valueBy = "busy"
                }

                filterOptions = @{
                    materialCentres = @($mcOptions)
                    itemGroups =
                        @(
                            $groupOptions |
                            Sort-Object pathLabel, name
                        )
                }

                summary = $snapshotSummary
                groupTotals = @($snapshotGroupTotals)
                data = @($sortedSnapshotRows)
            }
        }
        elseif ($FastSnapshot) {
            Write-ReportEngineLog `
                -Stage "STOCK-SNAPSHOT-FAST" `
                -Message (
                    "Direct DB connection unavailable; " +
                    "falling back to legacy stock engine"
                ) `
                -Color "Yellow"
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
        # TRUE FAST PAGINATION FOR PAGE-ORIENTED VIEWS
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
        # Balances, Detailed and Columnar can value only the requested page.
        # Grouped/Hierarchical remain global views and keep the complete
        # calculation path, but are now requested lazily only when clicked.
        # PageSize=0 ("All") and very large pages also keep the full path.
        # ========================================================
        $fastPageModeActive = (
            $viewMode -in @("balances", "detailed", "columnar") -and
            $PageSize -gt 0 -and
            $PageSize -le 250 -and
            -not $stockGeneralApplyFullScanRequested
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

            # Persistent quantity index survives bridge restarts.
            #
            # Access/BDS invalidation:
            #   physical DB file timestamp + file size
            #
            # SQL Server invalidation:
            #   cheap voucher/activity signature from Tran1 + Tran2
            #
            # The SQL signature intentionally avoids COUNT(*) over Tran2.
            # MAX(VchCode) and MAX(Tran2.VchCode) are enough to invalidate on
            # normal new-voucher activity, while the 30-minute in-memory TTL
            # still bounds staleness for unusual in-place edits.
            $fastDiskCachePath = ""
            $fastDiskDbStamp = ""
            $fastDiskIdentity = ""

            if ($null -ne $directStockContext) {
                $fastDiskDbStamp = (
                    [string]$directStockContext.writeTicks +
                    ":" +
                    [string]$directStockContext.length
                )

                $fastDiskIdentity =
                    [string]$directStockContext.path
            }
            elseif (
                $stockDbType -eq 1 -and
                $null -ne $directStockConn
            ) {
                try {
                    $sqlStampTable =
                        Invoke-StockStatusDirectTable `
                            -Connection $directStockConn `
                            -Timeout 30 `
                            -Sql @"
SELECT
    ISNULL(MAX(VchCode), 0) AS MaxTran1VchCode,
    ISNULL(MAX(CreationTime), 0) AS MaxTran1CreationTime
FROM Tran1
"@

                    $sqlTran2StampTable =
                        Invoke-StockStatusDirectTable `
                            -Connection $directStockConn `
                            -Timeout 30 `
                            -Sql @"
SELECT ISNULL(MAX(VchCode), 0) AS MaxTran2VchCode
FROM Tran2
WHERE RecType = 2
"@

                    $stampTran1Code = 0
                    $stampTran2Code = 0
                    $stampCreation = ""

                    if ($sqlStampTable.Rows.Count -gt 0) {
                        $stampTran1Code =
                            ConvertTo-ReportInt (
                                $sqlStampTable.Rows[0]["MaxTran1VchCode"]
                            )

                        try {
                            $rawCreation =
                                $sqlStampTable.Rows[0][
                                    "MaxTran1CreationTime"
                                ]

                            if (
                                $null -ne $rawCreation -and
                                $rawCreation -ne
                                [System.DBNull]::Value
                            ) {
                                $stampCreation =
                                    [string]$rawCreation
                            }
                        }
                        catch {}
                    }

                    if ($sqlTran2StampTable.Rows.Count -gt 0) {
                        $stampTran2Code =
                            ConvertTo-ReportInt (
                                $sqlTran2StampTable.Rows[0][
                                    "MaxTran2VchCode"
                                ]
                            )
                    }

                    $fastDiskDbStamp = (
                        [string]$stampTran1Code +
                        ":" +
                        [string]$stampTran2Code +
                        ":" +
                        $stampCreation
                    )

                    $fastDiskIdentity = (
                        "sql|" +
                        [string]$InstanceId +
                        "|" +
                        [string]$CompanyCode
                    )
                }
                catch {
                    # Persistent cache is optional. Direct SQL aggregation still
                    # works even if the lightweight DB signature cannot be read.
                    $fastDiskDbStamp = ""
                    $fastDiskIdentity = ""
                }
            }

            if (
                -not [string]::IsNullOrWhiteSpace(
                    $fastDiskDbStamp
                ) -and
                -not [string]::IsNullOrWhiteSpace(
                    $fastDiskIdentity
                )
            ) {
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
                                $fastDiskIdentity
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
                            "stock-fast-index-v11-single-snapshot" -and
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
                            engine = "stock-fast-index-v11-single-snapshot"
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

        # Disabled by default. Requires deliberate process-level opt-in.
        # This lab is deliberately restricted to the verified GS item 1198.
        $stockWaLabEnabled = ([string]$env:BUSY_STOCK_WA_LAB_1198 -eq '1')
        $stockWaLabByItem = @{}
        # GS SQL, 2026-10-14 ALL-ITEMS VALUATION EXPERIMENT.
        # OFF (default)    : no all-item replay or override.
        # observe          : replay candidates, log status, NEVER alter report values.
        # apply-verified   : override only exact names/aliases with native-matched
        #                    2026-10-14 reference qty/value, with MC reconciliation.
        # Safety: exact SQL database identity. Other companies / Access never affected.
        $stockAllWaRequestedMode = ([string]$env:BUSY_STOCK_WA_LAB_ALL_ITEMS).Trim().ToLowerInvariant()
        $stockAllWaMode = 'off'
        if ($stockAllWaRequestedMode -in @('observe','apply-verified')) {
            $isGsStockSql = $false
            if ($stockDbType -eq 1 -and $null -ne $directStockConn) {
                try {
                    $isGsStockSql = (
                        [string]$directStockConn.Database -eq 'BusyComp0005_db12026'
                    )
                }
                catch { $isGsStockSql = $false }
            }
            if (
                $isGsStockSql -and
                ([string]$CompanyCode).Trim() -ieq 'Comp0005' -and
                ([string]$env:BUSY_STOCK_WA_LAB_1198 -ne '1') -and
                $fromDate.ToString('yyyy-MM-dd') -eq '2026-04-01' -and
                $endDate.ToString('yyyy-MM-dd') -eq '2026-10-14' -and
                $asOfDate.ToString('yyyy-MM-dd') -eq '2026-10-14' -and
                $IncludeStockTransfers -and
                $valueMode -eq 'busy'
            ) {
                $stockAllWaMode = $stockAllWaRequestedMode
            }
            else {
                Write-ReportEngineLog -Stage 'STOCK-ALLWA-GUARD' -Message 'Lab refused: GS SQL database + Comp0005, FY dates, transfers, BUSY valuation and no simultaneous item-1198 lab required' -Color 'Yellow'
            }
        }
        # SEPTEMBER 30 TWO-DATE TEST: independent, explicit opt-in. The
        # original October 14 BUSY_STOCK_WA_LAB_ALL_ITEMS experiment above
        # stays byte-for-byte unchanged and remains date restricted.
        # NEVER allow a date range or arbitrary dates through this gate.
        $stockSep30RequestedMode = ([string]$env:BUSY_STOCK_WA_LAB_SEP30).Trim().ToLowerInvariant()
        if ($stockAllWaMode -eq 'off' -and $stockSep30RequestedMode -in @('observe','apply-verified')) {
            $stockSep30DatabaseOK = $false
            if ($stockDbType -eq 1 -and $null -ne $directStockConn) {
                try {
                    $stockSep30DatabaseOK = ([string]$directStockConn.Database -ceq 'BusyComp0005_db12026')
                }
                catch { $stockSep30DatabaseOK = $false }
            }
            if (
                $stockSep30DatabaseOK -and
                ([string]$CompanyCode).Trim() -ieq 'Comp0005' -and
                ([string]$env:BUSY_STOCK_WA_LAB_1198 -ne '1') -and
                $fromDate.ToString('yyyy-MM-dd') -eq '2026-04-01' -and
                $endDate.ToString('yyyy-MM-dd') -eq '2026-09-30' -and
                $asOfDate.ToString('yyyy-MM-dd') -eq '2026-09-30' -and
                $IncludeStockTransfers -and
                $valueMode -eq 'busy'
            ) {
                $stockAllWaMode = $stockSep30RequestedMode
            }
            else {
                Write-ReportEngineLog -Stage 'STOCK-SEP30-GUARD' -Message 'Sep30 lab refused: exact GS SQL identity/date, BUSY valuation, stock transfers and no 1198 lab required' -Color 'Yellow'
            }
        }
        # AUGUST 31 THREE-DATE TEST: a separate, explicit opt-in. October 14
        # and September 30 paths above are unchanged. No arbitrary date support.
        $stockAug31RequestedMode = ([string]$env:BUSY_STOCK_WA_LAB_AUG31).Trim().ToLowerInvariant()
        if ($stockAllWaMode -eq 'off' -and $stockAug31RequestedMode -in @('observe','apply-verified')) {
            $stockAug31DatabaseOK = $false
            if ($stockDbType -eq 1 -and $null -ne $directStockConn) {
                try {
                    $stockAug31DatabaseOK = ([string]$directStockConn.Database -ceq 'BusyComp0005_db12026')
                }
                catch { $stockAug31DatabaseOK = $false }
            }
            if (
                $stockAug31DatabaseOK -and
                ([string]$CompanyCode).Trim() -ieq 'Comp0005' -and
                ([string]$env:BUSY_STOCK_WA_LAB_1198 -ne '1') -and
                $fromDate.ToString('yyyy-MM-dd') -eq '2026-04-01' -and
                $endDate.ToString('yyyy-MM-dd') -eq '2026-08-31' -and
                $asOfDate.ToString('yyyy-MM-dd') -eq '2026-08-31' -and
                $IncludeStockTransfers -and
                $valueMode -eq 'busy'
            ) {
                $stockAllWaMode = $stockAug31RequestedMode
            }
            else {
                Write-ReportEngineLog -Stage 'STOCK-AUG31-GUARD' -Message 'Aug31 lab refused: exact GS SQL identity/date, BUSY valuation, stock transfers and no 1198 lab required' -Color 'Yellow'
            }
        }
        # [STOCK-GENERAL-OBSERVE-V1:GATE:BEGIN]
        # Observation is exclusive to the GS SQL FY and does not grant any
        # new apply permission. Old date-specific modes always take priority.
        $stockGeneralObserveMode = 'off'
        if (
            $stockAllWaMode -eq 'off' -and
            ([string]$env:BUSY_STOCK_WA_GENERAL_OBSERVE).Trim().ToLowerInvariant() -eq 'observe'
        ) {
            $stockGeneralObserveDbOK = $false
            if ($stockDbType -eq 1 -and $null -ne $directStockConn) {
                try {
                    $stockGeneralObserveDbOK = (
                        [string]$directStockConn.Database -ceq 'BusyComp0005_db12026'
                    )
                }
                catch { $stockGeneralObserveDbOK = $false }
            }
            if (
                $stockGeneralObserveDbOK -and
                ([string]$CompanyCode).Trim() -ieq 'Comp0005' -and
                ([string]$env:BUSY_STOCK_WA_LAB_1198 -ne '1') -and
                $fromDate.ToString('yyyy-MM-dd') -eq '2026-04-01' -and
                $endDate.Date -ge [datetime]'2026-04-01' -and
                $endDate.Date -le [datetime]'2027-03-31' -and
                $asOfDate.Date -eq $endDate.Date -and
                $IncludeStockTransfers -and
                $valueMode -eq 'busy'
            ) {
                # Reuse proven replay branch; 'observe' can NEVER enter the
                # existing reference-matched apply-verified override branch.
                $stockAllWaMode = 'observe'
                $stockGeneralObserveMode = 'observe'
            }
            else {
                Write-ReportEngineLog -Stage 'STOCK-GENERAL-OBSERVE-GUARD' -Message 'General observation refused: requires exact GS SQL/FY/as-of/company, transfers, BUSY valuation and no item-1198 lab' -Color 'Yellow'
            }
        }
        # [STOCK-GENERAL-OBSERVE-V1:GATE:END]
        # [STOCK-GENERAL-APPLY-V4:GATE:BEGIN]
        # A distinct OFF-by-default test switch. The three older checkpoint
        # apply modes retain priority. The general engine is not an alias
        # for apply-verified and does not consult native reference values.
        function Test-StockStatusGeneralApplyGuard {
            param(
                [string]$RequestedMode,
                [int]$DbType,
                [string]$DatabaseName,
                [string]$TargetInstance,
                [string]$TargetCompany,
                [datetime]$StartDate,
                [datetime]$FinishDate,
                [datetime]$CutoffDate,
                [bool]$TransfersIncluded,
                [string]$ValuationMode,
                [string]$ReportView,
                [string]$Other1198,
                [string]$OtherOct,
                [string]$OtherSep,
                [string]$OtherAug,
                [string]$OtherGeneralObserve
            )
            return (
                $RequestedMode -ceq 'apply-tested' -and
                $DbType -eq 1 -and
                $DatabaseName -ceq 'BusyComp0005_db12026' -and
                $TargetInstance -ceq 'instance_2' -and
                $TargetCompany -ceq 'Comp0005' -and
                $StartDate.Date -eq [datetime]'2026-04-01' -and
                $FinishDate.Date -ge [datetime]'2026-04-01' -and
                $FinishDate.Date -le [datetime]'2027-03-31' -and
                $CutoffDate.Date -eq $FinishDate.Date -and
                $TransfersIncluded -and
                $ValuationMode -ceq 'busy' -and
                $ReportView -ceq 'balances' -and
                [string]::IsNullOrWhiteSpace($Other1198) -and
                [string]::IsNullOrWhiteSpace($OtherOct) -and
                [string]::IsNullOrWhiteSpace($OtherSep) -and
                [string]::IsNullOrWhiteSpace($OtherAug) -and
                [string]::IsNullOrWhiteSpace($OtherGeneralObserve)
            )
        }
        $stockGeneralApplyMode = 'off'
        if ($stockGeneralApplyRequestedMode.Length -gt 0) {
            $generalApplyDbName = ''
            if ($stockDbType -eq 1 -and $null -ne $directStockConn) {
                try { $generalApplyDbName = [string]$directStockConn.Database }
                catch { $generalApplyDbName = '' }
            }
            if (
                $stockAllWaMode -eq 'off' -and
                (Test-StockStatusGeneralApplyGuard `
                    -RequestedMode $stockGeneralApplyRequestedMode `
                    -DbType $stockDbType `
                    -DatabaseName $generalApplyDbName `
                    -TargetInstance ([string]$InstanceId) `
                    -TargetCompany ([string]$CompanyCode) `
                    -StartDate $fromDate `
                    -FinishDate $endDate `
                    -CutoffDate $asOfDate `
                    -TransfersIncluded ([bool]$IncludeStockTransfers) `
                    -ValuationMode ([string]$valueMode) `
                    -ReportView ([string]$viewMode) `
                    -Other1198 ([string]$env:BUSY_STOCK_WA_LAB_1198) `
                    -OtherOct ([string]$env:BUSY_STOCK_WA_LAB_ALL_ITEMS) `
                    -OtherSep ([string]$env:BUSY_STOCK_WA_LAB_SEP30) `
                    -OtherAug ([string]$env:BUSY_STOCK_WA_LAB_AUG31) `
                    -OtherGeneralObserve ([string]$env:BUSY_STOCK_WA_GENERAL_OBSERVE))
            ) {
                # Reuse existing replay-only path; never enter old native
                # reference-table apply branch.
                $stockAllWaMode = 'observe'
                $stockGeneralApplyMode = 'apply-tested'
            }
            else {
                Write-ReportEngineLog -Stage 'STOCK-GENERAL-APPLY-GUARD' -Message 'Refused general test apply: exact GS SQL company, FY, BUSY balances view, transfers and exclusive opt-in required' -Color 'Yellow'
            }
        }
        $stockGeneralCandidateMode = if ($stockGeneralApplyMode -eq 'apply-tested') {
            'apply-tested'
        }
        elseif ($stockGeneralObserveMode -eq 'observe') { 'observe' }
        else { 'off' }
        # [STOCK-GENERAL-APPLY-V4:GATE:END]
        # Verified historical identity + reference, NOT invented rates.
        # The lab calculates stock value afresh; these values are guards only.
        $stockAllWaVerified = @{}
        if ($stockAllWaMode -eq 'apply-verified') {
            $stockAllWaVerified = @{
            '1178' = '30/70 GREEN 2L|TOTACHI--44102|0.000000|0.00'
            '1227' = 'FINE GASOLINE 10W-30 4L|TOTACHI10204|1453.000000|7950200.18'
            '1225' = 'FINE GASOLINE 10W-30 1L|TOTACHI10201|0.000000|0.00'
            '1235' = 'HEAVY DUTY 15W40 60 L|TOTACHI10760|0.000000|0.00'
            '1240' = 'HYDRAULIC OIL NRO 68 - 20L|TOTACHI50620|0.000000|0.00'
            '1244' = 'HYPER ECODRIVE 0W-20 3L SURPRISE PACK|PRO001|0.000000|0.00'
            '1253' = 'SUPER LONG LIFE COOLANT -40C 1L|TOTACHI41601|13.000000|20243.63'
            '1261' = 'TOTACHI-ALLION/PREMIO 260 1NZ,2ZR COROLL|TA-1343|451.000000|368299.61'
            '1271' = 'TOTACHI- FIT L13B(1300,1500 Cc) VEZEL L1|TA-3708|491.000000|640343.23'
            '1313' = 'TOTACHI-HIACE TRH200-2TR,PRADO TRJ120-2T|TO-1189|32.000000|23926.00'
            '2558' = 'HEAVY DUTY 15W40 20L|TOTACHI10720|0.000000|0.00'
            '3004' = 'DENTO PAWA 4T 20W-50 1L|TOTACHI13601|0.000000|0.00'
            '3313' = 'HYPER ECODRIVE 0W20 3L|TOTACHIE0103|3.000000|22271.72'
            '3583' = 'C525J OIL FILTER|C525J|-4.000000|-17288.12'
            '4084' = 'HYPER ECODRIVE 0W20 200 L|TOTACHIE012Z|0.000000|0.00'
            '4181' = 'DENTO HEAVY DUTY DIESEL 15W40  200L|TOTACHIE342Z|0.000000|0.00'
            '1180' = 'ATF DEX III 20L|TOTACHI20720|45.000000|989402.18'
            '1181' = 'ATF DEX-III 1L|TOTACHI20701|160.000000|260172.61'
            '1182' = 'ATF DEX-III 4 L|TOTACHI20704|209.000000|1195477.27'
            '1183' = 'ATF MULTI VEHICLE LV 4L|TOTACHIA3004|173.000000|1526386.78'
            '1184' = 'ATF MULTI-VEHICLE 1L|TOTACHI20601|177.000000|445741.36'
            '1216' = 'ECO GASOLINE 5W30 4 L|TOTACHI10804|3.000000|18695.46'
            '1248' = 'HYPER ULV 0W16,PURPLE 4L|TOTACHIE0004|441.000000|3258092.64'
            '4062' = 'HYPER ULV 0W8 4L|TOTACHIE5504|1105.000000|8915216.20'
            '1185' = 'ATF MULTI-VEHICLE 4L|TOTACHI20604|154.000000|1375589.52'
            '1186' = 'ATF SP IV 4L|TOTACHI21004|177.000000|1708856.53'
            '1187' = 'ATF SP-III 1 L|TOTACHI20401|145.000000|362768.70'
            '1188' = 'ATF SP-III 4 L|TOTACHI20404|84.000000|645871.80'
            '1189' = 'ATF SP-IV 1L|TOTACHI21001|423.000000|1269702.06'
            '1190' = 'ATF TYPE T-IV 1 L|TOTACHI20201|830.000000|1837572.75'
            '1191' = 'ATF TYPE T-IV 4 L|TOTACHI20204|1016.000000|6805692.80'
            '1192' = 'ATF WS  4L|TOTACHI20804|374.000000|3134704.87'
            '1193' = 'ATF WS 1L|TOTACHI20801|756.000000|2206646.36'
            '1194' = 'ATF Z-1 1L|TOTACHI20301|416.000000|1186230.18'
            '1195' = 'Concentrated Engine Power Flush 300 ML-T|TOTACHIP7130|403.000000|668648.65'
            '1197' = 'CVT MULTI-TYPE 1 L|TOTACHI20501|27.000000|97613.21'
            '1198' = 'CVT MULTI-TYPE 4 L|TOTACHI20504|206.000000|2599029.34'
            '1199' = 'CVT NS-3, BLUE 1 L|TOTACHI21101|986.000000|3472846.62'
            '1200' = 'CVT NS-3, BLUE 4 L|TOTACHI21104|47.000000|503823.75'
            '1201' = 'DCTF MULTI TYPE 4L|TOTACHIA8204|327.000000|4629262.50'
            '1202' = 'DCTF MULTI-TYPE 1 L|TOTACHIA8201|179.000000|639798.53'
            '1204' = 'DENTO ECO GASOLINE 15W-40 200 L|TOTACHI13922|11.000000|3140795.47'
            '1205' = 'DENTO ECO GASOLINE 15W-40 60 L|TOTACHI13960|0.000000|0.00'
            '1207' = 'DENTO HEAVY DUTY 15W40 CI-4 1 L|TOTACHI13201|1.000000|1818.47'
            '1208' = 'DENTO HEAVY DUTY 15W40 CI-4 4 L|TOTACHI13204|0.000000|0.00'
            '1209' = 'Diesel Cetane & Fuel System Boost 250 ML|TOTACHIP7025|73.000000|115967.49'
            '1210' = 'ECO DIESEL 10W-40 1 L|TOTACHIE1301|0.000000|0.00'
            '1211' = 'ECO DIESEL 10W-40 4 L|TOTACHIE1304|1.000000|6623.98'
            '1212' = 'FINE GASOLINE 10W-30 3L SURPRISE PACK|PRO004|0.000000|0.00'
            '1213' = 'ECO GASOLINE 10W-40 1 L|TOTACHI10901|1079.000000|3115903.46'
            '1214' = 'ECO GASOLINE 10W-40 4 L|TOTACHI10904|796.000000|4948845.14'
            '1215' = 'ECO GASOLINE 5W-30 1 L|TOTACHI10801|69.000000|117011.21'
            '1217' = 'EURODRIVE 5W40 4L|TOTACHIE6304|364.000000|3623703.87'
            '1218' = 'EURODRIVE ECO 5W30 4 L|TOTACHIE6604|303.000000|3378057.63'
            '1219' = 'EURODRIVE PRO LONG LIFE 5W30 4 L|TOTACHIE7804|0.000000|0.00'
            '1220' = 'Extended Life Coolant Blue 30/70 1 L|TOTACHIT8701|58.000000|52631.65'
            '1221' = 'Extended Life Coolant Green 30/70  1L|TOTACHIT8501|2912.000000|2598744.88'
            '1222' = 'Extended Life Coolant Green 30/70 4 L|TOTACHIT8504|104.000000|295243.31'
            '1223' = 'Extended Life Coolant Red 30/70 1 L|TOTACHIT7301|565.000000|415137.62'
            '1224' = 'Extended Life Coolant Red 30/70 4 L|TOTACHIT7304|19.000000|45665.40'
            '1226' = 'FINE GASOLINE 10W-30 200 L|TOTACHI10222|0.000000|0.00'
            '1228' = 'FINE GASOLINE 10W-30 60 L|TOTACHI10260|0.000000|0.00'
            '1229' = 'FLUSH OUT 4 L-TOTACHI12304|TOTACHI12304|137.000000|591200.81'
            '1230' = 'Fuel Injector Cleaner 300 ML|TOTACHIP8830|5862.000000|10835913.81'
            '1231' = 'GREEN LONG LIFE COOLANT 30/70 2L|TOTACHI44102|0.000000|0.00'
            '1232' = 'HEAVY DUTY 15W40 1L|TOTACHI10701|69.000000|123962.24'
            '1233' = 'HEAVY DUTY 15W-40 200L|TOTACHI10722|0.000000|0.00'
            '1234' = 'HEAVY DUTY 15W-40 4 L|TOTACHI10704|227.000000|1045852.82'
            '1236' = 'HEAVY DUTY 15W40 6L|TOTACHI10706|93.000000|953963.24'
            '1237' = 'HEAVY DUTY 15W-40 7L SURPRISE PACK|PRO003|0.000000|0.00'
            '1238' = 'HEAVY DUTY 15W-40 6L SURPRISE PACK|PRO002|0.000000|0.00'
            '1241' = 'HYPER ECODRIVE 0W20 1L|TOTACHIE0101|4.000000|9087.52'
            '1242' = 'HYPER ECODRIVE 0W-20 4L|TOTACHIE0104|4.000000|74609.35'
            '1243' = 'HYPER ECODRIVE 0W20 60 L|TOTACHIE0160|0.000000|0.00'
            '1245' = 'HYPER ECODRIVE 0W-40 1 L|TOTACHI12801|20.000000|52713.15'
            '1246' = 'HYPER ECODRIVE 0W-40 4 L|TOTACHI12804|17.000000|184260.30'
            '1247' = 'HYPER ULV 0W16,PURPLE 1L|TOTACHIE0001|1993.000000|4899769.13'
            '1249' = 'MIX-TYPE COOLANT PINK -40C G12EV 1L|TOTACHI46801|1.000000|2711.86'
            '1250' = 'POWER STEERING FLUID 1L|TOTACHI90301|201.000000|459277.57'
            '1251' = 'RED LONG LIFE COOLANT -40C 2L|TOTACHI41802|0.000000|0.00'
            '1252' = 'SUPER LONG LIFE COOLAN GREEN 1L||0.000000|0.00'
            '1254' = 'SUPER LONG LIFE COOLANT -40C 2L|TOTACHI44602|0.000000|0.00'
            '1255' = 'SUPER LONG LIFE COOLANT HYBRID-40 1L|TOTACHI44601|966.000000|1601677.17'
            '1260' = 'TOTACHI-COROLLA NZE121-TA-1334|TA-1334|380.000000|387185.98'
            '1262' = 'TOTACHI-HIACE KDH-TA-1380|TA-1380|137.000000|376776.37'
            '1263' = 'TOTACHI-LEXUS-TA-1387|TA-1387|0.000000|0.00'
            '1264' = 'TOTACHI-LAND CRUISER UZJ120/150 2TR-TA-1|TA-1427|58.000000|121178.94'
            '1265' = 'TOTACHI-VITZ-TA-1430|TA-1430|103.000000|67467.62'
            '1266' = 'TOTACHI- SUNNY QG13,QG15,QG16,CD17 X-TRA|TA-1465|105.000000|126889.43'
            '1267' = 'TOTACHI X-TRAIL NT32-TA-1873|TA-1873|64.000000|57537.62'
            '1268' = 'TOTACHI-PRIUS III-TA-1877|TA-1877|236.000000|245235.18'
            '1269' = 'TOTACHI-HONDA FIT GP5-TA-1882|TA-1882|435.000000|398071.66'
            '1270' = 'TOTACHI-YARIS HYBRID,AQUA-TA-2843|TA-2843|491.000000|456433.93'
            '1272' = 'TOTACHI - RANGER BT50-TA-3855|TA-3855|0.000000|0.00'
            '1274' = 'TOTACHI-LAND CRUISER 200 (A1533)-TA-3856|TA-3856|109.000000|226349.72'
            '1275' = 'TOTACHI C-HR NGX10,NGX50-TA-3859|TA-3859|113.000000|274823.17'
            '1277' = 'TOTACHI- HIACE,2L,3L-TC-1024|TC-1024|250.000000|826051.95'
            '1278' = 'TOTACHI-AQUA,AVANZA,COROLLA,IST,PRIUS (I|TC-1030|11503.000000|6437735.42'
            '1279' = 'TOTACHI-DYNA 1KD,PRADO 120,HIACE 1KD-TC-|TC-1031|3213.000000|2337383.32'
            '1280' = 'TOTACHI-LITEACE,TOWNACE,CR27-TC-1032|TC-1032|1299.000000|2633619.00'
            '1281' = 'TOTACHI-LAND CRUISER 100,120,CROWN,HILUX|TC-1034|200.000000|194744.32'
            '1282' = 'TOTACHI-COASTER 1HZ,1HD,HIACE,TOYOACE,HI|TC-1035|1469.000000|2483954.37'
            '1283' = 'TOTACHI - SUNNY B15,N16-TC-1047|TC-1047|383.000000|297206.06'
            '1285' = 'TOTACHI-PICK-UP,L200,L300,PAJERO-TC-1054|TC-1054|105.000000|221593.82'
            '1286' = 'TOTACHI-LANCER CS7-TC-1055|TC-1055|163.000000|275470.00'
            '1287' = 'TOTACHI- LANCER 4G15,L300,L400 DELICA-TC|TC-1065|399.000000|495250.77'
            '1288' = 'TOTACHI- FIT AREA L13 VEZEL L15-TC-1096|TC-1096|6996.000000|3166912.92'
            '1289' = 'TOTACHI-EVERY,MINICAB-TC-1097|TC-1097|382.000000|224435.98'
            '1292' = 'TOTACHI-NOAH,HILUX,KUN26-TCA-101|TCA-101|255.000000|531694.71'
            '1293' = 'TOTACHI-COROLLA NZE121 (AC101EX)-TCA-160|TCA-160|191.000000|336283.56'
            '1294' = 'TOTACHI-PRADO KDJ150,TRJ150-TCA-161|TCA-161|171.000000|125473.20'
            '1296' = 'TOTACHI-ALLION 260,BELTA SCP,NCP92,96-TC|TCA-166|440.000000|287143.27'
            '1297' = 'TOTACHI-ALLION260, BELTA SCP-TCA-166K|TCA-166K|6.000000|6113.07'
            '1298' = 'TOTACHI-SUNNY FB15-TCA-176|TCA-176|391.000000|920935.76'
            '1299' = 'TOTACHI-FIT JAZZ-TCA-229|TCA-229|105.000000|110054.92'
            '1300' = 'TOTACHI-VEZEL RU1-TCA-233|TCA-233|495.000000|524334.55'
            '1301' = 'TOTACHI-CIVIC,CRV-TCA-237|TCA-237|124.000000|118139.56'
            '1302' = 'TOTACHI X-TRAIL NT32 (AC2514)-TCA-566|TCA-566|295.000000|207458.64'
            '1303' = 'TOTACHI-HILUX VIGO,C-HR-TCA-806K|TCA-806K|555.000000|649350.14'
            '1304' = 'TOTACHI-MINI CAB-TCA-817|TCA-817|88.000000|96231.87'
            '1306' = 'TOTACHI HIACE,LITEACE,TOWNACE FILTER-TF-|TF-1016|829.000000|998791.29'
            '1307' = 'TOTACHI TOWNACE,RAV 4,YARIS,LITEACE-TF-1|TF-1019|1150.000000|1966662.60'
            '1308' = 'TOTACHI - TF-1020-TF-1020|TF-1020|1125.000000|1047901.76'
            '1309' = 'TOTACHI-1KD,2KD,KDH TF-1090-TF-1090|TF-1090|177.000000|102247.17'
            '1310' = 'TOTACHI- PRADO 120,150-TF-1316|TF-1316|185.000000|190546.63'
            '1311' = 'TOTACHI-LAND CRUISER 200-TF-1731|TF-1731|115.000000|119997.59'
            '1312' = 'TOTACHI LAND CRUISER 200(0121)-TO-1162|TO-1162|42.000000|40290.16'
            '1314' = 'TOTACHI VITZ(1KR),PASSO,ALLION PREMIO ZR|TO-1190|3261.000000|1011714.17'
            '1315' = 'TOTACHI-RANGER BT50-TO-1299|TO-1299|36.000000|23606.66'
            '1316' = 'TOTACHI-JUKE,QASHQAL-TO-1672|TO-1672|0.000000|0.00'
            '1318' = 'TOTACHI-SWIFT,ALTO-TC-1100|TC-1100|7654.000000|3979289.35'
            '2193' = 'HEAVY DUTY 15W-40 8L SURPRISE PACK|PRO005|0.000000|0.00'
            '2409' = 'EXTENDED LIFE COOL RED -40C (50/50) 1L|TOTACHIT7201|3010.000000|3229381.95'
            '2410' = 'EXTENDED LIFE COOL RED -40C (50/50) 4L|TOTACHIT7204|221.000000|795400.52'
            '2411' = 'SUPER LONG LIFE COOLANT -37C (50/50) 1L|TOTACHIT7401|1655.000000|2275565.27'
            '2412' = 'SUPER LONG LIFE COOLANT -37C (50/50) 4L|TOTACHIT7404|793.000000|3791506.18'
            '2460' = 'HYPER ECODRIVE 0W-20 PROMO (25 CANS + 9|PRO006|0.000000|0.00'
            '2461' = 'HYPER ECODRIVE 0W-20 PROMO (50 CANS + 20|PRO007|0.000000|0.00'
            '2462' = 'HYPER ECODRIVE 0W-20 PROMO (100 CANS + 4|PRO008|0.000000|0.00'
            '2463' = 'FINE GASOLINE 10W-30 PROMO (50 CANS+20 C|PRO009|0.000000|0.00'
            '2464' = 'FINE GASOLINE 10W-30 4L WITH TC-1030|PRO010|0.000000|0.00'
            '2465' = 'FINE GASOLINE 10W-30 PROMO (100 CANS +42|PRO011|0.000000|0.00'
            '2466' = 'DENTO HEAVY DUTY 15W-40 CI-4 7L SURPRISE|PRO012|0.000000|0.00'
            '2467' = 'DENTO HEAVY DUTY 15W-40 4L PROMO (25 CAN|PRO013|0.000000|0.00'
            '2468' = 'DENTO HEAVY DUTY 15W-40 4L PROMO (50 CAN|PRO014|0.000000|0.00'
            '2469' = 'DENTO HEAVY DUTY 15W-40 4L PROMO (100 CA|PRO015|0.000000|0.00'
            '2470' = 'DENTO HEAVY DUTY 15W-40 4L PROMO (200 CA|PRO016|0.000000|0.00'
            '2471' = 'DENTO HD 15W-40 5L PROMO (25 CANS+10 FRE|PRO017|0.000000|0.00'
            '2472' = 'DENTO HEAVY DUTY 15W-40 5L PROMO (50 CAN|PRO018|0.000000|0.00'
            '2473' = 'DENTO HEAVY DUTY 15W-40 5L PROMO (100 CA|PRO019|0.000000|0.00'
            '2474' = 'DENTO HEAVY DUTY 15W-40 5L PROMO (200 CA|PRO020|0.000000|0.00'
            '2475' = 'FINE GASOLINE 10W-30 PROMO (30 CANS +11|PRO021|0.000000|0.00'
            '2511' = 'DENTO HEAVY DUTY 15W-40 6L SURPRISE PACK|PRO024|0.000000|0.00'
            '2512' = 'DENTO HEAVY DUTY 15W40 CI-4 5L|TOTACHI13205|1.000000|10911.61'
            '2586' = 'FINE GASOLINE 10W-30 PROMO (25 CANS +10|PRO022|0.000000|0.00'
            '2745' = 'FINE GASOLINE 10W-30 PROMO (50 CANS+22 C|PRO023|0.000000|0.00'
            '2748' = 'FINE GASOLINE 10W-30 3L PACK (25 + 9) FR|PRO025|0.000000|0.00'
            '2786' = 'FINE GASOLINE 10W-30 3L PACK (25 + 10) F|PRO026|0.000000|0.00'
            '2794' = 'FINE GASOLINE 10W-30 3L PACK (50 +21) FR|PRO027|0.000000|0.00'
            '2801' = 'FINE GASOLINE 10W-30 PROMO(100CANS+54 C|PRO033|0.000000|0.00'
            '2824' = 'FINE GASOLINE 10W-30 PROMO (50 CANS+24 C|PRO029|0.000000|0.00'
            '2997' = 'Extended Life Coolant Blue 30/70 4 L|TOTACHIT8704|485.000000|1093306.46'
            '3003' = 'DENTO PAWA 4T 20W-50 1L (3-WHEEL)|TOTACHIE4201|0.000000|0.00'
            '3005' = 'ATF MULTI VEHICLE LV 1L|TOTACHIA3001|3637.000000|9295607.29'
            '3006' = 'EURODRIVE 5W-40 1L|TOTACHIE6301|213.000000|632758.81'
            '3007' = 'EURODRIVE ECO 5W-30 1 L|TOTACHIE6601|325.000000|1042466.27'
            '3008' = 'EURODRIVE PRO FUEL EFFICIENCY 5W-30 4L|TOTACHIE7904|529.000000|5169560.62'
            '3009' = 'HYPER ECODRIVE 10W-30 1L|TOTACHIE2601|205.000000|479184.74'
            '3010' = 'HYPER ECODRIVE 10W-30 4L|TOTACHIE2604|339.000000|2479922.23'
            '3011' = 'FINE GASOLINE 15W-40 4L (DON''T USE)||0.000000|0.00'
            '3012' = 'ULTIMA RACING 10W-60 1L|TOTACHIE5001|0.000000|0.00'
            '3075' = 'FINE GASOLINE 10W-30 PROMO (100 CANS +43|PRO030|0.000000|0.00'
            '3084' = 'FINE GASOLINE 10W-30 PROMO (25 CANS +9 F|PRO031|0.000000|0.00'
            '3090' = 'FINE GASOLINE 10W-30 PROMO (50 CANS+25 C|PRO032|0.000000|0.00'
            '3091' = 'FINE GASOLINE 10W-30 PROMO (25 CANS +11|PRO034|0.000000|0.00'
            '3134' = 'HYDRAULIC OIL NRO 68-200L|TOTACHI50622|0.000000|0.00'
            '3135' = 'HYDRAULIC OIL NRO 46-20L|TOTACHI50520|74.000000|1268490.98'
            '3150' = 'FINE GASOLINE 10W-30 4L PROMO (200+110)|PRO035|0.000000|0.00'
            '3152' = 'FINE GASOLINE 10W-30 4L FREE ISSUE|10W30-4L-FREE ISSUE|0.000000|0.00'
            '3190' = 'FINE GASOLINE 15W40 SP 1L|TOTACHIE1901|0.000000|0.00'
            '3191' = 'FINE GASOLINE 15W40 SP 4L|TOTACHIE1904|3.000000|20067.84'
            '3192' = 'FINE GASOLINE 15W40 SP 60L|TOTACHIE1960|47.000000|3557923.97'
            '3193' = 'FINE GASOLINE 15W40 SP 200L|TOTACHIE192Z|0.000000|0.00'
            '3314' = 'FINE GASOLINE 10W30 SP 3L|TOTACHIE1203|1.000000|6777.31'
            '3315' = 'FINE GASOLINE 15W40 SP 3L|TOTACHIE1903|25.000000|141816.44'
            '3322' = 'HYPER ECODRIVE 5W-30 4L|TOTACHIE0304|2.000000|15468.76'
            '3405' = 'DENTO GL-5 80W-90 GEAR OIL 1 L|TOTACHIG5701|625.000000|1046145.57'
            '3422' = 'SUPER LONG LIFE COOLANT GREEN -40C 2L|TOTACHI41602|0.000000|0.00'
            '3519' = 'TA-1418 AIT FILTER|TA-1418|290.000000|294309.40'
            '3520' = 'TA-1379 AIR FILTER|TA1379|219.000000|254473.62'
            '3521' = 'TA-1624 AIR FILTER|TA-1624|6.000000|7132.63'
            '3553' = 'DENTO BRAKE FLUID DOT 4 - 1L|TOTACHIZ9901|879.000000|2298701.14'
            '3584' = 'FC607 FUEL FILTER|FC607|0.000000|0.00'
            '3585' = 'C526J OIL FILTER|C526J|0.000000|0.00'
            '3710' = 'OE23010 OIL FILTER|OE23010|0.000000|0.00'
            '3761' = 'PREMEIUM DIESEL 5W-40 1L|TOTACHI11701|7.000000|0.00'
            '3762' = 'PREMIUM DIESEL 5W-40 4L|TOTACHI11704|1.000000|0.00'
            '3765' = 'ECO GASOLINE 5W-30 3L|TOTACHI10803|2.000000|9374.59'
            '3769' = 'FUEL PUMP - FREE ISSUE||0.000000|0.00'
            '3771' = 'ATF TYPE T-IV 20 L|TOTACHI20220|3.000000|108393.11'
            '3772' = 'EXTENDED LIFE COOLANT PINK -40C(50/50)4L|TOTACHIT0104|58.000000|227933.75'
            '3773' = 'EXTENDED LIFE COOLANT PINK -40C(50/50)1L|TOTACHIT0101|5257.000000|6138940.82'
            '3804' = 'DISC BRAKE PADS|TBM-2040|0.000000|0.00'
            '3891' = 'DENTO PAWA 4T 10W-40 1L|TOTACHIE3701|2046.000000|2597655.70'
            '3926' = 'FINE GASOLINE 15W-40 4L (Dont Use)|TOTACHIE9704|0.000000|0.00'
            '3957' = 'FINE GASOLINE 20W50 1L|TOTACHI10601|65.000000|129429.62'
            '3958' = 'FINE GASOLINE 20W50 4L|TOTACHI10604|432.000000|2190876.01'
            '3959' = 'ULTIMA RACING 5W50 4L|TOTACHIE4804|60.000000|901395.55'
            '3960' = 'ULTIMA RACING 10W60 4L|TOTACHIE5004|101.000000|1458836.01'
            '3961' = 'DENTO HEAVY DUTY DIESEL 20W50 200L|TOTACHIE3520|33.000000|7342680.02'
            '3984' = 'DENTO HYPOID 80W-90 GL-5  20L|TOTACHIG5720|0.000000|0.00'
            '3985' = 'DENTO HYPOID 80W-90 GL-5  200L|TOTACHIG572Z|0.000000|0.00'
            '4060' = 'HEAVY DUTY DIESEL 20W50 20L|TOTACHI18820|69.000000|1618975.98'
            '4061' = 'HEAVY DUTY DIESEL 20W50 60L|TOTACHI18860|4.000000|279636.16'
            '4063' = 'HYPER ULV 0W8 1L|TOTACHIE5501|450.000000|1136857.83'
            '4108' = 'HYBRID WIPER BLADE WBH-350 14"|WBH00710|0.000000|0.00'
            '4110' = 'HYBRID WIPER BLADE WBH-400 16"|WBH00160|628.000000|798888.35'
            '4111' = 'HYBRID WIPER BLADE WBH-450 18"|WBH00260|307.000000|403619.24'
            '4112' = 'HYBRID WIPER BLADE WBH-480 19"|WBH00810|40.000000|53100.90'
            '4113' = 'HYBRID WIPER BLADE WBH-500 20"|WBH00360|138.000000|205242.37'
            '4114' = 'HYBRID WIPER BLADE WBH-530 21"|WBH00910|124.000000|207411.27'
            '4115' = 'HYBRID WIPER BLADE WBH-550 22"|WBH00460|441.000000|686315.65'
            '4116' = 'HYBRID WIPER BLADE WBH-600 24"|WBH00560|1702.000000|2855140.67'
            '4117' = 'HYBRID WIPER BLADE WBH-650 26"|WBH00660|801.000000|1395017.20'
            '4180' = 'ECO DIESEL 15W-40 1 L|TOTACHIE5601|329.000000|673089.68'
            '4209' = 'C110 - SRR OIL FILTER|C110|0.000000|0.00'
            '4227' = 'FINE GASOLINE 15W 40 SP 4L PACK|PRO036|0.000000|0.00'
            '4231' = 'HYPER ECODRIVE 0W20 4L PACK|PRO040|0.000000|0.00'
            '4253' = 'EURODRIVE ECO 0W20 4 L|TOTACHIE6404|114.000000|1099615.42'
            '4254' = 'ECO DIESEL 5W-30 1 L|TOTACHIE2201|133.000000|276071.80'
            '4255' = 'ECO DIESEL 5W-30 6 L|TOTACHIE2206|86.000000|946615.68'
            '4256' = 'DENTO HEAVY DUTY DIESEL 15W40  25 L|TOTACHIE341Z|0.000000|0.00'
            '4257' = 'DENTO HEAVY DUTY DIESEL 20W50  25 L|TOTACHIE351Z|15.000000|442031.74'
            '4397' = 'ECO DIESEL 5W30 200 L|TOTACHIE222Z|3.000000|1139604.80'
            '4426' = 'ECO GASOLINE 5W30 4 L PACK|PRO037|0.000000|0.00'
            }
            # Date-specific known native snapshot is only a fail-closed
            # acceptance guard; never used as calculated response values.
            if (
                $endDate.ToString('yyyy-MM-dd') -eq '2026-09-30' -and
                $asOfDate.ToString('yyyy-MM-dd') -eq '2026-09-30' -and
                $stockSep30RequestedMode -eq 'apply-verified'
            ) {
                $stockAllWaVerified = @{
                '1178' = '30/70 GREEN 2L|TOTACHI--44102|1.000000|2564.62'
                '1180' = 'ATF DEX III 20L|TOTACHI20720|45.000000|956852.00'
                '1181' = 'ATF DEX-III 1L|TOTACHI20701|168.000000|271563.87'
                '1182' = 'ATF DEX-III 4 L|TOTACHI20704|224.000000|1272334.33'
                '1183' = 'ATF MULTI VEHICLE LV 4L|TOTACHIA3004|172.000000|1485098.13'
                '1184' = 'ATF MULTI-VEHICLE 1L|TOTACHI20601|178.000000|448259.67'
                '1185' = 'ATF MULTI-VEHICLE 4L|TOTACHI20604|154.000000|1375589.52'
                '1186' = 'ATF SP IV 4L|TOTACHI21004|176.000000|1688687.04'
                '1187' = 'ATF SP-III 1 L|TOTACHI20401|150.000000|375277.97'
                '1188' = 'ATF SP-III 4 L|TOTACHI20404|88.000000|664794.51'
                '1189' = 'ATF SP-IV 1L|TOTACHI21001|422.000000|1263240.20'
                '1190' = 'ATF TYPE T-IV 1 L|TOTACHI20201|834.000000|1844406.99'
                '1191' = 'ATF TYPE T-IV 4 L|TOTACHI20204|1040.000000|6957355.23'
                '1192' = 'ATF WS  4L|TOTACHI20804|375.000000|3130770.97'
                '1193' = 'ATF WS 1L|TOTACHI20801|754.000000|2197270.35'
                '1194' = 'ATF Z-1 1L|TOTACHI20301|414.000000|1173603.06'
                '1195' = 'Concentrated Engine Power Flush 300 ML-T|TOTACHIP7130|403.000000|668648.65'
                '1197' = 'CVT MULTI-TYPE 1 L|TOTACHI20501|30.000000|103441.42'
                '1198' = 'CVT MULTI-TYPE 4 L|TOTACHI20504|223.000000|2800617.84'
                '1199' = 'CVT NS-3, BLUE 1 L|TOTACHI21101|985.000000|3467227.98'
                '1200' = 'CVT NS-3, BLUE 4 L|TOTACHI21104|46.000000|481366.12'
                '1201' = 'DCTF MULTI TYPE 4L|TOTACHIA8204|336.000000|4742864.06'
                '1202' = 'DCTF MULTI-TYPE 1 L|TOTACHIA8201|208.000000|738764.62'
                '1204' = 'DENTO ECO GASOLINE 15W-40 200 L|TOTACHI13922|14.000000|3997376.05'
                '1205' = 'DENTO ECO GASOLINE 15W-40 60 L|TOTACHI13960|0.000000|0.00'
                '1207' = 'DENTO HEAVY DUTY 15W40 CI-4 1 L|TOTACHI13201|1.000000|1818.47'
                '1208' = 'DENTO HEAVY DUTY 15W40 CI-4 4 L|TOTACHI13204|0.000000|0.00'
                '1209' = 'Diesel Cetane & Fuel System Boost 250 ML|TOTACHIP7025|88.000000|132412.85'
                '1210' = 'ECO DIESEL 10W-40 1 L|TOTACHIE1301|0.000000|0.00'
                '1211' = 'ECO DIESEL 10W-40 4 L|TOTACHIE1304|1.000000|6623.98'
                '1212' = 'FINE GASOLINE 10W-30 3L SURPRISE PACK|PRO004|0.000000|0.00'
                '1213' = 'ECO GASOLINE 10W-40 1 L|TOTACHI10901|1083.000000|3127454.54'
                '1214' = 'ECO GASOLINE 10W-40 4 L|TOTACHI10904|820.000000|5098056.55'
                '1215' = 'ECO GASOLINE 5W-30 1 L|TOTACHI10801|66.000000|111923.77'
                '1216' = 'ECO GASOLINE 5W30 4 L|TOTACHI10804|1.000000|6231.82'
                '1217' = 'EURODRIVE 5W40 4L|TOTACHIE6304|362.000000|3572856.41'
                '1218' = 'EURODRIVE ECO 5W30 4 L|TOTACHIE6604|301.000000|3328905.09'
                '1219' = 'EURODRIVE PRO LONG LIFE 5W30 4 L|TOTACHIE7804|0.000000|0.00'
                '1220' = 'Extended Life Coolant Blue 30/70 1 L|TOTACHIT8701|58.000000|52631.65'
                '1221' = 'Extended Life Coolant Green 30/70  1L|TOTACHIT8501|2912.000000|2598744.88'
                '1222' = 'Extended Life Coolant Green 30/70 4 L|TOTACHIT8504|104.000000|295243.31'
                '1223' = 'Extended Life Coolant Red 30/70 1 L|TOTACHIT7301|577.000000|423954.70'
                '1224' = 'Extended Life Coolant Red 30/70 4 L|TOTACHIT7304|19.000000|45665.40'
                '1225' = 'FINE GASOLINE 10W-30 1L|TOTACHI10201|0.000000|0.00'
                '1226' = 'FINE GASOLINE 10W-30 200 L|TOTACHI10222|5.000000|1177907.78'
                '1227' = 'FINE GASOLINE 10W-30 4L|TOTACHI10204|1629.000000|8873619.80'
                '1228' = 'FINE GASOLINE 10W-30 60 L|TOTACHI10260|0.000000|0.00'
                '1229' = 'FLUSH OUT 4 L-TOTACHI12304|TOTACHI12304|137.000000|591200.81'
                '1230' = 'Fuel Injector Cleaner 300 ML|TOTACHIP8830|5887.000000|10882126.33'
                '1231' = 'GREEN LONG LIFE COOLANT 30/70 2L|TOTACHI44102|0.000000|0.00'
                '1232' = 'HEAVY DUTY 15W40 1L|TOTACHI10701|119.000000|203598.19'
                '1233' = 'HEAVY DUTY 15W-40 200L|TOTACHI10722|0.000000|0.00'
                '1234' = 'HEAVY DUTY 15W-40 4 L|TOTACHI10704|233.000000|1073496.51'
                '1235' = 'HEAVY DUTY 15W40 60 L|TOTACHI10760|0.000000|0.00'
                '1236' = 'HEAVY DUTY 15W40 6L|TOTACHI10706|160.000000|1600257.75'
                '1237' = 'HEAVY DUTY 15W-40 7L SURPRISE PACK|PRO003|0.000000|0.00'
                '1238' = 'HEAVY DUTY 15W-40 6L SURPRISE PACK|PRO002|0.000000|0.00'
                '1240' = 'HYDRAULIC OIL NRO 68 - 20L|TOTACHI50620|0.000000|0.00'
                '1241' = 'HYPER ECODRIVE 0W20 1L|TOTACHIE0101|2.000000|4543.76'
                '1242' = 'HYPER ECODRIVE 0W-20 4L|TOTACHIE0104|2.000000|16811.89'
                '1243' = 'HYPER ECODRIVE 0W20 60 L|TOTACHIE0160|0.000000|0.00'
                '1244' = 'HYPER ECODRIVE 0W-20 3L SURPRISE PACK|PRO001|0.000000|0.00'
                '1245' = 'HYPER ECODRIVE 0W-40 1 L|TOTACHI12801|20.000000|52713.15'
                '1246' = 'HYPER ECODRIVE 0W-40 4 L|TOTACHI12804|17.000000|184260.30'
                '1247' = 'HYPER ULV 0W16,PURPLE 1L|TOTACHIE0001|1993.000000|4899769.13'
                '1248' = 'HYPER ULV 0W16,PURPLE 4L|TOTACHIE0004|457.000000|3342511.95'
                '1249' = 'MIX-TYPE COOLANT PINK -40C G12EV 1L|TOTACHI46801|1.000000|2711.86'
                '1250' = 'POWER STEERING FLUID 1L|TOTACHI90301|213.000000|486697.13'
                '1251' = 'RED LONG LIFE COOLANT -40C 2L|TOTACHI41802|0.000000|0.00'
                '1252' = 'SUPER LONG LIFE COOLAN GREEN 1L||0.000000|0.00'
                '1253' = 'SUPER LONG LIFE COOLANT -40C 1L|TOTACHI41601|13.000000|20243.63'
                '1254' = 'SUPER LONG LIFE COOLANT -40C 2L|TOTACHI44602|0.000000|0.00'
                '1255' = 'SUPER LONG LIFE COOLANT HYBRID-40 1L|TOTACHI44601|966.000000|1601677.17'
                '1260' = 'TOTACHI-COROLLA NZE121-TA-1334|TA-1334|380.000000|387185.98'
                '1261' = 'TOTACHI-ALLION/PREMIO 260 1NZ,2ZR COROLL|TA-1343|451.000000|368299.61'
                '1262' = 'TOTACHI-HIACE KDH-TA-1380|TA-1380|137.000000|376776.37'
                '1263' = 'TOTACHI-LEXUS-TA-1387|TA-1387|0.000000|0.00'
                '1264' = 'TOTACHI-LAND CRUISER UZJ120/150 2TR-TA-1|TA-1427|58.000000|121178.94'
                '1265' = 'TOTACHI-VITZ-TA-1430|TA-1430|103.000000|67467.62'
                '1266' = 'TOTACHI- SUNNY QG13,QG15,QG16,CD17 X-TRA|TA-1465|105.000000|126889.43'
                '1267' = 'TOTACHI X-TRAIL NT32-TA-1873|TA-1873|64.000000|57537.62'
                '1268' = 'TOTACHI-PRIUS III-TA-1877|TA-1877|236.000000|245235.18'
                '1269' = 'TOTACHI-HONDA FIT GP5-TA-1882|TA-1882|435.000000|398071.66'
                '1270' = 'TOTACHI-YARIS HYBRID,AQUA-TA-2843|TA-2843|491.000000|455030.16'
                '1271' = 'TOTACHI- FIT L13B(1300,1500 Cc) VEZEL L1|TA-3708|491.000000|640343.23'
                '1272' = 'TOTACHI - RANGER BT50-TA-3855|TA-3855|0.000000|0.00'
                '1274' = 'TOTACHI-LAND CRUISER 200 (A1533)-TA-3856|TA-3856|109.000000|226349.72'
                '1275' = 'TOTACHI C-HR NGX10,NGX50-TA-3859|TA-3859|111.000000|264229.95'
                '1277' = 'TOTACHI- HIACE,2L,3L-TC-1024|TC-1024|250.000000|826051.95'
                '1278' = 'TOTACHI-AQUA,AVANZA,COROLLA,IST,PRIUS (I|TC-1030|11525.000000|6448248.18'
                '1279' = 'TOTACHI-DYNA 1KD,PRADO 120,HIACE 1KD-TC-|TC-1031|3215.000000|2338838.27'
                '1280' = 'TOTACHI-LITEACE,TOWNACE,CR27-TC-1032|TC-1032|1299.000000|2633619.00'
                '1281' = 'TOTACHI-LAND CRUISER 100,120,CROWN,HILUX|TC-1034|200.000000|194744.32'
                '1282' = 'TOTACHI-COASTER 1HZ,1HD,HIACE,TOYOACE,HI|TC-1035|1469.000000|2483954.37'
                '1283' = 'TOTACHI - SUNNY B15,N16-TC-1047|TC-1047|383.000000|297206.06'
                '1285' = 'TOTACHI-PICK-UP,L200,L300,PAJERO-TC-1054|TC-1054|105.000000|221593.82'
                '1286' = 'TOTACHI-LANCER CS7-TC-1055|TC-1055|163.000000|275470.00'
                '1287' = 'TOTACHI- LANCER 4G15,L300,L400 DELICA-TC|TC-1065|399.000000|495250.77'
                '1288' = 'TOTACHI- FIT AREA L13 VEZEL L15-TC-1096|TC-1096|7026.000000|3180493.16'
                '1289' = 'TOTACHI-EVERY,MINICAB-TC-1097|TC-1097|380.000000|221300.38'
                '1292' = 'TOTACHI-NOAH,HILUX,KUN26-TCA-101|TCA-101|255.000000|531694.71'
                '1293' = 'TOTACHI-COROLLA NZE121 (AC101EX)-TCA-160|TCA-160|191.000000|336283.56'
                '1294' = 'TOTACHI-PRADO KDJ150,TRJ150-TCA-161|TCA-161|171.000000|125473.20'
                '1296' = 'TOTACHI-ALLION 260,BELTA SCP,NCP92,96-TC|TCA-166|471.000000|307373.82'
                '1297' = 'TOTACHI-ALLION260, BELTA SCP-TCA-166K|TCA-166K|6.000000|6113.07'
                '1298' = 'TOTACHI-SUNNY FB15-TCA-176|TCA-176|391.000000|920935.76'
                '1299' = 'TOTACHI-FIT JAZZ-TCA-229|TCA-229|105.000000|110054.92'
                '1300' = 'TOTACHI-VEZEL RU1-TCA-233|TCA-233|500.000000|529630.86'
                '1301' = 'TOTACHI-CIVIC,CRV-TCA-237|TCA-237|124.000000|118139.56'
                '1302' = 'TOTACHI X-TRAIL NT32 (AC2514)-TCA-566|TCA-566|292.000000|201992.55'
                '1303' = 'TOTACHI-HILUX VIGO,C-HR-TCA-806K|TCA-806K|555.000000|649350.14'
                '1304' = 'TOTACHI-MINI CAB-TCA-817|TCA-817|93.000000|101699.59'
                '1306' = 'TOTACHI HIACE,LITEACE,TOWNACE FILTER-TF-|TF-1016|844.000000|1016863.51'
                '1307' = 'TOTACHI TOWNACE,RAV 4,YARIS,LITEACE-TF-1|TF-1019|1150.000000|1966662.60'
                '1308' = 'TOTACHI - TF-1020-TF-1020|TF-1020|1125.000000|1047901.76'
                '1309' = 'TOTACHI-1KD,2KD,KDH TF-1090-TF-1090|TF-1090|177.000000|102247.17'
                '1310' = 'TOTACHI- PRADO 120,150-TF-1316|TF-1316|185.000000|190546.63'
                '1311' = 'TOTACHI-LAND CRUISER 200-TF-1731|TF-1731|115.000000|119997.59'
                '1312' = 'TOTACHI LAND CRUISER 200(0121)-TO-1162|TO-1162|42.000000|40290.16'
                '1313' = 'TOTACHI-HIACE TRH200-2TR,PRADO TRJ120-2T|TO-1189|42.000000|31402.88'
                '1314' = 'TOTACHI VITZ(1KR),PASSO,ALLION PREMIO ZR|TO-1190|3261.000000|1011714.17'
                '1315' = 'TOTACHI-RANGER BT50-TO-1299|TO-1299|36.000000|23606.66'
                '1316' = 'TOTACHI-JUKE,QASHQAL-TO-1672|TO-1672|0.000000|0.00'
                '1318' = 'TOTACHI-SWIFT,ALTO-TC-1100|TC-1100|7706.000000|4006323.98'
                '2193' = 'HEAVY DUTY 15W-40 8L SURPRISE PACK|PRO005|0.000000|0.00'
                '2409' = 'EXTENDED LIFE COOL RED -40C (50/50) 1L|TOTACHIT7201|3010.000000|3229381.95'
                '2410' = 'EXTENDED LIFE COOL RED -40C (50/50) 4L|TOTACHIT7204|219.000000|788202.32'
                '2411' = 'SUPER LONG LIFE COOLANT -37C (50/50) 1L|TOTACHIT7401|1655.000000|2275565.27'
                '2412' = 'SUPER LONG LIFE COOLANT -37C (50/50) 4L|TOTACHIT7404|793.000000|3791506.18'
                '2460' = 'HYPER ECODRIVE 0W-20 PROMO (25 CANS + 9|PRO006|0.000000|0.00'
                '2461' = 'HYPER ECODRIVE 0W-20 PROMO (50 CANS + 20|PRO007|0.000000|0.00'
                '2462' = 'HYPER ECODRIVE 0W-20 PROMO (100 CANS + 4|PRO008|0.000000|0.00'
                '2463' = 'FINE GASOLINE 10W-30 PROMO (50 CANS+20 C|PRO009|0.000000|0.00'
                '2464' = 'FINE GASOLINE 10W-30 4L WITH TC-1030|PRO010|0.000000|0.00'
                '2465' = 'FINE GASOLINE 10W-30 PROMO (100 CANS +42|PRO011|0.000000|0.00'
                '2466' = 'DENTO HEAVY DUTY 15W-40 CI-4 7L SURPRISE|PRO012|0.000000|0.00'
                '2467' = 'DENTO HEAVY DUTY 15W-40 4L PROMO (25 CAN|PRO013|0.000000|0.00'
                '2468' = 'DENTO HEAVY DUTY 15W-40 4L PROMO (50 CAN|PRO014|0.000000|0.00'
                '2469' = 'DENTO HEAVY DUTY 15W-40 4L PROMO (100 CA|PRO015|0.000000|0.00'
                '2470' = 'DENTO HEAVY DUTY 15W-40 4L PROMO (200 CA|PRO016|0.000000|0.00'
                '2471' = 'DENTO HD 15W-40 5L PROMO (25 CANS+10 FRE|PRO017|0.000000|0.00'
                '2472' = 'DENTO HEAVY DUTY 15W-40 5L PROMO (50 CAN|PRO018|0.000000|0.00'
                '2473' = 'DENTO HEAVY DUTY 15W-40 5L PROMO (100 CA|PRO019|0.000000|0.00'
                '2474' = 'DENTO HEAVY DUTY 15W-40 5L PROMO (200 CA|PRO020|0.000000|0.00'
                '2475' = 'FINE GASOLINE 10W-30 PROMO (30 CANS +11|PRO021|0.000000|0.00'
                '2511' = 'DENTO HEAVY DUTY 15W-40 6L SURPRISE PACK|PRO024|0.000000|0.00'
                '2512' = 'DENTO HEAVY DUTY 15W40 CI-4 5L|TOTACHI13205|1.000000|10911.61'
                '2558' = 'HEAVY DUTY 15W40 20L|TOTACHI10720|0.000000|0.00'
                '2586' = 'FINE GASOLINE 10W-30 PROMO (25 CANS +10|PRO022|0.000000|0.00'
                '2745' = 'FINE GASOLINE 10W-30 PROMO (50 CANS+22 C|PRO023|0.000000|0.00'
                '2748' = 'FINE GASOLINE 10W-30 3L PACK (25 + 9) FR|PRO025|0.000000|0.00'
                '2786' = 'FINE GASOLINE 10W-30 3L PACK (25 + 10) F|PRO026|0.000000|0.00'
                '2794' = 'FINE GASOLINE 10W-30 3L PACK (50 +21) FR|PRO027|0.000000|0.00'
                '2801' = 'FINE GASOLINE 10W-30 PROMO(100CANS+54 C|PRO033|0.000000|0.00'
                '2824' = 'FINE GASOLINE 10W-30 PROMO (50 CANS+24 C|PRO029|0.000000|0.00'
                '2997' = 'Extended Life Coolant Blue 30/70 4 L|TOTACHIT8704|485.000000|1093306.46'
                '3003' = 'DENTO PAWA 4T 20W-50 1L (3-WHEEL)|TOTACHIE4201|0.000000|0.00'
                '3004' = 'DENTO PAWA 4T 20W-50 1L|TOTACHI13601|0.000000|0.00'
                '3005' = 'ATF MULTI VEHICLE LV 1L|TOTACHIA3001|3640.000000|9298626.00'
                '3006' = 'EURODRIVE 5W-40 1L|TOTACHIE6301|213.000000|632758.81'
                '3007' = 'EURODRIVE ECO 5W-30 1 L|TOTACHIE6601|325.000000|1042466.27'
                '3008' = 'EURODRIVE PRO FUEL EFFICIENCY 5W-30 4L|TOTACHIE7904|540.000000|5245849.60'
                '3009' = 'HYPER ECODRIVE 10W-30 1L|TOTACHIE2601|205.000000|479184.74'
                '3010' = 'HYPER ECODRIVE 10W-30 4L|TOTACHIE2604|340.000000|2487237.63'
                '3011' = 'FINE GASOLINE 15W-40 4L (DON''T USE)||0.000000|0.00'
                '3012' = 'ULTIMA RACING 10W-60 1L|TOTACHIE5001|0.000000|0.00'
                '3075' = 'FINE GASOLINE 10W-30 PROMO (100 CANS +43|PRO030|0.000000|0.00'
                '3084' = 'FINE GASOLINE 10W-30 PROMO (25 CANS +9 F|PRO031|0.000000|0.00'
                '3090' = 'FINE GASOLINE 10W-30 PROMO (50 CANS+25 C|PRO032|0.000000|0.00'
                '3091' = 'FINE GASOLINE 10W-30 PROMO (25 CANS +11|PRO034|0.000000|0.00'
                '3134' = 'HYDRAULIC OIL NRO 68-200L|TOTACHI50622|0.000000|0.00'
                '3135' = 'HYDRAULIC OIL NRO 46-20L|TOTACHI50520|54.000000|925655.58'
                '3150' = 'FINE GASOLINE 10W-30 4L PROMO (200+110)|PRO035|0.000000|0.00'
                '3152' = 'FINE GASOLINE 10W-30 4L FREE ISSUE|10W30-4L-FREE ISSUE|0.000000|0.00'
                '3190' = 'FINE GASOLINE 15W40 SP 1L|TOTACHIE1901|0.000000|0.00'
                '3191' = 'FINE GASOLINE 15W40 SP 4L|TOTACHIE1904|3.000000|20067.84'
                '3192' = 'FINE GASOLINE 15W40 SP 60L|TOTACHIE1960|46.000000|3396907.02'
                '3193' = 'FINE GASOLINE 15W40 SP 200L|TOTACHIE192Z|0.000000|0.00'
                '3313' = 'HYPER ECODRIVE 0W20 3L|TOTACHIE0103|3.000000|22271.72'
                '3314' = 'FINE GASOLINE 10W30 SP 3L|TOTACHIE1203|1.000000|6777.31'
                '3315' = 'FINE GASOLINE 15W40 SP 3L|TOTACHIE1903|21.000000|99816.44'
                '3322' = 'HYPER ECODRIVE 5W-30 4L|TOTACHIE0304|2.000000|15468.76'
                '3405' = 'DENTO GL-5 80W-90 GEAR OIL 1 L|TOTACHIG5701|649.000000|1076966.99'
                '3422' = 'SUPER LONG LIFE COOLANT GREEN -40C 2L|TOTACHI41602|0.000000|0.00'
                '3519' = 'TA-1418 AIT FILTER|TA-1418|290.000000|294309.40'
                '3520' = 'TA-1379 AIR FILTER|TA1379|219.000000|254473.62'
                '3521' = 'TA-1624 AIR FILTER|TA-1624|6.000000|7132.63'
                '3553' = 'DENTO BRAKE FLUID DOT 4 - 1L|TOTACHIZ9901|877.000000|2281864.22'
                '3583' = 'C525J OIL FILTER|C525J|0.000000|0.00'
                '3584' = 'FC607 FUEL FILTER|FC607|0.000000|0.00'
                '3585' = 'C526J OIL FILTER|C526J|0.000000|0.00'
                '3710' = 'OE23010 OIL FILTER|OE23010|0.000000|0.00'
                '3761' = 'PREMEIUM DIESEL 5W-40 1L|TOTACHI11701|7.000000|0.00'
                '3762' = 'PREMIUM DIESEL 5W-40 4L|TOTACHI11704|1.000000|0.00'
                '3765' = 'ECO GASOLINE 5W-30 3L|TOTACHI10803|2.000000|9374.59'
                '3769' = 'FUEL PUMP - FREE ISSUE||0.000000|0.00'
                '3771' = 'ATF TYPE T-IV 20 L|TOTACHI20220|7.000000|252917.25'
                '3772' = 'EXTENDED LIFE COOLANT PINK -40C(50/50)4L|TOTACHIT0104|58.000000|227933.75'
                '3773' = 'EXTENDED LIFE COOLANT PINK -40C(50/50)1L|TOTACHIT0101|5257.000000|6138940.82'
                '3804' = 'DISC BRAKE PADS|TBM-2040|0.000000|0.00'
                '3891' = 'DENTO PAWA 4T 10W-40 1L|TOTACHIE3701|2121.000000|2689970.87'
                '3926' = 'FINE GASOLINE 15W-40 4L (Dont Use)|TOTACHIE9704|0.000000|0.00'
                '3957' = 'FINE GASOLINE 20W50 1L|TOTACHI10601|65.000000|129429.62'
                '3958' = 'FINE GASOLINE 20W50 4L|TOTACHI10604|433.000000|2188459.29'
                '3959' = 'ULTIMA RACING 5W50 4L|TOTACHIE4804|59.000000|859870.13'
                '3960' = 'ULTIMA RACING 10W60 4L|TOTACHIE5004|106.000000|1502349.23'
                '3961' = 'DENTO HEAVY DUTY DIESEL 20W50 200L|TOTACHIE3520|36.000000|8010196.38'
                '3984' = 'DENTO HYPOID 80W-90 GL-5  20L|TOTACHIG5720|0.000000|0.00'
                '3985' = 'DENTO HYPOID 80W-90 GL-5  200L|TOTACHIG572Z|0.000000|0.00'
                '4060' = 'HEAVY DUTY DIESEL 20W50 20L|TOTACHI18820|69.000000|1618975.98'
                '4061' = 'HEAVY DUTY DIESEL 20W50 60L|TOTACHI18860|4.000000|279636.16'
                '4062' = 'HYPER ULV 0W8 4L|TOTACHIE5504|1101.000000|8867582.55'
                '4063' = 'HYPER ULV 0W8 1L|TOTACHIE5501|450.000000|1136857.83'
                '4084' = 'HYPER ECODRIVE 0W20 200 L|TOTACHIE012Z|0.000000|0.00'
                '4108' = 'HYBRID WIPER BLADE WBH-350 14"|WBH00710|0.000000|0.00'
                '4110' = 'HYBRID WIPER BLADE WBH-400 16"|WBH00160|628.000000|798888.35'
                '4111' = 'HYBRID WIPER BLADE WBH-450 18"|WBH00260|315.000000|414137.00'
                '4112' = 'HYBRID WIPER BLADE WBH-480 19"|WBH00810|40.000000|53100.90'
                '4113' = 'HYBRID WIPER BLADE WBH-500 20"|WBH00360|138.000000|205242.37'
                '4114' = 'HYBRID WIPER BLADE WBH-530 21"|WBH00910|124.000000|207411.27'
                '4115' = 'HYBRID WIPER BLADE WBH-550 22"|WBH00460|441.000000|686315.65'
                '4116' = 'HYBRID WIPER BLADE WBH-600 24"|WBH00560|1702.000000|2855140.67'
                '4117' = 'HYBRID WIPER BLADE WBH-650 26"|WBH00660|801.000000|1395017.20'
                '4180' = 'ECO DIESEL 15W-40 1 L|TOTACHIE5601|329.000000|673089.68'
                '4181' = 'DENTO HEAVY DUTY DIESEL 15W40  200L|TOTACHIE342Z|0.000000|0.00'
                '4209' = 'C110 - SRR OIL FILTER|C110|0.000000|0.00'
                '4227' = 'FINE GASOLINE 15W 40 SP 4L PACK|PRO036|0.000000|0.00'
                '4231' = 'HYPER ECODRIVE 0W20 4L PACK|PRO040|0.000000|0.00'
                '4253' = 'EURODRIVE ECO 0W20 4 L|TOTACHIE6404|114.000000|1099615.42'
                '4254' = 'ECO DIESEL 5W-30 1 L|TOTACHIE2201|132.000000|265419.28'
                '4255' = 'ECO DIESEL 5W-30 6 L|TOTACHIE2206|85.000000|893943.76'
                '4256' = 'DENTO HEAVY DUTY DIESEL 15W40  25 L|TOTACHIE341Z|0.000000|0.00'
                '4257' = 'DENTO HEAVY DUTY DIESEL 20W50  25 L|TOTACHIE351Z|15.000000|442031.74'
                '4397' = 'ECO DIESEL 5W30 200 L|TOTACHIE222Z|3.000000|1139604.80'
                '4426' = 'ECO GASOLINE 5W30 4 L PACK|PRO037|0.000000|0.00'
                }
            }
            # August 31 BUSY references are acceptance checks only; all
            # returned stock values are computed from transactional replay.
            # Omitted native items are explicitly marked as inferred zeros
            # in evidence/Aug31_Native_AllMC_235.csv.
            if (
                $endDate.ToString('yyyy-MM-dd') -eq '2026-08-31' -and
                $asOfDate.ToString('yyyy-MM-dd') -eq '2026-08-31' -and
                $stockAug31RequestedMode -eq 'apply-verified'
            ) {
                $stockAllWaVerified = @{
                '1178' = '30/70 GREEN 2L|TOTACHI--44102|1.000000|2564.62'
                '1180' = 'ATF DEX III 20L|TOTACHI20720|48.000000|1020642.14'
                '1181' = 'ATF DEX-III 1L|TOTACHI20701|289.000000|465473.36'
                '1182' = 'ATF DEX-III 4 L|TOTACHI20704|312.000000|1763619.00'
                '1183' = 'ATF MULTI VEHICLE LV 4L|TOTACHIA3004|185.000000|1597343.91'
                '1184' = 'ATF MULTI-VEHICLE 1L|TOTACHI20601|201.000000|502851.72'
                '1185' = 'ATF MULTI-VEHICLE 4L|TOTACHI20604|201.000000|1784821.13'
                '1186' = 'ATF SP IV 4L|TOTACHI21004|178.000000|1707876.66'
                '1187' = 'ATF SP-III 1 L|TOTACHI20401|149.000000|369133.90'
                '1188' = 'ATF SP-III 4 L|TOTACHI20404|88.000000|653230.35'
                '1189' = 'ATF SP-IV 1L|TOTACHI21001|423.000000|1266233.66'
                '1190' = 'ATF TYPE T-IV 1 L|TOTACHI20201|858.000000|1895454.83'
                '1191' = 'ATF TYPE T-IV 4 L|TOTACHI20204|1059.000000|7075438.45'
                '1192' = 'ATF WS  4L|TOTACHI20804|412.000000|3427259.72'
                '1193' = 'ATF WS 1L|TOTACHI20801|754.000000|2193722.64'
                '1194' = 'ATF Z-1 1L|TOTACHI20301|416.000000|1179272.64'
                '1195' = 'Concentrated Engine Power Flush 300 ML-T|TOTACHIP7130|402.000000|664835.09'
                '1197' = 'CVT MULTI-TYPE 1 L|TOTACHI20501|92.000000|312704.43'
                '1198' = 'CVT MULTI-TYPE 4 L|TOTACHI20504|415.000000|5200000.58'
                '1199' = 'CVT NS-3, BLUE 1 L|TOTACHI21101|986.000000|3468649.40'
                '1200' = 'CVT NS-3, BLUE 4 L|TOTACHI21104|57.000000|584482.25'
                '1201' = 'DCTF MULTI TYPE 4L|TOTACHIA8204|342.000000|4827558.07'
                '1202' = 'DCTF MULTI-TYPE 1 L|TOTACHIA8201|210.000000|745868.12'
                '1204' = 'DENTO ECO GASOLINE 15W-40 200 L|TOTACHI13922|16.000000|4568429.77'
                '1205' = 'DENTO ECO GASOLINE 15W-40 60 L|TOTACHI13960|0.000000|0.00'
                '1207' = 'DENTO HEAVY DUTY 15W40 CI-4 1 L|TOTACHI13201|1.000000|1818.47'
                '1208' = 'DENTO HEAVY DUTY 15W40 CI-4 4 L|TOTACHI13204|0.000000|0.00'
                '1209' = 'Diesel Cetane & Fuel System Boost 250 ML|TOTACHIP7025|102.000000|151678.14'
                '1210' = 'ECO DIESEL 10W-40 1 L|TOTACHIE1301|0.000000|0.00'
                '1211' = 'ECO DIESEL 10W-40 4 L|TOTACHIE1304|1.000000|6623.98'
                '1212' = 'FINE GASOLINE 10W-30 3L SURPRISE PACK|PRO004|0.000000|0.00'
                '1213' = 'ECO GASOLINE 10W-40 1 L|TOTACHI10901|1087.000000|3137995.08'
                '1214' = 'ECO GASOLINE 10W-40 4 L|TOTACHI10904|934.000000|5801139.04'
                '1215' = 'ECO GASOLINE 5W-30 1 L|TOTACHI10801|71.000000|120402.84'
                '1216' = 'ECO GASOLINE 5W30 4 L|TOTACHI10804|2.000000|12463.64'
                '1217' = 'EURODRIVE 5W40 4L|TOTACHIE6304|364.000000|3577041.99'
                '1218' = 'EURODRIVE ECO 5W30 4 L|TOTACHIE6604|312.000000|3450559.43'
                '1219' = 'EURODRIVE PRO LONG LIFE 5W30 4 L|TOTACHIE7804|0.000000|0.00'
                '1220' = 'Extended Life Coolant Blue 30/70 1 L|TOTACHIT8701|70.000000|62674.15'
                '1221' = 'Extended Life Coolant Green 30/70  1L|TOTACHIT8501|2980.000000|2658561.06'
                '1222' = 'Extended Life Coolant Green 30/70 4 L|TOTACHIT8504|122.000000|343537.92'
                '1223' = 'Extended Life Coolant Red 30/70 1 L|TOTACHIT7301|549.000000|402363.88'
                '1224' = 'Extended Life Coolant Red 30/70 4 L|TOTACHIT7304|46.000000|106173.99'
                '1225' = 'FINE GASOLINE 10W-30 1L|TOTACHI10201|0.000000|0.00'
                '1226' = 'FINE GASOLINE 10W-30 200 L|TOTACHI10222|12.000000|2826978.64'
                '1227' = 'FINE GASOLINE 10W-30 4L|TOTACHI10204|2456.000000|13370624.67'
                '1228' = 'FINE GASOLINE 10W-30 60 L|TOTACHI10260|6.000000|478480.28'
                '1229' = 'FLUSH OUT 4 L-TOTACHI12304|TOTACHI12304|136.000000|582810.98'
                '1230' = 'Fuel Injector Cleaner 300 ML|TOTACHIP8830|6003.000000|11096552.46'
                '1231' = 'GREEN LONG LIFE COOLANT 30/70 2L|TOTACHI44102|0.000000|0.00'
                '1232' = 'HEAVY DUTY 15W40 1L|TOTACHI10701|266.000000|455101.84'
                '1233' = 'HEAVY DUTY 15W-40 200L|TOTACHI10722|0.000000|0.00'
                '1234' = 'HEAVY DUTY 15W-40 4 L|TOTACHI10704|288.000000|1326896.95'
                '1235' = 'HEAVY DUTY 15W40 60 L|TOTACHI10760|0.000000|0.00'
                '1236' = 'HEAVY DUTY 15W40 6L|TOTACHI10706|430.000000|4300692.62'
                '1237' = 'HEAVY DUTY 15W-40 7L SURPRISE PACK|PRO003|0.000000|0.00'
                '1238' = 'HEAVY DUTY 15W-40 6L SURPRISE PACK|PRO002|0.000000|0.00'
                '1240' = 'HYDRAULIC OIL NRO 68 - 20L|TOTACHI50620|0.000000|0.00'
                '1241' = 'HYPER ECODRIVE 0W20 1L|TOTACHIE0101|2.000000|4543.76'
                '1242' = 'HYPER ECODRIVE 0W-20 4L|TOTACHIE0104|458.000000|3849923.53'
                '1243' = 'HYPER ECODRIVE 0W20 60 L|TOTACHIE0160|9.000000|1070169.65'
                '1244' = 'HYPER ECODRIVE 0W-20 3L SURPRISE PACK|PRO001|0.000000|0.00'
                '1245' = 'HYPER ECODRIVE 0W-40 1 L|TOTACHI12801|20.000000|52713.15'
                '1246' = 'HYPER ECODRIVE 0W-40 4 L|TOTACHI12804|17.000000|184260.30'
                '1247' = 'HYPER ULV 0W16,PURPLE 1L|TOTACHIE0001|1993.000000|4899769.13'
                '1248' = 'HYPER ULV 0W16,PURPLE 4L|TOTACHIE0004|547.000000|3992279.40'
                '1249' = 'MIX-TYPE COOLANT PINK -40C G12EV 1L|TOTACHI46801|1.000000|2711.86'
                '1250' = 'POWER STEERING FLUID 1L|TOTACHI90301|276.000000|629197.49'
                '1251' = 'RED LONG LIFE COOLANT -40C 2L|TOTACHI41802|0.000000|0.00'
                '1252' = 'SUPER LONG LIFE COOLAN GREEN 1L||0.000000|0.00'
                '1253' = 'SUPER LONG LIFE COOLANT -40C 1L|TOTACHI41601|13.000000|20243.63'
                '1254' = 'SUPER LONG LIFE COOLANT -40C 2L|TOTACHI44602|0.000000|0.00'
                '1255' = 'SUPER LONG LIFE COOLANT HYBRID-40 1L|TOTACHI44601|966.000000|1601677.17'
                '1260' = 'TOTACHI-COROLLA NZE121-TA-1334|TA-1334|380.000000|387185.98'
                '1261' = 'TOTACHI-ALLION/PREMIO 260 1NZ,2ZR COROLL|TA-1343|0.000000|0.00'
                '1262' = 'TOTACHI-HIACE KDH-TA-1380|TA-1380|137.000000|376776.37'
                '1263' = 'TOTACHI-LEXUS-TA-1387|TA-1387|0.000000|0.00'
                '1264' = 'TOTACHI-LAND CRUISER UZJ120/150 2TR-TA-1|TA-1427|67.000000|139360.00'
                '1265' = 'TOTACHI-VITZ-TA-1430|TA-1430|103.000000|67467.62'
                '1266' = 'TOTACHI- SUNNY QG13,QG15,QG16,CD17 X-TRA|TA-1465|4.000000|6575.66'
                '1267' = 'TOTACHI X-TRAIL NT32-TA-1873|TA-1873|64.000000|57537.62'
                '1268' = 'TOTACHI-PRIUS III-TA-1877|TA-1877|240.000000|246659.65'
                '1269' = 'TOTACHI-HONDA FIT GP5-TA-1882|TA-1882|434.000000|395550.47'
                '1270' = 'TOTACHI-YARIS HYBRID,AQUA-TA-2843|TA-2843|1.000000|1013.45'
                '1271' = 'TOTACHI- FIT L13B(1300,1500 Cc) VEZEL L1|TA-3708|1.000000|2288.14'
                '1272' = 'TOTACHI - RANGER BT50-TA-3855|TA-3855|0.000000|0.00'
                '1274' = 'TOTACHI-LAND CRUISER 200 (A1533)-TA-3856|TA-3856|109.000000|226349.72'
                '1275' = 'TOTACHI C-HR NGX10,NGX50-TA-3859|TA-3859|115.000000|273751.75'
                '1277' = 'TOTACHI- HIACE,2L,3L-TC-1024|TC-1024|253.000000|835964.57'
                '1278' = 'TOTACHI-AQUA,AVANZA,COROLLA,IST,PRIUS (I|TC-1030|4550.000000|2342147.25'
                '1279' = 'TOTACHI-DYNA 1KD,PRADO 120,HIACE 1KD-TC-|TC-1031|3215.000000|2338838.27'
                '1280' = 'TOTACHI-LITEACE,TOWNACE,CR27-TC-1032|TC-1032|1299.000000|2633619.00'
                '1281' = 'TOTACHI-LAND CRUISER 100,120,CROWN,HILUX|TC-1034|0.000000|0.00'
                '1282' = 'TOTACHI-COASTER 1HZ,1HD,HIACE,TOYOACE,HI|TC-1035|1469.000000|2483954.37'
                '1283' = 'TOTACHI - SUNNY B15,N16-TC-1047|TC-1047|382.000000|295680.64'
                '1285' = 'TOTACHI-PICK-UP,L200,L300,PAJERO-TC-1054|TC-1054|105.000000|221593.82'
                '1286' = 'TOTACHI-LANCER CS7-TC-1055|TC-1055|163.000000|275470.00'
                '1287' = 'TOTACHI- LANCER 4G15,L300,L400 DELICA-TC|TC-1065|399.000000|495250.77'
                '1288' = 'TOTACHI- FIT AREA L13 VEZEL L15-TC-1096|TC-1096|7086.000000|3206563.48'
                '1289' = 'TOTACHI-EVERY,MINICAB-TC-1097|TC-1097|387.000000|225376.97'
                '1292' = 'TOTACHI-NOAH,HILUX,KUN26-TCA-101|TCA-101|260.000000|542120.10'
                '1293' = 'TOTACHI-COROLLA NZE121 (AC101EX)-TCA-160|TCA-160|190.000000|334003.90'
                '1294' = 'TOTACHI-PRADO KDJ150,TRJ150-TCA-161|TCA-161|11.000000|7332.77'
                '1296' = 'TOTACHI-ALLION 260,BELTA SCP,NCP92,96-TC|TCA-166|1.000000|659.98'
                '1297' = 'TOTACHI-ALLION260, BELTA SCP-TCA-166K|TCA-166K|6.000000|6113.07'
                '1298' = 'TOTACHI-SUNNY FB15-TCA-176|TCA-176|390.000000|917876.44'
                '1299' = 'TOTACHI-FIT JAZZ-TCA-229|TCA-229|5.000000|8450.00'
                '1300' = 'TOTACHI-VEZEL RU1-TCA-233|TCA-233|0.000000|0.00'
                '1301' = 'TOTACHI-CIVIC,CRV-TCA-237|TCA-237|24.000000|13774.16'
                '1302' = 'TOTACHI X-TRAIL NT32 (AC2514)-TCA-566|TCA-566|294.000000|202245.79'
                '1303' = 'TOTACHI-HILUX VIGO,C-HR-TCA-806K|TCA-806K|555.000000|649350.14'
                '1304' = 'TOTACHI-MINI CAB-TCA-817|TCA-817|92.000000|100301.28'
                '1306' = 'TOTACHI HIACE,LITEACE,TOWNACE FILTER-TF-|TF-1016|844.000000|1016863.51'
                '1307' = 'TOTACHI TOWNACE,RAV 4,YARIS,LITEACE-TF-1|TF-1019|1150.000000|1966662.60'
                '1308' = 'TOTACHI - TF-1020-TF-1020|TF-1020|1125.000000|1047901.76'
                '1309' = 'TOTACHI-1KD,2KD,KDH TF-1090-TF-1090|TF-1090|179.000000|103402.51'
                '1310' = 'TOTACHI- PRADO 120,150-TF-1316|TF-1316|185.000000|190546.63'
                '1311' = 'TOTACHI-LAND CRUISER 200-TF-1731|TF-1731|115.000000|119997.59'
                '1312' = 'TOTACHI LAND CRUISER 200(0121)-TO-1162|TO-1162|41.000000|38383.38'
                '1313' = 'TOTACHI-HIACE TRH200-2TR,PRADO TRJ120-2T|TO-1189|2.000000|2529.66'
                '1314' = 'TOTACHI VITZ(1KR),PASSO,ALLION PREMIO ZR|TO-1190|3335.000000|1033584.36'
                '1315' = 'TOTACHI-RANGER BT50-TO-1299|TO-1299|36.000000|23606.66'
                '1316' = 'TOTACHI-JUKE,QASHQAL-TO-1672|TO-1672|0.000000|0.00'
                '1318' = 'TOTACHI-SWIFT,ALTO-TC-1100|TC-1100|7890.000000|4101043.91'
                '2193' = 'HEAVY DUTY 15W-40 8L SURPRISE PACK|PRO005|0.000000|0.00'
                '2409' = 'EXTENDED LIFE COOL RED -40C (50/50) 1L|TOTACHIT7201|3009.000000|3227220.93'
                '2410' = 'EXTENDED LIFE COOL RED -40C (50/50) 4L|TOTACHIT7204|226.000000|810639.16'
                '2411' = 'SUPER LONG LIFE COOLANT -37C (50/50) 1L|TOTACHIT7401|1710.000000|2350105.61'
                '2412' = 'SUPER LONG LIFE COOLANT -37C (50/50) 4L|TOTACHIT7404|814.000000|3888642.14'
                '2460' = 'HYPER ECODRIVE 0W-20 PROMO (25 CANS + 9|PRO006|0.000000|0.00'
                '2461' = 'HYPER ECODRIVE 0W-20 PROMO (50 CANS + 20|PRO007|0.000000|0.00'
                '2462' = 'HYPER ECODRIVE 0W-20 PROMO (100 CANS + 4|PRO008|0.000000|0.00'
                '2463' = 'FINE GASOLINE 10W-30 PROMO (50 CANS+20 C|PRO009|0.000000|0.00'
                '2464' = 'FINE GASOLINE 10W-30 4L WITH TC-1030|PRO010|0.000000|0.00'
                '2465' = 'FINE GASOLINE 10W-30 PROMO (100 CANS +42|PRO011|0.000000|0.00'
                '2466' = 'DENTO HEAVY DUTY 15W-40 CI-4 7L SURPRISE|PRO012|0.000000|0.00'
                '2467' = 'DENTO HEAVY DUTY 15W-40 4L PROMO (25 CAN|PRO013|0.000000|0.00'
                '2468' = 'DENTO HEAVY DUTY 15W-40 4L PROMO (50 CAN|PRO014|0.000000|0.00'
                '2469' = 'DENTO HEAVY DUTY 15W-40 4L PROMO (100 CA|PRO015|0.000000|0.00'
                '2470' = 'DENTO HEAVY DUTY 15W-40 4L PROMO (200 CA|PRO016|0.000000|0.00'
                '2471' = 'DENTO HD 15W-40 5L PROMO (25 CANS+10 FRE|PRO017|0.000000|0.00'
                '2472' = 'DENTO HEAVY DUTY 15W-40 5L PROMO (50 CAN|PRO018|0.000000|0.00'
                '2473' = 'DENTO HEAVY DUTY 15W-40 5L PROMO (100 CA|PRO019|0.000000|0.00'
                '2474' = 'DENTO HEAVY DUTY 15W-40 5L PROMO (200 CA|PRO020|0.000000|0.00'
                '2475' = 'FINE GASOLINE 10W-30 PROMO (30 CANS +11|PRO021|0.000000|0.00'
                '2511' = 'DENTO HEAVY DUTY 15W-40 6L SURPRISE PACK|PRO024|0.000000|0.00'
                '2512' = 'DENTO HEAVY DUTY 15W40 CI-4 5L|TOTACHI13205|1.000000|7141.02'
                '2558' = 'HEAVY DUTY 15W40 20L|TOTACHI10720|0.000000|0.00'
                '2586' = 'FINE GASOLINE 10W-30 PROMO (25 CANS +10|PRO022|0.000000|0.00'
                '2745' = 'FINE GASOLINE 10W-30 PROMO (50 CANS+22 C|PRO023|0.000000|0.00'
                '2748' = 'FINE GASOLINE 10W-30 3L PACK (25 + 9) FR|PRO025|0.000000|0.00'
                '2786' = 'FINE GASOLINE 10W-30 3L PACK (25 + 10) F|PRO026|0.000000|0.00'
                '2794' = 'FINE GASOLINE 10W-30 3L PACK (50 +21) FR|PRO027|0.000000|0.00'
                '2801' = 'FINE GASOLINE 10W-30 PROMO(100CANS+54 C|PRO033|0.000000|0.00'
                '2824' = 'FINE GASOLINE 10W-30 PROMO (50 CANS+24 C|PRO029|0.000000|0.00'
                '2997' = 'Extended Life Coolant Blue 30/70 4 L|TOTACHIT8704|551.000000|1238696.47'
                '3003' = 'DENTO PAWA 4T 20W-50 1L (3-WHEEL)|TOTACHIE4201|0.000000|0.00'
                '3004' = 'DENTO PAWA 4T 20W-50 1L|TOTACHI13601|0.000000|0.00'
                '3005' = 'ATF MULTI VEHICLE LV 1L|TOTACHIA3001|3657.000000|9342053.66'
                '3006' = 'EURODRIVE 5W-40 1L|TOTACHIE6301|213.000000|632758.81'
                '3007' = 'EURODRIVE ECO 5W-30 1 L|TOTACHIE6601|324.000000|1034415.42'
                '3008' = 'EURODRIVE PRO FUEL EFFICIENCY 5W-30 4L|TOTACHIE7904|568.000000|5502571.16'
                '3009' = 'HYPER ECODRIVE 10W-30 1L|TOTACHIE2601|205.000000|479184.74'
                '3010' = 'HYPER ECODRIVE 10W-30 4L|TOTACHIE2604|372.000000|2713476.50'
                '3011' = 'FINE GASOLINE 15W-40 4L (DON''T USE)||0.000000|0.00'
                '3012' = 'ULTIMA RACING 10W-60 1L|TOTACHIE5001|0.000000|0.00'
                '3075' = 'FINE GASOLINE 10W-30 PROMO (100 CANS +43|PRO030|0.000000|0.00'
                '3084' = 'FINE GASOLINE 10W-30 PROMO (25 CANS +9 F|PRO031|0.000000|0.00'
                '3090' = 'FINE GASOLINE 10W-30 PROMO (50 CANS+25 C|PRO032|0.000000|0.00'
                '3091' = 'FINE GASOLINE 10W-30 PROMO (25 CANS +11|PRO034|0.000000|0.00'
                '3134' = 'HYDRAULIC OIL NRO 68-200L|TOTACHI50622|0.000000|0.00'
                '3135' = 'HYDRAULIC OIL NRO 46-20L|TOTACHI50520|75.000000|1285632.75'
                '3150' = 'FINE GASOLINE 10W-30 4L PROMO (200+110)|PRO035|0.000000|0.00'
                '3152' = 'FINE GASOLINE 10W-30 4L FREE ISSUE|10W30-4L-FREE ISSUE|0.000000|0.00'
                '3190' = 'FINE GASOLINE 15W40 SP 1L|TOTACHIE1901|44.000000|85434.98'
                '3191' = 'FINE GASOLINE 15W40 SP 4L|TOTACHIE1904|53.000000|341322.19'
                '3192' = 'FINE GASOLINE 15W40 SP 60L|TOTACHIE1960|49.000000|3618444.42'
                '3193' = 'FINE GASOLINE 15W40 SP 200L|TOTACHIE192Z|0.000000|0.00'
                '3313' = 'HYPER ECODRIVE 0W20 3L|TOTACHIE0103|6.000000|44543.41'
                '3314' = 'FINE GASOLINE 10W30 SP 3L|TOTACHIE1203|2.000000|13554.63'
                '3315' = 'FINE GASOLINE 15W40 SP 3L|TOTACHIE1903|52.000000|247164.51'
                '3322' = 'HYPER ECODRIVE 5W-30 4L|TOTACHIE0304|7.000000|54140.70'
                '3405' = 'DENTO GL-5 80W-90 GEAR OIL 1 L|TOTACHIG5701|661.000000|1096880.09'
                '3422' = 'SUPER LONG LIFE COOLANT GREEN -40C 2L|TOTACHI41602|0.000000|0.00'
                '3519' = 'TA-1418 AIT FILTER|TA-1418|290.000000|294309.40'
                '3520' = 'TA-1379 AIR FILTER|TA1379|219.000000|254473.62'
                '3521' = 'TA-1624 AIR FILTER|TA-1624|6.000000|7132.63'
                '3553' = 'DENTO BRAKE FLUID DOT 4 - 1L|TOTACHIZ9901|881.000000|2290980.94'
                '3583' = 'C525J OIL FILTER|C525J|0.000000|0.00'
                '3584' = 'FC607 FUEL FILTER|FC607|0.000000|0.00'
                '3585' = 'C526J OIL FILTER|C526J|0.000000|0.00'
                '3710' = 'OE23010 OIL FILTER|OE23010|0.000000|0.00'
                '3761' = 'PREMEIUM DIESEL 5W-40 1L|TOTACHI11701|7.000000|0.00'
                '3762' = 'PREMIUM DIESEL 5W-40 4L|TOTACHI11704|1.000000|0.00'
                '3765' = 'ECO GASOLINE 5W-30 3L|TOTACHI10803|4.000000|18749.18'
                '3769' = 'FUEL PUMP - FREE ISSUE||0.000000|0.00'
                '3771' = 'ATF TYPE T-IV 20 L|TOTACHI20220|7.000000|252917.25'
                '3772' = 'EXTENDED LIFE COOLANT PINK -40C(50/50)4L|TOTACHIT0104|167.000000|653678.16'
                '3773' = 'EXTENDED LIFE COOLANT PINK -40C(50/50)1L|TOTACHIT0101|5277.000000|6161300.61'
                '3804' = 'DISC BRAKE PADS|TBM-2040|0.000000|0.00'
                '3891' = 'DENTO PAWA 4T 10W-40 1L|TOTACHIE3701|3913.000000|4962685.58'
                '3926' = 'FINE GASOLINE 15W-40 4L (Dont Use)|TOTACHIE9704|0.000000|0.00'
                '3957' = 'FINE GASOLINE 20W50 1L|TOTACHI10601|76.000000|151333.08'
                '3958' = 'FINE GASOLINE 20W50 4L|TOTACHI10604|459.000000|2319867.95'
                '3959' = 'ULTIMA RACING 5W50 4L|TOTACHIE4804|60.000000|874444.20'
                '3960' = 'ULTIMA RACING 10W60 4L|TOTACHIE5004|108.000000|1503343.14'
                '3961' = 'DENTO HEAVY DUTY DIESEL 20W50 200L|TOTACHIE3520|47.000000|10457756.38'
                '3984' = 'DENTO HYPOID 80W-90 GL-5  20L|TOTACHIG5720|0.000000|0.00'
                '3985' = 'DENTO HYPOID 80W-90 GL-5  200L|TOTACHIG572Z|0.000000|0.00'
                '4060' = 'HEAVY DUTY DIESEL 20W50 20L|TOTACHI18820|70.000000|1642439.40'
                '4061' = 'HEAVY DUTY DIESEL 20W50 60L|TOTACHI18860|4.000000|279636.16'
                '4062' = 'HYPER ULV 0W8 4L|TOTACHIE5504|1103.000000|8875982.19'
                '4063' = 'HYPER ULV 0W8 1L|TOTACHIE5501|450.000000|1136857.83'
                '4084' = 'HYPER ECODRIVE 0W20 200 L|TOTACHIE012Z|1.000000|401989.64'
                '4108' = 'HYBRID WIPER BLADE WBH-350 14"|WBH00710|0.000000|0.00'
                '4110' = 'HYBRID WIPER BLADE WBH-400 16"|WBH00160|690.000000|877759.50'
                '4111' = 'HYBRID WIPER BLADE WBH-450 18"|WBH00260|390.000000|512741.04'
                '4112' = 'HYBRID WIPER BLADE WBH-480 19"|WBH00810|65.000000|86288.98'
                '4113' = 'HYBRID WIPER BLADE WBH-500 20"|WBH00360|162.000000|240936.69'
                '4114' = 'HYBRID WIPER BLADE WBH-530 21"|WBH00910|134.000000|224137.99'
                '4115' = 'HYBRID WIPER BLADE WBH-550 22"|WBH00460|456.000000|709659.72'
                '4116' = 'HYBRID WIPER BLADE WBH-600 24"|WBH00560|1752.000000|2939016.72'
                '4117' = 'HYBRID WIPER BLADE WBH-650 26"|WBH00660|812.000000|1414174.73'
                '4180' = 'ECO DIESEL 15W-40 1 L|TOTACHIE5601|329.000000|673089.68'
                '4181' = 'DENTO HEAVY DUTY DIESEL 15W40  200L|TOTACHIE342Z|0.000000|0.00'
                '4209' = 'C110 - SRR OIL FILTER|C110|0.000000|0.00'
                '4227' = 'FINE GASOLINE 15W 40 SP 4L PACK|PRO036|0.000000|0.00'
                '4231' = 'HYPER ECODRIVE 0W20 4L PACK|PRO040|0.000000|0.00'
                '4253' = 'EURODRIVE ECO 0W20 4 L|TOTACHIE6404|116.000000|1118906.92'
                '4254' = 'ECO DIESEL 5W-30 1 L|TOTACHIE2201|142.000000|285526.79'
                '4255' = 'ECO DIESEL 5W-30 6 L|TOTACHIE2206|91.000000|946376.23'
                '4256' = 'DENTO HEAVY DUTY DIESEL 15W40  25 L|TOTACHIE341Z|0.000000|0.00'
                '4257' = 'DENTO HEAVY DUTY DIESEL 20W50  25 L|TOTACHIE351Z|15.000000|442031.74'
                '4397' = 'ECO DIESEL 5W30 200 L|TOTACHIE222Z|3.000000|1139604.80'
                '4426' = 'ECO GASOLINE 5W30 4 L PACK|PRO037|0.000000|0.00'
                }
            }
        }
        $stockAllWaMetrics = @{ computed = 0; supported = 0; unsupported = 0; allowed = 0; refused = 0 }
        $stockAllWaSelectedByItem = @{}

        # [STOCK-GENERAL-OBSERVE-V1:STATE:BEGIN]
        # Stored strictly for comparison after normal report valuation.
        # Never read by any report value/price override path.
        $stockGeneralObserveCandidates = @{}
        $stockGeneralObserveMovementSeen = @{}
        $stockGeneralObserveOpeningOnlySupported = 0
        $stockGeneralObserveOpeningOnlyUnsupported = 0
        $stockGeneralObserveComparisons = [System.Collections.ArrayList]::new()
        $stockGeneralObserveCentreComparisons = [System.Collections.ArrayList]::new()
        $stockGeneralObserveRejectedQty = 0
        # [STOCK-RETAINED-RETURN-V3:METRICS:BEGIN]
        $stockGeneralObserveRetainedMetric = [pscustomobject]@{
            events = 0
            items = [System.Collections.ArrayList]::new()
        }
        # [STOCK-RETAINED-RETURN-V3:METRICS:END]
        # [STOCK-GENERAL-OBSERVE-V1:STATE:END]
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

        # [COMP0004-V5-COMPANY-OBSERVE:CONFIG]
        # Observation only, OFF by default; does NOT affect API prices, totals,
        # databases, legacy Comp0005 experiments, or native BUSY settings.
        $v5CompanyObserverEnabled = (
            [string]$env:BUSY_COMP0004_V5_OBSERVE -eq '1' -and
            $gregoryNativeStockEligibility -and
            $asOfDate.Date -in @([datetime]'2026-09-30', [datetime]'2026-10-09') -and
            $IncludeStockTransfers -and
            $valueMode -eq 'busy'
        )
        $v5CompanyObserveResult = @{}
        $v5CompanyObserveErrors = 0
        $v5CompanyTraceItemCode = 0
        if ($v5CompanyObserverEnabled) {
            [void][int]::TryParse(
                [string]$env:BUSY_COMP0004_V5_TRACE_ITEM,
                [ref]$v5CompanyTraceItemCode
            )
            Write-ReportEngineLog -Stage 'COMP0004-V5-OBSERVE' -Message (
                'Read-only company-wide cost replay enabled for ' +
                $asOfDate.ToString('yyyy-MM-dd') + '. V4 report is unchanged.'
            ) -Color 'Yellow'
        }
        # [COMP0004-V5-COMPANY-OBSERVE:CONFIG-END]

        # V5 company-wide moving-weighted-average research engine.
        # NOT an authoritative implementation of BUSY's undocumented rate selection.
        # Read-only: accepts immutable voucher rows and opening snapshots; no DB writes.
        function Invoke-Comp0004V5CompanyReplay {
            param(
                [int]$ItemCode,
                [object[]]$Rows,
                [object[]]$Openings,
                [hashtable]$ParentMap,
                [int]$TraceItemCode = 0,
                [switch]$V511GeneratedFirst,
                [switch]$V511SkipAdaptiveSelection
            )
            $tol = 0.00000001
            $pool = [pscustomobject]@{ qty = 0.0; value = 0.0; last = 0.0 }
            foreach ($op in $Openings) {
                $pool.qty += [double]$op.qty
                $pool.value += [double]$op.value
            }
            if ([Math]::Abs($pool.qty) -gt $tol) {
                $pool.last = [Math]::Abs($pool.value / $pool.qty)
            }
            $physicalQty = [double]$pool.qty

            # V5.11 observation ONLY: do not globally change same-day ordering.
            # Examine both possible same-day orders only for items with a negative
            # generated Stock Journal receipt and sale on the same accounting date.
            # Select journal-first ONLY if sale-first makes positive physical stock
            # carry negative cost, while journal-first leaves it nonnegative.
            # Pure replay against voucher data, NEVER against BUSY reference values.
            $v511TriggerDay = ''
            if (-not $V511SkipAdaptiveSelection -and -not $V511GeneratedFirst) {
                $v511JournalDays = @{}
                $v511SaleDays = @{}
                foreach ($v511Row in $Rows) {
                    $v511Day = [string]$v511Row.voucherDate
                    if (
                        [int]$v511Row.vchType -eq 8 -and
                        [double]$v511Row.qty -gt $tol -and
                        [double]$v511Row.value3 -lt -$tol -and
                        [double]$v511Row.d5 -lt -$tol
                    ) { $v511JournalDays[$v511Day] = $true }
                    if (
                        [int]$v511Row.vchType -eq 9 -and
                        [double]$v511Row.qty -lt -$tol
                    ) { $v511SaleDays[$v511Day] = $true }
                }
                foreach ($v511Day in @($v511JournalDays.Keys | Sort-Object)) {
                    if (-not $v511SaleDays.ContainsKey($v511Day)) { continue }
                    $v511PrefixRows = @($Rows | Where-Object { [string]$_.voucherDate -cle $v511Day })
                    $v511Original = Invoke-Comp0004V5CompanyReplay -ItemCode $ItemCode `
                        -Rows $v511PrefixRows -Openings $Openings -ParentMap $ParentMap `
                        -V511SkipAdaptiveSelection
                    $v511Alternate = Invoke-Comp0004V5CompanyReplay -ItemCode $ItemCode `
                        -Rows $v511PrefixRows -Openings $Openings -ParentMap $ParentMap `
                        -V511GeneratedFirst -V511SkipAdaptiveSelection
                    if (
                        $v511Original.supported -and $v511Alternate.supported -and
                        [double]$v511Original.quantity -gt $tol -and
                        [double]$v511Alternate.quantity -gt $tol -and
                        [double]$v511Original.amount -lt -0.005 -and
                        [double]$v511Alternate.amount -ge -0.005
                    ) {
                        $V511GeneratedFirst = $true
                        $v511TriggerDay = $v511Day
                        break
                    }
                }
            }
            if ($V511GeneratedFirst) {
                $ordered = @($Rows | Sort-Object -Property `
                    @{Expression={ [string]$_.voucherDate }}, `
                    @{Expression={
                        if (
                            [int]$_.vchType -eq 8 -and [double]$_.qty -gt $tol -and
                            [double]$_.value3 -lt -$tol -and [double]$_.d5 -lt -$tol
                        ) { 0 } else { 1 }
                    }}, `
                    @{Expression={ [int]$_.vchCode }}, `
                    @{Expression={ [int]$_.srNo }} )
            }
            else {
                $ordered = @($Rows | Sort-Object voucherDate, vchCode, srNo)
            }
            $preFirstSigned = $null
            $probeQty = [double]$pool.qty
            foreach ($r in $ordered) {
                if (
                    [int]$r.vchType -eq 8 -and
                    [double]$r.qty -gt $tol -and
                    [double]$r.value3 -lt -$tol -and
                    [double]$r.d5 -lt -$tol -and
                    [Math]::Abs([double]$r.value3 - [double]$r.d5) -le 0.01
                ) {
                    $preFirstSigned = [double]$probeQty
                    break
                }
                $probeQty += [double]$r.qty
            }
            # Research hypothesis fixed before testing against both native dates:
            # signed journal inward when first qualifying entry had positive company stock;
            # otherwise use unsigned (ABS) as the initial exception rule.
            $useSignedJournals = ($null -ne $preFirstSigned -and [double]$preFirstSigned -gt $tol)
            $purchases = @{}
            $sales = @{}
            $trace = [System.Collections.ArrayList]::new()
            # V5.3 research: only balance-zero conditions caused by a signed generated Stock Journal.
            $zeroFromSignedJournal = $false
            # V5.4 observation hypothesis: a signed generated journal can leave
            # positive physical stock with negative cost; track this separately
            # from the zero-value journal replenishment logic in V5.3.
            $negativeFromSignedJournal = $false
            # V5.5 observation: detect a later stock-journal consumption while
            # a qualifying earlier generated signed journal is in the history.
            $seenSignedGeneratedJournal = $false
            $postNegativeConsumedJournal = $false
            $skippedTransfers = 0

            $i = 0
            while ($i -lt $ordered.Count) {
                $first = $ordered[$i]
                $voucherKey = [string]$first.vchCode
                $dateKey = [string]$first.voucherDate
                $group = [System.Collections.ArrayList]::new()
                while (
                    $i -lt $ordered.Count -and
                    [string]$ordered[$i].voucherDate -ceq $dateKey -and
                    [string]$ordered[$i].vchCode -ceq $voucherKey
                ) {
                    [void]$group.Add($ordered[$i]); $i++
                }
                $allTransfer = $true
                $transferQty = 0.0
                foreach ($g in $group) {
                    if ([int]$g.vchType -ne 5) { $allTransfer = $false }
                    $transferQty += [double]$g.qty
                    $physicalQty += [double]$g.qty
                }
                if ($allTransfer -and [Math]::Abs($transferQty) -lt 0.0000001) {
                    $skippedTransfers++
                    continue
                }
                $beforeQty = [double]$pool.qty
                $beforeValue = [double]$pool.value

                # Phase 1: priced inward movements and purchase returns.
                foreach ($r in $group) {
                    $q = [double]$r.qty
                    $typ = [int]$r.vchType
                    $amount = Get-Comp0004V5PostedValue $r ([Math]::Abs($q))
                    if ($q -gt $tol -and $typ -ne 3 -and $amount -gt $tol) {
                        $beforeInQty = [double]$pool.qty
                        $beforeInValue = [double]$pool.value
                        Add-Comp0004V5Pool $pool $q $amount
                        # RESEARCH ONLY: when signed journal had left a positive physical
                        # quantity with zero value, the later purchase may be costed
                        # against (inwardQty - zeroValueQty), but still valued over the
                        # full physical closing quantity. Never apply to zero-value purchases.
                        if (
                            $typ -eq 2 -and $zeroFromSignedJournal -and
                            $beforeInQty -gt $tol -and
                            [Math]::Abs($beforeInValue) -le $tol -and
                            $q -gt ($beforeInQty + $tol)
                        ) {
                            $denominator = $q - $beforeInQty
                            $newUnitCost = $amount / $denominator
                            $pool.value = [double]$pool.qty * $newUnitCost
                            $pool.last = [Math]::Abs($newUnitCost)
                            $zeroFromSignedJournal = $false
                        }
                        # V5.4 RESEARCH-ONLY: when a prior signed generated
                        # Stock Journal left a negative-valued positive-quantity
                        # pool, and a substantially larger purchase replenishes
                        # it, BUSY's native rate matches the new purchase rate
                        # rather than a blend with the negative prior balance.
                        # q>=5*oldQty is a conservative transaction-based guard.
                        # This affects only the V5 observer, never V4 output.
                        if (
                            $typ -eq 2 -and $negativeFromSignedJournal -and
                            $beforeInQty -gt $tol -and $beforeInValue -lt -$tol -and
                            $q -ge (5.0 * $beforeInQty)
                        ) {
                            $freshRate = $amount / $q
                            $pool.value = [double]$pool.qty * $freshRate
                            $pool.last = [Math]::Abs($freshRate)
                            $negativeFromSignedJournal = $false
                        }
                        if (
                            $useSignedJournals -and $typ -eq 8 -and
                            [double]$r.value3 -lt -$tol -and [double]$r.d5 -lt -$tol -and
                            [Math]::Abs([double]$r.value3 - [double]$r.d5) -le 0.01
                        ) {
                            $pool.value -= (2.0 * $amount)
                            $seenSignedGeneratedJournal = $true
                            if ([Math]::Abs($pool.qty) -gt $tol) {
                                $pool.last = [Math]::Abs($pool.value / $pool.qty)
                            }
                            if ($pool.qty -gt $tol -and [Math]::Abs($pool.value) -le $tol) {
                                $zeroFromSignedJournal = $true
                            }
                            if ($pool.qty -gt $tol -and $pool.value -lt -$tol) {
                                $negativeFromSignedJournal = $true
                            }
                        }
                        if ($typ -eq 2) {
                            if (-not $purchases.ContainsKey($voucherKey)) {
                                $purchases[$voucherKey] = [System.Collections.ArrayList]::new()
                            }
                            [void]$purchases[$voucherKey].Add([pscustomobject]@{
                                qty = $q; rate = ($amount / $q); returned = 0.0
                            })
                        }
                    }
                    if ($q -lt -$tol -and $typ -eq 10) {
                        $remaining = -$q
                        $parentKey = ''
                        if ($ParentMap.ContainsKey($voucherKey)) {
                            $parentKey = [string]$ParentMap[$voucherKey]
                        }
                        if ($parentKey -and $purchases.ContainsKey($parentKey)) {
                            foreach ($source in $purchases[$parentKey]) {
                                $take = [Math]::Min(
                                    [Math]::Max(0.0, [double]$source.qty - [double]$source.returned),
                                    $remaining
                                )
                                if ($take -gt $tol) {
                                    Remove-Comp0004V5Pool $pool $take ([double]$source.rate)
                                    $source.returned += $take
                                    $remaining -= $take
                                }
                                if ($remaining -le $tol) { break }
                            }
                        }
                        if ($remaining -gt $tol) {
                            $fallback = if ($amount -gt $tol) { $amount / (-$q) } else { Get-Comp0004V5Rate $pool }
                            Remove-Comp0004V5Pool $pool $remaining $fallback
                        }
                    }
                }
                # Phase 2: zero-valued inward movements, kept at zero.
                foreach ($r in $group) {
                    $q = [double]$r.qty
                    if (
                        $q -gt $tol -and [int]$r.vchType -ne 3 -and
                        (Get-Comp0004V5PostedValue $r $q) -le $tol
                    ) { Add-Comp0004V5Pool $pool $q 0.0 }
                }
                # Phase 3: weighted-average outward movements; use posted cost only
                # if the company-wide stock was already nonpositive.
                foreach ($r in $group) {
                    $q = [double]$r.qty
                    $typ = [int]$r.vchType
                    if ($q -ge -$tol -or $typ -eq 10) { continue }
                    $outQty = -$q
                    if ($pool.qty -gt $tol) {
                        $rate = if ([Math]::Abs($pool.value) -le $tol) { 0.0 } else { Get-Comp0004V5Rate $pool }
                    }
                    else {
                        $rate = Get-Comp0004V5NegativeOutRate $r $outQty
                        if ($rate -le $tol) { $rate = Get-Comp0004V5Rate $pool }
                    }
                    if ($typ -eq 9) {
                        if (-not $sales.ContainsKey($voucherKey)) {
                            $sales[$voucherKey] = [System.Collections.ArrayList]::new()
                        }
                        [void]$sales[$voucherKey].Add([pscustomobject]@{ qty = $outQty; rate = $rate })
                    }
                    Remove-Comp0004V5Pool $pool $outQty $rate
                    # V5.5 research only: a consumed stock-journal event following
                    # an earlier generated signed journal can leave a negative
                    # company-wide value even while the quantity remains positive.
                    if (
                        $typ -eq 8 -and $seenSignedGeneratedJournal -and
                        $pool.qty -gt $tol -and $pool.value -lt -$tol
                    ) { $postNegativeConsumedJournal = $true }
                }
                # Phase 4: Sale Returns, using linked source sale cost if available.
                foreach ($r in $group) {
                    $q = [double]$r.qty
                    if ($q -le $tol -or [int]$r.vchType -ne 3) { continue }
                    $parentKey = ''
                    if ($ParentMap.ContainsKey($voucherKey)) { $parentKey = [string]$ParentMap[$voucherKey] }
                    $rate = 0.0
                    if ($parentKey -and $sales.ContainsKey($parentKey)) {
                        $matched = 0.0; $matchedValue = 0.0
                        foreach ($source in $sales[$parentKey]) {
                            $take = [Math]::Min([double]$source.qty, [Math]::Max(0.0, $q - $matched))
                            $matched += $take; $matchedValue += $take * [double]$source.rate
                            if ($matched -ge $q - $tol) { break }
                        }
                        if ($matched -gt $tol) { $rate = $matchedValue / $matched }
                    }
                    if ($rate -le $tol) { $rate = Get-Comp0004V5Rate $pool }
                    if ($rate -le $tol) {
                        $rate = (Get-Comp0004V5PostedValue $r $q) / $q
                    }
                    # V5.5 RESEARCH-ONLY: after signed Stock Journal consumption
                    # with a negative-valued positive-quantity pool, BUSY Native
                    # is consistent with Sale Returns retaining the current
                    # negative average rate (rather than contributing positive
                    # linked or posted sale-return cost). Earlier sale returns
                    # retain V5.4 behavior. Only observer state is changed.
                    if (
                        $postNegativeConsumedJournal -and
                        $pool.qty -gt $tol -and $pool.value -lt -$tol
                    ) {
                        $negativeRate = [double]$pool.value / [double]$pool.qty
                        $pool.qty += $q
                        $pool.value += $q * $negativeRate
                        $pool.last = [Math]::Abs($negativeRate)
                    }
                    # V5.6 RESEARCH-ONLY: a Sales Return into a positive-quantity
                    # zero-valued pool created by a signed generated Stock Journal
                    # restores zero-valued physical stock rather than posting its
                    # voucher amount to company-wide inventory valuation.
                    elseif (
                        $seenSignedGeneratedJournal -and $zeroFromSignedJournal -and
                        $pool.qty -gt $tol -and
                        [Math]::Abs([double]$pool.value) -le $tol
                    ) {
                        $pool.qty += $q
                        $pool.value = 0.0
                        $pool.last = 0.0
                        # The zero pool has already been restored by the return;
                        # do not reuse V5.3's virtual-quantity purchase denominator.
                        $zeroFromSignedJournal = $false
                    }
                    else {
                        Add-Comp0004V5Pool $pool $q ($q * $rate)
                    }
                }
                if ($ItemCode -eq $TraceItemCode) {
                    [void]$trace.Add([pscustomobject]@{
                        date=$dateKey; voucher=$voucherKey; type=[int]$first.vchType
                        beforeQty=$beforeQty; beforeValue=$beforeValue
                        afterQty=[double]$pool.qty; afterValue=[double]$pool.value
                    })
                }
            }
            if ([Math]::Abs($physicalQty - $pool.qty) -gt 0.000001) {
                return [pscustomobject]@{supported=$false; reason='Company pool physical quantity mismatch'; itemCode=$ItemCode }
            }
            $rate = if ([Math]::Abs($pool.qty) -gt $tol) { $pool.value / $pool.qty } else { 0.0 }
            $journalMode = if ($useSignedJournals) { 'signed' } else { 'unsigned' }
            return [pscustomobject]@{
                supported = $true
                itemCode = $ItemCode
                quantity = [double]$pool.qty
                amount = [double]$pool.value
                rate = [double]$rate
                signedJournalMode = $journalMode
                firstSignedJournalPreQty = $preFirstSigned
                balancedTransfersSkipped = $skippedTransfers
                v511GeneratedJournalFirst = [bool]$V511GeneratedFirst
                v511TriggerDay = $v511TriggerDay
                trace = @($trace.ToArray())
            }
        }

        function Get-Comp0004V5Rate {
            param($Pool)
            if ([Math]::Abs([double]$Pool.qty) -gt 0.00000001) {
                return [Math]::Abs([double]$Pool.value / [double]$Pool.qty)
            }
            return [Math]::Abs([double]$Pool.last)
        }
        function Get-Comp0004V5PostedValue {
            param($Row, [double]$Quantity)
            if ([Math]::Abs([double]$Row.d5) -gt 0.00000001) { return [Math]::Abs([double]$Row.d5) }
            if ([Math]::Abs([double]$Row.d2) -gt 0.00000001) { return $Quantity * [Math]::Abs([double]$Row.d2) }
            return [Math]::Abs([double]$Row.value3)
        }
        function Get-Comp0004V5NegativeOutRate {
            param($Row, [double]$Quantity)
            if ([Math]::Abs([double]$Row.value3) -gt 0.00000001) { return [Math]::Abs([double]$Row.value3) / $Quantity }
            if ([Math]::Abs([double]$Row.d5) -gt 0.00000001) { return [Math]::Abs([double]$Row.d5) / $Quantity }
            if ([Math]::Abs([double]$Row.d2) -gt 0.00000001) { return [Math]::Abs([double]$Row.d2) }
            return 0.0
        }
        function Add-Comp0004V5Pool {
            param($Pool, [double]$Quantity, [double]$Value)
            if ($Quantity -le 0.00000001) { return }
            $amount = [Math]::Abs($Value)
            $rate = $amount / $Quantity
            if ($Pool.qty -lt -0.00000001 -and $amount -le 0.00000001) {
                $Pool.qty += $Quantity
                if ([Math]::Abs($Pool.qty) -le 0.00000001) { $Pool.qty = 0.0 }
                $Pool.value = 0.0; $Pool.last = 0.0
                return
            }
            if ($Pool.qty -lt -0.00000001) {
                $negativeRate = Get-Comp0004V5Rate $Pool
                $off = [Math]::Min($Quantity, [Math]::Abs($Pool.qty))
                $Pool.qty += $off; $Pool.value += ($off * $negativeRate)
                $Quantity -= $off
                if ([Math]::Abs($Pool.qty) -le 0.00000001) {
                    $Pool.qty = 0.0; $Pool.value = 0.0
                }
            }
            if ($Quantity -gt 0.00000001) {
                $Pool.qty += $Quantity
                $Pool.value += ($Quantity * $rate)
            }
            if ([Math]::Abs($Pool.qty) -gt 0.00000001) {
                $Pool.last = [Math]::Abs($Pool.value / $Pool.qty)
            }
            elseif ($rate -gt 0.00000001) { $Pool.last = $rate }
        }
        function Remove-Comp0004V5Pool {
            param($Pool, [double]$Quantity, [double]$UnitRate)
            if ($Quantity -le 0.00000001) { return }
            $rate = if ($UnitRate -gt 0.00000001) { $UnitRate } else { Get-Comp0004V5Rate $Pool }
            $Pool.qty -= $Quantity
            $Pool.value -= $Quantity * $rate
            if ([Math]::Abs($Pool.qty) -le 0.00000001) {
                $Pool.qty = 0.0; $Pool.value = 0.0
                if ($rate -gt 0.00000001) { $Pool.last = $rate }
            }
            else { $Pool.last = [Math]::Abs($Pool.value / $Pool.qty) }
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

            # [COMP0004-V5-COMPANY-OBSERVE:ITEM]
            if ($v5CompanyObserverEnabled) {
                try {
                    $v5Openings = [System.Collections.ArrayList]::new()
                    foreach ($v5McKey in @($items[$ItemCode].mc.Keys)) {
                        $v5Mc = $items[$ItemCode].mc[$v5McKey]
                        [void]$v5Openings.Add([pscustomobject]@{
                            qty = [double]$v5Mc.baseOpeningMain
                            value = [double]$v5Mc.baseOpeningValue
                        })
                    }
                    $v5Candidate = Invoke-Comp0004V5CompanyReplay `
                        -ItemCode $ItemCode `
                        -Rows @($Rows.ToArray()) `
                        -Openings @($v5Openings.ToArray()) `
                        -ParentMap $busyReturnParentMap `
                        -TraceItemCode $v5CompanyTraceItemCode
                    if ($v5Candidate.supported) {
                        $v5CompanyObserveResult[[string]$ItemCode] = $v5Candidate
                        if ($ItemCode -eq $v5CompanyTraceItemCode) {
                            Write-ReportEngineLog -Stage 'COMP0004-V5-TRACE' -Message (
                                'Item ' + $ItemCode + ', qty=' + $v5Candidate.quantity +
                                ', value=' + $v5Candidate.amount +
                                ', rate=' + $v5Candidate.rate +
                                ', signedJournal=' + $v5Candidate.signedJournalMode +
                                ', v511JournalFirst=' + $v5Candidate.v511GeneratedJournalFirst +
                                ', v511TriggerDay=' + $v5Candidate.v511TriggerDay
                            ) -Color 'Yellow'
                            $v5TraceTail = @($v5Candidate.trace | Select-Object -Last 25)
                            foreach ($v5Event in $v5TraceTail) {
                                Write-ReportEngineLog -Stage 'COMP0004-V5-TRACE' -Message (
                                    'Voucher ' + $v5Event.voucher + ' date ' + $v5Event.date +
                                    ' qty ' + $v5Event.beforeQty + ' -> ' + $v5Event.afterQty +
                                    ' value ' + $v5Event.beforeValue + ' -> ' + $v5Event.afterValue
                                ) -Color 'Yellow'
                            }
                        }
                    }
                    else { $v5CompanyObserveErrors++ }
                }
                catch {
                    $v5CompanyObserveErrors++
                    Write-ReportEngineLog -Stage 'COMP0004-V5-OBSERVE' -Message (
                        'Item ' + $ItemCode + ' diagnostic failure: ' + $_.Exception.Message
                    ) -Color 'Yellow'
                }
            }
            # [COMP0004-V5-COMPANY-OBSERVE:ITEM-END]

            # Independent date-level, item-wide replay for the verified item.
            # This does not replace legacy WA states used by other items.
            if (
                $stockWaLabEnabled -and
                $ItemCode -eq 1198 -and
                [string]$items[$ItemCode].name -eq 'CVT MULTI-TYPE 4 L' -and
                [string]$items[$ItemCode].alias -eq 'TOTACHI20504'
            ) {
                $stockWaLabOpenings = [System.Collections.ArrayList]::new()
                foreach ($labMcKey in @($items[$ItemCode].mc.Keys)) {
                    $labMc = $items[$ItemCode].mc[$labMcKey]
                    [void]$stockWaLabOpenings.Add([pscustomobject]@{
                        mcCode = [int]$labMcKey
                        qty = [double]$labMc.baseOpeningMain
                        value = [double]$labMc.baseOpeningValue
                    })
                }
                $labResult = Invoke-StockStatusWaLabReplay `
                    -Rows @($Rows.ToArray()) `
                    -Openings @($stockWaLabOpenings.ToArray()) `
                    -AsOf ($endDate.ToString('yyyy-MM-dd'))
                if ($labResult.supported) {
                    $stockWaLabByItem[[string]$ItemCode] = $labResult
                }
                else {
                    Write-ReportEngineLog `
                        -Stage 'STOCK-WA-LAB' `
                        -Message ('Test-only replay refused: ' + $labResult.reason) `
                        -Color 'Yellow'
                }
            }

            # Non-invasive ALL-ITEMS hypothesis: separate from the existing
            # item-1198 experiment and from legacy method-5 cost states.
            if ($stockAllWaMode -ne 'off') {
                if ($stockGeneralCandidateMode -ne 'off') {
                    $stockGeneralObserveMovementSeen[[string]$ItemCode] = $true
                }
                $stockAllWaMetrics.computed++
                $labOpens = [System.Collections.ArrayList]::new()
                foreach ($mcText in @($items[$ItemCode].mc.Keys)) {
                    $labMcState = $items[$ItemCode].mc[$mcText]
                    [void]$labOpens.Add([pscustomobject]@{
                        mcCode = [int]$mcText
                        qty = [double]$labMcState.baseOpeningMain
                        value = [double]$labMcState.baseOpeningValue
                    })
                }
                $labCandidate = $null
                try {
                    $labCandidate = Invoke-StockStatusWaLabReplay `
                        -Rows @($Rows.ToArray()) `
                        -Openings @($labOpens.ToArray()) `
                        -AsOf ($endDate.ToString('yyyy-MM-dd')) `
                        -SaleReturnsBeforeOtherInwards `
                        -UsePostedReturnWhenPoolEmpty `
                        -UseRetainedRateForEmptyReturn:($stockGeneralCandidateMode -ne 'off') `
                        -UsePostedIssueWhenPoolEmpty `
                        -PurchaseReceiptsBeforeReturns `
                        -RoundPoolValueAfterEachMovement
                }
                catch {
                    if ($stockGeneralCandidateMode -eq 'off') { throw }
                    $labCandidate = [pscustomobject]@{
                        supported = $false
                        reason = 'General observation replay exception'
                    }
                }
                if ($null -ne $labCandidate -and $labCandidate.supported) {
                    $stockAllWaMetrics.supported++
                    # General candidate is staged, then preflight-gated before any use.
                    if ($stockGeneralCandidateMode -ne 'off') {
                        $stockGeneralObserveCandidates[[string]$ItemCode] = $labCandidate
                        if ([int]$labCandidate.retainedEmptyReturns -gt 0) {
                            $stockGeneralObserveRetainedMetric.events += [int]$labCandidate.retainedEmptyReturns
                            [void]$stockGeneralObserveRetainedMetric.items.Add([int]$ItemCode)
                        }
                    }
                    if ($stockAllWaMode -eq 'apply-verified') {
                        $stockAllKey = [string]$ItemCode
                        if ($stockAllWaVerified.ContainsKey($stockAllKey)) {
                            $identityBits = ([string]$stockAllWaVerified[$stockAllKey]).Split('|')
                            $matchingIdentity = (
                                $identityBits.Count -eq 4 -and
                                [string]$items[$ItemCode].name -ceq $identityBits[0] -and
                                [string]$items[$ItemCode].alias -ceq $identityBits[1]
                            )
                            if ($matchingIdentity) {
                                [double]$refQty = 0.0
                                [double]$refAmount = 0.0
                                $qtyParsed = [double]::TryParse($identityBits[2], [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$refQty)
                                $amtParsed = [double]::TryParse($identityBits[3], [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$refAmount)
                                # Every item, including code 1227, uses the strict
                                # ten-cent reference guard at this checkpoint.
                                [double]$stockAllWaAmountTolerance = 0.100001
                                if (
                                    $qtyParsed -and $amtParsed -and
                                    [Math]::Abs([double]$labCandidate.quantity - $refQty) -le 0.000001 -and
                                    [Math]::Abs([double]$labCandidate.amount - $refAmount) -le $stockAllWaAmountTolerance -and
                                    -not [double]::IsNaN([double]$labCandidate.rate) -and
                                    -not [double]::IsInfinity([double]$labCandidate.rate)
                                ) {
                                    $stockAllWaSelectedByItem[$stockAllKey] = $labCandidate
                                    $stockAllWaMetrics.allowed++
                                }
                                else { $stockAllWaMetrics.refused++ }
                            }
                            else { $stockAllWaMetrics.refused++ }
                        }
                    }
                }
                else {
                    $stockAllWaMetrics.unsupported++
                    Write-ReportEngineLog -Stage 'STOCK-ALLWA' -Message ('Candidate item ' + $ItemCode + ' unsupported: ' + $labCandidate.reason) -Color 'Yellow'
                }
            }

            # [GREGORY-V3-SIGNED-JOURNAL-STRICT:ELIGIBILITY:BEGIN]
            # LOCAL Comp0004 / 2026-10-09 experiment only. Never infer signed
            # journal value simply from a negative D5 in arbitrary vouchers.
            # Only one negative-valued journal inward, up to six total item
            # movements, no returns, all MC closing quantities nonnegative,
            # and the receipt MC has positive opening quantity. A final guard
            # below checks that the current valuation pool is positive when
            # the signed entry is processed. These conservative conditions
            # were evaluated against the full Comp0004 V2 native comparison.
            $gregorySignedJournalV3Eligible = $false
            if (
                $gregoryNativeStockEligibility -and
                $asOfDate.Date -eq [datetime]'2026-10-09' -and
                $Rows.Count -le 6
            ) {
                $v3JournalRows = @(
                    $Rows | Where-Object {
                        [int]$_.vchType -eq 8
                    }
                )
                if ($v3JournalRows.Count -eq 1) {
                    $v3JournalRow = $v3JournalRows[0]
                    $v3JournalMcKey = [string][int]$v3JournalRow.mcCode
                    $v3AllNonNegative = $true
                    $v3HasReturn = $false
                    $v3FinalMcQty = @{}
                    foreach ($v3OpeningMcKey in @($items[$ItemCode].mc.Keys)) {
                        $v3FinalMcQty[[string]$v3OpeningMcKey] = [double](
                            $items[$ItemCode].mc[$v3OpeningMcKey].baseOpeningMain
                        )
                    }
                    foreach ($v3CandidateRow in $Rows) {
                        $v3CandidateMcKey = [string][int]$v3CandidateRow.mcCode
                        if (-not $v3FinalMcQty.ContainsKey($v3CandidateMcKey)) {
                            $v3FinalMcQty[$v3CandidateMcKey] = 0.0
                        }
                        $v3FinalMcQty[$v3CandidateMcKey] =
                            [double]$v3FinalMcQty[$v3CandidateMcKey] +
                            [double]$v3CandidateRow.qty
                        if (
                            [int]$v3CandidateRow.vchType -eq 3 -or
                            [int]$v3CandidateRow.vchType -eq 10
                        ) { $v3HasReturn = $true }
                    }
                    foreach ($v3FinalMcKey in @($v3FinalMcQty.Keys)) {
                        if ([double]$v3FinalMcQty[$v3FinalMcKey] -lt -$eps) {
                            $v3AllNonNegative = $false
                            break
                        }
                    }
                    $v3JournalMcOpeningQty = 0.0
                    if ($items[$ItemCode].mc.ContainsKey($v3JournalMcKey)) {
                        $v3JournalMcOpeningQty = [double](
                            $items[$ItemCode].mc[$v3JournalMcKey].baseOpeningMain
                        )
                    }
                    if (
                        [double]$v3JournalRow.qty -gt $eps -and
                        [double]$v3JournalRow.value3 -lt -$eps -and
                        [double]$v3JournalRow.d5 -lt -$eps -and
                        [Math]::Abs(
                            [double]$v3JournalRow.value3 -
                            [double]$v3JournalRow.d5
                        ) -le 0.01 -and
                        $v3JournalMcOpeningQty -gt $eps -and
                        $v3AllNonNegative -and
                        -not $v3HasReturn
                    ) {
                        $gregorySignedJournalV3Eligible = $true
                    }
                }
            }
            # [GREGORY-V3-SIGNED-JOURNAL-STRICT:ELIGIBILITY:END]

            # [GREGORY-V4-SIGNED-JOURNAL-EXTENSION:ELIGIBILITY:BEGIN]
            # Comp0004/2026-10-09 LOCAL TEST ONLY. Broaden the V3 rule
            # conservatively without modifying existing V3-eligible cases.
            # Case A: one negative-valued journal inward, 7..12 movements,
            # positive opening in receipt MC, no returns, nonnegative ending
            # balances in every MC. Case B: <=6 movements, receipt MC opening
            # nonpositive, no sales outside the journal MC, otherwise same.
            # These structural guards were regression-tested OFFLINE against
            # all 2,660 native-listed items, not validated as a general BUSY rule.
            $gregorySignedJournalV4Eligible = $false
            $gregorySignedJournalV4Case = ''
            if (
                $gregoryNativeStockEligibility -and
                $asOfDate.Date -eq [datetime]'2026-10-09' -and
                -not $gregorySignedJournalV3Eligible -and
                $Rows.Count -le 12
            ) {
                $v4JournalRows = @($Rows | Where-Object { [int]$_.vchType -eq 8 })
                if ($v4JournalRows.Count -eq 1) {
                    $v4Journal = $v4JournalRows[0]
                    $v4Mc = [string][int]$v4Journal.mcCode
                    $v4Opening = 0.0
                    if ($items[$ItemCode].mc.ContainsKey($v4Mc)) {
                        $v4Opening = [double]$items[$ItemCode].mc[$v4Mc].baseOpeningMain
                    }
                    $v4FinalQty = @{}
                    foreach ($v4OpeningMc in @($items[$ItemCode].mc.Keys)) {
                        $v4FinalQty[[string]$v4OpeningMc] = [double](
                            $items[$ItemCode].mc[$v4OpeningMc].baseOpeningMain
                        )
                    }
                    $v4AnyReturn = $false
                    $v4SaleOutsideJournalMc = $false
                    foreach ($v4Row in $Rows) {
                        $v4RowMc = [string][int]$v4Row.mcCode
                        if (-not $v4FinalQty.ContainsKey($v4RowMc)) {
                            $v4FinalQty[$v4RowMc] = 0.0
                        }
                        $v4FinalQty[$v4RowMc] = [double]$v4FinalQty[$v4RowMc] + [double]$v4Row.qty
                        if ([int]$v4Row.vchType -eq 3 -or [int]$v4Row.vchType -eq 10) {
                            $v4AnyReturn = $true
                        }
                        if ([int]$v4Row.vchType -eq 9 -and $v4RowMc -cne $v4Mc) {
                            $v4SaleOutsideJournalMc = $true
                        }
                    }
                    $v4AllFinalNonnegative = $true
                    foreach ($v4FinalMc in @($v4FinalQty.Keys)) {
                        if ([double]$v4FinalQty[$v4FinalMc] -lt -$eps) {
                            $v4AllFinalNonnegative = $false
                            break
                        }
                    }
                    if (
                        [double]$v4Journal.qty -gt $eps -and
                        [double]$v4Journal.value3 -lt -$eps -and
                        [double]$v4Journal.d5 -lt -$eps -and
                        [Math]::Abs([double]$v4Journal.value3 - [double]$v4Journal.d5) -le 0.01 -and
                        -not $v4AnyReturn -and
                        $v4AllFinalNonnegative
                    ) {
                        if ($v4Opening -gt $eps) {
                            $gregorySignedJournalV4Eligible = $true
                            $gregorySignedJournalV4Case = 'positive-opening'
                        }
                        elseif ($Rows.Count -le 6 -and -not $v4SaleOutsideJournalMc) {
                            $gregorySignedJournalV4Eligible = $true
                            $gregorySignedJournalV4Case = 'nonpositive-opening'
                        }
                    }
                }
            }
            # [GREGORY-V4-SIGNED-JOURNAL-EXTENSION:ELIGIBILITY:END]

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
                            # [GREGORY-V3-SIGNED-JOURNAL-STRICT:APPLY:BEGIN]
                            # The common Add-BusyWaPool intentionally takes
                            # Abs(Value) to support other vouchers. Preserve
                            # the negative signed journal value ONLY when the
                            # strict item test and a positive live MC pool
                            # agree; otherwise leave V2 behavior unchanged.
                            $v3ApplySignedJournal = $false
                            if (
                                $gregorySignedJournalV3Eligible -and
                                $vchType -eq 8 -and
                                [double]$row.value3 -lt -$eps -and
                                [double]$row.d5 -lt -$eps
                            ) {
                                $v3StateBefore = Get-BusyWaState `
                                    -ItemCode $ItemCode `
                                    -McCode $mcCode
                                $v3ApplySignedJournal = (
                                    [double]$v3StateBefore.poolQty -gt $eps
                                )
                            }
                            # [GREGORY-V3-SIGNED-JOURNAL-STRICT:APPLY:END]

                            # [GREGORY-V4-SIGNED-JOURNAL-EXTENSION:APPLY:BEGIN]
                            $v4ApplySignedJournal = $false
                            if (
                                $gregorySignedJournalV4Eligible -and
                                -not $v3ApplySignedJournal -and
                                $vchType -eq 8 -and
                                [double]$row.value3 -lt -$eps -and
                                [double]$row.d5 -lt -$eps
                            ) {
                                $v4Before = Get-BusyWaState -ItemCode $ItemCode -McCode $mcCode
                                if ($gregorySignedJournalV4Case -ceq 'positive-opening') {
                                    $v4ApplySignedJournal = ([double]$v4Before.poolQty -gt $eps)
                                }
                                elseif ($gregorySignedJournalV4Case -ceq 'nonpositive-opening') {
                                    $v4ApplySignedJournal = ([double]$v4Before.poolQty -le $eps)
                                }
                            }
                            # [GREGORY-V4-SIGNED-JOURNAL-EXTENSION:APPLY:END]

                            Add-BusyWaPool `
                                -ItemCode $ItemCode `
                                -McCode $mcCode `
                                -Qty $qty `
                                -Value $explicitValue

                            # [GREGORY-V3-SIGNED-JOURNAL-STRICT:ADJUST:BEGIN]
                            if ($v3ApplySignedJournal -or $v4ApplySignedJournal) {
                                $v3StateAfter = Get-BusyWaState `
                                    -ItemCode $ItemCode `
                                    -McCode $mcCode
                                $v3StateAfter.poolValue =
                                    [double]$v3StateAfter.poolValue -
                                    (2.0 * [double]$explicitValue)
                                if ([Math]::Abs([double]$v3StateAfter.poolQty) -gt $eps) {
                                    $v3StateAfter.lastRate = [Math]::Abs(
                                        [double]$v3StateAfter.poolValue /
                                        [double]$v3StateAfter.poolQty
                                    )
                                }
                            }
                            # [GREGORY-V3-SIGNED-JOURNAL-STRICT:ADJUST:END]

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
                            elseif (
                                # GREGORY METHOD-5 EXPERIMENTAL TEST ONLY.
                                # Scope to exact company/DB/instance/FY via the
                                # existing guarded eligibility switch, plus
                                # the single native-reference as-on date.
                                # Never apply to GS Comp0005 or other dates.
                                $gregoryNativeStockEligibility -and
                                $asOfDate.Date -eq [datetime]'2026-10-09' -and
                                $vchType -eq 9 -and
                                [double]$stateBeforeOut.poolQty + $eps -ge $outQty
                            ) {
                                # Sale with sufficient positively valued MC
                                # stock: issue at its weighted-average COST.
                                # Do not use Tran2.D5/D2 sale invoice rate.
                                # Uncovered / negative-stock sales retain the
                                # existing fallback behavior below.
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

                        # [GREGORY-METHOD5-RETURN-FALLBACK-V2:BEGIN]
                        # READ-ONLY TEST-GUARDED EXPERIMENT, not a general
                        # BUSY rule. A Sale Return entering an MC which has no
                        # recoverable local valuation may need the retained
                        # stock-cost rate from another MC, rather than the
                        # return document's sales price (D2/D5).
                        #
                        # Apply only when at least one OTHER MC has positive
                        # qty and value, and ALL eligible source MCs have the
                        # same valuation rate (within 0.000001). Multiple
                        # competing rates are ambiguous and remain unchanged.
                        # Parent-linked return rates and valued local pools
                        # already set $returnRate; leave them untouched.
                        #
                        # Same strict instance/company/SQL DB/FY eligibility
                        # as V1, and ONLY for the 2026-10-09 reference date.
                        if (
                            $returnRate -le $eps -and
                            $gregoryNativeStockEligibility -and
                            $asOfDate.Date -eq [datetime]'2026-10-09'
                        ) {
                            $otherPositiveMcCount = 0
                            $otherQtyTotal = 0.0
                            $otherValueTotal = 0.0
                            $otherMinimumRate = [double]::PositiveInfinity
                            $otherMaximumRate = [double]::NegativeInfinity

                            foreach ($otherMcKey in @($items[$ItemCode].mc.Keys)) {
                                $otherMcCode = 0
                                [void][int]::TryParse(
                                    [string]$otherMcKey,
                                    [ref]$otherMcCode
                                )
                                if ($otherMcCode -eq $mcCode) { continue }

                                $otherMcState = Get-BusyWaState `
                                    -ItemCode $ItemCode `
                                    -McCode $otherMcCode

                                if (
                                    [double]$otherMcState.poolQty -gt $eps -and
                                    [double]$otherMcState.poolValue -gt $eps
                                ) {
                                    $otherPositiveMcCount++
                                    $otherQtyTotal += [double]$otherMcState.poolQty
                                    $otherValueTotal += [double]$otherMcState.poolValue
                                    $otherRate = Get-BusyWaRate `
                                        -ItemCode $ItemCode `
                                        -McCode $otherMcCode
                                    $otherMinimumRate = [Math]::Min(
                                        $otherMinimumRate, $otherRate
                                    )
                                    $otherMaximumRate = [Math]::Max(
                                        $otherMaximumRate, $otherRate
                                    )
                                }
                            }

                            if (
                                $otherPositiveMcCount -gt 0 -and
                                $otherQtyTotal -gt $eps -and
                                $otherMaximumRate - $otherMinimumRate -le 0.000001
                            ) {
                                $returnRate = $otherValueTotal / $otherQtyTotal
                            }
                        }
                        # [GREGORY-METHOD5-RETURN-FALLBACK-V2:END]

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
        # BUSYWEB-STOCK-WA-PROFILE-V1
        $waProfile = [pscustomobject]@{
            OpenMs = 0.0
            ScanMs = 0.0
            MoveNextMs = 0.0
            ReplayMs = 0.0
            Rows = 0
            Flushes = 0
        }
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
                $waReplayWatch = [System.Diagnostics.Stopwatch]::StartNew()
                try {
                    Process-BusyWaItemRows `
                        -ItemCode ([int]$busyWaCursorState.itemCode) `
                        -Rows $busyWaCursorState.rows
                }
                finally {
                    $waProfile.ReplayMs += $waReplayWatch.Elapsed.TotalMilliseconds
                    $waProfile.Flushes++
                }

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
            $waOpenWatch = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                $busyWaTxnRst = $fi.GetRecordset($busyWaTxnQry)
            }
            finally {
                $waOpenWatch.Stop()
                $waProfile.OpenMs = $waOpenWatch.Elapsed.TotalMilliseconds
            }
            $waScanWatch = [System.Diagnostics.Stopwatch]::StartNew()
            $waMoveWatch = New-Object System.Diagnostics.Stopwatch

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
                    $waProfile.Rows++
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

                    $waMoveWatch.Start()
                    $busyWaTxnRst.MoveNext()
                    $waMoveWatch.Stop()
                }
            }

            $waScanWatch.Stop()
            $waProfile.ScanMs = $waScanWatch.Elapsed.TotalMilliseconds
            $waProfile.MoveNextMs = $waMoveWatch.Elapsed.TotalMilliseconds
            Close-ReportRecordset $busyWaTxnRst
        }

        if (
            [int]$busyWaCursorState.itemCode -gt 0 -and
            $busyWaCursorState.rows.Count -gt 0
        ) {
            $waReplayWatch = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                Process-BusyWaItemRows `
                    -ItemCode ([int]$busyWaCursorState.itemCode) `
                    -Rows $busyWaCursorState.rows
            }
            finally {
                $waProfile.ReplayMs += $waReplayWatch.Elapsed.TotalMilliseconds
                $waProfile.Flushes++
            }
        }

        # [COMP0004-V5-COMPANY-OBSERVE:SUMMARY]
        if ($v5CompanyObserverEnabled) {
            # Include opening-only items that had no Tran2 rows.
            foreach ($v5ItemRaw in @($itemOrder)) {
                $v5Key = [string][int]$v5ItemRaw
                if ($v5CompanyObserveResult.ContainsKey($v5Key)) { continue }
                if (-not $items.ContainsKey([int]$v5ItemRaw)) { continue }
                $v5OpeningRows = [System.Collections.ArrayList]::new()
                foreach ($v5McKey in @($items[[int]$v5ItemRaw].mc.Keys)) {
                    $v5Mc = $items[[int]$v5ItemRaw].mc[$v5McKey]
                    [void]$v5OpeningRows.Add([pscustomobject]@{
                        qty = [double]$v5Mc.baseOpeningMain
                        value = [double]$v5Mc.baseOpeningValue
                    })
                }
                try {
                    $v5Candidate = Invoke-Comp0004V5CompanyReplay `
                        -ItemCode ([int]$v5ItemRaw) `
                        -Rows @() `
                        -Openings @($v5OpeningRows.ToArray()) `
                        -ParentMap $busyReturnParentMap
                    if ($v5Candidate.supported) {
                        $v5CompanyObserveResult[$v5Key] = $v5Candidate
                    }
                    else { $v5CompanyObserveErrors++ }
                }
                catch { $v5CompanyObserveErrors++ }
            }
            $v5TotalValue = 0.0
            $v5TotalQty = 0.0
            foreach ($v5Result in @($v5CompanyObserveResult.Values)) {
                $v5TotalValue += [double]$v5Result.amount
                $v5TotalQty += [double]$v5Result.quantity
            }
            Write-ReportEngineLog -Stage 'COMP0004-V5-OBSERVE' -Message (
                'Candidate item count=' + $v5CompanyObserveResult.Count +
                ', qty=' + [Math]::Round($v5TotalQty,3) +
                ', value=' + [Math]::Round($v5TotalValue,2) +
                ', replay errors=' + $v5CompanyObserveErrors +
                '. V4 response prices and values NOT modified.'
            ) -Color 'Yellow'
        }
        # [COMP0004-V5-COMPANY-OBSERVE:SUMMARY-END]

        # [STOCK-FULL-OBSERVE-V2:OPENING-ONLY:BEGIN]
        # Every report item that had no voucher rows still needs a proposed
        # opening-stock valuation. Reuse the SAME item-wide replay with an
        # explicit empty movement list; never synthesize candidates for items
        # whose movement replay was attempted and failed its safety checks.
        # This block is observation-only and does not touch any output field.
        if ($stockGeneralCandidateMode -ne 'off') {
            foreach ($observeItemRaw in @($itemOrder)) {
                $observeItemCode = [int]$observeItemRaw
                $observeItemKey = [string]$observeItemCode
                if (
                    $stockGeneralObserveMovementSeen.ContainsKey($observeItemKey) -or
                    $stockGeneralObserveCandidates.ContainsKey($observeItemKey) -or
                    -not $items.ContainsKey($observeItemCode)
                ) { continue }
                try {
                    $observeOpenings = [System.Collections.ArrayList]::new()
                    foreach ($observeMcKey in @($items[$observeItemCode].mc.Keys)) {
                        $observeState = $items[$observeItemCode].mc[$observeMcKey]
                        [void]$observeOpenings.Add([pscustomobject]@{
                            mcCode = [int]$observeMcKey
                            qty = [double]$observeState.baseOpeningMain
                            value = [double]$observeState.baseOpeningValue
                        })
                    }
                    $observeOpeningCandidate = Invoke-StockStatusWaLabReplay `
                        -Rows @() `
                        -Openings @($observeOpenings.ToArray()) `
                        -AsOf ($endDate.ToString('yyyy-MM-dd')) `
                        -SaleReturnsBeforeOtherInwards `
                        -UsePostedReturnWhenPoolEmpty `
                        -UsePostedIssueWhenPoolEmpty `
                        -PurchaseReceiptsBeforeReturns `
                        -RoundPoolValueAfterEachMovement
                    if ($null -ne $observeOpeningCandidate -and $observeOpeningCandidate.supported) {
                        $stockGeneralObserveCandidates[$observeItemKey] = $observeOpeningCandidate
                        $stockGeneralObserveOpeningOnlySupported++
                    }
                    else { $stockGeneralObserveOpeningOnlyUnsupported++ }
                }
                catch { $stockGeneralObserveOpeningOnlyUnsupported++ }
            }
        }
        # [STOCK-FULL-OBSERVE-V2:OPENING-ONLY:END]

        # [STOCK-GENERAL-APPLY-V4:PREFLIGHT:BEGIN]
        $stockGeneralApplyReady = $false
        $stockGeneralApplyReason = 'off'
        if ($stockGeneralApplyMode -eq 'apply-tested') {
            $stockGeneralApplyReady = $true
            $stockGeneralApplyReason = 'all-requested-items-reconciled'
            if (
                @($itemOrder).Count -ne [int]$companyItemCount -or
                $stockGeneralObserveCandidates.Count -ne @($itemOrder).Count -or
                [int]$stockAllWaMetrics.unsupported -ne 0 -or
                [int]$stockGeneralObserveOpeningOnlyUnsupported -ne 0
            ) {
                $stockGeneralApplyReady = $false
                $stockGeneralApplyReason = 'missing-or-unsupported-candidate'
            }
            if ($stockGeneralApplyReady) {
                foreach ($applyRawItemCode in @($itemOrder)) {
                    try {
                        $applyCode = [int]$applyRawItemCode
                        $applyKey = [string]$applyCode
                        if (
                            -not $items.ContainsKey($applyCode) -or
                            -not $stockGeneralObserveCandidates.ContainsKey($applyKey)
                        ) { throw 'candidate-absent' }
                        $applyItem = $items[$applyCode]
                        $applyCandidate = $stockGeneralObserveCandidates[$applyKey]
                        if (-not $applyCandidate.supported) { throw 'candidate-unsupported' }
                        [double]$applyRate = [double]$applyCandidate.rate
                        [double]$applyAmount = [double]$applyCandidate.amount
                        if (
                            [double]::IsNaN($applyRate) -or
                            [double]::IsInfinity($applyRate) -or
                            [double]::IsNaN($applyAmount) -or
                            [double]::IsInfinity($applyAmount)
                        ) { throw 'invalid-candidate-rate' }
                        [double]$applyPhysicalTotal = 0.0
                        foreach ($applyMcKey in @($applyItem.mc.Keys)) {
                            $applyState = $applyItem.mc[$applyMcKey]
                            [double]$applyActual = (
                                [double]$applyState.baseOpeningMain +
                                [double]$applyState.priorMain +
                                [double]$applyState.inwardMain -
                                [double]$applyState.outwardMain
                            )
                            [double]$applyCandidateMcQty = 0.0
                            if ($applyCandidate.centres.ContainsKey([string]$applyMcKey)) {
                                $applyCandidateMcQty = [double]$applyCandidate.centres[[string]$applyMcKey].quantity
                            }
                            if (
                                [double]::IsNaN($applyCandidateMcQty) -or
                                [Math]::Abs($applyActual - $applyCandidateMcQty) -gt 0.000001
                            ) { throw 'mc-quantity-mismatch' }
                            $applyPhysicalTotal += $applyActual
                        }
                        foreach ($extraMcKey in @($applyCandidate.centres.Keys)) {
                            if (
                                -not $applyItem.mc.ContainsKey([string]$extraMcKey) -and
                                [Math]::Abs([double]$applyCandidate.centres[$extraMcKey].quantity) -gt 0.000001
                            ) { throw 'unexpected-candidate-centre' }
                        }
                        if (
                            [double]::IsNaN([double]$applyCandidate.quantity) -or
                            [Math]::Abs($applyPhysicalTotal - [double]$applyCandidate.quantity) -gt 0.000001
                        ) { throw 'item-quantity-mismatch' }
                    }
                    catch {
                        $stockGeneralApplyReady = $false
                        $stockGeneralApplyReason = 'item-' + [string]$applyRawItemCode + '-refused'
                        break
                    }
                }
            }
            if (-not $stockGeneralApplyReady) {
                Write-ReportEngineLog -Stage 'STOCK-GENERAL-APPLY-GUARD' -Message (
                    'REFUSED; whole request stays legacy; reason=' + $stockGeneralApplyReason +
                    '; request-items=' + @($itemOrder).Count
                ) -Color 'Yellow'
            }
        }
        # [STOCK-GENERAL-APPLY-V4:PREFLIGHT:END]

        if ($stockAllWaMode -ne 'off') {
            Write-ReportEngineLog `
                -Stage 'STOCK-ALLWA' `
                -Message (
                    'mode=' + $stockAllWaMode +
                    '; computed=' + $stockAllWaMetrics.computed +
                    '; supported=' + $stockAllWaMetrics.supported +
                    '; unsupported=' + $stockAllWaMetrics.unsupported +
                    '; reference-eligible=' + $stockAllWaMetrics.allowed +
                    '; refused=' + $stockAllWaMetrics.refused
                ) `
                -Color 'Cyan'
        }

        if (-not $directStockConn) {
            Write-ReportEngineLog `
                -Stage 'STOCK-WA-PROFILE' `
                -Message (
                    'recordsetOpenMs=' + [math]::Round($waProfile.OpenMs,1) +
                    '; scanMs=' + [math]::Round($waProfile.ScanMs,1) +
                    '; moveNextMs=' + [math]::Round($waProfile.MoveNextMs,1) +
                    '; replayMs=' + [math]::Round($waProfile.ReplayMs,1) +
                    '; rows=' + $waProfile.Rows +
                    '; itemFlushes=' + $waProfile.Flushes
                ) `
                -Color 'Cyan'
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

            # Apply a candidate only when ALL material-centre closing quantities
            # match the old independently accumulated physical ledger. Never
            # partially override an item when one centre disagrees.
            $stockAllWaItemCandidate = $null
            $stockAllKey = [string]$item.code
            if (
                $stockAllWaMode -eq 'apply-verified' -and
                $stockAllWaSelectedByItem.ContainsKey($stockAllKey)
            ) {
                $verifyCandidate = $stockAllWaSelectedByItem[$stockAllKey]
                $matchAllMc = $true
                [double]$sumPhysical = 0.0
                foreach ($physicalMcKey in @($item.mc.Keys)) {
                    $physicalState = $item.mc[$physicalMcKey]
                    [double]$actualMcQty = (
                        [double]$physicalState.baseOpeningMain +
                        [double]$physicalState.priorMain +
                        [double]$physicalState.inwardMain -
                        [double]$physicalState.outwardMain
                    )
                    $sumPhysical += $actualMcQty
                    [double]$candidateMcQty = 0.0
                    if ($verifyCandidate.centres.ContainsKey([string]$physicalMcKey)) {
                        $candidateMcQty = [double]$verifyCandidate.centres[[string]$physicalMcKey].quantity
                    }
                    if ([Math]::Abs($actualMcQty - $candidateMcQty) -gt 0.000001) {
                        $matchAllMc = $false
                        break
                    }
                }
                if (
                    $matchAllMc -and
                    [Math]::Abs($sumPhysical - [double]$verifyCandidate.quantity) -le 0.000001
                ) {
                    $stockAllWaItemCandidate = $verifyCandidate
                }
                else {
                    $stockAllWaMetrics.refused++
                    Write-ReportEngineLog -Stage 'STOCK-ALLWA-GUARD' -Message ('MC quantity reconciliation refused for item ' + $stockAllKey) -Color 'Yellow'
                }
            }

            # [STOCK-GENERAL-APPLY-V4:ROW:BEGIN]
            # Apply one shared item-wide rate across selected MCs, exactly as
            # existing guarded reference-gated branch does, ONLY after the
            # entire requested item set passed preflight.
            if ($stockGeneralApplyMode -eq 'apply-tested' -and $stockGeneralApplyReady) {
                $stockAllWaItemCandidate = $stockGeneralObserveCandidates[$stockAllKey]
            }
            # [STOCK-GENERAL-APPLY-V4:ROW:END]

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

                # Optional valuation LAB override (item 1198 only).
                # The regular path is exactly unchanged when the env switch
                # is absent; other items continue using existing BUSY rules.
                if (
                    $stockWaLabEnabled -and
                    [int]$item.code -eq 1198 -and
                    $stockWaLabByItem.ContainsKey('1198')
                ) {
                    $mcRate = [double]$stockWaLabByItem['1198'].rate
                    $mcClosingValue = [double]$mcClosingMain * $mcRate
                }

                # Apply only a fully guarded reference-matched item candidate.
                # Never change unsupported / unverified items. The rate comes
                # from calculation, NOT from the native reference whitelist.
                if ($null -ne $stockAllWaItemCandidate) {
                    $mcRate = [double]$stockAllWaItemCandidate.rate
                    $mcClosingValue = [double]$mcClosingMain * $mcRate
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

                    purchaseQuantity = [Math]::Round([double]$state.purchaseQty, 3)
                    saleQuantity = [Math]::Round([double]$state.saleQty, 3)
                    saleReturnQuantity = [Math]::Round([double]$state.saleReturnQty, 3)
                    purchaseReturnQuantity = [Math]::Round([double]$state.purchaseReturnQty, 3)
                    transferInQuantity = [Math]::Round([double]$state.transferInQty, 3)
                    transferOutQuantity = [Math]::Round([double]$state.transferOutQty, 3)
                    transferInAltQuantity = [Math]::Round([double]$state.transferInAltQty, 3)
                    transferOutAltQuantity = [Math]::Round([double]$state.transferOutAltQty, 3)
                    productionGeneratedQuantity = [Math]::Round([double]$state.productionGeneratedQty, 3)
                    productionConsumedQuantity = [Math]::Round([double]$state.productionConsumedQty, 3)
                    stockJournalGeneratedQuantity = [Math]::Round([double]$state.stockJournalGeneratedQty, 3)
                    stockJournalConsumedQuantity = [Math]::Round([double]$state.stockJournalConsumedQty, 3)
                    materialReceiptQuantity = [Math]::Round([double]$state.materialReceiptQty, 3)
                    materialIssueQuantity = [Math]::Round([double]$state.materialIssueQty, 3)
                    otherInQuantity = [Math]::Round([double]$state.otherInQty, 3)
                    otherOutQuantity = [Math]::Round([double]$state.otherOutQty, 3)
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

        # [STOCK-GENERAL-OBSERVE-V1:DIAGNOSTIC:BEGIN]
            # Calculate a diagnostic on the already-priced item, without
            # assigning to $busyClosingValue, $stockValue, or $valuationRate.
            # No native BUSY reference amount is consulted here.
            if ($stockGeneralObserveMode -eq 'observe') {
                try {
                    $observeKey = [string]$item.code
                    if ($stockGeneralObserveCandidates.ContainsKey($observeKey)) {
                        $observeCandidate = $stockGeneralObserveCandidates[$observeKey]
                        $observeQuantityOK = $true
                        [double]$observePhysicalTotal = 0.0
                        foreach ($observeMcKey in @($item.mc.Keys)) {
                            $observePhysical = $item.mc[$observeMcKey]
                            [double]$observePhysicalQty = (
                                [double]$observePhysical.baseOpeningMain +
                                [double]$observePhysical.priorMain +
                                [double]$observePhysical.inwardMain -
                                [double]$observePhysical.outwardMain
                            )
                            $observePhysicalTotal += $observePhysicalQty
                            [double]$observeCandidateMcQty = 0.0
                            if ($observeCandidate.centres.ContainsKey([string]$observeMcKey)) {
                                $observeCandidateMcQty = [double]$observeCandidate.centres[[string]$observeMcKey].quantity
                            }
                            if ([Math]::Abs($observePhysicalQty - $observeCandidateMcQty) -gt 0.000001) {
                                $observeQuantityOK = $false
                                break
                            }
                        }
                        if (
                            $observeQuantityOK -and
                            [Math]::Abs($observePhysicalTotal - [double]$observeCandidate.quantity) -le 0.000001 -and
                            -not [double]::IsNaN([double]$observeCandidate.rate) -and
                            -not [double]::IsInfinity([double]$observeCandidate.rate)
                        ) {
                            [double]$observeViewQty = 0.0
                            foreach ($observeCentre in @($centres)) {
                                $observeViewQty += [double]$observeCentre.closingQuantity
                            }
                            [double]$observeCurrent = [Math]::Round([double]$busyClosingValue, 2, [MidpointRounding]::ToEven)
                            [double]$observeProposed = [Math]::Round(($observeViewQty * [double]$observeCandidate.rate), 2, [MidpointRounding]::ToEven)
                            [void]$stockGeneralObserveComparisons.Add([pscustomobject]@{
                                itemCode = [int]$item.code
                                legacyAmount = $observeCurrent
                                candidateAmount = $observeProposed
                                delta = [Math]::Round(($observeProposed - $observeCurrent), 2, [MidpointRounding]::ToEven)
                            })
                            # Centres here are the EXACT centres selected by
                            # the caller, including STORE-only or Damage-only.
                            # Never use per-centre diagnostics to set values.
                            foreach ($observeCentre in @($centres)) {
                                [double]$observeMcLegacy = [double]$observeCentre.stockValue
                                [double]$observeMcCandidate = [Math]::Round(
                                    ([double]$observeCentre.closingQuantity * [double]$observeCandidate.rate),
                                    2, [MidpointRounding]::ToEven
                                )
                                [void]$stockGeneralObserveCentreComparisons.Add([pscustomobject]@{
                                    itemCode = [int]$item.code
                                    mcCode = [int]$observeCentre.code
                                    mcName = [string]$observeCentre.name
                                    closingQuantity = [double]$observeCentre.closingQuantity
                                    legacyAmount = $observeMcLegacy
                                    candidateAmount = $observeMcCandidate
                                    delta = [Math]::Round(($observeMcCandidate - $observeMcLegacy), 2, [MidpointRounding]::ToEven)
                                })
                            }
                        }
                        else {
                            $stockGeneralObserveRejectedQty++
                        }
                    }
                }
                catch {
                    # Observation diagnostics must never interrupt a report.
                    $stockGeneralObserveRejectedQty++
                }
            }
        # [STOCK-GENERAL-OBSERVE-V1:DIAGNOSTIC:END]
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
                transferInAltQuantity = [Math]::Round(
                    [double]((
                        @($centres) |
                        Measure-Object -Property transferInAltQuantity -Sum
                    ).Sum),
                    3
                )
                transferOutAltQuantity = [Math]::Round(
                    [double]((
                        @($centres) |
                        Measure-Object -Property transferOutAltQuantity -Sum
                    ).Sum),
                    3
                )
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

        # [STOCK-GENERAL-APPLY-V4:SUMMARY:BEGIN]
        if ($stockGeneralApplyMode -eq 'apply-tested') {
            $applyScope = if (-not $fastPageModeActive -and @($itemOrder).Count -eq [int]$companyItemCount) {
                'full-company'
            }
            else { 'partial-request' }
            $applyDisposition = if ($stockGeneralApplyReady) { 'TEST-APPLIED' } else { 'REFUSED-LEGACY' }
            Write-ReportEngineLog -Stage 'STOCK-GENERAL-APPLY' -Message (
                $applyDisposition + '; date=' + $endDate.ToString('yyyy-MM-dd') +
                '; scope=' + $applyScope +
                '; processed-masters=' + @($itemOrder).Count +
                '; company-masters=' + $companyItemCount +
                '; candidate-items=' + $stockGeneralObserveCandidates.Count +
                '; qty-preflight=' + $stockGeneralApplyReason
            ) -Color 'Yellow'
        }
        # [STOCK-GENERAL-APPLY-V4:SUMMARY:END]

        # [STOCK-FULL-OBSERVE-V2:SUMMARY:BEGIN]
        if ($stockGeneralObserveMode -eq 'observe') {
            try {
                [double]$observeLegacySum = 0.0
                [double]$observeCandidateSum = 0.0
                $observeChangedCount = 0
                foreach ($observeRow in @($stockGeneralObserveComparisons)) {
                    $observeLegacySum += [double]$observeRow.legacyAmount
                    $observeCandidateSum += [double]$observeRow.candidateAmount
                    if ([Math]::Abs([double]$observeRow.delta) -ge 0.005) {
                        $observeChangedCount++
                    }
                }
                $observeTop = @(
                    $stockGeneralObserveComparisons |
                    Sort-Object -Property @{
                        Expression = { [Math]::Abs([double]$_.delta) }
                        Descending = $true
                    } |
                    Select-Object -First 5
                )
                $observeTopText = @(
                    foreach ($observeWorst in $observeTop) {
                        ([string]$observeWorst.itemCode + ':' +
                            ([double]$observeWorst.delta).ToString('0.00', [Globalization.CultureInfo]::InvariantCulture))
                    }
                ) -join ','
                # A paginated or search-filtered request must never claim
                # all 235 company items are included in diagnostic totals.
                $observeProcessedItems = @($itemOrder).Count
                $observeComplete = (
                    -not $fastPageModeActive -and
                    $observeProcessedItems -eq [int]$companyItemCount -and
                    $stockGeneralObserveComparisons.Count -eq [int]$companyItemCount -and
                    $stockGeneralObserveRejectedQty -eq 0
                )
                $observeScope = if ($observeComplete) { 'full-company' } else { 'partial' }
                $observeMissing = $observeProcessedItems - $stockGeneralObserveComparisons.Count
                Write-ReportEngineLog -Stage 'STOCK-GENERAL-OBSERVE' -Message (
                    'OBSERVATION ONLY; date=' + $endDate.ToString('yyyy-MM-dd') +
                    '; scope=' + $observeScope +
                    '; processed-masters=' + $observeProcessedItems +
                    '; company-masters=' + $companyItemCount +
                    '; candidate-supported-movements=' + $stockAllWaMetrics.supported +
                    '; candidate-unsupported-movements=' + $stockAllWaMetrics.unsupported +
                    '; opening-only-supported=' + $stockGeneralObserveOpeningOnlySupported +
                    '; opening-only-unsupported=' + $stockGeneralObserveOpeningOnlyUnsupported +
                    '; compared=' + $stockGeneralObserveComparisons.Count +
                    '; unobserved=' + $observeMissing +
                    '; qty-refused=' + $stockGeneralObserveRejectedQty +
                    '; retained-empty-return-events=' + $stockGeneralObserveRetainedMetric.events +
                    '; retained-return-items=' + ($stockGeneralObserveRetainedMetric.items -join ',') +
                    '; changed=' + $observeChangedCount +
                    '; legacy-sum=' + $observeLegacySum.ToString('0.00', [Globalization.CultureInfo]::InvariantCulture) +
                    '; candidate-sum=' + $observeCandidateSum.ToString('0.00', [Globalization.CultureInfo]::InvariantCulture) +
                    '; candidate-minus-legacy=' + ([Math]::Round(($observeCandidateSum - $observeLegacySum),2, [MidpointRounding]::ToEven)).ToString('0.00', [Globalization.CultureInfo]::InvariantCulture) +
                    '; top-item-deltas=' + $observeTopText
                ) -Color 'Cyan'

                # Independently report STORE/Damage/other selected centres.
                # Subtotals are sums of rounded per-centre item values,
                # whereas the all-centres item total rounds after aggregation.
                $observeMcCodes = @($stockGeneralObserveCentreComparisons | Select-Object -ExpandProperty mcCode -Unique | Sort-Object)
                $observeMcLogged = 0
                foreach ($observeMcCode in $observeMcCodes) {
                    if ($observeMcLogged -ge 20) { break }
                    $observeMcRows = @($stockGeneralObserveCentreComparisons | Where-Object { [int]$_.mcCode -eq [int]$observeMcCode })
                    if ($observeMcRows.Count -eq 0) { continue }
                    [double]$observeMcQuantity = 0.0
                    [double]$observeMcLegacySum = 0.0
                    [double]$observeMcCandidateSum = 0.0
                    $observeMcChanged = 0
                    foreach ($observeMcRow in $observeMcRows) {
                        $observeMcQuantity += [double]$observeMcRow.closingQuantity
                        $observeMcLegacySum += [double]$observeMcRow.legacyAmount
                        $observeMcCandidateSum += [double]$observeMcRow.candidateAmount
                        if ([Math]::Abs([double]$observeMcRow.delta) -ge 0.005) { $observeMcChanged++ }
                    }
                    $observeMcName = ([string]$observeMcRows[0].mcName).Replace(';', ',')
                    Write-ReportEngineLog -Stage 'STOCK-GENERAL-OBSERVE-MC' -Message (
                        'OBSERVATION ONLY; date=' + $endDate.ToString('yyyy-MM-dd') +
                        '; scope=' + $observeScope +
                        '; mc-code=' + $observeMcCode +
                        '; mc-name=' + $observeMcName +
                        '; compared-items=' + $observeMcRows.Count +
                        '; quantity=' + $observeMcQuantity.ToString('0.000', [Globalization.CultureInfo]::InvariantCulture) +
                        '; changed=' + $observeMcChanged +
                        '; legacy-sum=' + $observeMcLegacySum.ToString('0.00', [Globalization.CultureInfo]::InvariantCulture) +
                        '; candidate-sum=' + $observeMcCandidateSum.ToString('0.00', [Globalization.CultureInfo]::InvariantCulture) +
                        '; candidate-minus-legacy=' + ([Math]::Round(($observeMcCandidateSum - $observeMcLegacySum),2, [MidpointRounding]::ToEven)).ToString('0.00', [Globalization.CultureInfo]::InvariantCulture)
                    ) -Color 'Cyan'
                    $observeMcLogged++
                }
                if ($observeMcCodes.Count -gt 20) {
                    Write-ReportEngineLog -Stage 'STOCK-GENERAL-OBSERVE-MC' -Message (
                        'OBSERVATION ONLY; material-centre log limit=20; available-centres=' + $observeMcCodes.Count
                    ) -Color 'Yellow'
                }
            }
            catch {
                # Diagnostic failure must never affect returned Stock Status.
            }
        }
        # [STOCK-FULL-OBSERVE-V2:SUMMARY:END]
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
                    if ($stockGeneralApplyMode -eq 'apply-tested' -and $stockGeneralApplyReady) {
                        'busy-general-wa-v23-guarded-test'
                    }
                    else { 'busy-stockvalmethod-5-v11-single-snapshot' }
                }
                else {
                    "item-master-$valueMode-price"
                }

                stockValuationMethod = 5
                stockValuationMode =
                    "busy-method5-v11-single-snapshot-by-material-centre"

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
            valuationSource = if ($valueMode -eq "busy") { if ($stockGeneralApplyMode -eq 'apply-tested' -and $stockGeneralApplyReady) {
                        'busy-general-wa-v23-guarded-test'
                    }
                    else { 'busy-stockvalmethod-5-v11-single-snapshot' } } else { "item-master-$valueMode-price" }
            stockValuationMethod = 5
            stockValuationMode = "busy-method5-v11-single-snapshot-by-material-centre"
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
