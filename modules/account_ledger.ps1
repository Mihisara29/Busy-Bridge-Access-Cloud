# modules/account_ledger.ps1
# Account Ledger - Phase 1 (corrected 2026-10-12): single-account postings + balances.
# Loaded AFTER reports.ps1 and BEFORE routes.ps1.
# SQL Server: direct company/year database connection for ledger rows.
# Access/BDS: BUSY COM GetRecordset. Master opening balance: BUSY GetMasterXML.
# Reference values from Gregory and GS Lubricants; live API verification still required.
# Fix: culture-independent yyyy-MM-dd / dd-MM-yyyy / dd/MM/yyyy parsing for Windows PowerShell 5.1.
# Fix: undefined/blank Account.OPBal resolves to zero for valid account masters.


# --------------------------------------------------------------------------
# Read-only performance instrumentation. One request is processed at a time
# in Start-BUSYServer, so a script-scoped active request is safe in this bridge.
# No credentials, JWTs, query text, or account balances are logged.
# --------------------------------------------------------------------------
$script:LedgerActivePerformance = $null

function Start-LedgerPerformance {
    param([string]$Mode, [string]$InstanceId, [string]$CompanyCode)
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $script:LedgerActivePerformance = [pscustomobject]@{
        Mode = $Mode
        Instance = $InstanceId
        Company = $CompanyCode
        Timer = $timer
        LastMs = [long]0
        Steps = [ordered]@{}
        Queries = @{}
        XmlCalls = 0
        XmlMs = [long]0
        CacheHits = 0
    }
    Write-Host ('[LEDGER-TIME] START {0}  mode={1}  company={2}/{3}' -f `
        (Get-Date -Format 'HH:mm:ss.fff'), $Mode, $InstanceId, $CompanyCode) -ForegroundColor Cyan
}

function Mark-LedgerPerformance {
    param([string]$Name)
    $ctx = $script:LedgerActivePerformance
    if ($null -eq $ctx) { return }
    $now = [long]$ctx.Timer.ElapsedMilliseconds
    $ctx.Steps[$Name] = $now - $ctx.LastMs
    $ctx.LastMs = $now
}

function Record-LedgerDatabaseTime {
    param([string]$Category, [long]$Milliseconds, [int]$Rows)
    $ctx = $script:LedgerActivePerformance
    if ($null -eq $ctx) { return }
    if (-not $ctx.Queries.ContainsKey($Category)) {
        $ctx.Queries[$Category] = [pscustomobject]@{ Count = 0; Milliseconds = [long]0; Rows = 0 }
    }
    $item = $ctx.Queries[$Category]
    $item.Count++
    $item.Milliseconds += $Milliseconds
    $item.Rows += $Rows
}

function Record-LedgerXmlTime {
    param([long]$Milliseconds)
    $ctx = $script:LedgerActivePerformance
    if ($null -eq $ctx) { return }
    $ctx.XmlCalls++
    $ctx.XmlMs += $Milliseconds
}

function Stop-LedgerPerformance {
    param([bool]$Succeeded)
    $ctx = $script:LedgerActivePerformance
    if ($null -eq $ctx) { return }
    $ctx.Timer.Stop()
    $total = [long]$ctx.Timer.ElapsedMilliseconds
    $state = if ($Succeeded) { 'SUCCESS' } else { 'FAILED' }
    Write-Host ('[LEDGER-TIME] {0} mode={1}  total={2:N0} ms ({3:N2} s)' -f `
        $state, $ctx.Mode, $total, ($total / 1000.0)) -ForegroundColor $(if ($Succeeded) { 'Green' } else { 'Red' })
    foreach ($stage in $ctx.Steps.Keys) {
        Write-Host ('[LEDGER-TIME]    {0,-29} {1,10:N0} ms' -f $stage, $ctx.Steps[$stage]) -ForegroundColor Gray
    }
    $otherMs = $total - $ctx.LastMs
    if ($otherMs -ge 0) {
        Write-Host ('[LEDGER-TIME]    {0,-29} {1,10:N0} ms' -f 'Return + cleanup', $otherMs) -ForegroundColor Gray
    }
    foreach ($category in @($ctx.Queries.Keys | Sort-Object)) {
        $q = $ctx.Queries[$category]
        Write-Host ('[LEDGER-TIME]    DB {0,-26} {1,9:N0} ms  queries={2} rows={3:N0}' -f `
            $category, $q.Milliseconds, $q.Count, $q.Rows) -ForegroundColor DarkCyan
    }
    Write-Host ('[LEDGER-TIME]    COM GetMasterXML                {0,10:N0} ms  calls={1} cacheHits={2}' -f `
        $ctx.XmlMs, $ctx.XmlCalls, $ctx.CacheHits) -ForegroundColor DarkCyan
    Write-Host '[LEDGER-TIME] NOTE: DB and COM subtotals are INCLUDED in the stages above (do not add them again).' -ForegroundColor DarkGray
    $script:LedgerActivePerformance = $null
}

function ConvertTo-AccountLedgerDate {
    param([string]$Value, [datetime]$Default)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $Default.Date }

    $dateText = $Value.Trim()

    # Parse the three explicit formats without culture-dependent overload resolution.
    # A numeric DateTime constructor also rejects impossible dates such as 2026-02-30.
    $year = 0
    $month = 0
    $day = 0

    if ($dateText -match '^(\d{4})-(\d{2})-(\d{2})$') {
        $year = [int]$Matches[1]
        $month = [int]$Matches[2]
        $day = [int]$Matches[3]
    }
    elseif ($dateText -match '^(\d{2})-(\d{2})-(\d{4})$' -or
            $dateText -match '^(\d{2})/(\d{2})/(\d{4})$') {
        $day = [int]$Matches[1]
        $month = [int]$Matches[2]
        $year = [int]$Matches[3]
    }
    else {
        throw "Invalid ledger date '$Value'. Use yyyy-MM-dd."
    }

    try {
        return [datetime]::new($year, $month, $day).Date
    }
    catch {
        throw "Invalid ledger date '$Value'. Use yyyy-MM-dd."
    }
}

function ConvertTo-AccountLedgerBalance {
    param([decimal]$Signed)
    $kind = 'Dr'
    if ($Signed -lt 0) { $kind = 'Cr' }
    if ($Signed -eq 0) { $kind = '-' }
    return [pscustomobject]@{ amount = [math]::Abs($Signed); type = $kind; signed = $Signed }
}

function Get-AccountLedgerRows {
    param(
        [string]$Query,
        [bool]$IsSql,
        $SqlConnection,
        $BusyConnection,
        [int]$Timeout = 120
    )
    $rows = [System.Collections.Generic.List[object]]::new()
    $queryTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $queryKind = if ($Query -match '(?i)FROM\s+Master1') { 'Master1' }
        elseif ($Query -match '(?i)FROM\s+Tran2\s+AS\s+P' -and
                $Query -match '(?i)GROUP\s+BY\s+P\.MasterCode1') { 'Tran2-aggregate' }
        elseif ($Query -match '(?i)FROM\s+Tran2\s+AS\s+P') { 'Tran2-postings' }
        elseif ($Query -match '(?i)FROM\s+Tran2\s+AS\s+O') { 'Tran2-counterparts' }
        else { 'other' }
    if ($IsSql) {
        $command = $null
        $reader = $null
        try {
            $command = $SqlConnection.CreateCommand()
            $command.CommandText = $Query
            $command.CommandTimeout = $Timeout
            $reader = $command.ExecuteReader()
            while ($reader.Read()) {
                $fields = [ordered]@{}
                for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                    $value = $reader.GetValue($i)
                    if ($value -eq [System.DBNull]::Value) { $value = $null }
                    $fields[$reader.GetName($i)] = $value
                }
                $rows.Add([pscustomobject]$fields)
            }
        }
        finally {
            if ($null -ne $reader) { $reader.Dispose() }
            if ($null -ne $command) { $command.Dispose() }
        }
    }
    else {
        $rst = $null
        try {
            $rst = $BusyConnection.GetRecordset($Query)
            if ($null -ne $rst -and -not $rst.EOF) {
                $rst.MoveFirst()
                while (-not $rst.EOF) {
                    $fields = [ordered]@{}
                    for ($i = 0; $i -lt $rst.Fields.Count; $i++) {
                        $field = $rst.Fields.Item($i)
                        $value = $field.Value
                        if ($value -eq [System.DBNull]::Value) { $value = $null }
                        $fields[[string]$field.Name] = $value
                    }
                    $rows.Add([pscustomobject]$fields)
                    $rst.MoveNext()
                }
            }
        }
        finally { if ($null -ne $rst) { try { $rst.Close() } catch {} } }
    }
    $queryTimer.Stop()
    Record-LedgerDatabaseTime -Category $queryKind -Milliseconds $queryTimer.ElapsedMilliseconds -Rows $rows.Count
    return $rows.ToArray()
}

# BUSY COM can serialize tiny floating-point residuals using E notation.
# Accept ordinary decimals AND exponent notation with invariant culture.
# Ignore only amounts less than one millionth of a rupee in magnitude:
# these are COM floating-point noise, not genuine cent-level balances.
# Missing/blank remains zero, whereas malformed/non-finite/overflow values fail.
function ConvertFrom-AccountLedgerXmlOpening {
    param(
        [AllowNull()][string]$OpeningText,
        [int]$AccountCode,
        [string]$AccountName = ''
    )
    $text = ([string]$OpeningText).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return [pscustomobject]@{
            raw = [decimal]0
            assumedZero = $true
            normalizedTiny = $false
        }
    }
    [decimal]$raw = 0
    $style = [System.Globalization.NumberStyles]::Number -bor
        [System.Globalization.NumberStyles]::AllowExponent
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    if (-not [decimal]::TryParse($text, $style, $culture, [ref]$raw)) {
        $accountLabel = if ([string]::IsNullOrWhiteSpace($AccountName)) {
            "account code $AccountCode"
        } else {
            "account $AccountCode ($AccountName)"
        }
        throw "Invalid Account.OPBal value for ${accountLabel}: '$text'."
    }
    $normalizedTiny = ($raw -ne [decimal]0 -and
        [math]::Abs($raw) -lt [decimal]0.000001)
    if ($normalizedTiny) { $raw = [decimal]0 }
    return [pscustomobject]@{
        raw = $raw
        assumedZero = $false
        normalizedTiny = $normalizedTiny
    }
}

function Get-AccountLedgerReport {
    param(
        [int]$AccountCode = 0,
        [string]$Account = '',
        [string]$From = '',
        [string]$To = '',
        [bool]$IncludePdc = $true,
        [int]$Page = 1,
        [int]$PageSize = 100,
        [string]$InstanceId = '',
        [string]$CompanyCode = ''
    )

    Start-LedgerPerformance -Mode 'one' -InstanceId $InstanceId -CompanyCode $CompanyCode
    $ledgerPerfSuccess = $false
    $fi = $null
    $sqlConnection = $null
    try {
        # In this rollout BUSY master XML reflects the *current* financial year.
        # Explicitly reject other years rather than display incorrect opening balances.
        $today = (Get-Date).Date
        $fyYear = if ($today.Month -lt 4) { $today.Year - 1 } else { $today.Year }
        $fyStart = [datetime]::new($fyYear, 4, 1)
        $fyEndExclusive = $fyStart.AddYears(1)
        $fromDate = ConvertTo-AccountLedgerDate -Value $From -Default $fyStart
        $toDate = ConvertTo-AccountLedgerDate -Value $To -Default $today
        if ($fromDate -gt $toDate) { throw 'From date must not be later than To date.' }
        if ($fromDate -lt $fyStart -or $toDate -ge $fyEndExclusive) {
            throw ('This first version supports the active BUSY financial year only: ' +
                $fyStart.ToString('yyyy-MM-dd') + ' through ' +
                $fyEndExclusive.AddDays(-1).ToString('yyyy-MM-dd') + '.')
        }
        if ($Page -lt 1) { $Page = 1 }
        if ($PageSize -lt 0) { $PageSize = 100 }
        if ($PageSize -gt 500) { $PageSize = 500 }

        Mark-LedgerPerformance -Name 'Validate parameters'
        $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
        if ($null -eq $fi) { throw 'Unable to connect to BUSY company.' }

        if ($AccountCode -le 0 -and -not [string]::IsNullOrWhiteSpace($Account)) {
            $AccountCode = [int]$fi.MasterName2Code($Account.Trim(), 2)
        }
        if ($AccountCode -le 0) { throw 'A valid accountCode (or exact account name) is required.' }

        Mark-LedgerPerformance -Name 'Connect BUSY'
        $xmlTimer = [System.Diagnostics.Stopwatch]::StartNew()
        try { $xmlText = $fi.GetMasterXML($AccountCode) }
        finally {
            $xmlTimer.Stop()
            Record-LedgerXmlTime -Milliseconds $xmlTimer.ElapsedMilliseconds
        }
        if ([string]::IsNullOrWhiteSpace([string]$xmlText)) {
            throw "BUSY did not return an account master XML for code $AccountCode."
        }
        $masterXml = [xml]$xmlText
        if ($null -eq $masterXml.Account -or
            [string]::IsNullOrWhiteSpace([string]$masterXml.Account.Name)) {
            throw "Master code $AccountCode is not a valid accounting account."
        }
        # BUSY sometimes omits Account.OPBal entirely for a valid zero-opening account.
        # Missing/blank opening is ZERO, not a fatal error. Invalid nonblank values
        # still fail rather than silently misstate financial balances.
        # BUSY XML: negative opening = debit; report: positive signed = debit.
        $openingText = ([string]$masterXml.Account.OPBal).Trim()
        $opening = ConvertFrom-AccountLedgerXmlOpening -OpeningText $openingText -AccountCode $AccountCode
        $openingBalanceAssumedZero = [bool]$opening.assumedZero
        [decimal]$rawOpening = [decimal]$opening.raw
        if ($openingBalanceAssumedZero) {
            Write-Host "[ACCOUNT-LEDGER] No OPBal for account $AccountCode; using zero opening." -ForegroundColor DarkYellow
        }
        $financialYearOpening = -$rawOpening
        $accountName = [string]$masterXml.Account.Name
        $groupName = [string]$masterXml.Account.ParentGroup

        Mark-LedgerPerformance -Name 'Account XML opening'
        $mode = Get-ReportDatabaseMode -InstanceId $InstanceId
        $isSql = [bool]$mode.isSql
        if ($isSql) {
            $direct = Get-DirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
            if ($null -eq $direct -or $null -eq $direct.connection) {
                throw 'SQL Server company connection could not be created.'
            }
            $sqlConnection = $direct.connection
            if ($sqlConnection.State.ToString() -ne 'Open') { $sqlConnection.Open() }
        }

        Mark-LedgerPerformance -Name 'Open fiscal database'
        $startSql = Get-ReportSqlDateLiteral -Date $fyStart -IsSql $isSql
        $endSql = Get-ReportSqlDateLiteral -Date $toDate.AddDays(1) -IsSql $isSql
        $recTypes = if ($IncludePdc) { '1,5' } else { '1' }

        # A voucher is uniquely keyed by BOTH VchType and VchCode.
        # Tran2.Value1: negative -> Debit, positive -> Credit.
        # Never use Tran2 RecType 17 (or inventory/bill reference rows) as posting rows.
        $query = @"
SELECT P.VchType, P.VchCode, P.SrNo, P.RecType,
       H.[Date] AS VoucherDate, H.VchNo AS VoucherNo,
       P.Value1 AS SignedPosting, P.ShortNar AS ShortNarration
FROM Tran2 AS P
INNER JOIN Tran1 AS H
    ON H.VchType = P.VchType AND H.VchCode = P.VchCode
WHERE P.MasterCode1 = $AccountCode
  AND P.RecType IN ($recTypes)
  AND P.Value1 <> 0
  AND (H.Cancelled = 0 OR H.Cancelled IS NULL)
  AND (H.VchCancelled = 0 OR H.VchCancelled IS NULL)
  AND H.[Date] >= $startSql
  AND H.[Date] < $endSql
ORDER BY H.[Date], H.VchType, H.VchCode, P.SrNo, P.RecType, P.Value1
"@
        $dbRows = @(Get-AccountLedgerRows -Query $query -IsSql $isSql `
            -SqlConnection $sqlConnection -BusyConnection $fi)

        Mark-LedgerPerformance -Name 'Voucher SQL / COM fetch'
        $periodOpening = $financialYearOpening
        $running = $financialYearOpening
        [decimal]$debitTotal = 0
        [decimal]$creditTotal = 0
        $periodRows = [System.Collections.Generic.List[object]]::new()

        foreach ($dbRow in $dbRows) {
            $date = ([datetime]$dbRow.VoucherDate).Date
            $signedDb = [decimal]$dbRow.SignedPosting
            $change = -$signedDb
            if ($date -lt $fromDate) {
                $periodOpening += $change
                $running += $change
                continue
            }

            $debit = [decimal]0
            $credit = [decimal]0
            if ($change -gt 0) { $debit = $change }
            if ($change -lt 0) { $credit = -$change }
            $debitTotal += $debit
            $creditTotal += $credit
            $running += $change
            $balance = ConvertTo-AccountLedgerBalance -Signed $running
            $voucherType = [int]$dbRow.VchType
            $voucherCode = [int]$dbRow.VchCode
            $periodRows.Add([pscustomobject]@{
                voucherDate       = $date.ToString('yyyy-MM-dd')
                voucherType       = $voucherType
                voucherCode       = $voucherCode
                voucherNo         = ([string]$dbRow.VoucherNo).Trim()
                postingSrNo       = [int]$dbRow.SrNo
                postingRecType    = [int]$dbRow.RecType
                isPdc             = ([int]$dbRow.RecType -eq 5)
                debit             = $debit
                credit            = $credit
                balance           = $balance.amount
                balanceType       = $balance.type
                signedBalance     = $balance.signed
                shortNarration    = [string]$dbRow.ShortNarration
                counterpartAccount = ''
                counterpartAccounts = @()
            })
        }

        $totalRows = $periodRows.Count
        $offset = if ($PageSize -eq 0) { 0 } else { [long]($Page - 1) * $PageSize }
        $pageRows = [System.Collections.Generic.List[object]]::new()
        for ($i = $offset; $i -lt $totalRows -and ($PageSize -eq 0 -or $i -lt ($offset + $PageSize)); $i++) {
            $pageRows.Add($periodRows[[int]$i])
        }

        Mark-LedgerPerformance -Name 'Balances + page selection'
        # Retrieve counterparty labels ONLY for the displayed page.
        # Keep supplementary postings separate so no row/money can multiply.
        if ($pageRows.Count -gt 0) {
            $needed = @{}
            foreach ($r in $pageRows) { $needed["$($r.voucherType)|$($r.voucherCode)"] = $true }
            $codes = @($pageRows | ForEach-Object { [int]$_.voucherCode } | Sort-Object -Unique)
            $counterNames = @{}
            for ($chunk = 0; $chunk -lt $codes.Count; $chunk += 100) {
                $last = [math]::Min($chunk + 99, $codes.Count - 1)
                $numbers = ($codes[$chunk..$last] -join ',')
                $othersQuery = @"
SELECT O.VchType, O.VchCode, O.MasterCode1, O.SrNo,
       M.Name AS CounterName
FROM Tran2 AS O
INNER JOIN Master1 AS M ON M.Code = O.MasterCode1
WHERE M.MasterType = 2
  AND O.RecType IN ($recTypes)
  AND O.Value1 <> 0
  AND O.VchCode IN ($numbers)
ORDER BY O.VchType, O.VchCode, O.SrNo, O.MasterCode1
"@
                $otherRows = @(Get-AccountLedgerRows -Query $othersQuery -IsSql $isSql `
                    -SqlConnection $sqlConnection -BusyConnection $fi)
                foreach ($o in $otherRows) {
                    $k = "$($o.VchType)|$($o.VchCode)"
                    if (-not $needed.ContainsKey($k) -or [int]$o.MasterCode1 -eq $AccountCode) { continue }
                    if (-not $counterNames.ContainsKey($k)) { $counterNames[$k] = [System.Collections.Generic.List[string]]::new() }
                    $nm = ([string]$o.CounterName).Trim()
                    if ($nm -and -not $counterNames[$k].Contains($nm)) { $counterNames[$k].Add($nm) }
                }
            }
            foreach ($r in $pageRows) {
                $k = "$($r.voucherType)|$($r.voucherCode)"
                if ($counterNames.ContainsKey($k)) {
                    $r.counterpartAccounts = $counterNames[$k].ToArray()
                    $r.counterpartAccount = $counterNames[$k].ToArray() -join ', '
                }
            }
        }

        Mark-LedgerPerformance -Name 'Counterparty lookup'
        $closing = ConvertTo-AccountLedgerBalance -Signed $running
        $openBal = ConvertTo-AccountLedgerBalance -Signed $periodOpening
        $fyOpenBal = ConvertTo-AccountLedgerBalance -Signed $financialYearOpening
        $totalPages = if ($PageSize -eq 0) { 1 } else { [math]::Max(1, [math]::Ceiling($totalRows / $PageSize)) }
        $postingTypes = @(1)
        if ($IncludePdc) { $postingTypes = @(1, 5) }
        $databaseLabel = if ($isSql) { 'sql-server-direct' } else { 'access-busy-com' }

        Mark-LedgerPerformance -Name 'Final summary'
        $ledgerPerfSuccess = $true
        return @{
            success = $true
            data = @{
                report = 'account-ledger'
                mode = 'one'
                account = @{ code = $AccountCode; name = $accountName; group = $groupName }
                from = $fromDate.ToString('yyyy-MM-dd')
                to = $toDate.ToString('yyyy-MM-dd')
                financialYearStart = $fyStart.ToString('yyyy-MM-dd')
                includePdc = $IncludePdc
                openingBalance = $openBal.amount
                openingBalanceType = $openBal.type
                financialYearOpeningBalance = $fyOpenBal.amount
                financialYearOpeningBalanceType = $fyOpenBal.type
                openingBalanceAssumedZero = $openingBalanceAssumedZero
                totalDebit = $debitTotal
                totalCredit = $creditTotal
                closingBalance = $closing.amount
                closingBalanceType = $closing.type
                signedClosingBalance = $closing.signed
                rows = $pageRows.ToArray()
                pagination = @{
                    page = $Page
                    pageSize = $PageSize
                    totalRows = $totalRows
                    totalPages = $totalPages
                    hasPreviousPage = ($Page -gt 1)
                    hasNextPage = ($PageSize -ne 0 -and $Page -lt $totalPages)
                }
                diagnostics = @{
                    postingRecTypes = $postingTypes
                    sorting = 'date-type-voucher-srno (deterministic; native within-day order pending validation)'
                    openingBalanceSource = 'BUSY.GetMasterXML(Account.OPBal)'
                    databaseMode = $databaseLabel
                }
            }
        }
    }
    catch {
        Write-Host "[ACCOUNT-LEDGER] $($_.Exception.Message)" -ForegroundColor Red
        return @{ success = $false; error = $_.Exception.Message }
    }
    finally {
        if ($null -ne $sqlConnection) {
            try { $sqlConnection.Close() } catch {}
            try { $sqlConnection.Dispose() } catch {}
        }
        # Disconnect-BUSY in this bridge is intentionally a no-op (pooled COM).
        if ($null -ne $fi) { Disconnect-BUSY $fi }
        Stop-LedgerPerformance -Succeeded $ledgerPerfSuccess
    }
}

# ==========================================================================
# ACCOUNT LEDGER PHASE 3 - GROUP OF ACCOUNTS
# Preserves Get-AccountLedgerReport (One Account) without modification.
# Account groups are Master1.MasterType=1 (NOT 5, which is Item Group).
# Account entries are Master1.MasterType=2 and link through ParentGrp.
# ==========================================================================

# Shared multi-account opening cache (current financial year only).
# BUSY master XML is the authoritative balance source; Master1.D1 is credit limit.
# TTL + refreshOpenings guard against stale changes to account masters.
if ($null -eq (Get-Variable -Name AccountLedgerOpeningCache -Scope Script -ErrorAction SilentlyContinue)) {
    $script:AccountLedgerOpeningCache = @{}
}

function Get-AccountLedgerCachedOpening {
    param(
        $BusyConnection,
        [int]$AccountCode,
        [string]$AccountName,
        [string]$InstanceId,
        [string]$CompanyCode,
        [datetime]$FinancialYearStart,
        [bool]$Refresh = $false
    )
    # One independent cache namespace for every BUSY instance/company/FY.
    $key = ('{0}|{1}|{2}|{3}' -f $InstanceId, $CompanyCode, $FinancialYearStart.ToString('yyyy-MM-dd'), $AccountCode)
    $now = [datetime]::UtcNow
    if (-not $Refresh -and $script:AccountLedgerOpeningCache.ContainsKey($key)) {
        $entry = $script:AccountLedgerOpeningCache[$key]
        if ($entry.expiresAt -gt $now) {
            if ($null -ne $script:LedgerActivePerformance) { $script:LedgerActivePerformance.CacheHits++ }
            return [pscustomobject]@{
                signed = [decimal]$entry.signed
                assumedZero = [bool]$entry.assumedZero
                cacheHit = $true
            }
        }
    }
    $xmlTimer = [System.Diagnostics.Stopwatch]::StartNew()
    try { $xmlText = [string]$BusyConnection.GetMasterXML($AccountCode) }
    finally {
        $xmlTimer.Stop()
        Record-LedgerXmlTime -Milliseconds $xmlTimer.ElapsedMilliseconds
    }
    $openingText = ''
    if (-not [string]::IsNullOrWhiteSpace($xmlText)) {
        $xml = [xml]$xmlText
        if ($null -ne $xml.Account) {
            $openingText = ([string]$xml.Account.OPBal).Trim()
        }
    }
    $opening = ConvertFrom-AccountLedgerXmlOpening -OpeningText $openingText `
        -AccountCode $AccountCode -AccountName $AccountName
    $missing = [bool]$opening.assumedZero
    [decimal]$signed = -[decimal]$opening.raw
    $script:AccountLedgerOpeningCache[$key] = [pscustomobject]@{
        signed = $signed
        assumedZero = $missing
        expiresAt = $now.AddMinutes(30)
    }
    return [pscustomobject]@{ signed = $signed; assumedZero = $missing; cacheHit = $false }
}

# Phase 7B: STRICT database pagination. No company-wide account rows,
# voucher aggregates, report-section arrays, or opening XML reads are performed.
# The only company-wide queries are COUNT(*) and (for Group mode) MasterType=1
# group metadata. Account details are read via SQL OFFSET/FETCH or Jet TOP/NOT IN.
# Deliberately reports PAGE financial totals. Full-company financial totals cannot
# be exact without reading all native BUSY opening balances: use pageSize=0 for
# a separate explicit all-record export, or use the legacy nonpaging endpoint.
function Get-AccountLedgerTrueDbPage {
    param(
        $BusyConnection,
        $SqlConnection,
        [bool]$IsSql,
        [bool]$AllAccounts,
        [int]$GroupCode,
        [string]$Group,
        [bool]$IncludeChildGroups,
        [bool]$IncludePdc,
        [string]$AccountsToInclude,
        [datetime]$FyStart,
        [datetime]$FromDate,
        [datetime]$ToDate,
        [int]$Page,
        [int]$PageSize,
        [bool]$RefreshOpenings,
        [string]$InstanceId,
        [string]$CompanyCode
    )
    # Filtering out dormant zero-closing accounts *before* paging requires
    # reading opening balances from earlier pages. Refuse the deceptive result.
    if ($AccountsToInclude -ne 'all') {
        throw 'Strict DB pagination requires accountsToInclude=all (including dormant accounts). moved-or-closing needs a full native opening-balance scan; use the legacy full-report operation instead.'
    }
    $groupRows = @()
    $groupCodes = [System.Collections.Generic.HashSet[int]]::new()
    $groupsByCode = @{}
    $groupRoot = $null
    if (-not $AllAccounts) {
        $groupRows = @(Get-AccountLedgerRows -Query 'SELECT Code, [Name], ParentGrp FROM Master1 WHERE MasterType=1 ORDER BY [Name], [Code]' `
            -IsSql $IsSql -SqlConnection $SqlConnection -BusyConnection $BusyConnection)
        $children = @{}
        foreach ($g in $groupRows) {
            if ($null -eq $g) { continue }
            $gc = [int]$g.Code
            $groupsByCode[$gc] = [pscustomobject]@{ code=$gc; name=([string]$g.Name).Trim(); parentCode=[int]$g.ParentGrp }
            $parent = [int]$g.ParentGrp
            if (-not $children.ContainsKey($parent)) { $children[$parent] = [System.Collections.Generic.List[int]]::new() }
            $children[$parent].Add($gc)
        }
        if ($GroupCode -le 0) {
            $matching = @($groupsByCode.Values | Where-Object { [string]::Equals($_.name, $Group.Trim(), [System.StringComparison]::OrdinalIgnoreCase) })
            if ($matching.Count -eq 0) { throw "Accounting group '$Group' was not found (MasterType=1)." }
            if ($matching.Count -gt 1) { throw "More than one accounting group named '$Group'; use groupCode." }
            $GroupCode = [int]$matching[0].code
        }
        if (-not $groupsByCode.ContainsKey($GroupCode)) { throw "Accounting group code $GroupCode was not found (MasterType=1)." }
        $groupRoot = $groupsByCode[$GroupCode]
        if (-not [string]::IsNullOrWhiteSpace($Group) -and
            -not [string]::Equals($Group.Trim(), $groupRoot.name, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw 'groupCode and group name identify different groups.'
        }
        $queue = [System.Collections.Generic.Queue[int]]::new()
        $queue.Enqueue($GroupCode)
        while ($queue.Count -gt 0) {
            $g = $queue.Dequeue()
            if (-not $groupCodes.Add($g)) { continue }
            if ($IncludeChildGroups -and $children.ContainsKey($g)) {
                foreach ($child in $children[$g]) { $queue.Enqueue([int]$child) }
            }
        }
    }

    # SQL and Jet interpret NULL in a negated conjunction differently.
    # Explicit IS NULL predicates preserve the verified linked-master rule.
    $predicate = 'MasterType=2 AND ((I10 IS NULL OR I10<>1) OR (CM5 IS NULL OR CM5<=0))'
    if (-not $AllAccounts) {
        $groupIds = (@($groupCodes | Sort-Object) -join ',')
        $predicate += " AND ParentGrp IN ($groupIds)"
    }
    $counts = @(Get-AccountLedgerRows -Query "SELECT COUNT(*) AS TotalCount FROM Master1 WHERE $predicate" `
        -IsSql $IsSql -SqlConnection $SqlConnection -BusyConnection $BusyConnection)
    $totalAccounts = if ($counts.Count -gt 0) { [int]$counts[0].TotalCount } else { 0 }
    $sourceCounts = @(Get-AccountLedgerRows -Query 'SELECT COUNT(*) AS TotalCount FROM Master1 WHERE MasterType=2' `
        -IsSql $IsSql -SqlConnection $SqlConnection -BusyConnection $BusyConnection)
    $sourceAccountCount = if ($sourceCounts.Count -gt 0) { [int]$sourceCounts[0].TotalCount } else { 0 }
    $rawIndependent = @(Get-AccountLedgerRows -Query 'SELECT COUNT(*) AS TotalCount FROM Master1 WHERE MasterType=2 AND ((I10 IS NULL OR I10<>1) OR (CM5 IS NULL OR CM5<=0))' `
        -IsSql $IsSql -SqlConnection $SqlConnection -BusyConnection $BusyConnection)
    $independentCount = if ($rawIndependent.Count -gt 0) { [int]$rawIndependent[0].TotalCount } else { 0 }
    $excludedCount = $sourceAccountCount - $independentCount
    Mark-LedgerPerformance -Name 'Group metadata + account COUNT'

    [long]$skip = ([long]$Page - 1L) * [long]$PageSize
    $pageRows = @()
    if ($skip -lt $totalAccounts) {
        $orderBy = '[Name], [Code]'
        if ($IsSql) {
            $pageQuery = "SELECT [Code], [Name], ParentGrp AS GroupCode FROM Master1 WHERE $predicate ORDER BY $orderBy OFFSET $skip ROWS FETCH NEXT $PageSize ROWS ONLY"
        }
        elseif ($skip -eq 0) {
            $pageQuery = "SELECT TOP $PageSize [Code], [Name], ParentGrp AS GroupCode FROM Master1 WHERE $predicate ORDER BY $orderBy"
        }
        else {
            # Jet/BDS supports TOP in a subquery; NOT IN excludes exactly the
            # previous page ranges. Unique [Code] prevents ambiguity on tied names.
            $pageQuery = "SELECT TOP $PageSize [Code], [Name], ParentGrp AS GroupCode FROM Master1 WHERE $predicate AND [Code] NOT IN (SELECT TOP $skip [Code] FROM Master1 WHERE $predicate ORDER BY $orderBy) ORDER BY $orderBy"
        }
        $pageRows = @(Get-AccountLedgerRows -Query $pageQuery -IsSql $IsSql `
            -SqlConnection $SqlConnection -BusyConnection $BusyConnection)
    }
    $pageMembers = @($pageRows | Where-Object { $null -ne $_ -and $null -ne $_.Code })
    if ($AllAccounts -and $pageMembers.Count -gt 0) {
        # Resolve group labels ONLY for parents referenced by this account page.
        $pageGroupIds = @($pageMembers | ForEach-Object { [int]$_.GroupCode } | Where-Object { $_ -gt 0 } | Sort-Object -Unique)
        for ($gAt = 0; $gAt -lt $pageGroupIds.Count; $gAt += 100) {
            $gEnd = [math]::Min($gAt + 100,$pageGroupIds.Count)
            $ids = (@($pageGroupIds[$gAt..($gEnd - 1)]) -join ',')
            foreach ($g in @(Get-AccountLedgerRows -Query "SELECT Code, [Name], ParentGrp FROM Master1 WHERE MasterType=1 AND Code IN ($ids)" `
                -IsSql $IsSql -SqlConnection $SqlConnection -BusyConnection $BusyConnection)) {
                $groupsByCode[[int]$g.Code] = [pscustomobject]@{
                    code=[int]$g.Code; name=([string]$g.Name).Trim(); parentCode=[int]$g.ParentGrp
                }
            }
        }
    }
    Write-Host ('[LEDGER-DB-PAGE] company={0}/{1} mode={2} page={3} size={4} offset={5} pageMasters={6} totalMasters={7}' -f `
        $InstanceId,$CompanyCode,$(if ($AllAccounts) {'all'} else {'group'}),$Page,$PageSize,$skip,$pageMembers.Count,$totalAccounts) -ForegroundColor Cyan
    Mark-LedgerPerformance -Name 'SELECT requested account page'

    $postings = @{}
    foreach ($m in $pageMembers) { $postings[[int]$m.Code] = [System.Collections.Generic.List[object]]::new() }
    $fySql = Get-ReportSqlDateLiteral -Date $FyStart -IsSql $IsSql
    $endSql = Get-ReportSqlDateLiteral -Date $ToDate.AddDays(1) -IsSql $IsSql
    $recTypes = if ($IncludePdc) { '1,5' } else { '1' }
    $postingQueryCount = 0
    $postingBatchSize = if ($IsSql) { 2000 } else { 120 }
    for ($at = 0; $at -lt $pageMembers.Count; $at += $postingBatchSize) {
        $last = [math]::Min($at+$postingBatchSize,$pageMembers.Count)
        $batch = @($pageMembers[$at..($last-1)])
        $ids = (@($batch | ForEach-Object { [int]$_.Code }) -join ',')
        $sql = @"
SELECT P.MasterCode1 AS AccountCode,
       P.VchType, P.VchCode, P.SrNo, P.RecType,
       H.[Date] AS VoucherDate, H.VchNo AS VoucherNo,
       P.Value1 AS SignedPosting, P.ShortNar AS ShortNarration
FROM Tran2 AS P
INNER JOIN Tran1 AS H ON H.VchType=P.VchType AND H.VchCode=P.VchCode
WHERE P.MasterCode1 IN ($ids)
  AND P.RecType IN ($recTypes)
  AND P.Value1<>0
  AND (H.Cancelled=0 OR H.Cancelled IS NULL)
  AND (H.VchCancelled=0 OR H.VchCancelled IS NULL)
  AND H.[Date] >= $fySql AND H.[Date] < $endSql
ORDER BY P.MasterCode1, H.[Date], H.VchType, H.VchCode, P.SrNo, P.RecType, P.Value1
"@
        $postingQueryCount++
        foreach ($p in @(Get-AccountLedgerRows -Query $sql -IsSql $IsSql `
            -SqlConnection $SqlConnection -BusyConnection $BusyConnection -Timeout 180)) {
            if ($null -ne $p -and $postings.ContainsKey([int]$p.AccountCode)) { $postings[[int]$p.AccountCode].Add($p) }
        }
    }
    Mark-LedgerPerformance -Name 'SELECT page account postings'

    [decimal]$netOpening = 0
    [decimal]$netClosing = 0
    [decimal]$totalDebit = 0
    [decimal]$totalCredit = 0
    [decimal]$openingDr = 0
    [decimal]$openingCr = 0
    [decimal]$closingDr = 0
    [decimal]$closingCr = 0
    $totalPostingCount = 0
    $openingLookups = 0
    $openingCacheHits = 0
    $assumedZeroCount = 0
    $sections = [System.Collections.Generic.List[object]]::new()
    foreach ($m in $pageMembers) {
        $code = [int]$m.Code
        $name = ([string]$m.Name).Trim()
        $groupId = [int]$m.GroupCode
        $opening = Get-AccountLedgerCachedOpening -BusyConnection $BusyConnection `
            -AccountCode $code -AccountName $name -InstanceId $InstanceId `
            -CompanyCode $CompanyCode -FinancialYearStart $FyStart -Refresh $RefreshOpenings
        if ($opening.cacheHit) { $openingCacheHits++ } else { $openingLookups++ }
        if ($opening.assumedZero) { $assumedZeroCount++ }
        [decimal]$fyOpening = [decimal]$opening.signed
        [decimal]$running = $fyOpening
        [decimal]$periodOpening = $fyOpening
        [decimal]$debit = 0
        [decimal]$credit = 0
        $rows = [System.Collections.Generic.List[object]]::new()
        foreach ($p in $postings[$code]) {
            $postDate = ([datetime]$p.VoucherDate).Date
            [decimal]$change = -[decimal]$p.SignedPosting
            $running += $change
            if ($postDate -lt $FromDate) {
                $periodOpening += $change
                continue
            }
            [decimal]$dr = 0
            [decimal]$cr = 0
            if ($change -gt 0) { $dr = $change }
            if ($change -lt 0) { $cr = -$change }
            $debit += $dr
            $credit += $cr
            $balance = ConvertTo-AccountLedgerBalance -Signed $running
            $rows.Add([pscustomobject]@{
                voucherDate=$postDate.ToString('yyyy-MM-dd')
                voucherType=[int]$p.VchType
                voucherCode=[int]$p.VchCode
                voucherNo=([string]$p.VoucherNo).Trim()
                postingSrNo=[int]$p.SrNo
                postingRecType=[int]$p.RecType
                isPdc=([int]$p.RecType -eq 5)
                debit=$dr
                credit=$cr
                balance=$balance.amount
                balanceType=$balance.type
                signedBalance=$balance.signed
                shortNarration=[string]$p.ShortNarration
                counterpartAccount=''
                counterpartAccounts=@()
            })
        }
        $pageOpeningBalance = ConvertTo-AccountLedgerBalance -Signed $periodOpening
        $fyOpenBalance = ConvertTo-AccountLedgerBalance -Signed $fyOpening
        $closingBalance = ConvertTo-AccountLedgerBalance -Signed $running
        $netOpening += $periodOpening
        $netClosing += $running
        $totalDebit += $debit
        $totalCredit += $credit
        $totalPostingCount += $rows.Count
        if ($periodOpening -ge 0) { $openingDr += $periodOpening } else { $openingCr -= $periodOpening }
        if ($running -ge 0) { $closingDr += $running } else { $closingCr -= $running }
        $sectionGroupName = if ($groupsByCode.ContainsKey($groupId)) { [string]$groupsByCode[$groupId].name } else { '' }
        $sections.Add([pscustomobject]@{
            account=@{ code=$code; name=$name; group=$sectionGroupName; groupCode=$groupId }
            openingBalance=$pageOpeningBalance.amount
            openingBalanceType=$pageOpeningBalance.type
            financialYearOpeningBalance=$fyOpenBalance.amount
            financialYearOpeningBalanceType=$fyOpenBalance.type
            openingBalanceAssumedZero=[bool]$opening.assumedZero
            totalDebit=$debit
            totalCredit=$credit
            closingBalance=$closingBalance.amount
            closingBalanceType=$closingBalance.type
            signedClosingBalance=$closingBalance.signed
            postingCount=$rows.Count
            rows=$rows.ToArray()
        })
    }
    Mark-LedgerPerformance -Name 'COM openings + page balances'

    # Counterparty names are fetched ONLY for vouchers present on the page.
    $vchCodes = @{}
    $voucherKeys = @{}
    foreach ($sec in $sections) {
        foreach ($p in $sec.rows) {
            $voucherKeys["$($p.voucherType)|$($p.voucherCode)"] = $true
            $vchCodes[[int]$p.voucherCode] = $true
        }
    }
    if ($vchCodes.Count -gt 0) {
        $nameByKey = @{}
        $codes = @($vchCodes.Keys | Sort-Object)
        for ($at = 0; $at -lt $codes.Count; $at += 100) {
            $end = [math]::Min($at+100, $codes.Count)
            $ids = (@($codes[$at..($end-1)]) -join ',')
            $query = @"
SELECT O.VchType, O.VchCode, O.MasterCode1, O.SrNo, M.Name AS CounterName
FROM Tran2 AS O
INNER JOIN Master1 AS M ON M.Code=O.MasterCode1
WHERE M.MasterType=2 AND O.RecType IN ($recTypes)
  AND O.Value1<>0 AND O.VchCode IN ($ids)
ORDER BY O.VchType, O.VchCode, O.SrNo, O.MasterCode1
"@
            foreach ($other in @(Get-AccountLedgerRows -Query $query -IsSql $IsSql `
                -SqlConnection $SqlConnection -BusyConnection $BusyConnection)) {
                $key = "$($other.VchType)|$($other.VchCode)"
                if (-not $voucherKeys.ContainsKey($key)) { continue }
                if (-not $nameByKey.ContainsKey($key)) { $nameByKey[$key] = [System.Collections.Generic.List[object]]::new() }
                $nameByKey[$key].Add($other)
            }
        }
        foreach ($sec in $sections) {
            foreach ($p in $sec.rows) {
                $key = "$($p.voucherType)|$($p.voucherCode)"
                $names = [System.Collections.Generic.List[string]]::new()
                if ($nameByKey.ContainsKey($key)) {
                    foreach ($other in $nameByKey[$key]) {
                        if ([int]$other.MasterCode1 -eq [int]$sec.account.code) { continue }
                        $candidate = ([string]$other.CounterName).Trim()
                        if ($candidate -and -not $names.Contains($candidate)) { $names.Add($candidate) }
                    }
                }
                $p.counterpartAccounts = $names.ToArray()
                $p.counterpartAccount = ($names.ToArray() -join ', ')
            }
        }
    }
    Mark-LedgerPerformance -Name 'SELECT page counterpart names'
    $openingBalance = ConvertTo-AccountLedgerBalance -Signed $netOpening
    $closingBalance = ConvertTo-AccountLedgerBalance -Signed $netClosing
    $pages = [math]::Max(1,[math]::Ceiling($totalAccounts / [double]$PageSize))
    return @{
        success=$true
        data=@{
            report='account-ledger'
            mode=$(if ($AllAccounts) { 'all' } else { 'group' })
            group=$(if ($AllAccounts) { $null } else { @{ code=$groupRoot.code; name=$groupRoot.name; parentCode=$groupRoot.parentCode } })
            includeChildGroups=$IncludeChildGroups
            accountsToInclude='all'
            from=$FromDate.ToString('yyyy-MM-dd')
            to=$ToDate.ToString('yyyy-MM-dd')
            financialYearStart=$FyStart.ToString('yyyy-MM-dd')
            includePdc=$IncludePdc
            selectedGroupCount=$groupCodes.Count
            selectedAccountCount=$totalAccounts
            sourceAccountMasterCount=$sourceAccountCount
            excludedLinkedAccountMasterCount=$excludedCount
            accountCount=$totalAccounts
            totalPostingCount=$totalPostingCount
            openingBalance=$openingBalance.amount
            openingBalanceType=$openingBalance.type
            totalOpeningDebit=$openingDr
            totalOpeningCredit=$openingCr
            totalDebit=$totalDebit
            totalCredit=$totalCredit
            closingBalance=$closingBalance.amount
            closingBalanceType=$closingBalance.type
            signedClosingBalance=$closingBalance.signed
            totalClosingDebit=$closingDr
            totalClosingCredit=$closingCr
            accounts=$sections.ToArray()
            totalsScope='page'
            pagination=@{
                unit='accounts'
                page=$Page
                pageSize=$PageSize
                totalRows=$totalAccounts
                totalPages=$pages
                hasPreviousPage=($Page -gt 1)
                hasNextPage=($Page -lt $pages)
            }
            diagnostics=@{
                databaseMode=$(if ($IsSql) {'sql-server-direct'} else {'access-busy-com'})
                openingBalanceSource='BUSY.GetMasterXML(Account.OPBal)'
                openingMasterLookups=$openingLookups
                openingCacheHits=$openingCacheHits
                openingCacheTtlMinutes=30
                missingOpeningAccountCount=$assumedZeroCount
                postingQueryCount=$postingQueryCount
                databaseLevelPagination=$true
                accountMastersFetched=$pageMembers.Count
                accountMasterQueryShape=$(if ($IsSql) {'OFFSET FETCH'} else {'Jet TOP/NOT IN'})
                fullCompanyTotalsCalculated=$false
                sorting='account-name then date/type/voucher/srno; stable Code tiebreak'
            }
        }
    }
}

function Get-AccountLedgerGroupReport {
    param(
        [int]$GroupCode = 0,
        [string]$Group = '',
        [string]$From = '',
        [string]$To = '',
        [bool]$IncludeChildGroups = $true,
        [bool]$IncludePdc = $true,
        [string]$AccountsToInclude = 'moved-or-closing',
        [bool]$AllAccounts = $false,
        [bool]$RefreshOpenings = $false,
        [bool]$UseDatabasePaging = $false,
        [int]$Page = 1,
        [int]$PageSize = 20,
        [string]$InstanceId = '',
        [string]$CompanyCode = ''
    )

    Start-LedgerPerformance -Mode $(if ($AllAccounts) { 'all' } else { 'group' }) -InstanceId $InstanceId -CompanyCode $CompanyCode
    $ledgerPerfSuccess = $false
    $fi = $null
    $sqlConnection = $null
    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $today = (Get-Date).Date
        $fyYear = if ($today.Month -lt 4) { $today.Year - 1 } else { $today.Year }
        $fyStart = [datetime]::new($fyYear, 4, 1)
        $fyEndExclusive = $fyStart.AddYears(1)
        $fromDate = ConvertTo-AccountLedgerDate -Value $From -Default $fyStart
        $toDate = ConvertTo-AccountLedgerDate -Value $To -Default $today
        if ($fromDate -gt $toDate) { throw 'From date must not be later than To date.' }
        if ($fromDate -lt $fyStart -or $toDate -ge $fyEndExclusive) {
            throw ('This first version supports the active BUSY financial year only: ' +
                $fyStart.ToString('yyyy-MM-dd') + ' through ' +
                $fyEndExclusive.AddDays(-1).ToString('yyyy-MM-dd') + '.')
        }
        if (-not $AllAccounts -and $GroupCode -le 0 -and [string]::IsNullOrWhiteSpace($Group)) {
            throw 'groupCode or group (exact group name) is required.'
        }
        $AccountsToInclude = $AccountsToInclude.Trim().ToLowerInvariant()
        if ($AccountsToInclude -notin @('all', 'moved-or-closing')) {
            throw 'accountsToInclude must be all or moved-or-closing.'
        }
        if ($Page -lt 1) { $Page = 1 }
        if ($PageSize -lt 0) { $PageSize = 20 }
        if ($PageSize -gt 2000) { $PageSize = 2000 }

        Mark-LedgerPerformance -Name 'Validate parameters'
        $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
        if ($null -eq $fi) { throw 'Unable to connect to BUSY company.' }
        $dbMode = Get-ReportDatabaseMode -InstanceId $InstanceId
        $isSql = [bool]$dbMode.isSql
        if ($isSql) {
            $direct = Get-DirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
            if ($null -eq $direct -or $null -eq $direct.connection) {
                throw 'SQL Server fiscal database connection could not be created.'
            }
            $sqlConnection = $direct.connection
            if ($sqlConnection.State.ToString() -ne 'Open') { $sqlConnection.Open() }
        }

        Mark-LedgerPerformance -Name 'Connect BUSY + fiscal DB'
        if ($UseDatabasePaging -and $PageSize -gt 0) {
            # Strict master-level paging: no SELECT of all MasterType=2 rows.
            $result = Get-AccountLedgerTrueDbPage -BusyConnection $fi `
                -SqlConnection $sqlConnection -IsSql $isSql -AllAccounts $AllAccounts `
                -GroupCode $GroupCode -Group $Group -IncludeChildGroups $IncludeChildGroups `
                -IncludePdc $IncludePdc -AccountsToInclude $AccountsToInclude `
                -FyStart $fyStart -FromDate $fromDate -ToDate $toDate `
                -Page $Page -PageSize $PageSize -RefreshOpenings $RefreshOpenings `
                -InstanceId $InstanceId -CompanyCode $CompanyCode
            $ledgerPerfSuccess = [bool]$result.success
            return $result
        }
        # Master1 code is company-local. Read both types from the SELECTED fiscal database.
        # A stock item group uses MasterType=5 and MUST NOT be a ledger group.
        # I10=1 and CM5>0 identify linked accounts excluded from native GS All Ledger.
        $masterQuery = 'SELECT Code, Name, ParentGrp, MasterType, I10, CM5 FROM Master1 ' +
            'WHERE MasterType IN (1, 2) ORDER BY MasterType, Name, Code'
        $masterRows = @(Get-AccountLedgerRows -Query $masterQuery -IsSql $isSql `
            -SqlConnection $sqlConnection -BusyConnection $fi)
        $groupsByCode = @{}
        $groupsByName = @{}
        $accountMasters = [System.Collections.Generic.List[object]]::new()
        $sourceAccountMasterCount = 0
        $excludedLinkedAccountMasterCount = 0
        $childGroups = @{}
        foreach ($m in $masterRows) {
            $code = [int]$m.Code
            $masterType = [int]$m.MasterType
            if ($masterType -eq 1) {
                $g = [pscustomobject]@{
                    code = $code
                    name = ([string]$m.Name).Trim()
                    parentCode = [int]$m.ParentGrp
                }
                $groupsByCode[$code] = $g
                $nameKey = $g.name.ToUpperInvariant()
                if (-not $groupsByName.ContainsKey($nameKey)) {
                    $groupsByName[$nameKey] = [System.Collections.Generic.List[int]]::new()
                }
                $groupsByName[$nameKey].Add($code)
                if (-not $childGroups.ContainsKey($g.parentCode)) {
                    $childGroups[$g.parentCode] = [System.Collections.Generic.List[int]]::new()
                }
                $childGroups[$g.parentCode].Add($code)
            }
            elseif ($masterType -eq 2) {
                $sourceAccountMasterCount++
                # Native BUSY GS export excludes I10/CM5-linked account masters.
                # They may have a nonzero XML OPBal despite no independent ledger.
                # Never filter ordinary opening-only accounts: those remain visible.
                if ([int]$m.I10 -eq 1 -and [int]$m.CM5 -gt 0) {
                    $excludedLinkedAccountMasterCount++
                    continue
                }
                $accountMasters.Add([pscustomobject]@{
                    code = $code
                    name = ([string]$m.Name).Trim()
                    groupCode = [int]$m.ParentGrp
                })
            }
        }

        $selectedGroupCodes = @{}
        $rootGroup = $null
        if ($AllAccounts) {
            # All Accounts means every independent accounting master, including orphaned groups.
            # Do not select stock masters, users, or groups themselves.
            foreach ($code in $groupsByCode.Keys) { $selectedGroupCodes[[int]$code] = $true }
            $members = @($accountMasters | Sort-Object -Property name, code)
        }
        else {
            if ($GroupCode -le 0) {
                $nameKey = $Group.Trim().ToUpperInvariant()
                if (-not $groupsByName.ContainsKey($nameKey)) {
                    throw "Accounting group '$Group' was not found (MasterType=1)."
                }
                if ($groupsByName[$nameKey].Count -ne 1) {
                    throw "More than one accounting group named '$Group'; use groupCode."
                }
                $GroupCode = [int]$groupsByName[$nameKey][0]
            }
            if (-not $groupsByCode.ContainsKey($GroupCode)) {
                throw "Accounting group code $GroupCode was not found (MasterType=1)."
            }
            $rootGroup = $groupsByCode[$GroupCode]
            if (-not [string]::IsNullOrWhiteSpace($Group) -and
                -not [string]::Equals($Group.Trim(), $rootGroup.name, [System.StringComparison]::OrdinalIgnoreCase)) {
                throw 'groupCode and group name identify different groups.'
            }

            # Traverse once with a visited set; handle nested groups and invalid loops.
            $pending = [System.Collections.Generic.Queue[int]]::new()
            $pending.Enqueue($GroupCode)
            while ($pending.Count -gt 0) {
                $nextCode = $pending.Dequeue()
                if ($selectedGroupCodes.ContainsKey($nextCode)) { continue }
                $selectedGroupCodes[$nextCode] = $true
                if ($IncludeChildGroups -and $childGroups.ContainsKey($nextCode)) {
                    foreach ($childCode in $childGroups[$nextCode]) {
                        if (-not $selectedGroupCodes.ContainsKey($childCode)) {
                            $pending.Enqueue([int]$childCode)
                        }
                    }
                }
            }
            $members = @($accountMasters | Where-Object {
                $selectedGroupCodes.ContainsKey([int]$_.groupCode)
            } | Sort-Object -Property name, code)
        }
        $memberCount = $members.Count
        Mark-LedgerPerformance -Name 'Master1 + group selection'
        # Phase 7: Opt-in server-side pagination. Summary uses set-based aggregates
        # for ALL qualifying accounts, but detailed posting rows are fetched only
        # AFTER the account page is selected. No accounting snapshot is cached:
        # each request re-reads voucher sums so posted edits are visible.
        $databasePaging = [bool]($UseDatabasePaging -and $PageSize -gt 0 -and $members.Count -gt 0)
        $summaryByCode = @{}
        $postingLists = @{}
        $postingBatchSize = 120
        if ($isSql -and $members.Count -gt 0 -and $members.Count -le 4000) {
            $postingBatchSize = $members.Count
        }
        $postingQueryCount = 0
        $summaryQueryCount = 0
        $startSql = Get-ReportSqlDateLiteral -Date $fyStart -IsSql $isSql
        $fromSql = Get-ReportSqlDateLiteral -Date $fromDate -IsSql $isSql
        $endSql = Get-ReportSqlDateLiteral -Date $toDate.AddDays(1) -IsSql $isSql
        $recTypes = if ($IncludePdc) { '1,5' } else { '1' }

        if ($databasePaging) {
            # SQL Server aggregates decimals explicitly to avoid summing FLOAT
            # ledger amounts and drifting at scale. Access uses Jet IIf/Sum.
            # Both branches keep original Tran1 cancellation/date/RecType rules.
            if ($isSql) {
                $numericValue = 'CONVERT(decimal(28,10), P.Value1)'
                $beforeSignedExpr = "SUM(CASE WHEN H.[Date] < $fromSql THEN $numericValue ELSE CONVERT(decimal(28,10),0) END)"
                $debitExpr = "SUM(CASE WHEN H.[Date] >= $fromSql AND P.Value1 < 0 THEN -$numericValue ELSE CONVERT(decimal(28,10),0) END)"
                $creditExpr = "SUM(CASE WHEN H.[Date] >= $fromSql AND P.Value1 > 0 THEN $numericValue ELSE CONVERT(decimal(28,10),0) END)"
                $countExpr = "SUM(CASE WHEN H.[Date] >= $fromSql THEN 1 ELSE 0 END)"
            }
            else {
                $beforeSignedExpr = "SUM(IIF(H.[Date] < $fromSql, P.Value1, 0))"
                $debitExpr = "SUM(IIF(H.[Date] >= $fromSql AND P.Value1 < 0, -P.Value1, 0))"
                $creditExpr = "SUM(IIF(H.[Date] >= $fromSql AND P.Value1 > 0, P.Value1, 0))"
                $countExpr = "SUM(IIF(H.[Date] >= $fromSql, 1, 0))"
            }
            # Jet imposes smaller SQL statement/IN-list limits. SQL Server can
            # aggregate the selected company's accounts in one statement.
            for ($start = 0; $start -lt $members.Count; $start += $postingBatchSize) {
                $end = [math]::Min($start + $postingBatchSize, $members.Count)
                $batch = @($members[$start..($end - 1)])
                $accountCodes = (@($batch | ForEach-Object { [int]$_.code }) -join ',')
                $summaryQuery = @"
SELECT P.MasterCode1 AS AccountCode,
       $beforeSignedExpr AS PreSigned,
       $debitExpr AS PeriodDebit,
       $creditExpr AS PeriodCredit,
       $countExpr AS PeriodCount
FROM Tran2 AS P
INNER JOIN Tran1 AS H
    ON H.VchType = P.VchType AND H.VchCode = P.VchCode
WHERE P.MasterCode1 IN ($accountCodes)
  AND P.RecType IN ($recTypes)
  AND P.Value1 <> 0
  AND (H.Cancelled = 0 OR H.Cancelled IS NULL)
  AND (H.VchCancelled = 0 OR H.VchCancelled IS NULL)
  AND H.[Date] >= $startSql AND H.[Date] < $endSql
GROUP BY P.MasterCode1
"@
                $summaryQueryCount++
                foreach ($a in @(Get-AccountLedgerRows -Query $summaryQuery -IsSql $isSql `
                    -SqlConnection $sqlConnection -BusyConnection $fi -Timeout 180)) {
                    $summaryByCode[[int]$a.AccountCode] = $a
                }
            }
            Write-Host ('[LEDGER-PAGE] aggregated accounts={0} summaryQueries={1} requestedPage={2} size={3}' -f `
                $members.Count, $summaryQueryCount, $Page, $PageSize) -ForegroundColor DarkCyan
        }
        else {
            $postingLists = @{}
            foreach ($m in $members) {
                $postingLists[[int]$m.code] = [System.Collections.Generic.List[object]]::new()
            }

            # Read each posting exactly once, in bounded chunks, for SQL and Jet/BDS.
            # Never JOIN Tran3 or Tran5 here; either would multiply ledger amounts.
            $startSql = Get-ReportSqlDateLiteral -Date $fyStart -IsSql $isSql
            $endSql = Get-ReportSqlDateLiteral -Date $toDate.AddDays(1) -IsSql $isSql
            $recTypes = if ($IncludePdc) { '1,5' } else { '1' }
            # SQL Server can retrieve the same ordered postings for a moderate-sized
            # account selection in one set-based query. This avoids 20 repeated
            # Tran2/Tran1 query executions for GS (2,360 independent masters).
            # The SQL uses validated integer codes from Master1, never user input.
            # Keep the already-verified 120-account batching for Jet/Access and for
            # exceptionally large SQL company account selections.
            $postingBatchSize = 120
            if ($isSql -and $members.Count -gt 0 -and $members.Count -le 4000) {
                $postingBatchSize = $members.Count
            }
            $postingQueryCount = 0
            for ($start = 0; $start -lt $members.Count; $start += $postingBatchSize) {
                $end = [math]::Min($start + $postingBatchSize, $members.Count)
                $batch = @($members[$start..($end - 1)])
                $accountCodes = (@($batch | ForEach-Object { [int]$_.code }) -join ',')
                $postingQuery = @"
    SELECT P.MasterCode1 AS AccountCode,
           P.VchType, P.VchCode, P.SrNo, P.RecType,
           H.[Date] AS VoucherDate, H.VchNo AS VoucherNo,
           P.Value1 AS SignedPosting, P.ShortNar AS ShortNarration
    FROM Tran2 AS P
    INNER JOIN Tran1 AS H
        ON H.VchType = P.VchType AND H.VchCode = P.VchCode
    WHERE P.MasterCode1 IN ($accountCodes)
      AND P.RecType IN ($recTypes)
      AND P.Value1 <> 0
      AND (H.Cancelled = 0 OR H.Cancelled IS NULL)
      AND (H.VchCancelled = 0 OR H.VchCancelled IS NULL)
      AND H.[Date] >= $startSql AND H.[Date] < $endSql
    ORDER BY P.MasterCode1, H.[Date], H.VchType, H.VchCode, P.SrNo, P.RecType, P.Value1
"@
                $postingQueryCount++
                $batchRows = @(Get-AccountLedgerRows -Query $postingQuery -IsSql $isSql `
                    -SqlConnection $sqlConnection -BusyConnection $fi -Timeout 180)
                foreach ($p in $batchRows) {
                    $postingLists[[int]$p.AccountCode].Add($p)
                }
            }

        }
        Mark-LedgerPerformance -Name $(if ($databasePaging) { 'Voucher aggregate queries' } else { 'Voucher batch queries' })
        [decimal]$netOpening = 0
        [decimal]$netClosing = 0
        [decimal]$totalDebit = 0
        [decimal]$totalCredit = 0
        [decimal]$openingDr = 0
        [decimal]$openingCr = 0
        [decimal]$closingDr = 0
        [decimal]$closingCr = 0
        $totalPostings = 0
        $sections = [System.Collections.Generic.List[object]]::new()
        $fyOpeningByCode = @{}
        # pageSize=0 is full CSV export: retain the original one-pass detail
        # creation. Paginated reports compute all totals but defer posting-row
        # objects until the requested account sections have been selected.
        $fullDetailRequested = ($PageSize -eq 0)
        $openingLookupCount = 0
        $openingCacheHitCount = 0
        $missingOpeningAccountCodes = [System.Collections.Generic.List[int]]::new()

        foreach ($m in $members) {
            $code = [int]$m.code
            # Missing/blank opening is zero ONLY for validated MasterType=2 masters.
            # Cache keeps repeated pages and group/all requests from re-reading COM XML.
            $opening = Get-AccountLedgerCachedOpening -BusyConnection $fi `
                -AccountCode $code -AccountName $m.name -InstanceId $InstanceId `
                -CompanyCode $CompanyCode -FinancialYearStart $fyStart `
                -Refresh $RefreshOpenings
            if ($opening.cacheHit) { $openingCacheHitCount++ }
            else { $openingLookupCount++ }
            $openingBalanceAssumedZero = [bool]$opening.assumedZero
            [decimal]$fyOpening = [decimal]$opening.signed
            if ($openingBalanceAssumedZero) { $missingOpeningAccountCodes.Add($code) }
            [decimal]$periodOpening = $fyOpening
            [decimal]$running = $fyOpening
            [decimal]$debit = 0
            [decimal]$credit = 0
            $postingCount = 0
            $rows = $null
            if ($fullDetailRequested) { $rows = [System.Collections.Generic.List[object]]::new() }

            if ($databasePaging) {
                # No need to read all voucher rows to compute full-report totals.
                # PreSigned is SUM(P.Value1), whose sign is opposite the change in
                # BUSY's running debit-positive balance.
                [decimal]$preSigned = 0
                if ($summaryByCode.ContainsKey($code)) {
                    $a = $summaryByCode[$code]
                    $preSigned = [decimal]$a.PreSigned
                    $debit = [decimal]$a.PeriodDebit
                    $credit = [decimal]$a.PeriodCredit
                    $postingCount = [int]$a.PeriodCount
                }
                $periodOpening = $fyOpening - $preSigned
                $running = $periodOpening + $debit - $credit
            }
            else {
                foreach ($p in $postingLists[$code]) {
                    $postingDate = ([datetime]$p.VoucherDate).Date
                    $change = -[decimal]$p.SignedPosting
                    if ($postingDate -lt $fromDate) {
                        $periodOpening += $change
                        $running += $change
                        continue
                    }
                    $rowDr = [decimal]0
                    $rowCr = [decimal]0
                    if ($change -gt 0) { $rowDr = $change }
                    if ($change -lt 0) { $rowCr = -$change }
                    $debit += $rowDr
                    $credit += $rowCr
                    $running += $change
                    $postingCount++
                    if ($fullDetailRequested) {
                        $balance = ConvertTo-AccountLedgerBalance -Signed $running
                        $rows.Add([pscustomobject]@{
                            voucherDate = $postingDate.ToString('yyyy-MM-dd')
                            voucherType = [int]$p.VchType
                            voucherCode = [int]$p.VchCode
                            voucherNo = ([string]$p.VoucherNo).Trim()
                            postingSrNo = [int]$p.SrNo
                            postingRecType = [int]$p.RecType
                            isPdc = ([int]$p.RecType -eq 5)
                            debit = $rowDr
                            credit = $rowCr
                            balance = $balance.amount
                            balanceType = $balance.type
                            signedBalance = $balance.signed
                            shortNarration = [string]$p.ShortNarration
                            counterpartAccount = ''
                            counterpartAccounts = @()
                        })
                    }
                }
            }

            # BUSY report option: Moved/Closing Balance omits dormant zero-balance masters.
            if ($AccountsToInclude -eq 'moved-or-closing' -and
                $postingCount -eq 0 -and $running -eq [decimal]0) { continue }
            $fyOpeningByCode[$code] = $fyOpening

            $periodOpenBalance = ConvertTo-AccountLedgerBalance -Signed $periodOpening
            $fyOpenBalance = ConvertTo-AccountLedgerBalance -Signed $fyOpening
            $closingBalance = ConvertTo-AccountLedgerBalance -Signed $running
            $netOpening += $periodOpening
            $netClosing += $running
            $totalDebit += $debit
            $totalCredit += $credit
            $totalPostings += $postingCount
            if ($periodOpening -ge 0) { $openingDr += $periodOpening }
            else { $openingCr += -$periodOpening }
            if ($running -ge 0) { $closingDr += $running }
            else { $closingCr += -$running }

            $parentName = ''
            if ($groupsByCode.ContainsKey([int]$m.groupCode)) {
                $parentName = $groupsByCode[[int]$m.groupCode].name
            }
            $sectionRows = @()
            if ($fullDetailRequested) { $sectionRows = $rows.ToArray() }
            $sections.Add([pscustomobject]@{
                account = @{ code = $code; name = $m.name; group = $parentName; groupCode = [int]$m.groupCode }
                openingBalance = $periodOpenBalance.amount
                openingBalanceType = $periodOpenBalance.type
                financialYearOpeningBalance = $fyOpenBalance.amount
                financialYearOpeningBalanceType = $fyOpenBalance.type
                openingBalanceAssumedZero = $openingBalanceAssumedZero
                totalDebit = $debit
                totalCredit = $credit
                closingBalance = $closingBalance.amount
                closingBalanceType = $closingBalance.type
                signedClosingBalance = $closingBalance.signed
                postingCount = $postingCount
                rows = $sectionRows
            })
        }

        Mark-LedgerPerformance -Name 'XML cache + balance loops'
        $accountCount = $sections.Count
        if ($missingOpeningAccountCodes.Count -gt 0) {
            Write-Host "[ACCOUNT-LEDGER-MULTI] Treated undefined opening as zero for $($missingOpeningAccountCodes.Count) validated account master(s)." -ForegroundColor DarkYellow
        }
        $pageOffset = if ($PageSize -eq 0) { 0 } else { [long]($Page - 1) * $PageSize }
        $paged = [System.Collections.Generic.List[object]]::new()
        for ($idx = $pageOffset;
             $idx -lt $accountCount -and ($PageSize -eq 0 -or $idx -lt $pageOffset + $PageSize);
             $idx++) {
            $paged.Add($sections[[int]$idx])
        }

        if ($databasePaging -and $paged.Count -gt 0) {
            $pagedMembers = @($paged | ForEach-Object { $_.account })
            $pageBatchSize = 120
            if ($isSql -and $pagedMembers.Count -le 4000) { $pageBatchSize = $pagedMembers.Count }
            foreach ($m in $pagedMembers) {
                $postingLists[[int]$m.code] = [System.Collections.Generic.List[object]]::new()
            }
            for ($batchAt = 0; $batchAt -lt $pagedMembers.Count; $batchAt += $pageBatchSize) {
                $batchEnd = [math]::Min($batchAt + $pageBatchSize, $pagedMembers.Count)
                $pageBatch = @($pagedMembers[$batchAt..($batchEnd - 1)])
                $pageCodes = (@($pageBatch | ForEach-Object { [int]$_.code }) -join ',')
                $pageDetailQuery = @"
SELECT P.MasterCode1 AS AccountCode,
       P.VchType, P.VchCode, P.SrNo, P.RecType,
       H.[Date] AS VoucherDate, H.VchNo AS VoucherNo,
       P.Value1 AS SignedPosting, P.ShortNar AS ShortNarration
FROM Tran2 AS P
INNER JOIN Tran1 AS H
    ON H.VchType = P.VchType AND H.VchCode = P.VchCode
WHERE P.MasterCode1 IN ($pageCodes)
  AND P.RecType IN ($recTypes)
  AND P.Value1 <> 0
  AND (H.Cancelled = 0 OR H.Cancelled IS NULL)
  AND (H.VchCancelled = 0 OR H.VchCancelled IS NULL)
  AND H.[Date] >= $startSql AND H.[Date] < $endSql
ORDER BY P.MasterCode1, H.[Date], H.VchType, H.VchCode, P.SrNo, P.RecType, P.Value1
"@
                $postingQueryCount++
                foreach ($p in @(Get-AccountLedgerRows -Query $pageDetailQuery -IsSql $isSql `
                    -SqlConnection $sqlConnection -BusyConnection $fi -Timeout 180)) {
                    $postingLists[[int]$p.AccountCode].Add($p)
                }
            }
            Write-Host ('[LEDGER-PAGE] detailAccounts={0} detailPostingQueries={1}' -f `
                $pagedMembers.Count, $postingQueryCount) -ForegroundColor DarkCyan
        }
        Mark-LedgerPerformance -Name 'Page voucher queries'

        # Materialize voucher rows ONLY for the requested account page.
        # Use exactly the original posting order, opening and running-balance
        # calculations. A pageSize=0 full export was materialized above.
        if (-not $fullDetailRequested) {
            foreach ($section in $paged) {
                $code = [int]$section.account.code
                [decimal]$running = [decimal]$fyOpeningByCode[$code]
                $detailRows = [System.Collections.Generic.List[object]]::new()
                foreach ($p in $postingLists[$code]) {
                    $postingDate = ([datetime]$p.VoucherDate).Date
                    $change = -[decimal]$p.SignedPosting
                    $running += $change
                    if ($postingDate -lt $fromDate) { continue }
                    $rowDr = [decimal]0
                    $rowCr = [decimal]0
                    if ($change -gt 0) { $rowDr = $change }
                    if ($change -lt 0) { $rowCr = -$change }
                    $balance = ConvertTo-AccountLedgerBalance -Signed $running
                    $detailRows.Add([pscustomobject]@{
                        voucherDate = $postingDate.ToString('yyyy-MM-dd')
                        voucherType = [int]$p.VchType
                        voucherCode = [int]$p.VchCode
                        voucherNo = ([string]$p.VoucherNo).Trim()
                        postingSrNo = [int]$p.SrNo
                        postingRecType = [int]$p.RecType
                        isPdc = ([int]$p.RecType -eq 5)
                        debit = $rowDr
                        credit = $rowCr
                        balance = $balance.amount
                        balanceType = $balance.type
                        signedBalance = $balance.signed
                        shortNarration = [string]$p.ShortNarration
                        counterpartAccount = ''
                        counterpartAccounts = @()
                    })
                }
                if ($databasePaging) {
                    [decimal]$detailDebit = 0
                    [decimal]$detailCredit = 0
                    foreach ($item in $detailRows) {
                        $detailDebit += [decimal]$item.debit
                        $detailCredit += [decimal]$item.credit
                    }
                    if ($detailRows.Count -ne [int]$section.postingCount -or
                        [math]::Abs($detailDebit - [decimal]$section.totalDebit) -gt [decimal]0.005 -or
                        [math]::Abs($detailCredit - [decimal]$section.totalCredit) -gt [decimal]0.005) {
                        throw "Account ${code}: aggregate totals do not match detail postings. Server pagination aborted."
                    }
                }
                $section.rows = $detailRows.ToArray()
            }
        }
        Mark-LedgerPerformance -Name 'Detail rows for account page'

        # Only fetch counterparty labels for the sections on this page.
        $voucherKeys = @{}
        $voucherCodes = @{}
        foreach ($section in $paged) {
            foreach ($posting in $section.rows) {
                $k = "$($posting.voucherType)|$($posting.voucherCode)"
                $voucherKeys[$k] = $true
                $voucherCodes[[int]$posting.voucherCode] = $true
            }
        }
        if ($voucherCodes.Count -gt 0) {
            $uniqueCodes = @($voucherCodes.Keys | Sort-Object)
            $nameByKey = @{}
            for ($batchAt = 0; $batchAt -lt $uniqueCodes.Count; $batchAt += 100) {
                $batchEnd = [math]::Min($batchAt + 100, $uniqueCodes.Count)
                $ids = (@($uniqueCodes[$batchAt..($batchEnd - 1)]) -join ',')
                $counterQuery = @"
SELECT O.VchType, O.VchCode, O.MasterCode1, O.SrNo,
       M.Name AS CounterName
FROM Tran2 AS O
INNER JOIN Master1 AS M ON M.Code = O.MasterCode1
WHERE M.MasterType = 2 AND O.RecType IN ($recTypes)
  AND O.Value1 <> 0 AND O.VchCode IN ($ids)
ORDER BY O.VchType, O.VchCode, O.SrNo, O.MasterCode1
"@
                foreach ($other in @(Get-AccountLedgerRows -Query $counterQuery -IsSql $isSql `
                    -SqlConnection $sqlConnection -BusyConnection $fi)) {
                    $k = "$($other.VchType)|$($other.VchCode)"
                    if (-not $voucherKeys.ContainsKey($k)) { continue }
                    if (-not $nameByKey.ContainsKey($k)) {
                        $nameByKey[$k] = [System.Collections.Generic.List[object]]::new()
                    }
                    $nameByKey[$k].Add($other)
                }
            }
            foreach ($section in $paged) {
                $sectionCode = [int]$section.account.code
                foreach ($posting in $section.rows) {
                    $k = "$($posting.voucherType)|$($posting.voucherCode)"
                    $names = [System.Collections.Generic.List[string]]::new()
                    if ($nameByKey.ContainsKey($k)) {
                        foreach ($counter in $nameByKey[$k]) {
                            if ([int]$counter.MasterCode1 -eq $sectionCode) { continue }
                            $counterName = ([string]$counter.CounterName).Trim()
                            if ($counterName -and -not $names.Contains($counterName)) {
                                $names.Add($counterName)
                            }
                        }
                    }
                    $posting.counterpartAccounts = $names.ToArray()
                    $posting.counterpartAccount = ($names.ToArray() -join ', ')
                }
            }
        }

        Mark-LedgerPerformance -Name 'Page + counterpart lookups'
        $netOpeningBalance = ConvertTo-AccountLedgerBalance -Signed $netOpening
        $netClosingBalance = ConvertTo-AccountLedgerBalance -Signed $netClosing
        $totalPages = if ($PageSize -eq 0) { 1 } else {
            [math]::Max(1, [math]::Ceiling($accountCount / $PageSize))
        }
        Mark-LedgerPerformance -Name 'Final report totals'
        $clock.Stop()
        $ledgerPerfSuccess = $true
        return @{
            success = $true
            data = @{
                report = 'account-ledger'
                mode = $(if ($AllAccounts) { 'all' } else { 'group' })
                group = $(if ($AllAccounts) { $null } else { @{ code = $rootGroup.code; name = $rootGroup.name; parentCode = $rootGroup.parentCode } })
                includeChildGroups = $IncludeChildGroups
                accountsToInclude = $AccountsToInclude
                from = $fromDate.ToString('yyyy-MM-dd')
                to = $toDate.ToString('yyyy-MM-dd')
                financialYearStart = $fyStart.ToString('yyyy-MM-dd')
                includePdc = $IncludePdc
                selectedGroupCount = $selectedGroupCodes.Count
                selectedAccountCount = $memberCount
                sourceAccountMasterCount = $sourceAccountMasterCount
                excludedLinkedAccountMasterCount = $excludedLinkedAccountMasterCount
                accountCount = $accountCount
                totalPostingCount = $totalPostings
                openingBalance = $netOpeningBalance.amount
                openingBalanceType = $netOpeningBalance.type
                totalOpeningDebit = $openingDr
                totalOpeningCredit = $openingCr
                totalDebit = $totalDebit
                totalCredit = $totalCredit
                closingBalance = $netClosingBalance.amount
                closingBalanceType = $netClosingBalance.type
                signedClosingBalance = $netClosingBalance.signed
                totalClosingDebit = $closingDr
                totalClosingCredit = $closingCr
                accounts = $paged.ToArray()
                pagination = @{
                    unit = 'accounts'
                    page = $Page
                    pageSize = $PageSize
                    totalRows = $accountCount
                    totalPages = $totalPages
                    hasPreviousPage = ($Page -gt 1)
                    hasNextPage = ($PageSize -ne 0 -and $Page -lt $totalPages)
                }
                diagnostics = @{
                    postingRecTypes = $(if ($IncludePdc) { @(1,5) } else { @(1) })
                    databaseMode = $(if ($isSql) { 'sql-server-direct' } else { 'access-busy-com' })
                    openingBalanceSource = 'BUSY.GetMasterXML(Account.OPBal)'
                    linkedAccountExclusion = 'Master1.I10=1 AND Master1.CM5>0'
                    openingMasterLookups = $openingLookupCount
                    openingCacheHits = $openingCacheHitCount
                    openingCacheTtlMinutes = 30
                    openingCacheRefreshRequested = $RefreshOpenings
                    postingQueryCount = $postingQueryCount
                    postingSummaryQueryCount = $summaryQueryCount
                    databaseLevelPagination = $databasePaging
                    postingQueryBatchSize = $postingBatchSize
                    detailRowsMaterializedOnlyForPage = (-not $fullDetailRequested)
                    missingOpeningAccountCount = $missingOpeningAccountCodes.Count
                    missingOpeningAccountCodesSample = @($missingOpeningAccountCodes | Select-Object -First 20)
                    elapsedMilliseconds = $clock.ElapsedMilliseconds
                    sorting = 'account-name then date/type/voucher/srno; native within-day ordering pending validation'
                }
            }
        }
    }
    catch {
        $clock.Stop()
        Write-Host "[ACCOUNT-LEDGER-GROUP] $($_.Exception.Message)" -ForegroundColor Red
        return @{ success = $false; error = $_.Exception.Message }
    }
    finally {
        if ($null -ne $sqlConnection) {
            try { $sqlConnection.Close() } catch {}
            try { $sqlConnection.Dispose() } catch {}
        }
        if ($null -ne $fi) { Disconnect-BUSY $fi }
        Stop-LedgerPerformance -Succeeded $ledgerPerfSuccess
    }
}



# Pre-warm one bounded batch of account openings, avoiding a multi-minute first
# report request on large companies. Repeated calls continue where the cache left off.
# Cache entries expire in 30 minutes; no balance is inferred from Master1.D1.
function Get-AccountLedgerWarmOpenings {
    param(
        [int]$BatchSize = 60,
        [string]$InstanceId = '',
        [string]$CompanyCode = ''
    )
    Start-LedgerPerformance -Mode 'warm-up' -InstanceId $InstanceId -CompanyCode $CompanyCode
    $ledgerPerfSuccess = $false
    if ($BatchSize -lt 1) { $BatchSize = 1 }
    if ($BatchSize -gt 100) { $BatchSize = 100 }
    $fi = $null
    $sqlConnection = $null
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $today = (Get-Date).Date
        $fyYear = if ($today.Month -lt 4) { $today.Year - 1 } else { $today.Year }
        $fyStart = [datetime]::new($fyYear, 4, 1)
        Mark-LedgerPerformance -Name 'Validate parameters'
        $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
        if ($null -eq $fi) { throw 'Unable to connect to BUSY company.' }
        $mode = Get-ReportDatabaseMode -InstanceId $InstanceId
        $isSql = [bool]$mode.isSql
        if ($isSql) {
            $direct = Get-DirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
            if ($null -eq $direct -or $null -eq $direct.connection) {
                throw 'SQL Server fiscal database connection could not be created.'
            }
            $sqlConnection = $direct.connection
            if ($sqlConnection.State.ToString() -ne 'Open') { $sqlConnection.Open() }
        }
        Mark-LedgerPerformance -Name 'Connect BUSY + fiscal DB'
        $masters = @(Get-AccountLedgerRows -Query 'SELECT Code, Name, I10, CM5 FROM Master1 WHERE MasterType = 2 ORDER BY Code' `
            -IsSql $isSql -SqlConnection $sqlConnection -BusyConnection $fi)
        # Warm only eligible independent account masters. Preserve the original
        # raw totalAccountMasters for diagnostic/backward compatibility.
        $eligibleMasters = @($masters | Where-Object {
            -not ([int]$_.I10 -eq 1 -and [int]$_.CM5 -gt 0)
        })
        $linkedExcluded = $masters.Count - $eligibleMasters.Count
        Mark-LedgerPerformance -Name 'Read + filter Master1'
        $prefix = '{0}|{1}|{2}|' -f $InstanceId, $CompanyCode, $fyStart.ToString('yyyy-MM-dd')
        $now = [datetime]::UtcNow
        $lookups = 0
        $hits = 0
        foreach ($m in $eligibleMasters) {
            $code = [int]$m.Code
            $key = $prefix + [string]$code
            if ($script:AccountLedgerOpeningCache.ContainsKey($key) -and
                $script:AccountLedgerOpeningCache[$key].expiresAt -gt $now) {
                $hits++
                continue
            }
            if ($lookups -ge $BatchSize) { continue }
            $null = Get-AccountLedgerCachedOpening -BusyConnection $fi `
                -AccountCode $code -AccountName ([string]$m.Name) `
                -InstanceId $InstanceId -CompanyCode $CompanyCode `
                -FinancialYearStart $fyStart
            $lookups++
        }
        Mark-LedgerPerformance -Name 'COM opening XML batch'
        # Do not claim a value has been cached without checking the TTL.
        $ready = 0
        $now = [datetime]::UtcNow
        foreach ($m in $eligibleMasters) {
            $key = $prefix + [string]([int]$m.Code)
            if ($script:AccountLedgerOpeningCache.ContainsKey($key) -and
                $script:AccountLedgerOpeningCache[$key].expiresAt -gt $now) { $ready++ }
        }
        Mark-LedgerPerformance -Name 'Count cached openings'
        $watch.Stop()
        $ledgerPerfSuccess = $true
        return @{
            success = $true
            data = @{
                report = 'account-ledger'
                mode = 'opening-warmup'
                totalAccountMasters = $masters.Count
                eligibleAccountMasters = $eligibleMasters.Count
                excludedLinkedAccountMasters = $linkedExcluded
                cachedAccountMasters = $ready
                remainingAccountMasters = ($eligibleMasters.Count - $ready)
                lookupsThisRequest = $lookups
                cacheHitsAtStart = $hits
                complete = ($ready -eq $eligibleMasters.Count)
                cacheTtlMinutes = 30
                elapsedMilliseconds = $watch.ElapsedMilliseconds
                databaseMode = $(if ($isSql) { 'sql-server-direct' } else { 'access-busy-com' })
            }
        }
    }
    catch {
        Write-Host "[ACCOUNT-LEDGER-WARM] $($_.Exception.Message)" -ForegroundColor Red
        return @{ success = $false; error = $_.Exception.Message }
    }
    finally {
        if ($null -ne $sqlConnection) {
            try { $sqlConnection.Close() } catch {}
            try { $sqlConnection.Dispose() } catch {}
        }
        if ($null -ne $fi) { Disconnect-BUSY $fi }
        Stop-LedgerPerformance -Succeeded $ledgerPerfSuccess
    }
}

# All-accounts wrapper reuses the Group Ledger's proven financial calculations.
# pageSize=0 requests complete account sections for export (may be slow/large).
function Get-AccountLedgerAllReport {
    param(
        [string]$From = '',
        [string]$To = '',
        [bool]$IncludePdc = $true,
        [string]$AccountsToInclude = 'moved-or-closing',
        [bool]$RefreshOpenings = $false,
        [bool]$UseDatabasePaging = $false,
        [int]$Page = 1,
        [int]$PageSize = 20,
        [string]$InstanceId = '',
        [string]$CompanyCode = ''
    )
    return Get-AccountLedgerGroupReport -AllAccounts $true `
        -From $From -To $To -IncludePdc $IncludePdc `
        -AccountsToInclude $AccountsToInclude `
        -RefreshOpenings $RefreshOpenings -UseDatabasePaging $UseDatabasePaging -Page $Page -PageSize $PageSize `
        -InstanceId $InstanceId -CompanyCode $CompanyCode
}
