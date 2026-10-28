# dashboard.ps1 — isolated, read-only executive reporting, PowerShell 5.1.
# Load AFTER reports.ps1 and account_ledger.ps1, BEFORE routes.ps1.
# All amounts are LKR. Voucher figures are gross posted voucher amounts, NOT profit or bank balances.
# Never use numeric zeros for unavailable data; null metrics remain unavailable in the UI.
$script:ExecutiveDashboardCache = @{}
$script:ExecutiveDashboardTtlSeconds = 180

function Get-ExecutiveNumber {
    param($Value)
    if ($null -eq $Value -or $Value -eq [DBNull]::Value) { return 0.0 }
    return [double]$Value
}

function Get-ExecutiveRows {
    param([string]$Query, [bool]$IsSql, $SqlConnection, $BusyConnection)
    $rows = [System.Collections.Generic.List[object]]::new()
    if ($IsSql) {
        $cmd = $null
        $reader = $null
        try {
            $cmd = $SqlConnection.CreateCommand()
            $cmd.CommandText = $Query
            $cmd.CommandTimeout = 120
            $reader = $cmd.ExecuteReader()
            while ($reader.Read()) {
                $fields = [ordered]@{}
                for ($i=0; $i -lt $reader.FieldCount; $i++) {
                    $v = $reader.GetValue($i)
                    if ($v -is [DBNull]) { $v = $null }
                    $fields[[string]$reader.GetName($i)] = $v
                }
                $rows.Add([pscustomobject]$fields)
            }
        } finally {
            if ($null -ne $reader) { $reader.Dispose() }
            if ($null -ne $cmd) { $cmd.Dispose() }
        }
    } else {
        $rst = $null
        try {
            $rst = $BusyConnection.GetRecordset($Query)
            if ($null -ne $rst -and -not $rst.EOF) {
                $rst.MoveFirst()
                while (-not $rst.EOF) {
                    $fields = [ordered]@{}
                    for ($i=0; $i -lt $rst.Fields.Count; $i++) {
                        $f = $rst.Fields.Item($i)
                        $v = $f.Value
                        if ($v -is [DBNull]) { $v = $null }
                        $fields[[string]$f.Name] = $v
                    }
                    $rows.Add([pscustomobject]$fields)
                    $rst.MoveNext()
                }
            }
        } finally {
            if ($null -ne $rst) { try { $rst.Close() } catch {} }
        }
    }
    # Enumerate row objects on the PowerShell pipeline; callers collect them with @(...).
    return $rows.ToArray()
}

function Get-ExecutiveDashboard {
    param(
        [string]$InstanceId = '',
        [string]$CompanyCode = '',
        [string]$AsOf = '',
        [ValidateSet('core','exposure')][string]$Section = 'core',
        [bool]$ForceRefresh = $false
    )
    try {
        if ([string]::IsNullOrWhiteSpace($InstanceId) -or [string]::IsNullOrWhiteSpace($CompanyCode)) {
            throw 'Select a BUSY instance and company.'
        }
        if (-not [string]::IsNullOrWhiteSpace($AsOf) -and $AsOf -notmatch '^\d{4}-\d{2}-\d{2}$') {
            throw 'asOf must use yyyy-MM-dd.'
        }
        $today = (Get-Date).Date
        $asOfDate = ConvertTo-ReportDate -Value $AsOf -Fallback $today
        $fyYear = if ($asOfDate.Month -lt 4) { $asOfDate.Year - 1 } else { $asOfDate.Year }
        $fyStart = [datetime]::new($fyYear,4,1)
        $fyEnd = $fyStart.AddYears(1)
        if ($asOfDate -lt $fyStart -or $asOfDate -ge $fyEnd -or $asOfDate -gt $today) {
            throw 'Dashboard date must not be in the future.'
        }
        # Exposed report valuation/openings are currently valid for the active financial year.
        $activeFy = if ($today.Month -lt 4) { $today.Year - 1 } else { $today.Year }
        if ($fyYear -ne $activeFy) { throw 'The dashboard currently supports only the active financial year.' }

        $key = ($InstanceId+'|'+$CompanyCode+'|'+$Section+'|'+$asOfDate.ToString('yyyy-MM-dd')).ToLowerInvariant()
        if (-not $ForceRefresh -and $script:ExecutiveDashboardCache.ContainsKey($key)) {
            $hit = $script:ExecutiveDashboardCache[$key]
            if ([datetime]::UtcNow -lt $hit.expiresAt) { return $hit.response }
        }
        $dateText = $asOfDate.ToString('yyyy-MM-dd')
        $base = @{
            success = $true
            section = $Section
            asOf = $dateText
            financialYearStart = $fyStart.ToString('yyyy-MM-dd')
            generatedAt = [datetime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
            companyCode = $CompanyCode
            instanceId = $InstanceId
            currency = 'LKR'
            warnings = @()
        }

        if ($Section -eq 'core') {
            $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
            if ($null -eq $fi) { throw 'BUSY company connection failed.' }
            $mode = Get-ReportDatabaseMode -InstanceId $InstanceId
            $isSql = [bool]$mode.isSql
            $directConn = $null
            try {
                if ($isSql) {
                    $direct = Get-DirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
                    if ($null -eq $direct -or $null -eq $direct.connection) { throw 'SQL fiscal connection unavailable.' }
                    $directConn = $direct.connection
                    if ($directConn.State.ToString() -ne 'Open') { $directConn.Open() }
                }
                $first = Get-ReportSqlDateLiteral -Date $fyStart -IsSql $isSql
                $last = Get-ReportSqlDateLiteral -Date $asOfDate.AddDays(1) -IsSql $isSql
                # Year/Month/Day/Abs are available in T-SQL and Access SQL.
                # Transaction type comes only from Tran1; excludes canceled documents.
                $dailySql = @"
SELECT Year([Date]) AS DocYear, Month([Date]) AS DocMonth,
       Day([Date]) AS DocDay, VchType,
       COUNT(*) AS VoucherCount,
       SUM(Abs(VchAmtBaseCur)) AS GrossAmount
FROM Tran1
WHERE [Date] >= $first AND [Date] < $last
  AND VchType IN (2,3,9,10,14,19)
  AND (Cancelled = 0 OR Cancelled IS NULL)
  AND (VchCancelled = 0 OR VchCancelled IS NULL)
GROUP BY Year([Date]), Month([Date]), Day([Date]), VchType
ORDER BY Year([Date]), Month([Date]), Day([Date]), VchType
"@
                $rawDaily = @(Get-ExecutiveRows -Query $dailySql -IsSql $isSql -SqlConnection $directConn -BusyConnection $fi)
                $monthlyStart = [datetime]::new($asOfDate.Year,$asOfDate.Month,1)
                $trendStart = $asOfDate.AddDays(-29)
                if ($trendStart -lt $fyStart) { $trendStart = $fyStart }
                $totals = @{}
                foreach ($scope in @('today','mtd','fytd')) {
                    foreach ($vch in @(2,3,9,10,14,19)) { $totals["$scope/$vch"] = @{ value = 0.0; count = 0 } }
                }
                $trend = @{}
                for ($d=$trendStart; $d -le $asOfDate; $d=$d.AddDays(1)) {
                    $ds=$d.ToString('yyyy-MM-dd')
                    $trend[$ds]=@{ date=$ds; sales=0.0; purchases=0.0; receipts=0.0; payments=0.0; salesCount=0 }
                }
                foreach ($r in $rawDaily) {
                    $vch=[int]$r.VchType
                    $d=[datetime]::new([int]$r.DocYear,[int]$r.DocMonth,[int]$r.DocDay)
                    $amount=Get-ExecutiveNumber $r.GrossAmount
                    $cnt=[int]$r.VoucherCount
                    $periods=@('fytd')
                    if ($d -ge $monthlyStart) { $periods+= 'mtd' }
                    if ($d -eq $asOfDate) { $periods+= 'today' }
                    foreach ($period in $periods) {
                        $entry=$totals["$period/$vch"]
                        $entry.value += $amount
                        $entry.count += $cnt
                    }
                    $ds=$d.ToString('yyyy-MM-dd')
                    if ($trend.ContainsKey($ds)) {
                        $target=$trend[$ds]
                        switch ($vch) {
                            9 { $target.sales += $amount; $target.salesCount += $cnt }
                            2 { $target.purchases += $amount }
                            14 { $target.receipts += $amount }
                            19 { $target.payments += $amount }
                        }
                    }
                }
                $periodsResult=@{}
                foreach ($scope in @('today','mtd','fytd')) {
                    $sale=$totals["$scope/9"]
                    $saleReturn=$totals["$scope/3"]
                    $purchase=$totals["$scope/2"]
                    $purchaseReturn=$totals["$scope/10"]
                    $receipt=$totals["$scope/14"]
                    $payment=$totals["$scope/19"]
                    $periodsResult[$scope]=@{
                        salesGross=[math]::Round($sale.value,2); salesReturns=[math]::Round($saleReturn.value,2)
                        netSales=[math]::Round(($sale.value-$saleReturn.value),2)
                        purchasesGross=[math]::Round($purchase.value,2); purchaseReturns=[math]::Round($purchaseReturn.value,2)
                        netPurchases=[math]::Round(($purchase.value-$purchaseReturn.value),2)
                        receiptVouchers=[math]::Round($receipt.value,2); paymentVouchers=[math]::Round($payment.value,2)
                        saleVouchers=[int]$sale.count; purchaseVouchers=[int]$purchase.count
                        receiptCount=[int]$receipt.count; paymentCount=[int]$payment.count
                    }
                }

                $recentFrom=Get-ReportSqlDateLiteral -Date $fyStart -IsSql $isSql
                $recentSql=@"
SELECT TOP 12 H.VchType, H.VchCode, H.VchNo, H.[Date] AS VoucherDate,
       H.VchAmtBaseCur AS Amount, M.Name AS PartyName
FROM Tran1 AS H
LEFT JOIN Master1 AS M ON M.Code = H.MasterCode1 AND M.MasterType = 2
WHERE H.[Date] >= $recentFrom AND H.[Date] < $last
  AND H.VchType IN (2,3,9,10,14,19)
  AND (H.Cancelled = 0 OR H.Cancelled IS NULL)
  AND (H.VchCancelled = 0 OR H.VchCancelled IS NULL)
ORDER BY H.[Date] DESC, H.VchCode DESC
"@
                $rawRecent=@(Get-ExecutiveRows -Query $recentSql -IsSql $isSql -SqlConnection $directConn -BusyConnection $fi)
                $recentList=[System.Collections.Generic.List[object]]::new()
                foreach($r in $rawRecent) {
                    $dateValue=if ($null -ne $r.VoucherDate) { ([datetime]$r.VoucherDate).ToString('yyyy-MM-dd') } else { $dateText }
                    $recentList.Add(@{
                        voucherType=[int]$r.VchType; voucherCode=[int]$r.VchCode
                        voucherNo=[string]$r.VchNo; date=$dateValue; party=[string]$r.PartyName
                        amount=[math]::Round([math]::Abs((Get-ExecutiveNumber $r.Amount)),2)
                    })
                }
                $base.periods=$periodsResult
                $base.trend=@($trend.Keys | Sort-Object | ForEach-Object { $trend[$_] })
                $base.recent=@($recentList.ToArray())
                $base.source='Tran1 fiscal posted voucher headers'
                $base.databaseMode=if ($isSql) { 'sql' } else { 'access' }
                $base.definitions=@{
                    netSales='Sales voucher gross amounts less Sale Return voucher gross amounts. Tax-inclusive voucher totals may not equal recognized accounting revenue.'
                    netPurchases='Purchase voucher gross amounts less Purchase Return voucher gross amounts; not an expense or COGS calculation.'
                    receiptVouchers='Receipt voucher face amounts; not verified cash inflow or available bank balance.'
                    paymentVouchers='Payment voucher face amounts; not verified cash outflow or available bank balance.'
                }
            } finally {
                if ($null -ne $directConn) { $directConn.Dispose() }
            }
        }
        else {
            # Each existing report remains the single source of truth. Failures are
            # isolated: never convert failed financial sections into false zeros.
            $base.stock=$null
            $base.outstanding=$null
            $warnings=[System.Collections.Generic.List[string]]::new()
            try {
                $stock=Get-StockStatusReport -From ($fyStart.ToString('yyyy-MM-dd')) `
                    -To $dateText -AsOf $dateText -View 'balances' -ValueBy 'busy' `
                    -MastersMode 'moved-closing' -Page 1 -PageSize 0 `
                    -IncludeZero $true -FastSnapshot $false `
                    -InstanceId $InstanceId -CompanyCode $CompanyCode
                if (-not $stock.success) { throw ('Stock Status: '+[string]$stock.error) }
                if ([string]$stock.summaryScope -ne 'all') { throw 'Stock Status returned only page totals.' }
                $s=$stock.summary
                $stockItems=@($stock.data)
                $topValued = @($stockItems | Where-Object { [double]$_.stockValue -gt 0 } |
                    Sort-Object -Property stockValue -Descending | Select-Object -First 5 |
                    ForEach-Object { @{ code=[int]$_.itemCode; name=[string]$_.itemName; quantity=[double]$_.closingQuantity; stockValue=[double]$_.stockValue } })
                $negativeStock = @($stockItems | Where-Object { [double]$_.closingQuantity -lt 0 } |
                    Sort-Object -Property stockValue | Select-Object -First 5 |
                    ForEach-Object { @{ code=[int]$_.itemCode; name=[string]$_.itemName; quantity=[double]$_.closingQuantity; stockValue=[double]$_.stockValue } })
                $base.stock=@{
                    topValuedItems=$topValued
                    negativeItemExamples=$negativeStock
                    stockValue=[double]$s.stockValue
                    quantity=[double]$s.totalQuantity
                    totalItems=[int]$s.totalItems
                    negativeItems=[int]$s.negativeStockItems
                    lowStockItems=[int]$s.lowStockItems
                    outOfStockItems=[int]$s.outOfStockItems
                    valuationSource=[string]$stock.valuationSource
                    valuationMethod=[string]$stock.stockValuationMode
                    asOf=$dateText
                }
            } catch {
                $warnings.Add('Stock valuation unavailable: '+$_.Exception.Message)
            }
            try {
                $outstanding=Get-OutstandingReport -From ($fyStart.ToString('yyyy-MM-dd')) `
                    -To $dateText -AsOf $dateText -Page 1 -PageSize 1 `
                    -InstanceId $InstanceId -CompanyCode $CompanyCode
                if (-not $outstanding.success) { throw ('Outstanding: '+[string]$outstanding.error) }
                if ([string]$outstanding.summaryScope -ne 'all') { throw 'Outstanding report returned partial totals.' }
                $s=$outstanding.summary
                $base.outstanding=@{
                    receivable=[double]$s.receivablePendingTotal
                    payable=[double]$s.payablePendingTotal
                    overdueReceivable=[double]$s.overdueReceivable
                    overduePayable=[double]$s.overduePayable
                    netPosition=[double]$s.netPosition
                    receivableAccounts=[int]$s.receivableAccounts
                    payableAccounts=[int]$s.payableAccounts
                    totalBills=[int]$s.totalBills
                    asOf=$dateText
                }
            } catch {
                $warnings.Add('Outstanding balances unavailable: '+$_.Exception.Message)
            }
            $base.warnings=@($warnings.ToArray())
        }
        $script:ExecutiveDashboardCache[$key]=@{
            expiresAt=[datetime]::UtcNow.AddSeconds($script:ExecutiveDashboardTtlSeconds)
            response=$base
        }
        # Bound cache memory across many company/date combinations.
        if ($script:ExecutiveDashboardCache.Count -gt 48) {
            foreach($oldKey in @($script:ExecutiveDashboardCache.Keys)) {
                if ([datetime]::UtcNow -gt $script:ExecutiveDashboardCache[$oldKey].expiresAt) {
                    $script:ExecutiveDashboardCache.Remove($oldKey)
                }
            }
        }
        return $base
    } catch {
        Write-Host ('[EXEC-DASHBOARD] '+$_.Exception.Message) -ForegroundColor DarkYellow
        return @{ success=$false; error=$_.Exception.Message; section=$Section }
    }
}
