# Executive sales rankings — read-only, isolated worker computation.
# All amounts are BUSY voucher face values, not recognized accounting revenue.
# Sales: VchType 9, Sale Return: VchType 3. Inventory: Tran2 RecType=2.
function Get-ExecutiveSalesInsights {
    param(
        [Parameter(Mandatory=$true)][string]$InstanceId,
        [Parameter(Mandatory=$true)][string]$CompanyCode,
        [Parameter(Mandatory=$true)][string]$AsOf,
        [ValidateSet('today','mtd')][string]$Period='mtd'
    )
    $conn=$null
    $fi=$null
    try {
        $asOfDate=[datetime]::ParseExact($AsOf,'yyyy-MM-dd',[System.Globalization.CultureInfo]::InvariantCulture)
        $now=(Get-Date).Date
        if($asOfDate -gt $now) { throw 'Future report date not allowed.' }
        $thisFy=if($now.Month -lt 4){$now.Year-1}else{$now.Year}
        $asOfFy=if($asOfDate.Month -lt 4){$asOfDate.Year-1}else{$asOfDate.Year}
        if($asOfFy -ne $thisFy){throw 'Only the current active fiscal year is supported.'}
        $fromDate=if($Period -eq 'today'){$asOfDate}else{[datetime]::new($asOfDate.Year,$asOfDate.Month,1)}
        $toExclusive=$asOfDate.AddDays(1)
        $fi=Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
        if($null -eq $fi){throw 'BUSY connection unavailable.'}
        $mode=Get-ReportDatabaseMode -InstanceId $InstanceId
        $isSql=[bool]$mode.isSql
        if($isSql){
            $direct=Get-DirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
            if($null -eq $direct -or $null -eq $direct.connection){throw 'Fiscal SQL connection unavailable.'}
            $conn=$direct.connection
            if($conn.State.ToString() -ne 'Open'){$conn.Open()}
        }
        $fromLit=Get-ReportSqlDateLiteral -Date $fromDate -IsSql $isSql
        $toLit=Get-ReportSqlDateLiteral -Date $toExclusive -IsSql $isSql
        # Avoid a detail-level join for customers: count each posted voucher once.
        $customerSql=@"
SELECT TOP 5 H.MasterCode1 AS CustomerCode, M.Name AS CustomerName,
 SUM(CASE WHEN H.VchType = 9 THEN ABS(H.VchAmtBaseCur)
          WHEN H.VchType = 3 THEN -ABS(H.VchAmtBaseCur) ELSE 0 END) AS NetVoucherValue,
 SUM(CASE WHEN H.VchType = 9 THEN 1 ELSE 0 END) AS SaleVouchers,
 SUM(CASE WHEN H.VchType = 3 THEN 1 ELSE 0 END) AS ReturnVouchers
FROM Tran1 AS H
INNER JOIN Master1 AS M ON M.Code = H.MasterCode1 AND M.MasterType = 2
WHERE H.[Date] >= $fromLit AND H.[Date] < $toLit
 AND H.VchType IN (9,3)
 AND (H.Cancelled = 0 OR H.Cancelled IS NULL)
 AND (H.VchCancelled = 0 OR H.VchCancelled IS NULL)
GROUP BY H.MasterCode1, M.Name
HAVING SUM(CASE WHEN H.VchType = 9 THEN ABS(H.VchAmtBaseCur)
                WHEN H.VchType = 3 THEN -ABS(H.VchAmtBaseCur) ELSE 0 END) > 0
ORDER BY NetVoucherValue DESC, M.Name
"@
        # Net base-item quantity: sale outward less sale-return inward.
        # NOT closing stock quantity. Avoid mixing unlike units into a combined total.
        $itemSql=@"
SELECT TOP 5 T.MasterCode1 AS ItemCode, M.Name AS ItemName,
 SUM(CASE WHEN H.VchType = 9 AND T.Value1 < 0 THEN -T.Value1
          WHEN H.VchType = 3 AND T.Value1 > 0 THEN -T.Value1 ELSE 0 END) AS NetSoldQuantity,
 SUM(CASE WHEN H.VchType = 9 AND T.Value1 < 0 THEN -T.Value1 ELSE 0 END) AS GrossSoldQuantity,
 SUM(CASE WHEN H.VchType = 3 AND T.Value1 > 0 THEN T.Value1 ELSE 0 END) AS ReturnedQuantity
FROM Tran2 AS T
INNER JOIN Tran1 AS H ON H.VchCode = T.VchCode
INNER JOIN Master1 AS M ON M.Code = T.MasterCode1 AND M.MasterType = 6
WHERE T.RecType = 2 AND H.[Date] >= $fromLit AND H.[Date] < $toLit
 AND H.VchType IN (9,3)
 AND (H.Cancelled = 0 OR H.Cancelled IS NULL)
 AND (H.VchCancelled = 0 OR H.VchCancelled IS NULL)
GROUP BY T.MasterCode1, M.Name
HAVING SUM(CASE WHEN H.VchType = 9 AND T.Value1 < 0 THEN -T.Value1
                WHEN H.VchType = 3 AND T.Value1 > 0 THEN -T.Value1 ELSE 0 END) > 0
ORDER BY NetSoldQuantity DESC, M.Name
"@
        # BUSYWEB_INSIGHTS_ACCESS_ROW_AGG_V1
        # Access/BDS: use supported raw joins and worker-side ranking.
        # SQL Server: preserve the existing query implementation.
        if (-not $isSql) {
            $accessCustomerSql = @"
SELECT TOP 100001
 H.MasterCode1 AS CustomerCode,
 M.Name AS CustomerName,
 H.VchType AS VoucherType,
 H.VchAmtBaseCur AS Amount
FROM Tran1 AS H
INNER JOIN Master1 AS M ON H.MasterCode1 = M.Code
WHERE H.[Date] >= $fromLit
 AND H.[Date] < $toLit
 AND H.VchType IN (9,3)
 AND M.MasterType = 2
 AND (H.Cancelled = 0 OR H.Cancelled IS NULL)
 AND (H.VchCancelled = 0 OR H.VchCancelled IS NULL)
"@

            $accessItemSql = @"
SELECT TOP 100001
 T.MasterCode1 AS ItemCode,
 M.Name AS ItemName,
 H.VchType AS VoucherType,
 T.Value1 AS Quantity
FROM (Tran1 AS H
 INNER JOIN Tran2 AS T
 ON H.VchCode = T.VchCode
 AND H.VchType = T.VchType)
INNER JOIN Master1 AS M ON T.MasterCode1 = M.Code
WHERE H.[Date] >= $fromLit
 AND H.[Date] < $toLit
 AND H.VchType IN (9,3)
 AND T.RecType = 2
 AND M.MasterType = 6
 AND (H.Cancelled = 0 OR H.Cancelled IS NULL)
 AND (H.VchCancelled = 0 OR H.VchCancelled IS NULL)
"@

            $customerMap = @{}
            $itemMap = @{}

            foreach ($kind in @('customer','item')) {
                $query = if ($kind -eq 'customer') {
                    $accessCustomerSql
                } else {
                    $accessItemSql
                }

                $rs = $null
                $readCount = 0

                try {
                    $rs = $fi.GetRecordset([string]$query)

                    if ($null -eq $rs) {
                        throw "BUSY Access $kind ranking query returned no recordset."
                    }

                    while (-not $rs.EOF) {
                        $readCount++

                        # Never publish incomplete rankings from a truncated source.
                        if ($readCount -gt 100000) {
                            throw "Access $kind ranking exceeded 100000 rows. Optimized aggregation is required."
                        }

                        if ($kind -eq 'customer') {
                            $code = [int]$rs.Fields.Item('CustomerCode').Value
                            $key = [string]$code
                            $name = [string]$rs.Fields.Item('CustomerName').Value
                            $type = [int]$rs.Fields.Item('VoucherType').Value
                            $amountValue = $rs.Fields.Item('Amount').Value

                            if ($null -eq $amountValue -or
                                $amountValue -is [System.DBNull]) {
                                throw 'Customer ranking encountered a missing voucher amount.'
                            }

                            $amount = [decimal]$amountValue

                            if (-not $customerMap.ContainsKey($key)) {
                                $customerMap[$key] = [pscustomobject]@{
                                    CustomerCode = $code
                                    CustomerName = $name
                                    NetVoucherValue = [decimal]0
                                    SaleVouchers = 0
                                    ReturnVouchers = 0
                                }
                            }

                            $c = $customerMap[$key]

                            if ($type -eq 9) {
                                $c.NetVoucherValue += [math]::Abs($amount)
                                $c.SaleVouchers++
                            }
                            elseif ($type -eq 3) {
                                $c.NetVoucherValue -= [math]::Abs($amount)
                                $c.ReturnVouchers++
                            }
                        }
                        else {
                            $code = [int]$rs.Fields.Item('ItemCode').Value
                            $key = [string]$code
                            $name = [string]$rs.Fields.Item('ItemName').Value
                            $type = [int]$rs.Fields.Item('VoucherType').Value
                            $quantityValue = $rs.Fields.Item('Quantity').Value

                            if ($null -eq $quantityValue -or
                                $quantityValue -is [System.DBNull]) {
                                throw 'Item ranking encountered a missing quantity.'
                            }

                            $qty = [decimal]$quantityValue

                            if (-not $itemMap.ContainsKey($key)) {
                                $itemMap[$key] = [pscustomobject]@{
                                    ItemCode = $code
                                    ItemName = $name
                                    NetSoldQuantity = [decimal]0
                                    GrossSoldQuantity = [decimal]0
                                    ReturnedQuantity = [decimal]0
                                }
                            }

                            $item = $itemMap[$key]

                            if ($type -eq 9 -and $qty -lt 0) {
                                $sold = -$qty
                                $item.GrossSoldQuantity += $sold
                                $item.NetSoldQuantity += $sold
                            }
                            elseif ($type -eq 3 -and $qty -gt 0) {
                                $item.ReturnedQuantity += $qty
                                $item.NetSoldQuantity -= $qty
                            }
                        }

                        $rs.MoveNext()
                    }
                }
                finally {
                    if ($null -ne $rs) {
                        try { $rs.Close() } catch {}
                    }
                }

                Write-Host (
                    '[EXEC-INSIGHTS-ACCESS] ' + $kind +
                    ' source rows=' + $readCount
                )
            }

            $rawCustomers = @(
                $customerMap.Values |
                    Where-Object { $_.NetVoucherValue -gt 0 } |
                    Sort-Object -Property @(
                        @{Expression='NetVoucherValue';Descending=$true},
                        @{Expression='CustomerName';Descending=$false}
                    ) |
                    Select-Object -First 5
            )

            $rawItems = @(
                $itemMap.Values |
                    Where-Object { $_.NetSoldQuantity -gt 0 } |
                    Sort-Object -Property @(
                        @{Expression='NetSoldQuantity';Descending=$true},
                        @{Expression='ItemName';Descending=$false}
                    ) |
                    Select-Object -First 5
            )

            Write-Host (
                '[EXEC-INSIGHTS-ACCESS] Ranked customers=' +
                $rawCustomers.Count + ', items=' + $rawItems.Count
            )
        }
        else {
            $rawCustomers=@(Get-ExecutiveRows -Query $customerSql -IsSql $isSql -SqlConnection $conn -BusyConnection $fi)
            $rawItems=@(Get-ExecutiveRows -Query $itemSql -IsSql $isSql -SqlConnection $conn -BusyConnection $fi)
        }

        $customers=@(foreach($r in $rawCustomers){
            @{ code=[int]$r.CustomerCode; name=[string]$r.CustomerName;
               amount=[math]::Round([double]$r.NetVoucherValue,2);
               voucherCount=[int]$r.SaleVouchers; returnCount=[int]$r.ReturnVouchers }
        })
        $items=@(foreach($r in $rawItems){
            @{ code=[int]$r.ItemCode; name=[string]$r.ItemName;
               quantity=[math]::Round([double]$r.NetSoldQuantity,3);
               grossSoldQuantity=[math]::Round([double]$r.GrossSoldQuantity,3);
               returnedQuantity=[math]::Round([double]$r.ReturnedQuantity,3);
               returnRate=if([double]$r.GrossSoldQuantity -gt 0){[math]::Round((100*[double]$r.ReturnedQuantity/[double]$r.GrossSoldQuantity),2)}else{$null} }
        })
        # BUSYWEB_SCOPED_CUSTOMER_BALANCES_V5
        # Calculate BUSY reference metrics ONCE for up to five ranked customers.
        # Uses the native Outstanding engine's Method 1/2/3 aggregation, status,
        # signed pending, cheque-return, due date and closed-reference rules.
        # Never run a full-company report from this rankings calculation.
        $outstandingReady=$false
        $outstandingWarning=$null
        $asOfOutstanding=$null
        $topAccountCodes=@(
            $customers | ForEach-Object { [int]$_.code } |
                Where-Object { $_ -gt 0 } | Sort-Object -Unique
        )
        if ($topAccountCodes.Count -eq 0) {
            $outstandingReady=$true
        } else {
            $scopeTimer=[System.Diagnostics.Stopwatch]::StartNew()
            try {
                $asOfOutstanding=Get-OutstandingReport `
                    -InstanceId $InstanceId `
                    -CompanyCode $CompanyCode `
                    -AsOf $AsOf `
                    -Type 'receivable' `
                    -AccountCodes $topAccountCodes `
                    -Page 1 -PageSize 1
                if ($null -ne $asOfOutstanding -and
                    $asOfOutstanding.success -eq $true -and
                    $null -ne $asOfOutstanding.accountTotals) {
                    $outstandingReady=$true
                } else {
                    $outstandingWarning='Scoped BUSY reference calculation failed; customer balances unavailable.'
                    Write-Warning '[EXEC-360] Scoped Outstanding did not return valid account totals.'
                }
            } catch {
                # Customer/item rankings remain valid if optional enrichment fails.
                $outstandingWarning='Scoped BUSY reference calculation failed; customer balances unavailable.'
                Write-Warning ('[EXEC-360] Scoped Outstanding error: '+$_.Exception.Message)
            } finally {
                $scopeTimer.Stop()
                Write-Host ('[EXEC-360] Account-scoped Outstanding for ' +
                    $topAccountCodes.Count + ' customer(s) completed in ' +
                    [Math]::Round($scopeTimer.Elapsed.TotalMilliseconds) +
                    ' ms; balances available=' + $outstandingReady)
            }
        }
        foreach($c in $customers){
            $c.pendingAmount=$null; $c.adjustedAmount=$null; $c.overdueAmount=$null
            $c.openReferenceCount=$null; $c.referenceScope='unavailable'
            if($outstandingReady){
                $key=[string]$c.code
                $a=$null
                if($asOfOutstanding.accountTotals -is [System.Collections.IDictionary]){
                    $a=$asOfOutstanding.accountTotals[$key]
                }
                # An account absent from the outstanding snapshot may have no open
                # references, but don't infer whether it has historic settlements.
                if($null -ne $a -and [string]$a.nature -eq 'receivable'){
                    $c.pendingAmount=[math]::Round([double]$a.pendingTotal,2)
                    $c.adjustedAmount=[math]::Round([double]$a.adjustedTotal,2)
                    $c.overdueAmount=[math]::Round([double]$a.overdueTotal,2)
                    $c.openReferenceCount=[int]$a.billCount
                    $c.referenceScope='open-references-as-of'
                } else {
                    $c.pendingAmount=0.0
                    $c.adjustedAmount=0.0
                    $c.overdueAmount=0.0
                    $c.openReferenceCount=0
                    $c.referenceScope='no-open-references-as-of'
                }
            }
        }
        return @{
            success=$true;section=('insights_'+$Period);asOf=$asOfDate.ToString('yyyy-MM-dd');period=$Period
            instanceId=$InstanceId;companyCode=$CompanyCode
            generatedAt=[datetime]::UtcNow.ToString('o')
            topCustomers=$customers;bestSellingItems=$items;outstandingAvailable=$outstandingReady;outstandingWarning=$outstandingWarning
            source='Posted Tran1 voucher headers; Tran2 RecType=2 item movements, current fiscal database'
            definitions=@{
                customer='Net sale voucher face value (sales less sale returns), includes taxes if present in voucher totals; NOT accounting revenue.'
                item='Net sold units = gross sale outward units minus sale return inward units. Return rate = returned / gross sold; verify return movement signs against native BUSY.'
                outstanding='Pending, adjusted and overdue are as-of account totals of currently open receivable references from native Outstanding. Adjusted is NOT cash collected, and does NOT necessarily include closed references. It is NOT the difference between period sales and receipts.'
            }
        }
    } catch {
        Write-Warning ('[EXEC-INSIGHTS] '+$_.Exception.Message)
        return @{success=$false;error='Sales insights calculation failed. Check the worker log.'}
    } finally {
        if($null -ne $conn){try{$conn.Dispose()}catch{}}
    }
}
