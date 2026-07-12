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
        [string]$AsOf          = "",
        [string]$Type          = "all",
        [string]$Search        = "",
        [string]$Group         = "",
        [string]$Status        = "all",
        [double]$MinAmount     = 0,
        [bool]$IncludeZero     = $false,
        [int]$Page             = 1,
        [int]$PageSize         = 50,
        [string]$InstanceId    = "",
        [string]$CompanyCode   = ""
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
        $masterPage = Get-ReportPagedMasterCodes `
            -Fi $fi `
            -MasterType 2 `
            -Search $Search `
            -Group $Group `
            -Page $Page `
            -PageSize $PageSize

        $accounts = @($masterPage.rows)

        if ($accounts.Count -eq 0) {
            return @{
                success = $true
                asOfDate = $asOfDate.ToString("yyyy-MM-dd")
                generatedAt = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
                page = $Page
                pageSize = $PageSize
                total = $masterPage.total
                totalPages = [Math]::Max(1, [Math]::Ceiling($masterPage.total / $PageSize))
                summaryScope = "page"
                summary = @{
                    totalReceivable = 0
                    totalPayable = 0
                    netPosition = 0
                    receivableParties = 0
                    payableParties = 0
                    overdueReceivable = 0
                    overduePayable = 0
                    totalOutstanding = 0
                    totalBills = 0
                    totalOnAccountSigned = 0
                    totalOnAccountAmount = 0
                    aging = @{
                        notDue = 0
                        days0To30 = 0
                        days31To60 = 0
                        days61To90 = 0
                        days91To180 = 0
                        above180 = 0
                    }
                }
                data = @()
            }
        }

        $masterMap = @{}
        foreach ($account in $accounts) {
            $masterMap[[int]$account.code] = @{
                code = [int]$account.code
                name = [string]$account.name
                groupName = [string]$account.groupName
            }
        }

        function Resolve-AccountType {
            param([string]$GroupName)
            $name = ([string]$GroupName).Trim().ToLowerInvariant()

            foreach ($token in @(
                "customer", "sundry debtor", "debtor",
                "trade debtor", "receivable"
            )) {
                if ($name.Contains($token)) { return "receivable" }
            }

            foreach ($token in @(
                "supplier", "sundry creditor", "creditor",
                "trade creditor", "payable"
            )) {
                if ($name.Contains($token)) { return "payable" }
            }

            return "unknown"
        }

        # Apply type filter before loading references.
        $allowedCodes = @()
        foreach ($account in $accounts) {
            $resolvedType = Resolve-AccountType $account.groupName
            if ($Type -eq "all" -or $Type -eq $resolvedType -or $resolvedType -eq "unknown") {
                $allowedCodes += [int]$account.code
            }
        }

        if ($allowedCodes.Count -eq 0) {
            return @{
                success = $true
                asOfDate = $asOfDate.ToString("yyyy-MM-dd")
                generatedAt = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
                page = $Page
                pageSize = $PageSize
                total = $masterPage.total
                totalPages = [Math]::Max(1, [Math]::Ceiling($masterPage.total / $PageSize))
                summaryScope = "page"
                summary = @{
                    totalReceivable = 0; totalPayable = 0; netPosition = 0
                    receivableParties = 0; payableParties = 0
                    overdueReceivable = 0; overduePayable = 0
                    totalOutstanding = 0; totalBills = 0
                    totalOnAccountSigned = 0; totalOnAccountAmount = 0
                    aging = @{
                        notDue = 0; days0To30 = 0; days31To60 = 0
                        days61To90 = 0; days91To180 = 0; above180 = 0
                    }
                }
                data = @()
            }
        }

        $codeList = ($allowedCodes | Sort-Object -Unique) -join ","

        $referenceQry = @"
SELECT
    T.RefCode,
    T.Method,
    T.MasterCode1,
    T.MasterCode2,
    T.[No] AS RefNo,
    T.[Date] AS RefDate,
    T.DueDate,
    T.Value1,
    T.NewRefAmount,
    T.VchCode,
    T.VchType,
    V.VchNo,
    V.[Date] AS VoucherDate,
    V.MasterCode1 AS VoucherPartyCode
FROM Tran3 T
LEFT JOIN Tran1 V ON V.VchCode = T.VchCode
WHERE
    T.RecType = 1
    AND T.[Date] <= $dateLiteral
    AND (
        T.MasterCode1 IN ($codeList)
        OR V.MasterCode1 IN ($codeList)
    )
ORDER BY T.MasterCode1, T.RefCode, T.Method, T.[Date], T.VchCode
"@

        $rst = $fi.GetRecordset($referenceQry)
        $references = @{}
        $onAccountByParty = @{}

        if ($rst -and -not $rst.EOF) {
            $rst.MoveFirst()

            while (-not $rst.EOF) {
                $accountCode = ConvertTo-ReportInt (
                    Get-ReportFieldValue $rst "MasterCode1" 0
                )

                if (-not $masterMap.ContainsKey($accountCode)) {
                    $candidate = ConvertTo-ReportInt (
                        Get-ReportFieldValue $rst "VoucherPartyCode" 0
                    )
                    if ($masterMap.ContainsKey($candidate)) {
                        $accountCode = $candidate
                    }
                }

                if ($masterMap.ContainsKey($accountCode)) {
                    $method = ConvertTo-ReportInt (
                        Get-ReportFieldValue $rst "Method" 0
                    )
                    $refCode = ConvertTo-ReportInt (
                        Get-ReportFieldValue $rst "RefCode" 0
                    )
                    $refNo = [string](
                        Get-ReportFieldValue $rst "RefNo" ""
                    )
                    $refNo = $refNo.Trim()
                    $value = ConvertTo-ReportDouble (
                        Get-ReportFieldValue $rst "Value1" 0
                    )

                    $isOnAccount =
                        ($refCode -eq 0) -or
                        ([string]::IsNullOrWhiteSpace($refNo)) -or
                        ($method -notin @(1, 2))

                    if ($isOnAccount) {
                        if (-not $onAccountByParty.ContainsKey($accountCode)) {
                            $onAccountByParty[$accountCode] = 0.0
                        }
                        $onAccountByParty[$accountCode] += $value
                    }
                    else {
                        $key = "$accountCode|$refCode"

                        if (-not $references.ContainsKey($key)) {
                            $references[$key] = @{
                                accountCode = $accountCode
                                refCode = $refCode
                                refNo = $refNo
                                refDate = $asOfDate
                                dueDate = $asOfDate
                                voucherNo = ""
                                voucherType = 0
                                originalSigned = 0.0
                                adjustmentSigned = 0.0
                                pendingSigned = 0.0
                            }
                        }

                        $ref = $references[$key]

                        if ($method -eq 1) {
                            $ref.originalSigned += $value
                            $ref.refNo = $refNo
                            $ref.voucherNo = [string](
                                Get-ReportFieldValue $rst "VchNo" ""
                            )
                            $ref.voucherType = ConvertTo-ReportInt (
                                Get-ReportFieldValue $rst "VchType" 0
                            )

                            try {
                                $ref.refDate = [datetime](
                                    Get-ReportFieldValue $rst "RefDate" $asOfDate
                                )
                            } catch {}

                            try {
                                $ref.dueDate = [datetime](
                                    Get-ReportFieldValue $rst "DueDate" $ref.refDate
                                )
                            } catch {
                                $ref.dueDate = $ref.refDate
                            }
                        }
                        elseif ($method -eq 2) {
                            $ref.adjustmentSigned += $value
                        }

                        $ref.pendingSigned =
                            $ref.originalSigned +
                            $ref.adjustmentSigned
                    }
                }

                $rst.MoveNext()
            }
        }

        Close-ReportRecordset $rst

        $accountRows = @{}
        $detailsByAccount = @{}

        foreach ($ref in $references.Values) {
            $account = $masterMap[[int]$ref.accountCode]
            if ($null -eq $account) { continue }

            $accountType = Resolve-AccountType $account.groupName
            if ($accountType -eq "unknown") {
                $accountType =
                    if ([double]$ref.originalSigned -lt 0) {
                        "receivable"
                    } else {
                        "payable"
                    }
            }

            if ($Type -ne "all" -and $Type -ne $accountType) { continue }

            $pendingSigned = [Math]::Round([double]$ref.pendingSigned, 2)
            $pendingAbs = [Math]::Abs($pendingSigned)

            if (-not $IncludeZero -and $pendingAbs -lt 0.005) { continue }
            if ($pendingAbs -lt $MinAmount) { continue }

            $dueDate = [datetime]$ref.dueDate
            $isNotDue = $dueDate.Date -gt $asOfDate.Date
            $daysOverdue =
                if ($isNotDue) { 0 }
                else { [Math]::Max(0, ($asOfDate.Date - $dueDate.Date).Days) }

            $entryStatus =
                if ($isNotDue) { "not-due" }
                elseif ($daysOverdue -gt 0) { "overdue" }
                else { "due" }

            if ($Status -ne "all" -and $Status -ne $entryStatus) { continue }

            $bucket = Get-OutstandingAgeBucket `
                -DaysOverdue $daysOverdue `
                -IsNotDue $isNotDue

            if (-not $accountRows.ContainsKey($account.code)) {
                $accountRows[$account.code] = @{
                    accountCode = $account.code
                    accountName = $account.name
                    accountGroup = $account.groupName
                    type = $accountType
                    billCount = 0
                    billOriginalTotal = 0.0
                    billPendingSigned = 0.0
                    billPendingAbs = 0.0
                    receivable = 0.0
                    payable = 0.0
                    notDueAmount = 0.0
                    overdueAmount = 0.0
                    onAccountSigned =
                        if ($onAccountByParty.ContainsKey($account.code)) {
                            [double]$onAccountByParty[$account.code]
                        } else { 0.0 }
                    oldestDueDate = $null
                    maxDaysOverdue = 0
                    aging = @{
                        notDue = 0.0
                        days0To30 = 0.0
                        days31To60 = 0.0
                        days61To90 = 0.0
                        days91To180 = 0.0
                        above180 = 0.0
                    }
                }
                $detailsByAccount[$account.code] =
                    [System.Collections.Generic.List[object]]::new()
            }

            $row = $accountRows[$account.code]
            $row.billCount++
            $row.billOriginalTotal += [Math]::Abs([double]$ref.originalSigned)
            $row.billPendingSigned += $pendingSigned
            $row.billPendingAbs += $pendingAbs

            if ($accountType -eq "receivable") { $row.receivable += $pendingAbs }
            else { $row.payable += $pendingAbs }

            if ($isNotDue) { $row.notDueAmount += $pendingAbs }
            else {
                $row.overdueAmount += $pendingAbs
                if ($daysOverdue -gt $row.maxDaysOverdue) {
                    $row.maxDaysOverdue = $daysOverdue
                }
            }

            $row.aging[$bucket] += $pendingAbs

            if ($null -eq $row.oldestDueDate -or $dueDate -lt $row.oldestDueDate) {
                $row.oldestDueDate = $dueDate
            }

            $detailsByAccount[$account.code].Add([pscustomobject]@{
                accountCode = $account.code
                refCode = $ref.refCode
                refNo = if ($ref.refNo) { $ref.refNo } else { $ref.voucherNo }
                voucherNo = $ref.voucherNo
                voucherType = $ref.voucherType
                refDate = ([datetime]$ref.refDate).ToString("yyyy-MM-dd")
                dueDate = $dueDate.ToString("yyyy-MM-dd")
                originalAmount = [Math]::Round(
                    [Math]::Abs([double]$ref.originalSigned), 2
                )
                adjustedAmount = [Math]::Round(
                    [Math]::Abs([double]$ref.adjustmentSigned), 2
                )
                outstandingAmount = $pendingAbs
                pendingSigned = $pendingSigned
                status = $entryStatus
                daysOverdue = $daysOverdue
                agingBucket = $bucket
            })
        }

        $rows = [System.Collections.Generic.List[object]]::new()

        foreach ($account in $accounts) {
            if (-not $accountRows.ContainsKey([int]$account.code)) { continue }

            $row = $accountRows[[int]$account.code]
            $onAccountSigned = [Math]::Round([double]$row.onAccountSigned, 2)
            $netSigned = [Math]::Round(
                [double]$row.billPendingSigned + $onAccountSigned, 2
            )

            $rows.Add([pscustomobject]@{
                accountCode = $row.accountCode
                accountName = $row.accountName
                accountGroup = $row.accountGroup
                type = $row.type
                billCount = $row.billCount
                billOriginalTotal = [Math]::Round($row.billOriginalTotal, 2)
                billPendingSigned = [Math]::Round($row.billPendingSigned, 2)
                billPendingAmount = [Math]::Abs(
                    [Math]::Round($row.billPendingSigned, 2)
                )
                onAccountSigned = $onAccountSigned
                onAccountAmount = [Math]::Abs($onAccountSigned)
                netSigned = $netSigned
                netOutstanding = [Math]::Abs($netSigned)
                receivable = [Math]::Round($row.receivable, 2)
                payable = [Math]::Round($row.payable, 2)
                totalOutstanding = [Math]::Round($row.billPendingAbs, 2)
                notDueAmount = [Math]::Round($row.notDueAmount, 2)
                overdueAmount = [Math]::Round($row.overdueAmount, 2)
                oldestDueDate =
                    if ($row.oldestDueDate) {
                        ([datetime]$row.oldestDueDate).ToString("yyyy-MM-dd")
                    } else { "" }
                daysOverdue = $row.maxDaysOverdue
                aging = @{
                    notDue = [Math]::Round($row.aging.notDue, 2)
                    days0To30 = [Math]::Round($row.aging.days0To30, 2)
                    days31To60 = [Math]::Round($row.aging.days31To60, 2)
                    days61To90 = [Math]::Round($row.aging.days61To90, 2)
                    days91To180 = [Math]::Round($row.aging.days91To180, 2)
                    above180 = [Math]::Round($row.aging.above180, 2)
                }
            })
        }

        $pageRows = @(
            $rows |
                Sort-Object accountName
        )

        $pageDetails = @{}
        foreach ($row in $pageRows) {
            $pageDetails[[string]$row.accountCode] =
                @($detailsByAccount[[int]$row.accountCode])
        }

        $totalReceivable = [Math]::Round(
            (@($pageRows | Measure-Object receivable -Sum).Sum), 2
        )
        $totalPayable = [Math]::Round(
            (@($pageRows | Measure-Object payable -Sum).Sum), 2
        )
        $overdueReceivable = [Math]::Round(
            (@(
                $pageRows |
                    Where-Object type -eq "receivable" |
                    Measure-Object overdueAmount -Sum
            ).Sum), 2
        )
        $overduePayable = [Math]::Round(
            (@(
                $pageRows |
                    Where-Object type -eq "payable" |
                    Measure-Object overdueAmount -Sum
            ).Sum), 2
        )

        return @{
            success = $true
            asOfDate = $asOfDate.ToString("yyyy-MM-dd")
            generatedAt = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
            page = $Page
            pageSize = $PageSize
            total = $masterPage.total
            totalPages = [Math]::Max(
                1, [Math]::Ceiling($masterPage.total / $PageSize)
            )
            summaryScope = "page"
            summary = @{
                totalReceivable = $totalReceivable
                totalPayable = $totalPayable
                netPosition = [Math]::Round(
                    $totalReceivable - $totalPayable, 2
                )
                receivableParties =
                    @($pageRows | Where-Object type -eq "receivable").Count
                payableParties =
                    @($pageRows | Where-Object type -eq "payable").Count
                overdueReceivable = $overdueReceivable
                overduePayable = $overduePayable
                totalOutstanding = [Math]::Round(
                    $totalReceivable + $totalPayable, 2
                )
                totalBills = @(
                    $pageRows | Measure-Object billCount -Sum
                ).Sum
                totalOnAccountSigned = [Math]::Round(
                    (@($pageRows | Measure-Object onAccountSigned -Sum).Sum), 2
                )
                totalOnAccountAmount = [Math]::Round(
                    (@($pageRows | Measure-Object onAccountAmount -Sum).Sum), 2
                )
                aging = @{
                    notDue = [Math]::Round(
                        (@($pageRows | ForEach-Object { $_.aging.notDue } |
                            Measure-Object -Sum).Sum), 2
                    )
                    days0To30 = [Math]::Round(
                        (@($pageRows | ForEach-Object { $_.aging.days0To30 } |
                            Measure-Object -Sum).Sum), 2
                    )
                    days31To60 = [Math]::Round(
                        (@($pageRows | ForEach-Object { $_.aging.days31To60 } |
                            Measure-Object -Sum).Sum), 2
                    )
                    days61To90 = [Math]::Round(
                        (@($pageRows | ForEach-Object { $_.aging.days61To90 } |
                            Measure-Object -Sum).Sum), 2
                    )
                    days91To180 = [Math]::Round(
                        (@($pageRows | ForEach-Object { $_.aging.days91To180 } |
                            Measure-Object -Sum).Sum), 2
                    )
                    above180 = [Math]::Round(
                        (@($pageRows | ForEach-Object { $_.aging.above180 } |
                            Measure-Object -Sum).Sum), 2
                    )
                }
            }
            data = $pageRows
            detailsByAccount = $pageDetails
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