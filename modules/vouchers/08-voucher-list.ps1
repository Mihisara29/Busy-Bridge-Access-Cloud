# 08-voucher-list.ps1
# Fast voucher list implementation.
# SQL path avoids COM, full Master1 scans, unbounded Tran1 reads, and PowerShell += loops.

if ($null -eq $script:VoucherListBurstCache) { $script:VoucherListBurstCache = @{} }

function Get-Vouchers {
    param(
        [int]$VchType,
        [string]$From = "",
        [string]$To = "",
        [string]$Party = "",
        [string]$Search = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $cfg = $script:VoucherConfig[$VchType]
    if (-not $cfg) { return @{ success = $false; error = "Unsupported vchType" } }

    # Coalesce duplicate requests caused by React renders/filter transitions.
    $cacheKey = "$InstanceId|$CompanyCode|$VchType|$From|$To|$Party|$Search"
    try {
        if ($script:VoucherListBurstCache.ContainsKey($cacheKey)) {
            $cached = $script:VoucherListBurstCache[$cacheKey]
            if ($null -ne $cached -and ((Get-Date) - $cached.at).TotalSeconds -lt 2) {
                return $cached.value
            }
        }
    } catch {}

    $targetInst = Get-InstanceConfig -InstanceId $InstanceId
    $dbType = 0
    if ($null -ne $targetInst -and $null -ne $targetInst.dbType) { $dbType = [int]$targetInst.dbType }
    $isSql = ($dbType -eq 1)

    if ($isSql) {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $dirConn = Get-FastVoucherDirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
        if ($null -eq $dirConn) {
            return @{ success = $false; error = "Could not build direct database connection" }
        }

        $sqlConn = $dirConn.connection
        try {
            $sqlConn.Open()

            $stptColExists = $false
            try {
                $colCmd = $sqlConn.CreateCommand()
                $colCmd.CommandTimeout = 5
                $colCmd.CommandText = "SELECT CASE WHEN COL_LENGTH('dbo.Tran1','STPTCode') IS NULL THEN 0 ELSE 1 END"
                $stptColExists = ([int]$colCmd.ExecuteScalar() -eq 1)
            } catch {}

            $typeCodeExpr = if ($stptColExists) { "T.[STPTCode]" } else { "T.[CM1]" }

            $where = [System.Collections.Generic.List[string]]::new()
            $where.Add("T.[VchType] = @VchType")

            $cmd = $sqlConn.CreateCommand()
            $cmd.CommandTimeout = 15
            $null = $cmd.Parameters.Add("@VchType", [System.Data.SqlDbType]::Int)
            $cmd.Parameters["@VchType"].Value = $VchType

            if ($From) {
                $d = Safe-ParseDate -dateStr $From
                if ($null -ne $d) {
                    $where.Add("T.[Date] >= @FromDate")
                    $null = $cmd.Parameters.Add("@FromDate", [System.Data.SqlDbType]::DateTime)
                    $cmd.Parameters["@FromDate"].Value = $d.Date
                }
            }
            if ($To) {
                $d = Safe-ParseDate -dateStr $To
                if ($null -ne $d) {
                    $where.Add("T.[Date] < @ToExclusive")
                    $null = $cmd.Parameters.Add("@ToExclusive", [System.Data.SqlDbType]::DateTime)
                    $cmd.Parameters["@ToExclusive"].Value = $d.Date.AddDays(1)
                }
            }
            if ($Party) {
                $where.Add("ISNULL(P.[Name], '') = @Party")
                $null = $cmd.Parameters.Add("@Party", [System.Data.SqlDbType]::NVarChar, 200)
                $cmd.Parameters["@Party"].Value = $Party.Trim()
            }
            if ($Search) {
                $where.Add("(ISNULL(CONVERT(nvarchar(100), T.[VchNo]), '') LIKE @Search OR ISNULL(P.[Name], '') LIKE @Search OR ISNULL(S.[Name], '') LIKE @Search OR ISNULL(VT.[Name], '') LIKE @Search)")
                $null = $cmd.Parameters.Add("@Search", [System.Data.SqlDbType]::NVarChar, 220)
                $cmd.Parameters["@Search"].Value = "%$($Search.Trim())%"
            }

            $whereSql = [string]::Join(" AND ", $where)
            $cmd.CommandText = @"
SELECT TOP (1000)
    T.[VchCode],
    T.[VchNo],
    T.[AutoVchNo],
    T.[Date],
    T.[VchAmtBaseCur],
    ISNULL(P.[Name], '') AS [PartyName],
    ISNULL(MC.[Name], '') AS [MatCentreName],
    ISNULL(S.[Name], '') AS [SeriesName],
    ISNULL(VT.[Name], '') AS [VoucherTypeName]
FROM [dbo].[Tran1] T WITH (NOLOCK)
LEFT JOIN [dbo].[Master1] P WITH (NOLOCK) ON P.[Code] = T.[MasterCode1]
LEFT JOIN [dbo].[Master1] MC WITH (NOLOCK) ON MC.[Code] = T.[MasterCode2]
LEFT JOIN [dbo].[Master1] S WITH (NOLOCK) ON S.[Code] = T.[VchSeriesCode]
LEFT JOIN [dbo].[Master1] VT WITH (NOLOCK) ON VT.[Code] = $typeCodeExpr
WHERE $whereSql
ORDER BY T.[Date] DESC, T.[VchCode] DESC
"@

            $rows = [System.Collections.Generic.List[object]]::new()
            $codes = [System.Collections.Generic.List[int]]::new()
            $rdr = $cmd.ExecuteReader()
            try {
                while ($rdr.Read()) {
                    $vchCode = 0
                    try { $vchCode = [int]$rdr["VchCode"] } catch {}
                    if ($vchCode -gt 0) { $codes.Add($vchCode) }

                    $vNo = ""
                    try { if ($rdr["VchNo"] -ne [DBNull]::Value) { $vNo = [string]$rdr["VchNo"] } } catch {}
                    if ([string]::IsNullOrWhiteSpace($vNo)) {
                        try { if ($rdr["AutoVchNo"] -ne [DBNull]::Value) { $vNo = [string]$rdr["AutoVchNo"] } } catch {}
                    }

                    $dateText = ""
                    try {
                        if ($rdr["Date"] -ne [DBNull]::Value) { $dateText = ([datetime]$rdr["Date"]).ToString("dd-MM-yyyy") }
                    } catch {}

                    $amt = 0.0
                    try { if ($rdr["VchAmtBaseCur"] -ne [DBNull]::Value) { $amt = [Math]::Abs([double]$rdr["VchAmtBaseCur"]) } } catch {}

                    $rows.Add([ordered]@{
                        vchNo = $vNo.Trim()
                        vchSeries = ([string]$rdr["SeriesName"]).Trim()
                        date = $dateText
                        party = ([string]$rdr["PartyName"]).Trim()
                        matCentre = ([string]$rdr["MatCentreName"]).Trim()
                        debitAccount = "-"
                        creditAccount = "-"
                        voucherType = ([string]$rdr["VoucherTypeName"]).Trim()
                        amount = $amt
                        vchType = $VchType
                        __vchCode = $vchCode
                    })
                }
            } finally {
                $rdr.Close()
            }

            # One batched account query instead of one Tran2 query per voucher.
            if ($VchType -in @(14,15,16,19) -and $codes.Count -gt 0) {
                $codeSql = [string]::Join(",", $codes)
                $accCmd = $sqlConn.CreateCommand()
                $accCmd.CommandTimeout = 10
                $accCmd.CommandText = @"
SELECT T2.[VchCode], T2.[Value1], ISNULL(M.[Name], CONVERT(nvarchar(50), T2.[MasterCode1])) AS [AccountName]
FROM [dbo].[Tran2] T2 WITH (NOLOCK)
LEFT JOIN [dbo].[Master1] M WITH (NOLOCK) ON M.[Code] = T2.[MasterCode1]
WHERE T2.[VchCode] IN ($codeSql) AND T2.[RecType] IN (1,2)
"@
                $debits = @{}
                $credits = @{}
                $ar = $accCmd.ExecuteReader()
                try {
                    while ($ar.Read()) {
                        $code = [int]$ar["VchCode"]
                        $value = 0.0
                        try { $value = [double]$ar["Value1"] } catch {}
                        $name = ([string]$ar["AccountName"]).Trim()
                        if ($value -gt 0) {
                            if (-not $debits.ContainsKey($code)) { $debits[$code] = [System.Collections.Generic.List[string]]::new() }
                            if ($name) { $debits[$code].Add($name) }
                        } elseif ($value -lt 0) {
                            if (-not $credits.ContainsKey($code)) { $credits[$code] = [System.Collections.Generic.List[string]]::new() }
                            if ($name) { $credits[$code].Add($name) }
                        }
                    }
                } finally { $ar.Close() }

                foreach ($row in $rows) {
                    $code = [int]$row.__vchCode
                    if ($debits.ContainsKey($code)) { $row.debitAccount = [string]::Join(", ", $debits[$code]) }
                    if ($credits.ContainsKey($code)) { $row.creditAccount = [string]::Join(", ", $credits[$code]) }
                }
            }

            $output = [System.Collections.Generic.List[object]]::new()
            foreach ($row in $rows) {
                $row.Remove("__vchCode")
                $output.Add($row)
            }

            $sw.Stop()
            Write-Host "  [VOUCHER-LIST-FAST] type=$VchType rows=$($output.Count) elapsedMs=$($sw.ElapsedMilliseconds)" -ForegroundColor DarkCyan
            $result = @{ success = $true; count = $output.Count; vchType = $VchType; data = @($output); capped = ($output.Count -ge 1000) }
            try { $script:VoucherListBurstCache[$cacheKey] = @{ at = Get-Date; value = $result } } catch {}
            return $result
        } catch {
            $sw.Stop()
            return @{ success = $false; error = $_.Exception.Message }
        } finally {
            if ($null -ne $sqlConn) { try { $sqlConn.Close() } catch {} }
        }
    } else {
        # -- ACCESS PATH: Original stable COM recordset query --
        $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
        if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
        
        try {
            $masterMap = @{}
            try {
                $mRst = $fi.GetRecordset("SELECT Code, Name FROM Master1")
                if ($null -ne $mRst -and -not $mRst.EOF) {
                    $mRst.MoveFirst()
                    while (-not $mRst.EOF) {
                        $c = $mRst.Fields.Item("Code").Value
                        $n = $mRst.Fields.Item("Name").Value
                        if ($c -ne [System.DBNull]::Value -and $null -ne $c) {
                            $cStr = $c.ToString().Trim()
                            $nStr = if ($n -ne [System.DBNull]::Value -and $null -ne $n) { $n.ToString().Trim() } else { "" }
                            $masterMap[$cStr] = $nStr
                        }
                        $mRst.MoveNext()
                    }
                }
                if ($null -ne $mRst) { try { $mRst.Close() } catch {} }
            } catch {}

            # 2. Build Where Filter (Using culture-insensitive parser with hash '#' marks for Access)
            $where = "VchType = $VchType"
            if ($From -ne "") {
                $dFrom = Safe-ParseDate -dateStr $From
                if ($null -ne $dFrom) {
                    $where += " AND [Date] >= #$($dFrom.ToString('yyyy-MM-dd'))#"
                }
            }
            if ($To -ne "") {
                $dTo = Safe-ParseDate -dateStr $To
                if ($null -ne $dTo) {
                    $where += " AND [Date] <= #$($dTo.ToString('yyyy-MM-dd'))#"
                }
            }
            if ($Party -ne "") {
                $pCode = ""
                foreach ($k in $masterMap.Keys) {
                    if ($masterMap[$k] -eq $Party) { $pCode = $k; break }
                }
                if ($pCode -ne "") { $where += " AND MasterCode1 = $pCode" }
            }

            $qry = "SELECT * FROM Tran1 WHERE $where ORDER BY VAL(VchNo) DESC"
            $rst = $fi.GetRecordset($qry)
            if ($null -eq $rst) {
                $rst = $fi.GetRecordset("SELECT * FROM Tran1 WHERE VchType = $VchType ORDER BY VAL(VchNo) DESC")
            }

            $vouchers = @()
            if ($null -ne $rst -and -not $rst.EOF) {
                $rst.MoveFirst()
                while (-not $rst.EOF) {
                    $vNo = ""
                    try {
                        $v = $rst.Fields.Item("VchNo").Value
                        if ($v -ne [System.DBNull]::Value -and $null -ne $v) { $vNo = $v.ToString().Trim() }
                        if ($vNo -eq "") {
                            $av = $rst.Fields.Item("AutoVchNo").Value
                            if ($av -ne [System.DBNull]::Value -and $null -ne $av) { $vNo = $av.ToString().Trim() }
                        }
                    } catch {}

                    $strDate = ""
                    try {
                        $rawDate = $rst.Fields.Item("Date").Value
                        if ($rawDate -is [datetime]) { $strDate = $rawDate.ToString("dd-MM-yyyy") }
                        elseif ($rawDate -ne [System.DBNull]::Value -and $rawDate) {
                            $strDate = [datetime]::Parse($rawDate.ToString()).ToString("dd-MM-yyyy")
                        }
                    } catch {}

                    $vchCode = 0
                    try { $vchCode = [int]$rst.Fields.Item("VchCode").Value } catch {}

                    $debitAcc = ""
                    $creditAcc = ""
                    $partyName = ""

                    if ($VchType -in @(14, 15, 16, 19)) {
                        try {
                            $r2 = $fi.GetRecordset("SELECT MasterCode1, Value1 FROM Tran2 WHERE VchCode = $vchCode AND RecType IN (1,2)")
                            if ($null -ne $r2 -and -not $r2.EOF) {
                                $r2.MoveFirst()
                                $dAccList = @()
                                $cAccList = @()
                                while (-not $r2.EOF) {
                                    $mc1 = $r2.Fields.Item("MasterCode1").Value
                                    $val1 = $r2.Fields.Item("Value1").Value
                                    if ($null -ne $mc1 -and $mc1 -ne [System.DBNull]::Value -and $null -ne $val1 -and $val1 -ne [System.DBNull]::Value) {
                                        $amt2 = [double]$val1
                                        $accStr = $mc1.ToString().Trim()
                                        $accName = if ($masterMap.ContainsKey($accStr)) { $masterMap[$accStr] } else { $accStr }
                                        if ($amt2 -gt 0) { $dAccList += $accName } elseif ($amt2 -lt 0) { $cAccList += $accName }
                                    }
                                    $r2.MoveNext()
                                }
                                $r2.Close()
                                $debitAcc = $dAccList -join ", "
                                $creditAcc = $cAccList -join ", "
                            }
                        } catch {}
                    } else {
                        $pCode = ""
                        try {
                            $v = $rst.Fields.Item("MasterCode1").Value
                            if ($v -ne [System.DBNull]::Value -and $null -ne $v) { $pCode = $v.ToString().Trim() }
                        } catch {}
                        $partyName = if ($masterMap.ContainsKey($pCode)) { $masterMap[$pCode] } else { $pCode }

                        $p2Code = ""
                        try {
                            $v2 = $rst.Fields.Item("MasterCode2").Value
                            if ($v2 -ne [System.DBNull]::Value -and $null -ne $v2) { $p2Code = $v2.ToString().Trim() }
                        } catch {}
                        $toCenterName = if ($masterMap.ContainsKey($p2Code)) { $masterMap[$p2Code] } else { $p2Code }
                    }

                    $sCode = ""
                    try {
                        $v = $rst.Fields.Item("VchSeriesCode").Value
                        if ($v -ne [System.DBNull]::Value -and $null -ne $v) { $sCode = $v.ToString().Trim() }
                    } catch {}
                    $seriesName = if ($masterMap.ContainsKey($sCode)) { $masterMap[$sCode] } else { $sCode }

                    $tCode = ""
                    try {
                        $v = $rst.Fields.Item("STPTCode").Value
                        if ($v -ne [System.DBNull]::Value -and $null -ne $v) { $tCode = $v.ToString().Trim() }
                    } catch {
                        try {
                            $v = $rst.Fields.Item("CM1").Value
                            if ($v -ne [System.DBNull]::Value -and $null -ne $v) { $tCode = $v.ToString().Trim() }
                        } catch {}
                    }
                    $typeName = if ($masterMap.ContainsKey($tCode)) { $masterMap[$tCode] } else { $tCode }

                    $amt = 0
                    $amountCols = @("VchAmtBaseCur","OrgVchAmtBaseCur","VchSalePurcAmt","Value1","Value2","Value3","VchAmt","TotalAmt","NettAmt")
                    foreach ($col in $amountCols) {
                        try {
                            $v = $rst.Fields.Item($col).Value
                            if ($v -ne [System.DBNull]::Value -and $null -ne $v) {
                                $dVal = [Convert]::ToDouble($v.ToString(),[System.Globalization.CultureInfo]::InvariantCulture)
                                if ($dVal -ne 0) {
                                    $amt = [math]::Abs($dVal)
                                    break
                                }
                            }
                        } catch {}
                    }

                    $vouchers += @{
                        vchNo = $vNo
                        vchSeries = $seriesName
                        date = $strDate
                        party = $partyName
                        matCentre = $toCenterName
                        debitAccount = if ($debitAcc -ne "") { $debitAcc } else { "-" }
                        creditAccount = if ($creditAcc -ne "") { $creditAcc } else { "-" }
                        voucherType = $typeName
                        amount = $amt
                        vchType = $VchType
                    }
                    $rst.MoveNext()
                }
                try { $rst.Close() } catch {}
            }
            return @{ success = $true; count = $vouchers.Count; vchType = $VchType; data = $vouchers }
        } catch {
            return @{ success = $false; error = $_.Exception.Message }
        } finally {
            Disconnect-BUSY $fi
        }
    }
}

# ===============================================================
#  MATERIAL-CENTRE-SPECIFIC CURRENT AVERAGE ITEM COST
# ===============================================================
# Replays BUSY stock layers from opening stock + inventory vouchers.
# FIFO is used to maintain the remaining layers; the UI cost is the
# weighted average of those remaining layers, not the next FIFO rate.
# ===============================================================

