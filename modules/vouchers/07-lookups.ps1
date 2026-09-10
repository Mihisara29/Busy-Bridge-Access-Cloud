# 07-lookups.ps1
# Extracted from vouchers.ps1. Keep functions behavior-compatible with the monolith.

function Get-VoucherSeries {
    param(
        [int]$VchType = 0,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $requestCacheKey = (
        "{0}|{1}|voucherseries|{2}" -f
        $InstanceId,
        $CompanyCode,
        $VchType
    ).ToLowerInvariant()

    $cached = Get-Cache $requestCacheKey
    if ($cached) {
        Write-Host (
            "  [SERIES-FAST-V2] cache HIT {0}/{1} type={2}" -f
            $InstanceId,
            $CompanyCode,
            $VchType
        ) -ForegroundColor DarkCyan

        return $cached
    }

    # Cache the RAW series master once per company. Every voucher type can then
    # be filtered in memory, avoiding repeated scans of Master1.
    if ($null -eq $script:BusyCloudVoucherSeriesMasterCache) {
        $script:BusyCloudVoucherSeriesMasterCache = @{}
    }

    $masterCacheKey = (
        "{0}|{1}|all-series" -f
        $InstanceId,
        $CompanyCode
    ).ToLowerInvariant()

    $allSeries = $null

    if ($script:BusyCloudVoucherSeriesMasterCache.ContainsKey($masterCacheKey)) {
        $entry = $script:BusyCloudVoucherSeriesMasterCache[$masterCacheKey]

        if (
            $entry -and
            $entry.expires -and
            (Get-Date) -lt $entry.expires
        ) {
            $allSeries = @($entry.data)
        }
        else {
            try {
                $script:BusyCloudVoucherSeriesMasterCache.Remove(
                    $masterCacheKey
                )
            }
            catch {
            }
        }
    }

    $startedAt = [System.Diagnostics.Stopwatch]::StartNew()
    $ctx = $null
    $reader = $null
    $cmd = $null

    try {
        if ($null -eq $allSeries) {
            $ctx = Get-BusyCloudFastConfigDbContext `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode

            if (
                $null -eq $ctx -or
                $null -eq $ctx.connection
            ) {
                throw "Direct fiscal database connection is unavailable."
            }

            $conn = $ctx.connection
            $dbType = [int]$ctx.dbType

            $cmd = $conn.CreateCommand()

            # 10 seconds is only the hard ceiling. SQL NOLOCK below prevents
            # normal write locks from making the first Sale form fail after 5 s.
            try { $cmd.CommandTimeout = 10 } catch {}

            if ($dbType -eq 1) {
                $cmd.CommandText = @"
SELECT
    Code,
    Name,
    I1,
    ParentGrp,
    CM1,
    CM2
FROM Master1 WITH (NOLOCK)
WHERE MasterType = 21
"@
            }
            else {
                $cmd.CommandText = @"
SELECT
    Code,
    Name,
    I1,
    ParentGrp,
    CM1,
    CM2
FROM Master1
WHERE MasterType = 21
"@
            }

            $reader = $cmd.ExecuteReader()

            $loaded = @()

            while ($reader.Read()) {
                $code = 0
                $name = ""
                $i1 = 0
                $parentGrp = 0
                $cm1 = 0
                $cm2 = 0

                try {
                    if (-not $reader.IsDBNull(0)) {
                        $code = [int][string]$reader.GetValue(0)
                    }
                }
                catch {}

                try {
                    if (-not $reader.IsDBNull(1)) {
                        $name = ([string]$reader.GetValue(1)).Trim()
                    }
                }
                catch {}

                foreach ($pair in @(
                    @{ index = 2; target = "i1" },
                    @{ index = 3; target = "parentGrp" },
                    @{ index = 4; target = "cm1" },
                    @{ index = 5; target = "cm2" }
                )) {
                    $value = 0
                    try {
                        if (-not $reader.IsDBNull([int]$pair.index)) {
                            $value = [int][string]$reader.GetValue(
                                [int]$pair.index
                            )
                        }
                    }
                    catch {
                        $value = 0
                    }

                    switch ([string]$pair.target) {
                        "i1"        { $i1 = $value }
                        "parentGrp" { $parentGrp = $value }
                        "cm1"       { $cm1 = $value }
                        "cm2"       { $cm2 = $value }
                    }
                }

                if ($code -gt 0 -and -not [string]::IsNullOrWhiteSpace($name)) {
                    $loaded += @{
                        code      = $code
                        rawName   = $name
                        I1        = $i1
                        ParentGrp = $parentGrp
                        CM1       = $cm1
                        CM2       = $cm2
                    }
                }
            }

            try { $reader.Close() } catch {}
            try { $reader.Dispose() } catch {}
            $reader = $null

            try { $cmd.Dispose() } catch {}
            $cmd = $null

            $allSeries = @($loaded)

            $script:BusyCloudVoucherSeriesMasterCache[$masterCacheKey] = @{
                expires = (Get-Date).AddMinutes(5)
                data    = @($allSeries)
            }
        }

        $filtered = @()

        foreach ($series in @($allSeries)) {
            $seriesVchType = 0

            foreach ($candidate in @(
                $series.I1,
                $series.ParentGrp,
                $series.CM1,
                $series.CM2
            )) {
                $parsed = 0

                if (
                    [int]::TryParse([string]$candidate, [ref]$parsed) -and
                    $parsed -gt 0
                ) {
                    $seriesVchType = $parsed
                    break
                }
            }

            $cleanName = ([string]$series.rawName).Trim()

            if (
                $seriesVchType -gt 0 -and
                $cleanName.StartsWith(
                    ("{0:D2}" -f $seriesVchType),
                    [System.StringComparison]::OrdinalIgnoreCase
                )
            ) {
                $cleanName = $cleanName.Substring(2)
            }

            if ($VchType -le 0 -or $seriesVchType -eq $VchType) {
                $filtered += @{
                    code    = [int]$series.code
                    name    = $cleanName
                    vchType = $seriesVchType
                }
            }
        }

        # Preserve the previous compatibility fallback.
        if ($VchType -gt 0 -and $filtered.Count -eq 0 -and $allSeries.Count -gt 0) {
            foreach ($series in @($allSeries)) {
                $cleanName = ([string]$series.rawName).Trim()

                if (
                    $cleanName.Length -ge 2 -and
                    $cleanName.Substring(0, 2) -match '^\d{2}$'
                ) {
                    $cleanName = $cleanName.Substring(2)
                }

                $filtered += @{
                    code    = [int]$series.code
                    name    = $cleanName
                    vchType = 0
                }
            }
        }

        $filtered = @(
            $filtered |
            Sort-Object name
        )

        $result = @{
            success = $true
            data    = @($filtered)
        }

        Set-Cache $requestCacheKey $result

        $startedAt.Stop()

        Write-Host (
            "  [SERIES-FAST-V2] {0}/{1} type={2} rows={3} elapsedMs={4}" -f
            $InstanceId,
            $CompanyCode,
            $VchType,
            @($filtered).Count,
            [int]$startedAt.ElapsedMilliseconds
        ) -ForegroundColor DarkCyan

        return $result
    }
    catch {
        if ($startedAt.IsRunning) {
            $startedAt.Stop()
        }

        Write-Host (
            "  [SERIES-FAST-V2 FAIL] {0}/{1} type={2} elapsedMs={3} error={4}" -f
            $InstanceId,
            $CompanyCode,
            $VchType,
            [int]$startedAt.ElapsedMilliseconds,
            $_.Exception.Message
        ) -ForegroundColor Red

        return @{
            success = $false
            error = $_.Exception.Message
        }
    }
    finally {
        if ($reader) {
            try { $reader.Close() } catch {}
            try { $reader.Dispose() } catch {}
        }

        if ($cmd) {
            try { $cmd.Dispose() } catch {}
        }

        if (
            $ctx -and
            $ctx.connection
        ) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}

function Get-VoucherInputTypes {
    return @{
        success = $true
        data = @(
            @{ code = 1; name = "Direct" }
            @{ code = 2; name = "Against Challan" }
        )
    }
}

function Get-PendingOrders {
    param(
        [int]$VchType,[string]$Party = "",[string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    # Detect dbType
    $targetInst = Get-InstanceConfig -InstanceId $InstanceId
    $dbType = 0
    if ($null -ne $targetInst -and $null -ne $targetInst.dbType) { $dbType = [int]$targetInst.dbType }
    $isSql = ($dbType -eq 1)

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) {
        return @{ success = $false; error = "BUSY connection failed" }
    }
    
    try {
        $masterMap = @{}
        $mRst = $fi.GetRecordset("SELECT Code, Name FROM Master1")
        
        if ($null -ne $mRst -and -not $mRst.EOF) {
            $mRst.MoveFirst()
            while (-not $mRst.EOF) {
                $c = $mRst.Fields.Item("Code").Value
                $n = $mRst.Fields.Item("Name").Value
                if ($c -ne [System.DBNull]::Value) {
                    $masterMap[$c.ToString().Trim()] = if ($n -ne [System.DBNull]::Value) { $n.ToString().Trim() } else { "" }
                }
                $mRst.MoveNext()
            }
            $mRst.Close()
        }
        
        $partyCode = ""
        if ($Party -ne "") {
            $cleanParty = $Party.Trim().ToLower()
            foreach ($k in $masterMap.Keys) {
                if ($masterMap[$k].Trim().ToLower() -eq $cleanParty) {
                    $partyCode = $k
                    break
                }
            }
        }
        
        $whereOrder = "VchType = $VchType"
        if ($partyCode -ne "") {
            $whereOrder += " AND MasterCode1 = $partyCode"
        }

        # Access uses VAL(), SQL Server uses ISNUMERIC()/CAST()
        $orderSort = if ($isSql) {
            "Date DESC, CASE WHEN ISNUMERIC(VchNo)=1 THEN CAST(VchNo AS INT) ELSE 0 END DESC"
        } else {
            "Date DESC, VAL(VchNo) DESC"
        }
        
        $orderRst = $fi.GetRecordset("SELECT VchCode, VchNo, VchSeriesCode, Date, MasterCode1 FROM Tran1 WHERE $whereOrder ORDER BY $orderSort")
        $pendingOrders = @()
        $completedCount = 0

        if ($null -ne $orderRst -and -not $orderRst.EOF) {
            $orderRst.MoveFirst()
            while (-not $orderRst.EOF) {
                $oCode = [int]$orderRst.Fields.Item("VchCode").Value
                $vNo = $orderRst.Fields.Item("VchNo").Value.ToString().Trim()
                $sCode = $orderRst.Fields.Item("VchSeriesCode").Value.ToString().Trim()
                $mc1 = $orderRst.Fields.Item("MasterCode1").Value.ToString().Trim()
                
                $strDate = ""
                try {
                    $strDate = [datetime]::Parse($orderRst.Fields.Item("Date").Value.ToString()).ToString("dd-MM-yyyy")
                } catch {}
                
                $orderedQty = 0.0
                # SUM(Value1) works in both Access and SQL Server - no VAL() needed
                $r1 = $fi.GetRecordset("SELECT SUM(Value1) AS TQ FROM Tran3 WHERE Method=1 AND VchCode=$oCode AND RecType IN (4,5)")
                if ($r1 -and -not $r1.EOF) {
                    $v = $r1.Fields.Item("TQ").Value
                    if ($v -ne [System.DBNull]::Value) {
                        $orderedQty = [Math]::Abs([double]$v)  
                    }
                }
                if ($r1) { $r1.Close() }
                
                $consumedQty = 0.0
                $r2 = $fi.GetRecordset("SELECT SUM(T2.Value1) AS CQ FROM Tran3 T1, Tran3 T2 WHERE T1.RefCode = T2.RefCode AND T1.MasterCode1=T2.MasterCode1 AND T1.Method=1 AND T1.VchCode=$oCode AND T1.RecType IN (4,5) AND T2.Method=2")
                if ($r2 -and -not $r2.EOF) {
                    $v = $r2.Fields.Item("CQ").Value
                    if ($v -ne [System.DBNull]::Value) {
                        $consumedQty = [double]$v
                    }
                }
                if ($r2) { $r2.Close() }
                
                $remainingQty = [Math]::Round($orderedQty - [Math]::Abs($consumedQty), 3)
                
                if ($orderedQty -gt 0.001) {
                    $isComp = ($remainingQty -le 0.001)
                    
                    if ($isComp) {
                        if ($completedCount -lt 6) {
                            $pendingOrders += @{
                                vchCode = $oCode
                                vchNo = $vNo
                                vchSeries = if ($masterMap.ContainsKey($sCode)) { $masterMap[$sCode] } else { $sCode }
                                date = $strDate
                                party = if ($masterMap.ContainsKey($mc1)) { $masterMap[$mc1] } else { $mc1 }
                                remainingQty = 0
                                originalQty = $orderedQty
                                isCompleted = $true
                            }
                            $completedCount++
                        }
                    } else {
                        $pendingOrders += @{
                            vchCode = $oCode
                            vchNo = $vNo
                            vchSeries = if ($masterMap.ContainsKey($sCode)) { $masterMap[$sCode] } else { $sCode }
                            date = $strDate
                            party = if ($masterMap.ContainsKey($mc1)) { $masterMap[$mc1] } else { $mc1 }
                            remainingQty = $remainingQty
                            originalQty = $orderedQty
                            isCompleted = $false
                        }
                    }
                }
                $orderRst.MoveNext()
            }
            $orderRst.Close()
        }
        return @{ success = $true; count = $pendingOrders.Count; data = $pendingOrders }
    } catch {
        return @{ success = $false; error = $_.Exception.Message }
    } finally {
        Disconnect-BUSY $fi
    }
}

function Get-PendingChallans {
    param([int]$VchType, [string]$Party = "", [string]$InstanceId = "", [string]$CompanyCode = "")

    # Detect dbType
    $targetInst = Get-InstanceConfig -InstanceId $InstanceId
    $dbType = 0
    if ($null -ne $targetInst -and $null -ne $targetInst.dbType) { $dbType = [int]$targetInst.dbType }
    $isSql = ($dbType -eq 1)

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }

    try {
        $masterMap = @{}
        $mRst = $fi.GetRecordset("SELECT Code, Name FROM Master1")
        if ($null -ne $mRst -and -not $mRst.EOF) {
            $mRst.MoveFirst()
            while (-not $mRst.EOF) {
                $c = $mRst.Fields.Item("Code").Value
                $n = $mRst.Fields.Item("Name").Value
                if ($c -ne [System.DBNull]::Value) {
                    $masterMap[$c.ToString().Trim()] = if ($n -ne [System.DBNull]::Value) { $n.ToString().Trim() } else { "" }
                }
                $mRst.MoveNext()
            }
            $mRst.Close()
        }

        $partyCode = ""
        if ($Party -ne "") {
            $cleanParty = $Party.Trim().ToLower()
            foreach ($k in $masterMap.Keys) {
                if ($masterMap[$k].Trim().ToLower() -eq $cleanParty) { 
                    $partyCode = $k
                    break 
                }
            }
        }

        $whereChallan = "VchType = $VchType"
        if ($partyCode -ne "") { $whereChallan += " AND MasterCode1 = $partyCode" }

        # Access uses VAL(), SQL Server uses ISNUMERIC()/CAST()
        $challanSort = if ($isSql) {
            "Date DESC, CASE WHEN ISNUMERIC(VchNo)=1 THEN CAST(VchNo AS INT) ELSE 0 END DESC"
        } else {
            "Date DESC, VAL(VchNo) DESC"
        }

        $challanRst = $fi.GetRecordset("SELECT VchCode, VchNo, VchSeriesCode, Date, MasterCode1 FROM Tran1 WHERE $whereChallan ORDER BY $challanSort")
        $pendingChallans = @()

        $targetRecType = 3 

        if ($null -ne $challanRst -and -not $challanRst.EOF) {
            $challanRst.MoveFirst()
            while (-not $challanRst.EOF) {
                $cCode  = [int]$challanRst.Fields.Item("VchCode").Value
                $vNo    = $challanRst.Fields.Item("VchNo").Value.ToString().Trim()
                $sCode  = $challanRst.Fields.Item("VchSeriesCode").Value.ToString().Trim()
                $mc1    = $challanRst.Fields.Item("MasterCode1").Value.ToString().Trim()
                $strDate = ""
                try { $strDate = [datetime]::Parse($challanRst.Fields.Item("Date").Value.ToString()).ToString("dd-MM-yyyy") } catch {}

                $totalRemaining = 0.0
                $refRst = $fi.GetRecordset("SELECT RefCode, Value1 FROM Tran3 WHERE Method=1 AND VchCode=$cCode AND RecType=$targetRecType")
                if ($refRst -and -not $refRst.EOF) {
                    $refRst.MoveFirst()
                    while (-not $refRst.EOF) {
                        $refCode = [int]$refRst.Fields.Item("RefCode").Value
                        $origQty = [Math]::Abs([double]$refRst.Fields.Item("Value1").Value)

                        $consumedQty = 0.0
                        # ABS() works in both Access and SQL Server
                        $consRst = $fi.GetRecordset("SELECT SUM(ABS(Value1)) AS CQ FROM Tran3 WHERE Method=2 AND RefCode=$refCode AND RecType=$targetRecType")
                        if ($consRst -and -not $consRst.EOF) {
                            $v = $consRst.Fields.Item("CQ").Value
                            if ($v -ne [System.DBNull]::Value) { $consumedQty = [Math]::Abs([double]$v) }
                        }
                        if ($consRst) { $consRst.Close() }

                        $totalRemaining += ($origQty - $consumedQty)
                        $refRst.MoveNext()
                    }
                    $refRst.Close()
                }

                $remainingQty = [Math]::Round($totalRemaining, 3)

                if ($remainingQty -gt 0.001) {
                    $pendingChallans += @{
                        vchCode      = $cCode
                        vchNo        = $vNo
                        vchSeries    = if ($masterMap.ContainsKey($sCode)) { $masterMap[$sCode] } else { $sCode }
                        date         = $strDate
                        party        = if ($masterMap.ContainsKey($mc1)) { $masterMap[$mc1] } else { $mc1 }
                        remainingQty = $remainingQty
                        originalQty  = $remainingQty
                        isCompleted  = $false
                    }
                }
                $challanRst.MoveNext()
            }
            $challanRst.Close()
        }
        return @{ success = $true; count = $pendingChallans.Count; data = $pendingChallans }
    } catch {
        return @{ success = $false; error = $_.Exception.Message }
    } finally {
        Disconnect-BUSY $fi
    }
}

