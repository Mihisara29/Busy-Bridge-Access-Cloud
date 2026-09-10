# 10-voucher-details.ps1
# Extracted from vouchers.ps1. Keep functions behavior-compatible with the monolith.

function Get-VoucherDetail {
    param(
        [int]$VchType,
        [string]$VchNo,
        [string]$VchSeries,
        [string]$VchDate,
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        [switch]$DebugStockJournal
    )

    function Write-SJDebug {
        param([string]$Message)
        if ($VchType -eq 8 -or $DebugStockJournal) {
            try { Write-Host "[StockJournal Detail DEBUG] $Message" -ForegroundColor Cyan } catch {}
        }
    }

    function Get-NodeTextSafe {
        param($Node)
        try {
            if ($null -eq $Node) { return "" }
            $txt = ([string]$Node).Trim()
            return $txt
        } catch { return "" }
    }

    function Get-DoubleSafe {
        param($Node, [double]$DefaultValue = 0.0)
        try {
            if ($null -eq $Node) { return $DefaultValue }
            $txt = ([string]$Node).Trim()
            if ($txt -eq "") { return $DefaultValue }
            return [double]::Parse($txt, [System.Globalization.CultureInfo]::InvariantCulture)
        } catch {
            try { return [double]$Node } catch { return $DefaultValue }
        }
    }

    function Get-IntSafe {
        param($Node, [int]$DefaultValue = 0)
        try {
            if ($null -eq $Node) { return $DefaultValue }
            $txt = ([string]$Node).Trim()
            if ($txt -eq "") { return $DefaultValue }
            return [int]$txt
        } catch {
            try { return [int]$Node } catch { return $DefaultValue }
        }
    }

    # Item master lookup populated after the BUSY connection is opened.
    # Key: lower-case item name
    # Value: @{ code = <item master code>; alias = <item alias> }
    $itemMasterByName = @{}

    function Convert-VoucherXmlItemToHash {
        param(
            $d,
            [int]$ForcedItemType,
            [int]$OutputSrNo,
            $consumptionMap,
            $challanItemRemainingMap
        )

        $itemName = Get-NodeTextSafe $d.ItemName
        $unitName = Get-NodeTextSafe $d.UnitName
        $mcName   = Get-NodeTextSafe $d.MC

        # Resolve the BUSY item master code and alias from the lookup.
        $itemCode = 0
        $itemAlias = ""

        try {
            $itemKey = $itemName.Trim().ToLowerInvariant()

            if (
                $itemKey -ne "" -and
                $null -ne $itemMasterByName -and
                $itemMasterByName.ContainsKey($itemKey)
            ) {
                $itemMaster = $itemMasterByName[$itemKey]
                $itemCode = [int]$itemMaster.code
                $itemAlias = [string]$itemMaster.alias
            }
        } catch {
            $itemCode = 0
            $itemAlias = ""
        }

        $qty      = Get-DoubleSafe $d.Qty 0.0
        $price    = Get-DoubleSafe $d.Price 0.0
        $amount   = Get-DoubleSafe $d.Amt 0.0
        $mainQty  = Get-DoubleSafe $d.QtyMainUnit $qty

        $altUnit   = Get-NodeTextSafe $d.AltUnitName
        $altQty    = Get-DoubleSafe $d.QtyAltUnit 0.0
        $altPrice  = Get-DoubleSafe $d.PriceAltUnit 0.0
        $conFactor = Get-DoubleSafe $d.ConFactor 1.0
        $conType   = Get-IntSafe $d.ConFactorType 1

        $itemType = $ForcedItemType
        if ($ForcedItemType -eq 0) {
            $itemType = Get-IntSafe $d.ItemType 1
        }

        $listPrice = Get-DoubleSafe $d.ListPrice $price
        $discVal = 0.0
        $discType = "U"
        try {
            if ($d.DiscountPercent) {
                $discType = "%"
                $discVal = Get-DoubleSafe $d.DiscountPercent 0.0
            } elseif ($d.Discount) {
                $discVal = Get-DoubleSafe $d.Discount 0.0
            }
        } catch {}

        $itemXmlSrNo = Get-IntSafe $d.SrNo 1
        $pendingQty = $qty

        try {
            if ($null -ne $challanItemRemainingMap -and $challanItemRemainingMap.ContainsKey($itemXmlSrNo)) {
                $pendingQty = $challanItemRemainingMap[$itemXmlSrNo]
            }
        } catch {}

        try {
            if ($null -ne $consumptionMap -and $consumptionMap.ContainsKey($itemName)) {
                $availableToDeduct = $consumptionMap[$itemName]
                if ($availableToDeduct -lt 0) {
                    if (([Math]::Abs($availableToDeduct)) -ge $qty) {
                        $pendingQty = 0
                        $consumptionMap[$itemName] = $availableToDeduct + $qty
                    } else {
                        $pendingQty = [Math]::Round($qty + $availableToDeduct, 3)
                        $consumptionMap[$itemName] = 0
                    }
                }
            }
        } catch {}

        return @{
            srNo            = $OutputSrNo
            itemCode        = $itemCode
            code            = $itemCode
            itemName        = $itemName
            alias           = $itemAlias
            itemAlias       = $itemAlias
            unit            = $unitName
            qty             = $qty
            listPrice       = $listPrice
            discVal         = $discVal
            discType        = $discType
            price           = $price
            amount          = $amount
            itemType        = $itemType
            mc              = $mcName
            altUnit         = $altUnit
            conFactor       = $conFactor
            conType         = $conType
            altQtyConFactor = $conFactor
            altQty          = $altQty
            altPrice        = $altPrice
            mainQty         = $mainQty
            pendingQty      = $pendingQty
        }
    }

    # Resolve if SQL or Access from instances.json
    $targetInst = Get-InstanceConfig -InstanceId $InstanceId
    $dbType = 0
    if ($null -ne $targetInst -and $null -ne $targetInst.dbType) { $dbType = [int]$targetInst.dbType }
    $isSql = ($dbType -eq 1)

    $vchCode = 0
    Write-SJDebug "Start Get-VoucherDetail. VchType=$VchType VchNo=$VchNo VchSeries=$VchSeries VchDate=$VchDate InstanceId=$InstanceId CompanyCode=$CompanyCode IsSql=$isSql"

    if ($isSql) {
        $dirConn = Get-FastVoucherDirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
        if ($null -eq $dirConn) { return @{ success = $false; error = "Could not build direct database connection" } }
        $conn = $dirConn.connection
        try {
            $conn.Open()
            $vchCode = Get-VchCode-Direct -conn $conn -VchType $VchType -VchNo $VchNo -VchDate $VchDate -isSql $true
            Write-SJDebug "SQL resolved VchCode=$vchCode"
        } finally {
            if ($null -ne $conn) { try { $conn.Close() } catch {} }
        }
    } else {
        $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
        if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
        try {
            $dateFilter = ""
            if ($VchDate -match "\d{2}-\d{2}-\d{4}") {
                $d = [datetime]::ParseExact($VchDate, "dd-MM-yyyy", $null)
                $dateFilter = " AND [Date]=#" + $d.ToString("MM/dd/yyyy") + "#"
            } elseif ($VchDate -match "\d{4}-\d{2}-\d{2}") {
                $d = [datetime]::ParseExact($VchDate, "yyyy-MM-dd", $null)
                $dateFilter = " AND [Date]=#" + $d.ToString("MM/dd/yyyy") + "#"
            }

            $targetNo = $VchNo.Trim().ToLower()
            $qry = "SELECT VchCode, VchNo, AutoVchNo FROM Tran1 WHERE VchType=$VchType" + $dateFilter
            Write-SJDebug "Access lookup query: $qry"
            $rst = $fi.GetRecordset($qry)

            if ($rst -and -not $rst.EOF) {
                $rst.MoveFirst()
                while (-not $rst.EOF) {
                    $dbVchNo = ""
                    try {
                        $v = $rst.Fields.Item("VchNo").Value
                        if ($v -ne [System.DBNull]::Value) { $dbVchNo = $v.ToString().Trim().ToLower() }
                        if ($dbVchNo -eq "") {
                            $av = $rst.Fields.Item("AutoVchNo").Value
                            if ($av -ne [System.DBNull]::Value) { $dbVchNo = $av.ToString().Trim().ToLower() }
                        }
                    } catch {}

                    if ($dbVchNo -eq $targetNo) {
                        $vchCode = [int]$rst.Fields.Item("VchCode").Value
                        break
                    }
                    $rst.MoveNext()
                }
                try { $rst.Close() } catch {}
            }

            if ($vchCode -eq 0) {
                $qryAll = "SELECT VchCode, VchNo, AutoVchNo FROM Tran1 WHERE VchType=$VchType"
                Write-SJDebug "Date-filter lookup failed. Fallback query: $qryAll"
                $rstAll = $fi.GetRecordset($qryAll)
                if ($rstAll -and -not $rstAll.EOF) {
                    $rstAll.MoveFirst()
                    while (-not $rstAll.EOF) {
                        $dbVchNo = ""
                        try {
                            $v = $rstAll.Fields.Item("VchNo").Value
                            if ($v -ne [System.DBNull]::Value) { $dbVchNo = $v.ToString().Trim().ToLower() }
                            if ($dbVchNo -eq "") {
                                $av = $rstAll.Fields.Item("AutoVchNo").Value
                                if ($av -ne [System.DBNull]::Value) { $dbVchNo = $av.ToString().Trim().ToLower() }
                            }
                        } catch {}

                        if ($dbVchNo -eq $targetNo) {
                            $vchCode = [int]$rstAll.Fields.Item("VchCode").Value
                            break
                        }
                        $rstAll.MoveNext()
                    }
                    try { $rstAll.Close() } catch {}
                }
            }
            Write-SJDebug "Access resolved VchCode=$vchCode"
        } finally {
            Disconnect-BUSY $fi
        }
    }

    if ($vchCode -eq 0) {
        return @{ success = $false; error = "Could not find Voucher in database (VchType=$VchType, No=$VchNo)" }
    }

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }

    try {
        # Load all item codes and aliases once.
        # This avoids one database query per voucher row.
        try {
            $itemRst = $fi.GetRecordset(@"
SELECT
    Code,
    Name,
    Alias
FROM Master1
WHERE MasterType = 6
"@)

            if ($itemRst -and -not $itemRst.EOF) {
                $itemRst.MoveFirst()

                while (-not $itemRst.EOF) {
                    $codeRaw = $itemRst.Fields.Item("Code").Value
                    $nameRaw = $itemRst.Fields.Item("Name").Value
                    $aliasRaw = $itemRst.Fields.Item("Alias").Value

                    $masterName = ""
                    $masterCode = 0
                    $masterAlias = ""

                    if (
                        $null -ne $nameRaw -and
                        $nameRaw -ne [System.DBNull]::Value
                    ) {
                        $masterName = $nameRaw.ToString().Trim()
                    }

                    if (
                        $null -ne $codeRaw -and
                        $codeRaw -ne [System.DBNull]::Value
                    ) {
                        try {
                            $masterCode = [int][string]$codeRaw
                        } catch {
                            $masterCode = 0
                        }
                    }

                    if (
                        $null -ne $aliasRaw -and
                        $aliasRaw -ne [System.DBNull]::Value
                    ) {
                        $masterAlias = $aliasRaw.ToString().Trim()
                    }

                    if (-not [string]::IsNullOrWhiteSpace($masterName)) {
                        $itemKey = $masterName.ToLowerInvariant()

                        $itemMasterByName[$itemKey] = @{
                            code  = $masterCode
                            alias = $masterAlias
                        }
                    }

                    $itemRst.MoveNext()
                }

                try { $itemRst.Close() } catch {}
            }

            Write-SJDebug "Loaded item alias lookup count=$($itemMasterByName.Count)"
        } catch {
            Write-SJDebug "Could not load item alias lookup: $($_.Exception.Message)"
            $itemMasterByName = @{}
        }

        $errMsg = ""
        $xmlStr = ""
        try {
            $xmlStr = $fi.GetVchXML($vchCode, [ref]$errMsg)
        } catch {
            $xmlStr = $fi.GetVchXML($vchCode)
        }
        if ($VchType -eq 6) {
    Write-Host ""
    Write-Host "========== PRODUCTION RAW XML ==========" -ForegroundColor Yellow
    Write-Host $xmlStr
    Write-Host "========== END PRODUCTION RAW XML ======" -ForegroundColor Yellow
    Write-Host ""
}

        if (-not $xmlStr) {
            return @{ success = $false; error = if ($errMsg) { $errMsg } else { "Voucher XML is empty" } }
        }

        Write-SJDebug "GetVchXML length=$($xmlStr.Length)"

        $xml = [xml]$xmlStr
        $root = $xml.DocumentElement

        if ($VchType -eq 8) {
            try {
                $childNames = @($root.ChildNodes | ForEach-Object { $_.Name }) -join ", "
                Write-SJDebug "Root=$($root.Name). Child nodes: $childNames"
                Write-SJDebug "ItemEntries nodes=$(@($root.SelectNodes('ItemEntries/ItemDetail')).Count), ConsumedItemEntries nodes=$(@($root.SelectNodes('ConsumedItemEntries/ItemDetail')).Count), ItemEntries1 nodes=$(@($root.SelectNodes('ItemEntries1/ItemDetail')).Count)"
                Write-SJDebug "Anywhere consumed nodes: ConsumedItemEntries=$(@($root.SelectNodes('//ConsumedItemEntries/ItemDetail')).Count), ItemEntries1=$(@($root.SelectNodes('//ItemEntries1/ItemDetail')).Count)"
            } catch {}
        }

        $vchDateStr = ""
        try { if ($root.Date) { $vchDateStr = ([string]$root.Date).Trim() } } catch {}

        $inputType = 1
        try { if ($root.InputType) { $inputType = [int]$root.InputType } } catch {}

        $narration = ""
        try { if ($root.VchOtherInfoDetails.Narration1) { $narration = ([string]$root.VchOtherInfoDetails.Narration1).Trim() } } catch {}

        $supplierBillNo = ""
        try { if ($root.VchOtherInfoDetails.PurchaseBillNo) { $supplierBillNo = ([string]$root.VchOtherInfoDetails.PurchaseBillNo).Trim() } } catch {}

        $stptName = ""
        try { if ($root.STPTName) { $stptName = ([string]$root.STPTName).Trim() } } catch {}

        $salesmanName = ""
        $salesmanCode = 0

        if (@(26, 12, 9, 3) -contains $VchType) {
            try {
                if ($root.BrokerName) {
                    $salesmanName = ([string]$root.BrokerName).Trim()
                }
            }
            catch {}

            try {
                if ($root.tmpBrokerCode) {
                    $salesmanCode = [int]$root.tmpBrokerCode
                }
            }
            catch { $salesmanCode = 0 }
        }

        $matCentre = ""
        $party = ""
        $bomName = ""
        $extraExpenseInBOM = 0.0

        if ($VchType -eq 6) {

            # ---------------------------------------------
            # PRODUCTION VOUCHER
            #
            # MasterName1 = Generated Material Centre
            # ConMCName   = Consumed Material Centre
            # MasterName2 = BOM Name
            # ---------------------------------------------

            try {
                if ($root.MasterName1) {
                    $matCentre = ([string]$root.MasterName1).Trim()
                }
            } catch {}

            try {
                if ($root.ConMCName) {
                    $party = ([string]$root.ConMCName).Trim()
                }
            } catch {}

            try {
                if ($root.MasterName2) {
                    $bomName = ([string]$root.MasterName2).Trim()
                }
            } catch {}

            try {
                if ($root.ExtraExpenseInBOM) {
                    $extraExpenseInBOM =
                        Get-DoubleSafe $root.ExtraExpenseInBOM 0.0
                }
            } catch {}

        }
        else {

            # Existing voucher behaviour
            try {
                if ($root.MasterName2) {
                    $matCentre = ([string]$root.MasterName2).Trim()
                }
            } catch {}

            try {
                if ($root.MasterName1) {
                    $party = ([string]$root.MasterName1).Trim()
                }
            } catch {}
        }

        $consumptionMap = @{}
        if ($VchType -eq 12 -or $VchType -eq 13) {
            try {
                $dirConn = Get-FastVoucherDirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
                if ($dirConn) {
                    $conn = $dirConn.connection
                    $conn.Open()
                    $qryCons = "SELECT M.Name AS ItemName, SUM(T2.Value1) AS CQ FROM Tran3 T1, Tran3 T2, Master1 M WHERE T1.RefCode = T2.RefCode AND T1.MasterCode1 = M.Code AND T1.Method=1 AND T1.VchCode=$vchCode AND T1.RecType IN (4,5) AND T2.Method=2 GROUP BY M.Name"
                    $ccCmd = $conn.CreateCommand()
                    $ccCmd.CommandText = $qryCons
                    $rCons = $ccCmd.ExecuteReader()
                    while ($rCons.Read()) {
                        $iName = $rCons["ItemName"].ToString().Trim()
                        $cVal = $rCons["CQ"]
                        if ($null -ne $cVal -and "$cVal" -ne "") { $consumptionMap[$iName] = [double]$cVal }
                    }
                    $rCons.Close()
                }
            } catch {} finally {
                if ($null -ne $conn) { try { $conn.Close() } catch {} }
            }
        }

        $challanItemRemainingMap = @{}
        if ($VchType -eq 11 -or $VchType -eq 4) {
            $targetRecType = 3
            try {
                $dirConn = Get-FastVoucherDirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
                if ($dirConn) {
                    $conn = $dirConn.connection
                    $conn.Open()
                    $cCmd = $conn.CreateCommand()
                    $cCmd.CommandText = "SELECT ItemSrNo, RefCode, Value1 FROM Tran3 WHERE Method=1 AND VchCode=$vchCode AND RecType=$targetRecType"
                    $cRdr = $cCmd.ExecuteReader()
                    while ($cRdr.Read()) {
                        $iSrNo   = [int]$cRdr["ItemSrNo"]
                        $refCode = [int]$cRdr["RefCode"]
                        $origQty = [Math]::Abs([double]$cRdr["Value1"])

                        $consumedQty = 0.0
                        $consCmd = $conn.CreateCommand()
                        $consCmd.CommandText = "SELECT SUM(ABS(Value1)) AS CQ FROM Tran3 WHERE Method=2 AND RefCode=$refCode AND RecType=$targetRecType"
                        $consVal = $consCmd.ExecuteScalar()
                        if ($null -ne $consVal -and "$consVal" -ne "") { $consumedQty = [Math]::Abs([double]$consVal) }

                        $challanItemRemainingMap[$iSrNo] = [Math]::Max(0, [Math]::Round($origQty - $consumedQty, 3))
                    }
                    $cRdr.Close()
                }
            } catch {} finally {
                if ($null -ne $conn) { try { $conn.Close() } catch {} }
            }
        }

        $items = @()
        $srNo = 1

        if ($VchType -eq 8 -or $VchType -eq 6) {
            # Stock Journal is special:
            #   ItemEntries          = Items Generated
            #   ConsumedItemEntries  = Items Consumed
            $generatedNodes = @()
            $consumedNodes = @()

            try { $generatedNodes = @($root.SelectNodes("ItemEntries/ItemDetail")) } catch {}
            if ($generatedNodes.Count -eq 0) {
                try { $generatedNodes = @($root.SelectNodes("//ItemEntries/ItemDetail")) } catch {}
            }

            # BUSY Stock Journal consumed/bottom grid may be exported as either:
            #   ConsumedItemEntries  OR  ItemEntries1
            # In your debug XML, BUSY returned ItemEntries1.
            try { $consumedNodes = @($root.SelectNodes("ConsumedItemEntries/ItemDetail")) } catch {}
            if ($consumedNodes.Count -eq 0) {
                try { $consumedNodes = @($root.SelectNodes("ItemEntries1/ItemDetail")) } catch {}
                if ($consumedNodes.Count -gt 0) { Write-SJDebug "Using ItemEntries1 as consumed Stock Journal rows" }
            }
            if ($consumedNodes.Count -eq 0) {
                try { $consumedNodes = @($root.SelectNodes("//ConsumedItemEntries/ItemDetail")) } catch {}
            }
            if ($consumedNodes.Count -eq 0) {
                try { $consumedNodes = @($root.SelectNodes("//ItemEntries1/ItemDetail")) } catch {}
                if ($consumedNodes.Count -gt 0) { Write-SJDebug "Using //ItemEntries1 fallback as consumed Stock Journal rows" }
            }

            Write-SJDebug "Final parsed generatedNodes=$($generatedNodes.Count), consumedNodes=$($consumedNodes.Count)"

            foreach ($d in $generatedNodes) {
                $items += Convert-VoucherXmlItemToHash -d $d -ForcedItemType 1 -OutputSrNo $srNo -consumptionMap $consumptionMap -challanItemRemainingMap $challanItemRemainingMap
                $srNo++
            }

            foreach ($d in $consumedNodes) {
                $items += Convert-VoucherXmlItemToHash -d $d -ForcedItemType 2 -OutputSrNo $srNo -consumptionMap $consumptionMap -challanItemRemainingMap $challanItemRemainingMap
                $srNo++
            }

            Write-SJDebug "Returned item count=$($items.Count), generated=$(@($items | Where-Object { [int]$_.itemType -eq 1 }).Count), consumed=$(@($items | Where-Object { [int]$_.itemType -eq 2 }).Count)"
        } else {
            try {
                if ($root.ItemEntries -and $root.ItemEntries.ItemDetail) {
                    foreach ($d in @($root.ItemEntries.ItemDetail)) {
                        $items += Convert-VoucherXmlItemToHash -d $d -ForcedItemType 0 -OutputSrNo $srNo -consumptionMap $consumptionMap -challanItemRemainingMap $challanItemRemainingMap
                        $srNo++
                    }
                }
            } catch {}
        }

        $billSundries = @()
        $bsr = 1
        try {
            foreach ($bs in $root.BillSundries.BSDetail) {
                $percentVal = 0
                try { $percentVal = [double]$bs.PercentVal } catch {}
                $billSundries += @{
                    srNo       = $bsr
                    name       = if ($bs.BSName) { ([string]$bs.BSName).Trim() } else { "" }
                    amount     = [Math]::Abs([double]$bs.Amt)
                    percentVal = $percentVal
                }
                $bsr++
            }
        } catch {}

        $settlements = $null
        try {
            if ($root.POSVchData) {
                $sd = $root.POSVchData
                $cashAmt = 0; $cardAmt = 0; $giftAmt = 0; $tendered = 0
                try { $cashAmt = [double]$sd.CashAmt } catch {}
                try { $cardAmt = [double]$sd.CreditCardAmt1 } catch {}
                try { $giftAmt = [double]$sd.CreditCardAmt2 } catch {}
                try { $tendered = [double]$sd.CashRecvdAmt } catch {}

                if ($cashAmt -gt 0 -or $cardAmt -gt 0 -or $giftAmt -gt 0) {
                    $settlements = @{
                        cash = @{ amount = $cashAmt; account = if ($sd.CashAccName) { ([string]$sd.CashAccName).Trim() } else { "Cash" } }
                        card = @{ amount = $cardAmt; account = if ($sd.CCAccName1) { ([string]$sd.CCAccName1).Trim() } else { "" } }
                        gift = @{ amount = $giftAmt; account = if ($sd.CCAccName2) { ([string]$sd.CCAccName2).Trim() } else { "" } }
                        cashTendered = $tendered
                    }
                }
            }
        } catch {}

        $refEntries = @()
        try {
            if ($root.PendingBillDetails.BillDetail) {
                foreach ($bd in $root.PendingBillDetails.BillDetail) {
                    $refSrNo = 1
                    foreach ($ref in $bd.BillRefs) {
                        $amt = [double]$ref.Value1
                        if ($amt -lt 0) { $amt = -$amt }
                        $refType = 1
                        try { $refType = [int]$ref.Method } catch {}

                        $refEntries += @{
                            srNo    = $refSrNo
                            refType = $refType
                            refNo   = if ($ref.RefNo) { ([string]$ref.RefNo).Trim() } else { "" }
                            amount  = $amt
                            dueDate = if ($ref.dueDate) { ([string]$ref.DueDate).Trim() } else { "" }
                        }
                        $refSrNo++
                    }
                }
            }
        } catch {}

        $linkedChallans = @()
        try {
            $dirConn = Get-FastVoucherDirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
            if ($dirConn) {
                $conn = $dirConn.connection
                $conn.Open()
                $lcCmd = $conn.CreateCommand()
                $lcCmd.CommandText = "SELECT DISTINCT [No], [Date] FROM [Tran3] WHERE [VchCode] = $vchCode AND [Method] = 2 AND [RecType] = 3"
                $lcRdr = $lcCmd.ExecuteReader()
                while ($lcRdr.Read()) {
                    $cNo = $lcRdr["No"].ToString().Trim()
                    $dVal = $lcRdr["Date"]
                    $cDate = ""
                    if ($dVal -is [datetime]) { $cDate = $dVal.ToString("dd-MM-yyyy") }
                    else { $cDate = [datetime]::Parse($dVal.ToString()).ToString("dd-MM-yyyy") }
                    $linkedChallans += @{ vchNo = $cNo; date = $cDate }
                }
                $lcRdr.Close()
            }
        } catch {} finally {
            if ($null -ne $conn) { try { $conn.Close() } catch {} }
        }

        $optionalFields = @{}
        try {
            if ($root.VchOtherInfoDetails.OFInfo) {
                $ofNode = $root.VchOtherInfoDetails.OFInfo
                for ($i = 1; $i -le 20; $i++) {
                    $nodeName = "OF$i"
                    if ($ofNode.$nodeName) {
                        $val = ([string]$ofNode.$nodeName).Trim()
                        if ($val -ne "") {
                            if ($val -match "^(\d{2})-(\d{2})-(\d{4})$") {
                                $val = "$($Matches[3])-$($Matches[2])-$($Matches[1])"
                            }
                            $optionalFields["OptionField$i"] = $val
                        }
                    }
                }
            }
        } catch {}

        return @{
            success = $true
            data = @{
                vchType          = $VchType
                inputType        = $inputType
                vchNo            = $VchNo
                vchSeries        = $VchSeries
                date             = $vchDateStr

                # party is reused by the shared frontend as Consumed MC.
                party            = $party

                # matCentre is reused by the shared frontend as Generated MC.
                matCentre        = $matCentre

                # Production-specific fields.
                bomName          = $bomName
                extraExpenseInBOM = $extraExpenseInBOM

                saleType         = $stptName
                purchaseType     = $stptName
                salesmanName     = $salesmanName
                salesmanCode     = $salesmanCode
                narration        = $narration
                supplierBillNo   = $supplierBillNo
                items            = @($items)
                billSundries     = @($billSundries)
                settlements      = $settlements
                refEntries       = @($refEntries)
                linkedChallans   = @($linkedChallans)
                optionalFields   = $optionalFields
            }
        }
    } catch {
        Write-SJDebug "ERROR: $($_.Exception.Message)"
        return @{ success = $false; error = $_.Exception.Message }
    } finally {
        Disconnect-BUSY $fi
    }
}

function Get-AccountVoucherDetail {
    param(
        [int]$VchType,[string]$VchNo,[string]$VchSeries,
        [string]$VchDate,
        [string]$InstanceId = "",[string]$CompanyCode = "")

    # Resolve if SQL or Access from instances.json
    $targetInst = Get-InstanceConfig -InstanceId $InstanceId
    $dbType = 0
    if ($null -ne $targetInst -and $null -ne $targetInst.dbType) { $dbType = [int]$targetInst.dbType }
    $isSql = ($dbType -eq 1)

    $vchCode = 0
    
    if ($isSql) {
        # -- SQL PATH: Resolve VchCode natively using direct SQL connection --
        $dirConn = Get-FastVoucherDirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
        if ($null -eq $dirConn) { return @{ success = $false; error = "Could not build direct database connection" } }
        $conn = $dirConn.connection
        try {
            $conn.Open()
            $vchCode = Get-VchCode-Direct -conn $conn -VchType $VchType -VchNo $VchNo -VchDate $VchDate -isSql $true
        } finally {
            if ($null -ne $conn) { try { $conn.Close() } catch {} }
        }
    } else {
        # -- ACCESS PATH: Original stable COM recordset query --
        $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
        if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
        try {
            $dateFilter = ""
            if ($VchDate -match "\d{2}-\d{2}-\d{4}") {
                $d = [datetime]::ParseExact($VchDate, "dd-MM-yyyy", $null)
                $dateFilter = " AND [Date]=#" + $d.ToString("MM/dd/yyyy") + "#"
            } elseif ($VchDate -match "\d{4}-\d{2}-\d{2}") {
                $d = [datetime]::ParseExact($VchDate, "yyyy-MM-dd", $null)
                $dateFilter = " AND [Date]=#" + $d.ToString("MM/dd/yyyy") + "#"
            }

            $targetNo = $VchNo.Trim().ToLower()
            $qry = "SELECT VchCode, VchNo, AutoVchNo FROM Tran1 WHERE VchType=$VchType" + $dateFilter
            $rst = $fi.GetRecordset($qry)

            if ($rst -and -not $rst.EOF) {
                $rst.MoveFirst()
                while (-not $rst.EOF) {
                    $dbVchNo = ""
                    try {
                        $v = $rst.Fields.Item("VchNo").Value
                        if ($v -ne [System.DBNull]::Value) { $dbVchNo = $v.ToString().Trim().ToLower() }
                        if ($dbVchNo -eq "") {
                            $av = $rst.Fields.Item("AutoVchNo").Value
                            if ($av -ne [System.DBNull]::Value) { $dbVchNo = $av.ToString().Trim().ToLower() }
                        }
                    } catch {}
                    
                    if ($dbVchNo -eq $targetNo) {
                        $vchCode = [int]$rst.Fields.Item("VchCode").Value
                        break
                    }
                    $rst.MoveNext()
                }
                try { $rst.Close() } catch {}
            }

            if ($vchCode -eq 0) {
                $qryAll = "SELECT VchCode, VchNo, AutoVchNo FROM Tran1 WHERE VchType=$VchType"
                $rstAll = $fi.GetRecordset($qryAll)
                if ($rstAll -and -not $rstAll.EOF) {
                    $rstAll.MoveFirst()
                    while (-not $rstAll.EOF) {
                        $dbVchNo = ""
                        try {
                            $v = $rstAll.Fields.Item("VchNo").Value
                            if ($v -ne [System.DBNull]::Value) { $dbVchNo = $v.ToString().Trim().ToLower() }
                            if ($dbVchNo -eq "") {
                                $av = $rstAll.Fields.Item("AutoVchNo").Value
                                if ($av -ne [System.DBNull]::Value) { $dbVchNo = $av.ToString().Trim().ToLower() }
                            }
                        } catch {}

                        if ($dbVchNo -eq $targetNo) {
                            $vchCode = [int]$rstAll.Fields.Item("VchCode").Value
                            break
                        }
                        $rstAll.MoveNext()
                    }
                    try { $rstAll.Close() } catch {}
                }
            }
        } finally {
            Disconnect-BUSY $fi
        }
    }

    if ($vchCode -eq 0) {
        return @{ success = $false; error = "Could not find Voucher in database (VchType=$VchType, No=$VchNo)" }
    }

    # 2. Invoke COM only to fetch the parsed XML
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }

    try {
        $errMsg = ""
        $xmlStr = ""
        try {
            $xmlStr = $fi.GetVchXML($vchCode, [ref]$errMsg)
        } catch {
            $xmlStr = $fi.GetVchXML($vchCode)
        }

        if (-not $xmlStr) {
            return @{ success = $false; error = if ($errMsg) { $errMsg } else { "Voucher XML is empty" } }
        }

        $xml = [xml]$xmlStr
        $root = $xml.DocumentElement

        $vchDateStr = ""
        try { if ($root.Date) { $vchDateStr = ([string]$root.Date).Trim() } } catch {}

        $narration = ""
        try { if ($root.VchOtherInfoDetails.Narration1) { $narration = ([string]$root.VchOtherInfoDetails.Narration1).Trim() } } catch {}

        $bbaMap = @{}
        try {
            if ($root.PendingBillDetails.BillDetail) {
                foreach ($bd in $root.PendingBillDetails.BillDetail) {
                    $pName = ([string]$bd.MasterName1).Trim()
                    $refs = @()
                    $sr = 1
                    foreach ($ref in $bd.BillRefs) {
                        $amt = [double]$ref.Value1
                        if ($amt -lt 0) { $amt = -$amt }
                        
                        $refType = 1
                        try { $refType = [int]$ref.Method } catch {}
                        
                        $refs += @{
                            srNo    = $sr
                            refType = $refType
                            refNo   = if ($ref.RefNo) { ([string]$ref.RefNo).Trim() } else { "" }
                            amount  = $amt
                            dueDate = if ($ref.DueDate) { ([string]$ref.DueDate).Trim() } else { "" }
                        }
                        $sr++
                    }
                    $bbaMap[$pName] = $refs
                }
            }
        } catch {}

        $accounts = @()
        $srNo = 1
        try {
            if ($root.AccEntries.AccDetail) {
                foreach ($acc in $root.AccEntries.AccDetail) {
                    $accName = if ($acc.AccountName) { ([string]$acc.AccountName).Trim() } else { "" }
                    $amtType = [int]$acc.AmountType
                    $dc = if ($amtType -eq 1) { "D" } else { "C" }
                    $amt = [double]$acc.AmtMainCur
                    $shortNar = if ($acc.ShortNar) { ([string]$acc.ShortNar).Trim() } else { "" }
                    
                    $bbaEntries = @()
                    
                    if ($acc.BillRefs.BillDetails) {
                        $bsr = 1
                        foreach ($bref in $acc.BillRefs.BillDetails) {
                            $bAmt = [double]$bref.Value1
                            if ($bAmt -lt 0) { $bAmt = -$bAmt }
                            
                            $bRefType = 1
                            try { $bRefType = [int]$bref.Method } catch {}
                            
                            $bbaEntries += @{
                                srNo    = $bsr
                                refType = $bRefType
                                refNo   = if ($bref.RefNo) { ([string]$bref.RefNo).Trim() } else { "" }
                                amount  = $bAmt
                                dueDate = if ($bref.DueDate) { ([string]$bref.DueDate).Trim() } else { "" }
                            }
                            $bsr++
                        }
                    } elseif ($bbaMap.ContainsKey($accName)) {
                        $bbaEntries = $bbaMap[$accName]
                        $bbaMap.Remove($accName) 
                    }

                    $accounts += @{
                        srNo           = $srNo
                        dc             = $dc
                        accountName    = $accName
                        amount         = $amt
                        shortNarration = $shortNar
                        bbaEntries     = @($bbaEntries)
                    }
                    $srNo++
                }
            }
        } catch {}

        return @{
            success = $true
            data = @{
                vchType   = $VchType
                vchNo     = $VchNo
                vchSeries = $VchSeries
                date      = $vchDateStr
                narration = $narration
                accounts  = @($accounts)
            }
        }
    } catch {
        return @{ success = $false; error = $_.Exception.Message }
    } finally {
        Disconnect-BUSY $fi
    }
}

# ===============================================================
#  SEARCH ORIGINAL VOUCHERS (Specifically for BBA)
# ===============================================================
# ===============================================================
# SEARCH ORIGINAL/PENDING VOUCHERS FOR RECEIPT AND PAYMENT BBA
#
# Supported parent voucher types:
#   9  = Sale
#   10 = Purchase Return
#   2  = Purchase
#   3  = Sale Return
#
# Receipt searches:
#   Sale + Purchase Return
#
# Payment searches:
#   Purchase + Sale Return
# ===============================================================

