# 09-item-cost.ps1
# Extracted from vouchers.ps1. Keep functions behavior-compatible with the monolith.

function Get-ItemMaterialCentreAverageCost {
    param(
        [int]$ItemCode,
        [string]$MaterialCentre,
        [string]$AsOf = "",
        [int]$ExcludeVchCode = 0,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    # ============================================================
    # MATERIAL-CENTRE-SPECIFIC CURRENT AVERAGE COST
    #
    # Current Average Cost = Remaining Layer Value / Remaining Qty
    #
    # Layer rules used here:
    #   * Opening stock creates opening layers from Tran4 D1 / D3.
    #   * Positive inventory transactions create inward layers.
    #   * Normal stock-out consumes layers FIFO.
    #   * Parent-linked Sale Return restores the exact allocations
    #     consumed by the parent Sale (reverse allocation order).
    #   * Parent-linked Purchase Return reduces only the parent
    #     Purchase layer(s).
    #   * Unlinked Purchase Return uses FIFO.
    #   * Unlinked Sale Return restores prior Sale allocations FIFO;
    #     if historical allocation cannot be resolved, a clearly
    #     flagged transaction-rate fallback layer is used.
    #
    # IMPORTANT: FIFO determines which layers remain.  The value
    # returned to the UI is NOT the next FIFO issue rate.  It is the
    # weighted average of all remaining layers in the selected MC.
    # ============================================================

    $eps = 0.000001

    function Convert-CostDouble {
        param($Value)
        if ($null -eq $Value -or $Value -eq [System.DBNull]::Value) { return 0.0 }
        try {
            return [Convert]::ToDouble(
                $Value.ToString(),
                [System.Globalization.CultureInfo]::InvariantCulture
            )
        }
        catch {
            try { return [double]$Value } catch { return 0.0 }
        }
    }

    function Convert-CostInt {
        param($Value)
        if ($null -eq $Value -or $Value -eq [System.DBNull]::Value) { return 0 }
        try { return [int]$Value } catch { return 0 }
    }

    function Close-CostRecordset {
        param($Rst)
        if ($null -ne $Rst) {
            try { $Rst.Close() } catch {}
            try {
                if ([System.Runtime.InteropServices.Marshal]::IsComObject($Rst)) {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Rst) | Out-Null
                }
            }
            catch {}
        }
    }

    if ($ItemCode -le 0) {
        return @{ success = $false; error = "Valid item code is required" }
    }

    if ([string]::IsNullOrWhiteSpace($MaterialCentre)) {
        return @{ success = $false; error = "Material centre is required" }
    }

    # Resolve date using the bridge helper so yyyy-MM-dd and dd-MM-yyyy
    # continue to behave like the rest of the voucher module.
    $costDate = $null
    if (-not [string]::IsNullOrWhiteSpace($AsOf)) {
        try { $costDate = Safe-ParseDate $AsOf } catch { $costDate = $null }
        if ($null -eq $costDate) {
            return @{
                success = $false
                error   = "Invalid asOf date. Use yyyy-MM-dd or dd-MM-yyyy."
            }
        }
    }
    else {
        $costDate = Get-Date
    }

    $targetInst = Get-InstanceConfig -InstanceId $InstanceId
    $dbType = 0
    if ($null -ne $targetInst -and $null -ne $targetInst.dbType) {
        try { $dbType = [int]$targetInst.dbType } catch { $dbType = 0 }
    }
    $isSql = ($dbType -eq 1)

    $dateLiteral = if ($isSql) {
        "'" + $costDate.ToString("yyyy-MM-dd") + "'"
    }
    else {
        "#" + $costDate.ToString("MM/dd/yyyy") + "#"
    }

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) {
        return @{ success = $false; error = "BUSY connection failed" }
    }

    try {
        $warnings = [System.Collections.ArrayList]::new()
        $layersByMc = @{}
        $ledgerQtyByMc = @{}
        $mcNames = @{}
        $saleAllocationsByVch = @{}
        $saleConsumptionHistory = [System.Collections.ArrayList]::new()
        $returnParentMap = @{}
        $costState = [pscustomobject]@{ layerSequence = 0 }

        function Ensure-LayerList {
            param([int]$McCode)

            # Do not return the ArrayList from a PowerShell function.
            # PowerShell pipeline enumeration can turn an empty list into $null
            # or reshape a non-empty list.  We only ensure the hashtable entry
            # exists, then callers access the ArrayList directly from the table.
            $key = [string]$McCode
            if (-not $layersByMc.ContainsKey($key)) {
                $layersByMc[$key] = [System.Collections.ArrayList]::new()
            }
        }

        function Add-LedgerQty {
            param([int]$McCode, [double]$Qty)
            $key = [string]$McCode
            if (-not $ledgerQtyByMc.ContainsKey($key)) { $ledgerQtyByMc[$key] = 0.0 }
            $ledgerQtyByMc[$key] = [double]$ledgerQtyByMc[$key] + $Qty
        }

        function Add-StockLayer {
            param(
                [int]$McCode,
                [double]$Qty,
                [double]$Value,
                [int]$SourceVchCode,
                [int]$SourceVchType,
                [string]$SourceType,
                [string]$SourceDate,
                [int]$SourceSrNo
            )

            if ($Qty -le $eps) { return $null }

            $safeValue = [Math]::Abs($Value)
            $unitCost = if ($Qty -gt $eps) { $safeValue / $Qty } else { 0.0 }

            $costState.layerSequence = [int]$costState.layerSequence + 1
            $newLayerId = "L$($costState.layerSequence)"

            $layer = [pscustomobject]@{
                layerId        = $newLayerId
                mcCode         = $McCode
                sourceVchCode  = $SourceVchCode
                sourceVchType  = $SourceVchType
                sourceType     = $SourceType
                sourceDate     = $SourceDate
                sourceSrNo     = $SourceSrNo
                originalQty    = [double]$Qty
                remainingQty   = [double]$Qty
                unitCost       = [double]$unitCost
                originalValue  = [double]$safeValue
                remainingValue = [double]$safeValue
            }

            Ensure-LayerList -McCode $McCode
            $list = $layersByMc[[string]$McCode]
            [void]$list.Add($layer)
            return $layer
        }

        function Get-McLayerTotals {
            param([int]$McCode)

            $qty = 0.0
            $value = 0.0
            Ensure-LayerList -McCode $McCode
            $list = $layersByMc[[string]$McCode]

            foreach ($layer in $list) {
                $rq = [double]$layer.remainingQty
                if ($rq -gt $eps) {
                    $qty += $rq
                    $value += [double]$layer.remainingValue
                }
            }

            return [pscustomobject]@{
                qty   = $qty
                value = $value
                rate  = if ($qty -gt $eps) { $value / $qty } else { 0.0 }
            }
        }

        function Remove-FifoLayers {
            param([int]$McCode, [double]$Qty)

            $remaining = [double]$Qty
            $allocations = [System.Collections.ArrayList]::new()
            $removedValue = 0.0
            $removedQty = 0.0

            Ensure-LayerList -McCode $McCode
            $list = $layersByMc[[string]$McCode]

            foreach ($layer in $list) {
                if ($remaining -le $eps) { break }

                $available = [double]$layer.remainingQty
                if ($available -le $eps) { continue }

                $take = [Math]::Min($available, $remaining)
                $takeValue = $take * [double]$layer.unitCost

                $layer.remainingQty = [Math]::Max(0.0, $available - $take)
                $layer.remainingValue = [Math]::Max(
                    0.0,
                    [double]$layer.remainingValue - $takeValue
                )

                $allocation = [pscustomobject]@{
                    layerRef     = $layer
                    layerId      = [string]$layer.layerId
                    mcCode       = $McCode
                    qty          = [double]$take
                    value        = [double]$takeValue
                    unitCost     = [double]$layer.unitCost
                    restoredQty  = 0.0
                }

                [void]$allocations.Add($allocation)
                $removedQty += $take
                $removedValue += $takeValue
                $remaining -= $take
            }

            return [pscustomobject]@{
                requestedQty = [double]$Qty
                removedQty   = [double]$removedQty
                removedValue = [double]$removedValue
                shortageQty  = [Math]::Max(0.0, $remaining)
                allocations  = $allocations
            }
        }

        function Remove-ParentPurchaseLayers {
            param([int]$McCode, [int]$ParentVchCode, [double]$Qty)

            $remaining = [double]$Qty
            $removedQty = 0.0
            $removedValue = 0.0
            $allocations = [System.Collections.ArrayList]::new()
            Ensure-LayerList -McCode $McCode
            $list = $layersByMc[[string]$McCode]

            foreach ($layer in $list) {
                if ($remaining -le $eps) { break }
                if ([int]$layer.sourceVchCode -ne $ParentVchCode) { continue }
                if ([int]$layer.sourceVchType -ne 2) { continue }

                $available = [double]$layer.remainingQty
                if ($available -le $eps) { continue }

                $take = [Math]::Min($available, $remaining)
                $takeValue = $take * [double]$layer.unitCost

                $layer.remainingQty = [Math]::Max(0.0, $available - $take)
                $layer.remainingValue = [Math]::Max(
                    0.0,
                    [double]$layer.remainingValue - $takeValue
                )

                [void]$allocations.Add([pscustomobject]@{
                    layerRef    = $layer
                    layerId     = [string]$layer.layerId
                    mcCode      = $McCode
                    qty         = [double]$take
                    value       = [double]$takeValue
                    unitCost    = [double]$layer.unitCost
                    restoredQty = 0.0
                })

                $removedQty += $take
                $removedValue += $takeValue
                $remaining -= $take
            }

            return [pscustomobject]@{
                requestedQty = [double]$Qty
                removedQty   = [double]$removedQty
                removedValue = [double]$removedValue
                shortageQty  = [Math]::Max(0.0, $remaining)
                allocations  = $allocations
            }
        }

        function Restore-SaleAllocations {
            param(
                $Allocations,
                [int]$McCode,
                [double]$Qty,
                [bool]$ReverseOrder
            )

            $remaining = [double]$Qty
            $restoredQty = 0.0
            $restoredValue = 0.0
            $ordered = @($Allocations)

            if ($ReverseOrder -and $ordered.Count -gt 1) {
                [array]::Reverse($ordered)
            }

            foreach ($allocation in $ordered) {
                if ($remaining -le $eps) { break }
                if ([int]$allocation.mcCode -ne $McCode) { continue }

                $availableToRestore =
                    [double]$allocation.qty - [double]$allocation.restoredQty

                if ($availableToRestore -le $eps) { continue }

                $restoreQty = [Math]::Min($availableToRestore, $remaining)
                $layer = $allocation.layerRef

                if ($null -eq $layer) { continue }

                $restoreValue = $restoreQty * [double]$allocation.unitCost

                $layer.remainingQty = [double]$layer.remainingQty + $restoreQty
                $layer.remainingValue = [double]$layer.remainingValue + $restoreValue
                $allocation.restoredQty = [double]$allocation.restoredQty + $restoreQty

                $restoredQty += $restoreQty
                $restoredValue += $restoreValue
                $remaining -= $restoreQty
            }

            return [pscustomobject]@{
                requestedQty  = [double]$Qty
                restoredQty   = [double]$restoredQty
                restoredValue = [double]$restoredValue
                shortageQty   = [Math]::Max(0.0, $remaining)
            }
        }

        function Get-RowInwardValue {
            param($Row, [double]$Qty, [double]$VoucherOutValue, [double]$VoucherInboundQty)

            $value3 = [Math]::Abs([double]$Row.value3)
            if ($value3 -gt $eps) { return $value3 }

            $d5 = [Math]::Abs([double]$Row.d5)
            if ($d5 -gt $eps) { return $d5 }

            $d2 = [Math]::Abs([double]$Row.d2)
            if ($d2 -gt $eps) { return $Qty * $d2 }

            # Stock Journal / Production transfers can occasionally carry a
            # zero inward amount while the same voucher contains the source
            # outflow.  In that case preserve the source consumed value.
            if ($VoucherOutValue -gt $eps -and $VoucherInboundQty -gt $eps) {
                return $Qty * ($VoucherOutValue / $VoucherInboundQty)
            }

            $current = Get-McLayerTotals -McCode ([int]$Row.mcCode)
            if ([double]$current.rate -gt $eps) {
                return $Qty * [double]$current.rate
            }

            return 0.0
        }

        function Get-SourceTypeName {
            param([int]$VchType, [bool]$IsInward)
            switch ($VchType) {
                2  { return "purchase" }
                3  { return "sale-return" }
                4  { return "material-receipt" }
                6  { if ($IsInward) { return "production-generated" } else { return "production-consumed" } }
                8  { if ($IsInward) { return "stock-journal-generated" } else { return "stock-journal-consumed" } }
                9  { return "sale" }
                10 { return "purchase-return" }
                11 { return "material-issue" }
                default { if ($IsInward) { return "stock-in" } else { return "stock-out" } }
            }
        }

        # ------------------------------------------------------------
        # Resolve item and material-centre masters.
        # ------------------------------------------------------------
        $itemName = ""
        $itemRst = $fi.GetRecordset(
            "SELECT Code, Name FROM Master1 WHERE MasterType=6 AND Code=$ItemCode"
        )
        if ($itemRst -and -not $itemRst.EOF) {
            try { $itemName = [string]$itemRst.Fields.Item("Name").Value } catch {}
        }
        Close-CostRecordset $itemRst

        if ([string]::IsNullOrWhiteSpace($itemName)) {
            return @{ success = $false; error = "Item code $ItemCode was not found" }
        }

        $selectedMcCode = -1
        $mcRst = $fi.GetRecordset("SELECT Code, Name FROM Master1 WHERE MasterType=11")
        if ($mcRst -and -not $mcRst.EOF) {
            $mcRst.MoveFirst()
            while (-not $mcRst.EOF) {
                $code = Convert-CostInt $mcRst.Fields.Item("Code").Value
                $name = ""
                try { $name = [string]$mcRst.Fields.Item("Name").Value } catch {}
                $mcNames[[string]$code] = $name

                if (
                    $name.Trim().ToLowerInvariant() -eq
                    $MaterialCentre.Trim().ToLowerInvariant()
                ) {
                    $selectedMcCode = $code
                }
                $mcRst.MoveNext()
            }
        }
        Close-CostRecordset $mcRst

        if ($selectedMcCode -lt 0) {
            return @{
                success = $false
                error   = "Material centre '$MaterialCentre' was not found"
            }
        }

        # ------------------------------------------------------------
        # Read BUSY's configured stock valuation method from the item
        # master. The observed BUSY data for StockValMethod=5 behaves
        # as weighted-average valuation, so that method is handled by
        # a BUSY-parity weighted-average strategy below.
        # Other methods continue through the existing layer-replay
        # strategy until their BUSY semantics are mapped explicitly.
        # ------------------------------------------------------------
        $stockValMethod = 0
        $masterXmlText = ""

        try {
            $masterXmlText = [string]$fi.GetMasterXML($ItemCode)

            if (-not [string]::IsNullOrWhiteSpace($masterXmlText)) {
                [xml]$masterXmlDoc = $masterXmlText
                $stockValNode = $masterXmlDoc.SelectSingleNode("//StockValMethod")

                if ($null -ne $stockValNode) {
                    $stockValMethod = Convert-CostInt $stockValNode.InnerText
                }
            }
        }
        catch {
            [void]$warnings.Add(
                "Could not read StockValMethod from item master XML: $($_.Exception.Message)"
            )
            $stockValMethod = 0
        }

        # ============================================================
        # BUSY STOCK VALUATION METHOD 5
        # Observed behaviour: weighted-average valuation.
        #
        # Important BUSY-parity rule:
        #   Normal outward quantity does NOT remove an old FIFO layer
        #   from the valuation pool. The unit valuation remains the
        #   weighted average of valuation-bearing inward stock.
        #
        # Example proven against BUSY:
        #   Opening 79 @ 275 = 21,725
        #   Purchase 10 @ 100 = 1,000
        #   Weighted average = 22,725 / 89 = 255.337...
        #   A normal Sale changes physical stock quantity, but the
        #   weighted-average unit cost remains 255.337...
        #
        # The response still returns remainingValue as:
        #   physical remaining qty * BUSY weighted-average rate.
        # ============================================================
        if ($stockValMethod -eq 5) {

            $waStates = @{}
            $waPurchaseSourceByVch = @{}
            $waSaleRateByVch = @{}
            $waReturnParentMap = @{}

            function Ensure-WaState {
                param([int]$McCode)

                $key = [string]$McCode

                if (-not $waStates.ContainsKey($key)) {
                    $waStates[$key] = [pscustomobject]@{
                        mcCode     = $McCode
                        ledgerQty  = 0.0
                        poolQty    = 0.0
                        poolValue  = 0.0
                        lastRate   = 0.0
                    }
                }
            }

            function Get-WaState {
                param([int]$McCode)

                Ensure-WaState -McCode $McCode
                return $waStates[[string]$McCode]
            }

            function Get-WaRate {
                param([int]$McCode)

                $state = Get-WaState -McCode $McCode

                if ([Math]::Abs([double]$state.poolQty) -gt $eps) {
                    return [Math]::Abs(
                        [double]$state.poolValue /
                        [double]$state.poolQty
                    )
                }

                return [Math]::Abs([double]$state.lastRate)
            }

            function Add-WaPool {
                param(
                    [int]$McCode,
                    [double]$Qty,
                    [double]$Value
                )

                if ($Qty -le $eps) { return }

                $state = Get-WaState -McCode $McCode
                $incomingQty = [double]$Qty
                $incomingValue = [Math]::Abs([double]$Value)
                $incomingRate = if ($incomingQty -gt $eps) {
                    $incomingValue / $incomingQty
                }
                else {
                    0.0
                }

                # BUSY negative-stock rule: inward quantity first clears an
                # existing negative balance at the CURRENT negative average.
                # This preserves the negative rate until stock reaches zero.
                if ([double]$state.poolQty -lt -$eps) {
                    $negativeQty = [Math]::Abs([double]$state.poolQty)
                    $negativeRate = Get-WaRate -McCode $McCode
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

                # Only quantity remaining after the negative balance has been
                # cleared is valued at the inward row's own rate.
                if ($incomingQty -gt $eps) {
                    $state.poolQty = [double]$state.poolQty + $incomingQty
                    $state.poolValue = [double]$state.poolValue + ($incomingQty * $incomingRate)
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

            function Remove-WaPool {
                param(
                    [int]$McCode,
                    [double]$Qty,
                    [double]$UnitRate
                )

                if ($Qty -le $eps) { return }

                $state = Get-WaState -McCode $McCode
                $outQty = [double]$Qty
                $currentQty = [double]$state.poolQty
                $currentRate = Get-WaRate -McCode $McCode
                $postedRate = [Math]::Max(0.0, [double]$UnitRate)

                if ($currentQty -gt $eps) {
                    # Positive before issue: BUSY uses current WA for the
                    # whole row, including a row that crosses below zero.
                    $useRate = $currentRate
                }
                else {
                    # Already zero/negative: use this row's posted stock rate.
                    $useRate = if ($postedRate -gt $eps) {
                        $postedRate
                    }
                    else {
                        $currentRate
                    }
                }

                $state.poolQty = [double]$state.poolQty - $outQty
                $state.poolValue = [double]$state.poolValue - ($outQty * $useRate)

                if ([Math]::Abs([double]$state.poolQty) -le $eps) {
                    $state.poolQty = 0.0
                    $state.poolValue = 0.0
                    if ($useRate -gt $eps) {
                        $state.lastRate = $useRate
                    }
                }
                else {
                    $state.lastRate = [Math]::Abs(
                        [double]$state.poolValue /
                        [double]$state.poolQty
                    )
                }
            }

            function Get-WaNegativeOutRate {
                param(
                    $Row,
                    [double]$Qty
                )

                if ($Qty -le $eps) { return 0.0 }

                # For stock that is already zero/negative, BUSY's signed
                # valuation follows Value3 first.  GF002 proves this because
                # its Sale row has Value3/Qty=1039.50 while D2/D5 are 1050.
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

            function Get-WaExplicitValue {
                param(
                    $Row,
                    [double]$Qty
                )

                $value3 = [Math]::Abs([double]$Row.value3)
                if ($value3 -gt $eps) {
                    return $value3
                }

                $d5 = [Math]::Abs([double]$Row.d5)
                if ($d5 -gt $eps) {
                    return $d5
                }

                $d2 = [Math]::Abs([double]$Row.d2)
                if ($d2 -gt $eps) {
                    return $Qty * $d2
                }

                return 0.0
            }

            # --------------------------------------------------------
            # Parent return mapping.
            # --------------------------------------------------------
            try {
                $waParentQry = @"
SELECT DISTINCT
    P.VchCode AS ParentVchCode,
    R.VchCode AS ReturnVchCode
FROM Tran3 P, Tran3 R, Tran1 H
WHERE
    P.RefCode = R.RefCode
    AND P.Method = 1
    AND R.Method = 2
    AND R.VchCode = H.VchCode
    AND (H.VchType = 3 OR H.VchType = 10)
"@

                $waParentRst = $fi.GetRecordset($waParentQry)

                if ($waParentRst -and -not $waParentRst.EOF) {
                    $waParentRst.MoveFirst()

                    while (-not $waParentRst.EOF) {
                        $parentCode =
                            Convert-CostInt $waParentRst.Fields.Item("ParentVchCode").Value

                        $returnCode =
                            Convert-CostInt $waParentRst.Fields.Item("ReturnVchCode").Value

                        if ($parentCode -gt 0 -and $returnCode -gt 0) {
                            $waReturnParentMap[[string]$returnCode] = $parentCode
                        }

                        $waParentRst.MoveNext()
                    }
                }

                Close-CostRecordset $waParentRst
            }
            catch {
                [void]$warnings.Add(
                    "Weighted-average valuation could not read parent-return links: $($_.Exception.Message)"
                )
            }

            # --------------------------------------------------------
            # Opening stock.
            # Tran4.D1 = opening quantity
            # Tran4.D3 = opening value
            # --------------------------------------------------------
            $waOpeningRst = $fi.GetRecordset(
                "SELECT MasterCode2, D1, D3 FROM Tran4 WHERE RecType=0 AND MasterCode1=$ItemCode"
            )

            if ($waOpeningRst -and -not $waOpeningRst.EOF) {
                $waOpeningRst.MoveFirst()

                while (-not $waOpeningRst.EOF) {
                    $mcCode =
                        Convert-CostInt $waOpeningRst.Fields.Item("MasterCode2").Value

                    $qty =
                        Convert-CostDouble $waOpeningRst.Fields.Item("D1").Value

                    $value =
                        Convert-CostDouble $waOpeningRst.Fields.Item("D3").Value

                    $state = Get-WaState -McCode $mcCode
                    $state.ledgerQty = [double]$state.ledgerQty + $qty

                    if ([Math]::Abs($qty) -gt $eps) {
                        # Preserve BUSY opening stock as a signed valuation
                        # balance, including negative opening quantity/value.
                        $state.poolQty = [double]$state.poolQty + $qty
                        $state.poolValue = [double]$state.poolValue + $value

                        if ([Math]::Abs([double]$state.poolQty) -gt $eps) {
                            $state.lastRate = [Math]::Abs(
                                [double]$state.poolValue /
                                [double]$state.poolQty
                            )
                        }
                    }

                    $waOpeningRst.MoveNext()
                }
            }

            Close-CostRecordset $waOpeningRst

            # --------------------------------------------------------
            # Load all inventory movements for this item.
            # --------------------------------------------------------
            $waExcludeSql = ""

            if ($ExcludeVchCode -gt 0) {
                $waExcludeSql = " AND T.VchCode <> $ExcludeVchCode"
            }

            $waTxnQry = @"
SELECT
    T.VchCode,
    T.MasterCode2 AS MCCode,
    T.SrNo,
    T.Value1,
    T.Value3,
    T.D2,
    T.D5,
    V.VchType,
    V.[Date] AS VoucherDate,
    V.VchNo
FROM Tran2 T
LEFT JOIN Tran1 V ON V.VchCode = T.VchCode
WHERE
    T.RecType = 2
    AND T.MasterCode1 = $ItemCode
    AND (V.[Date] IS NULL OR V.[Date] <= $dateLiteral)
    $waExcludeSql
ORDER BY V.[Date], T.VchCode, T.SrNo
"@

            $waRows = [System.Collections.ArrayList]::new()
            $waTxnRst = $fi.GetRecordset($waTxnQry)

            if ($waTxnRst -and -not $waTxnRst.EOF) {
                $waTxnRst.MoveFirst()

                while (-not $waTxnRst.EOF) {
                    $voucherDateText = ""

                    try {
                        $voucherDateText =
                            ([datetime]$waTxnRst.Fields.Item("VoucherDate").Value).
                                ToString("yyyy-MM-dd")
                    }
                    catch {}

                    [void]$waRows.Add([pscustomobject]@{
                        vchCode =
                            Convert-CostInt $waTxnRst.Fields.Item("VchCode").Value

                        mcCode =
                            Convert-CostInt $waTxnRst.Fields.Item("MCCode").Value

                        srNo =
                            Convert-CostInt $waTxnRst.Fields.Item("SrNo").Value

                        qty =
                            Convert-CostDouble $waTxnRst.Fields.Item("Value1").Value

                        value3 =
                            Convert-CostDouble $waTxnRst.Fields.Item("Value3").Value

                        d2 =
                            Convert-CostDouble $waTxnRst.Fields.Item("D2").Value

                        d5 =
                            Convert-CostDouble $waTxnRst.Fields.Item("D5").Value

                        vchType =
                            Convert-CostInt $waTxnRst.Fields.Item("VchType").Value

                        voucherDate =
                            $voucherDateText
                    })

                    $waTxnRst.MoveNext()
                }
            }

            Close-CostRecordset $waTxnRst

            # --------------------------------------------------------
            # Physical stock quantity always follows BUSY's signed Qty.
            # Valuation pool is updated separately.
            # --------------------------------------------------------
            foreach ($row in $waRows) {
                $state = Get-WaState -McCode ([int]$row.mcCode)

                $state.ledgerQty =
                    [double]$state.ledgerQty +
                    [double]$row.qty
            }

            # --------------------------------------------------------
            # BUSY method-5 parity uses date-level weighted valuation.
            # Process rate-changing inward/return transactions for a date
            # before recording normal Sales for that same date. This is
            # required to reproduce BUSY's 255.34 result in the supplied
            # Main Store example.
            # --------------------------------------------------------
            $waDates =
                @(
                    $waRows |
                    Group-Object voucherDate |
                    Sort-Object Name
                )

            foreach ($dateGroup in $waDates) {
                $dateRows = @($dateGroup.Group)

                # ====================================================
                # PHASE 1: Direct valuation-bearing inward transactions
                #          and Purchase Returns.
                # ====================================================
                foreach ($row in $dateRows) {
                    $qty = [double]$row.qty
                    $vchType = [int]$row.vchType
                    $mcCode = [int]$row.mcCode
                    $vchCode = [int]$row.vchCode

                    # Purchase / Material Receipt / other direct inward
                    # rows with a stored inventory value.
                    if ($qty -gt $eps -and $vchType -ne 3) {
                        $explicitValue =
                            Get-WaExplicitValue `
                                -Row $row `
                                -Qty $qty

                        if ($explicitValue -gt $eps) {
                            Add-WaPool `
                                -McCode $mcCode `
                                -Qty $qty `
                                -Value $explicitValue

                            if ($vchType -eq 2) {
                                $rate =
                                    if ($qty -gt $eps) {
                                        $explicitValue / $qty
                                    }
                                    else {
                                        0.0
                                    }

                                if (-not $waPurchaseSourceByVch.ContainsKey([string]$vchCode)) {
                                    $waPurchaseSourceByVch[[string]$vchCode] =
                                        [System.Collections.ArrayList]::new()
                                }

                                [void]$waPurchaseSourceByVch[[string]$vchCode].Add(
                                    [pscustomobject]@{
                                        mcCode      = $mcCode
                                        originalQty = [double]$qty
                                        rate        = [double]$rate
                                        returnedQty = 0.0
                                    }
                                )
                            }
                        }
                    }

                    # Purchase Return changes weighted-average valuation.
                    if ($qty -lt -$eps -and $vchType -eq 10) {
                        $returnQty = [Math]::Abs($qty)
                        $parentCode = 0

                        if ($waReturnParentMap.ContainsKey([string]$vchCode)) {
                            $parentCode =
                                [int]$waReturnParentMap[[string]$vchCode]
                        }

                        $remainingReturnQty = $returnQty

                        if (
                            $parentCode -gt 0 -and
                            $waPurchaseSourceByVch.ContainsKey([string]$parentCode)
                        ) {
                            foreach ($source in $waPurchaseSourceByVch[[string]$parentCode]) {
                                if ($remainingReturnQty -le $eps) {
                                    break
                                }

                                if ([int]$source.mcCode -ne $mcCode) {
                                    continue
                                }

                                $available =
                                    [double]$source.originalQty -
                                    [double]$source.returnedQty

                                if ($available -le $eps) {
                                    continue
                                }

                                $take =
                                    [Math]::Min(
                                        $available,
                                        $remainingReturnQty
                                    )

                                Remove-WaPool `
                                    -McCode $mcCode `
                                    -Qty $take `
                                    -UnitRate ([double]$source.rate)

                                $source.returnedQty =
                                    [double]$source.returnedQty + $take

                                $remainingReturnQty -= $take
                            }
                        }

                        if ($remainingReturnQty -gt $eps) {
                            $explicitReturnValue =
                                Get-WaExplicitValue `
                                    -Row $row `
                                    -Qty $remainingReturnQty

                            $returnRate = 0.0

                            if ($explicitReturnValue -gt $eps) {
                                $returnRate =
                                    $explicitReturnValue /
                                    $remainingReturnQty
                            }
                            else {
                                $returnRate =
                                    Get-WaRate -McCode $mcCode
                            }

                            Remove-WaPool `
                                -McCode $mcCode `
                                -Qty $remainingReturnQty `
                                -UnitRate $returnRate
                        }
                    }
                }

                # ====================================================
                # PHASE 2: Zero-value inward transfers / Stock Journal /
                #          Production rows.
                #
                # If BUSY has not stored an explicit inward amount for
                # the row, carry valuation from the same voucher's
                # outward material-centre weighted-average rate.
                # ====================================================
                $dateVoucherGroups =
                    @(
                        $dateRows |
                        Group-Object vchCode
                    )

                foreach ($voucherGroup in $dateVoucherGroups) {
                    $voucherRows = @($voucherGroup.Group)
                    $zeroValueInRows =
                        @(
                            $voucherRows |
                            Where-Object {
                                [double]$_.qty -gt $eps -and
                                [int]$_.vchType -ne 3 -and
                                (Get-WaExplicitValue -Row $_ -Qty ([double]$_.qty)) -le $eps
                            }
                        )

                    if ($zeroValueInRows.Count -eq 0) {
                        continue
                    }

                    $sourceValue = 0.0

                    foreach ($outRow in @(
                        $voucherRows |
                        Where-Object { [double]$_.qty -lt -$eps }
                    )) {
                        $outQty = [Math]::Abs([double]$outRow.qty)
                        $sourceRate =
                            Get-WaRate -McCode ([int]$outRow.mcCode)

                        $sourceValue +=
                            $outQty * $sourceRate
                    }

                    $inQtyTotal = 0.0

                    foreach ($inRow in $zeroValueInRows) {
                        $inQtyTotal += [double]$inRow.qty
                    }

                    foreach ($inRow in $zeroValueInRows) {
                        $qty = [double]$inRow.qty
                        $mcCode = [int]$inRow.mcCode

                        $value = 0.0

                        if (
                            $sourceValue -gt $eps -and
                            $inQtyTotal -gt $eps
                        ) {
                            # Transfer / Stock Journal style inward:
                            # when this same item's voucher also contains
                            # an outward row, carry the source MC valuation
                            # into the destination MC.
                            $value =
                                $sourceValue *
                                ($qty / $inQtyTotal)
                        }
                        else {
                            # IMPORTANT BUSY PARITY RULE:
                            #
                            # If BUSY stores an inward inventory row with
                            # Qty > 0 but Value3/D2/D5 = 0, and there is no
                            # same-voucher outward row for THIS ITEM, keep
                            # that inward valuation at ZERO.
                            #
                            # Do NOT manufacture a value using the current
                            # material-centre average. BUSY can intentionally
                            # create zero-valued Production/Stock-Journal
                            # inward quantity.
                            #
                            # Proven example:
                            #   Item: 20 SMD RESISTOR WE043
                            #   MC:   LP-1437
                            #
                            #   02-04-2026 transfer in:
                            #       10 qty, Value3=480  -> 48 each
                            #
                            #   21-04-2026 Production inward:
                            #       10 qty, Value3=0, D2=0, D5=0
                            #
                            #   BUSY valuation:
                            #       (480 + 0) / (10 + 10) = 24
                            #
                            # Previous code incorrectly assigned:
                            #       10 * currentRate(48) = 480
                            # which produced:
                            #       (480 + 480) / 20 = 48
                            $value = 0.0
                        }

                        Add-WaPool `
                            -McCode $mcCode `
                            -Qty $qty `
                            -Value $value

                        if ($value -le $eps) {
                            [void]$warnings.Add(
                                "Weighted-average inward VchCode=$([int]$inRow.vchCode) SrNo=$([int]$inRow.srNo) has zero stored value and no same-item transferable source; BUSY zero valuation was preserved."
                            )
                        }
                    }
                }

                # ====================================================
                # PHASE 3: Normal outward movements.
                #
                # IMPORTANT MOVING-WEIGHTED-AVERAGE RULE:
                #
                # A normal Sale / Transfer Out / Stock Journal Consumed /
                # Production Consumed DOES NOT change the unit average,
                # but it MUST reduce BOTH valuation-pool quantity and
                # valuation-pool value at the CURRENT average rate.
                #
                # Example:
                #   Pool before Sale = 89 qty / 22,725 value
                #   Average          = 255.337...
                #   Sale 1           = remove 1 qty + 255.337 value
                #   Pool after Sale  = 88 qty / 22,469.663 value
                #   Average remains  = 255.337...
                #
                # This matters when a LATER inward transaction happens:
                # the next weighted average must use the remaining 88,
                # not the historical 89.  The previous implementation
                # left outward qty/value inside the valuation pool, which
                # is why 21-04-2026 could match BUSY while 21-08-2026
                # drifted slightly (224.19 vs BUSY 224.05).
                #
                # Purchase Return (VchType 10) is excluded here because
                # Phase 1 already removes it using its parent/original
                # purchase valuation rule.
                # ====================================================
                foreach ($row in $dateRows) {
                    $rowQty = [double]$row.qty
                    $vchType = [int]$row.vchType

                    if (
                        $rowQty -lt -$eps -and
                        $vchType -ne 10
                    ) {
                        $vchCode = [int]$row.vchCode
                        $mcCode = [int]$row.mcCode
                        $outQty = [Math]::Abs($rowQty)

                        $stateBeforeOut = Get-WaState -McCode $mcCode

                        if ([double]$stateBeforeOut.poolQty -gt $eps) {
                            # Positive stock: issue the whole row at current WA.
                            $outRate = Get-WaRate -McCode $mcCode
                        }
                        else {
                            # Already zero/negative: use the row's posted stock
                            # value so negative issues form a signed WA balance.
                            $outRate =
                                Get-WaNegativeOutRate `
                                    -Row $row `
                                    -Qty $outQty

                            if ($outRate -le $eps) {
                                $outRate = Get-WaRate -McCode $mcCode
                            }
                        }

                        # Sale Return against a parent Sale must restore
                        # the parent's issue valuation, so remember the
                        # exact rate used by the Sale before removing it.
                        if ($vchType -eq 9) {
                            if (-not $waSaleRateByVch.ContainsKey([string]$vchCode)) {
                                $waSaleRateByVch[[string]$vchCode] =
                                    [System.Collections.ArrayList]::new()
                            }

                            [void]$waSaleRateByVch[[string]$vchCode].Add(
                                [pscustomobject]@{
                                    mcCode = $mcCode
                                    qty    = $outQty
                                    rate   = [double]$outRate
                                }
                            )
                        }

                        # Reduce valuation qty/value at current weighted
                        # average. This preserves the unit rate now while
                        # making future inward weighted averages correct.
                        Remove-WaPool `
                            -McCode $mcCode `
                            -Qty $outQty `
                            -UnitRate $outRate
                    }
                }

                # ====================================================
                # PHASE 4: Sale Returns.
                # Linked Sale Return uses the parent's valuation rate.
                # Unlinked Sale Return uses the current MC weighted rate,
                # so it does not artificially move the average.
                # ====================================================
                foreach ($row in $dateRows) {
                    if (
                        [double]$row.qty -gt $eps -and
                        [int]$row.vchType -eq 3
                    ) {
                        $qty = [double]$row.qty
                        $mcCode = [int]$row.mcCode
                        $vchCode = [int]$row.vchCode
                        $parentCode = 0

                        if ($waReturnParentMap.ContainsKey([string]$vchCode)) {
                            $parentCode =
                                [int]$waReturnParentMap[[string]$vchCode]
                        }

                        $returnRate = 0.0

                        if (
                            $parentCode -gt 0 -and
                            $waSaleRateByVch.ContainsKey([string]$parentCode)
                        ) {
                            $matchedQty = 0.0
                            $matchedValue = 0.0

                            foreach ($salePart in $waSaleRateByVch[[string]$parentCode]) {
                                if ([int]$salePart.mcCode -ne $mcCode) {
                                    continue
                                }

                                $take =
                                    [Math]::Min(
                                        [double]$salePart.qty,
                                        [Math]::Max(
                                            0.0,
                                            $qty - $matchedQty
                                        )
                                    )

                                if ($take -le $eps) {
                                    continue
                                }

                                $matchedQty += $take
                                $matchedValue +=
                                    $take *
                                    [double]$salePart.rate

                                if ($matchedQty -ge $qty - $eps) {
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
                                Get-WaRate -McCode $mcCode
                        }

                        Add-WaPool `
                            -McCode $mcCode `
                            -Qty $qty `
                            -Value ($qty * $returnRate)
                    }
                }
            }

            $selectedState =
                Get-WaState -McCode $selectedMcCode

            $remainingQty =
                [double]$selectedState.ledgerQty

            $averageCost =
                Get-WaRate -McCode $selectedMcCode

            # Signed BUSY valuation balance.  Negative stock must retain
            # its negative inventory value; do not clamp it to zero.
            $remainingValue = [double]$selectedState.poolValue

            if ($remainingQty -lt -$eps) {
                [void]$warnings.Add(
                    "Selected Material Centre has negative physical stock; signed BUSY method-5 valuation was preserved."
                )
            }

            if (Get-Command Write-SJDebug -ErrorAction SilentlyContinue) {
                Write-SJDebug (
                    "ItemCost Item=$ItemCode MC='$MaterialCentre'($selectedMcCode) " +
                    "AsOf=$($costDate.ToString('yyyy-MM-dd')) " +
                    "StockValMethod=$stockValMethod Mode=moving-weighted-average " +
                    "Qty=$remainingQty Avg=$averageCost Value=$remainingValue"
                )
            }

            return @{
                success = $true
                data = @{
                    itemCode           = $ItemCode
                    itemName           = $itemName.Trim()
                    materialCentre     = $MaterialCentre
                    materialCentreCode = $selectedMcCode
                    asOf               = $costDate.ToString("dd-MM-yyyy")

                    remainingQty =
                        [Math]::Round(
                            $remainingQty,
                            3
                        )

                    remainingValue =
                        [Math]::Round(
                            $remainingValue,
                            2
                        )

                    averageCost =
                        [Math]::Round(
                            $averageCost,
                            6
                        )

                    stockValMethod =
                        $stockValMethod

                    calculationMode =
                        "busy-signed-moving-weighted-average"

                    valuationSource =
                        "busy-stockvalmethod-5-signed-moving-weighted-average"

                    valuationPoolQty =
                        [Math]::Round(
                            [double]$selectedState.poolQty,
                            6
                        )

                    valuationPoolValue =
                        [Math]::Round(
                            [double]$selectedState.poolValue,
                            6
                        )

                    warnings =
                        @($warnings)
                }
            }
        }

        # ------------------------------------------------------------
        # Opening stock layers.
        # Diagnostic validation for BUSY Access schema:
        #   Tran4.D1 = opening main-unit quantity
        #   Tran4.D3 = opening stock amount/value
        # ------------------------------------------------------------
        $openingRst = $fi.GetRecordset(
            "SELECT MasterCode2, D1, D3 FROM Tran4 WHERE RecType=0 AND MasterCode1=$ItemCode"
        )

        $openingIndex = 0
        if ($openingRst -and -not $openingRst.EOF) {
            $openingRst.MoveFirst()
            while (-not $openingRst.EOF) {
                $openingIndex++
                $mcCode = Convert-CostInt $openingRst.Fields.Item("MasterCode2").Value
                $qty = Convert-CostDouble $openingRst.Fields.Item("D1").Value
                $value = Convert-CostDouble $openingRst.Fields.Item("D3").Value

                if ([Math]::Abs($qty) -gt $eps) {
                    Add-LedgerQty -McCode $mcCode -Qty $qty
                    if ($qty -gt 0) {
                        [void](Add-StockLayer `
                            -McCode $mcCode `
                            -Qty $qty `
                            -Value $value `
                            -SourceVchCode 0 `
                            -SourceVchType 0 `
                            -SourceType "opening" `
                            -SourceDate "opening" `
                            -SourceSrNo $openingIndex)
                    }
                    else {
                        [void]$warnings.Add("Negative opening stock found for MC code $mcCode; layer valuation cannot be fully reconstructed.")
                    }
                }

                $openingRst.MoveNext()
            }
        }
        Close-CostRecordset $openingRst

        # ------------------------------------------------------------
        # Parent return mapping through BUSY Tran3 references.
        # Method=1 = parent/original side, Method=2 = return/against side.
        # ------------------------------------------------------------
        try {
            $parentQry = @"
SELECT DISTINCT
    P.VchCode AS ParentVchCode,
    R.VchCode AS ReturnVchCode
FROM Tran3 P, Tran3 R, Tran1 H
WHERE
    P.RefCode = R.RefCode
    AND P.Method = 1
    AND R.Method = 2
    AND R.VchCode = H.VchCode
    AND (H.VchType = 3 OR H.VchType = 10)
"@
            $parentRst = $fi.GetRecordset($parentQry)
            if ($parentRst -and -not $parentRst.EOF) {
                $parentRst.MoveFirst()
                while (-not $parentRst.EOF) {
                    $parentCode = Convert-CostInt $parentRst.Fields.Item("ParentVchCode").Value
                    $returnCode = Convert-CostInt $parentRst.Fields.Item("ReturnVchCode").Value
                    if ($parentCode -gt 0 -and $returnCode -gt 0) {
                        $returnParentMap[[string]$returnCode] = $parentCode
                    }
                    $parentRst.MoveNext()
                }
            }
            Close-CostRecordset $parentRst
        }
        catch {
            [void]$warnings.Add("Could not read parent-return links from Tran3: $($_.Exception.Message)")
        }

        # ------------------------------------------------------------
        # Load every inventory movement for this item across ALL material
        # centres.  Tracking all MCs lets a transfer-in inherit the value of
        # the same voucher's FIFO source outflow when BUSY does not persist
        # an explicit inward amount.
        # ------------------------------------------------------------
        $excludeSql = ""
        if ($ExcludeVchCode -gt 0) {
            $excludeSql = " AND T.VchCode <> $ExcludeVchCode"
        }

        $txnQry = @"
SELECT
    T.VchCode,
    T.MasterCode2 AS MCCode,
    T.SrNo,
    T.Value1,
    T.Value3,
    T.D2,
    T.D5,
    V.VchType,
    V.[Date] AS VoucherDate,
    V.VchNo
FROM Tran2 T
LEFT JOIN Tran1 V ON V.VchCode = T.VchCode
WHERE
    T.RecType = 2
    AND T.MasterCode1 = $ItemCode
    AND (V.[Date] IS NULL OR V.[Date] <= $dateLiteral)
    $excludeSql
ORDER BY V.[Date], T.VchCode, T.SrNo
"@

        $txnRows = [System.Collections.ArrayList]::new()
        $txnRst = $fi.GetRecordset($txnQry)
        if ($txnRst -and -not $txnRst.EOF) {
            $txnRst.MoveFirst()
            while (-not $txnRst.EOF) {
                $voucherDateText = ""
                try {
                    $voucherDateText = ([datetime]$txnRst.Fields.Item("VoucherDate").Value).ToString("yyyy-MM-dd")
                }
                catch {}

                $voucherNoText = ""
                try { $voucherNoText = [string]$txnRst.Fields.Item("VchNo").Value } catch {}

                [void]$txnRows.Add([pscustomobject]@{
                    vchCode     = Convert-CostInt $txnRst.Fields.Item("VchCode").Value
                    mcCode      = Convert-CostInt $txnRst.Fields.Item("MCCode").Value
                    srNo        = Convert-CostInt $txnRst.Fields.Item("SrNo").Value
                    qty         = Convert-CostDouble $txnRst.Fields.Item("Value1").Value
                    value3      = Convert-CostDouble $txnRst.Fields.Item("Value3").Value
                    d2          = Convert-CostDouble $txnRst.Fields.Item("D2").Value
                    d5          = Convert-CostDouble $txnRst.Fields.Item("D5").Value
                    vchType     = Convert-CostInt $txnRst.Fields.Item("VchType").Value
                    voucherDate = $voucherDateText
                    vchNo       = $voucherNoText
                })

                $txnRst.MoveNext()
            }
        }
        Close-CostRecordset $txnRst

        function Process-VoucherGroup {
            param($Rows)

            if ($null -eq $Rows -or @($Rows).Count -eq 0) { return }

            $rowsArray = @($Rows)
            $vchCode = [int]$rowsArray[0].vchCode
            $vchType = [int]$rowsArray[0].vchType
            $parentCode = 0
            if ($returnParentMap.ContainsKey([string]$vchCode)) {
                $parentCode = [int]$returnParentMap[[string]$vchCode]
            }

            $outRows = @($rowsArray | Where-Object { [double]$_.qty -lt -$eps })
            $inRows = @($rowsArray | Where-Object { [double]$_.qty -gt $eps })

            $voucherOutValue = 0.0
            $voucherOutQty = 0.0

            # ----- stock out first -----
            foreach ($row in $outRows) {
                $mcCode = [int]$row.mcCode
                $qty = [Math]::Abs([double]$row.qty)
                Add-LedgerQty -McCode $mcCode -Qty ([double]$row.qty)

                $removal = $null

                if ($vchType -eq 10 -and $parentCode -gt 0) {
                    $removal = Remove-ParentPurchaseLayers `
                        -McCode $mcCode `
                        -ParentVchCode $parentCode `
                        -Qty $qty

                    if ([double]$removal.shortageQty -gt $eps) {
                        [void]$warnings.Add(
                            "Purchase Return VchCode=$vchCode references Purchase VchCode=$parentCode, but $([Math]::Round([double]$removal.shortageQty,3)) qty was no longer available in that parent purchase layer."
                        )
                    }
                }
                else {
                    # Includes Sale, unlinked Purchase Return, Stock Journal
                    # consumed, Production consumed, transfer-out, etc.
                    $removal = Remove-FifoLayers -McCode $mcCode -Qty $qty

                    if ([double]$removal.shortageQty -gt $eps) {
                        [void]$warnings.Add(
                            "FIFO stock-out VchCode=$vchCode exceeded reconstructed stock by $([Math]::Round([double]$removal.shortageQty,3)) qty in MC code $mcCode."
                        )
                    }
                }

                $voucherOutValue += [double]$removal.removedValue
                $voucherOutQty += [double]$removal.removedQty

                if ($vchType -eq 9) {
                    if (-not $saleAllocationsByVch.ContainsKey([string]$vchCode)) {
                        $saleAllocationsByVch[[string]$vchCode] = [System.Collections.ArrayList]::new()
                    }

                    foreach ($allocation in @($removal.allocations)) {
                        [void]$saleAllocationsByVch[[string]$vchCode].Add($allocation)
                        [void]$saleConsumptionHistory.Add($allocation)
                    }
                }
            }

            $voucherInboundQty = 0.0
            foreach ($r in $inRows) { $voucherInboundQty += [Math]::Abs([double]$r.qty) }

            # ----- stock in second -----
            foreach ($row in $inRows) {
                $mcCode = [int]$row.mcCode
                $qty = [Math]::Abs([double]$row.qty)
                Add-LedgerQty -McCode $mcCode -Qty ([double]$row.qty)

                if ($vchType -eq 3) {
                    $restore = $null

                    if ($parentCode -gt 0) {
                        if ($saleAllocationsByVch.ContainsKey([string]$parentCode)) {
                            # Parent-linked Sale Return: reverse only the exact
                            # layers used by the parent sale, starting from the
                            # last allocation made by that sale.
                            $restore = Restore-SaleAllocations `
                                -Allocations $saleAllocationsByVch[[string]$parentCode] `
                                -McCode $mcCode `
                                -Qty $qty `
                                -ReverseOrder $true
                        }
                        else {
                            $restore = [pscustomobject]@{
                                restoredQty = 0.0
                                restoredValue = 0.0
                                shortageQty = $qty
                            }
                        }
                    }
                    else {
                        # New-reference Sale Return: restore previously consumed
                        # Sale allocations FIFO across sale history.
                        $restore = Restore-SaleAllocations `
                            -Allocations $saleConsumptionHistory `
                            -McCode $mcCode `
                            -Qty $qty `
                            -ReverseOrder $false
                    }

                    $unresolved = [double]$restore.shortageQty
                    if ($unresolved -gt $eps) {
                        # This fallback is intentionally visible in warnings.
                        # It keeps posted stock quantity reconcilable when BUSY
                        # history does not expose enough sale-allocation detail.
                        $fallbackValue = Get-RowInwardValue `
                            -Row $row `
                            -Qty $unresolved `
                            -VoucherOutValue $voucherOutValue `
                            -VoucherInboundQty $voucherInboundQty

                        [void](Add-StockLayer `
                            -McCode $mcCode `
                            -Qty $unresolved `
                            -Value $fallbackValue `
                            -SourceVchCode $vchCode `
                            -SourceVchType $vchType `
                            -SourceType "sale-return-unresolved" `
                            -SourceDate ([string]$row.voucherDate) `
                            -SourceSrNo ([int]$row.srNo))

                        [void]$warnings.Add(
                            "Sale Return VchCode=$vchCode could not restore $([Math]::Round($unresolved,3)) qty from sale allocation history; transaction/current-average fallback value was used for that remainder."
                        )
                    }

                    continue
                }

                # Normal inward layer: Purchase, transfer-in, Stock Journal
                # generated, Production generated, Material Receipt, etc.
                $layerValue = Get-RowInwardValue `
                    -Row $row `
                    -Qty $qty `
                    -VoucherOutValue $voucherOutValue `
                    -VoucherInboundQty $voucherInboundQty

                if ($layerValue -le $eps -and $qty -gt $eps) {
                    [void]$warnings.Add(
                        "Inward VchCode=$vchCode row SrNo=$([int]$row.srNo) has qty $qty but no resolvable inventory value; a zero-value layer was created."
                    )
                }

                [void](Add-StockLayer `
                    -McCode $mcCode `
                    -Qty $qty `
                    -Value $layerValue `
                    -SourceVchCode $vchCode `
                    -SourceVchType $vchType `
                    -SourceType (Get-SourceTypeName -VchType $vchType -IsInward $true) `
                    -SourceDate ([string]$row.voucherDate) `
                    -SourceSrNo ([int]$row.srNo))
            }
        }

        # Replay in Date -> VchCode -> SrNo order, grouping rows by voucher so
        # stock-out value can be carried to same-voucher transfer/generated rows.
        $currentVchCode = -1
        $groupRows = [System.Collections.ArrayList]::new()

        foreach ($row in @($txnRows)) {
            if ($currentVchCode -eq -1) {
                $currentVchCode = [int]$row.vchCode
            }

            if ([int]$row.vchCode -ne $currentVchCode) {
                Process-VoucherGroup -Rows $groupRows
                $groupRows = [System.Collections.ArrayList]::new()
                $currentVchCode = [int]$row.vchCode
            }

            [void]$groupRows.Add($row)
        }

        if ($groupRows.Count -gt 0) {
            Process-VoucherGroup -Rows $groupRows
        }

        $totals = Get-McLayerTotals -McCode $selectedMcCode
        $remainingQty = [double]$totals.qty
        $remainingValue = [double]$totals.value
        $averageCost = if ($remainingQty -gt $eps) {
            $remainingValue / $remainingQty
        }
        else { 0.0 }

        $ledgerKey = [string]$selectedMcCode
        $ledgerQty = if ($ledgerQtyByMc.ContainsKey($ledgerKey)) {
            [double]$ledgerQtyByMc[$ledgerKey]
        }
        else { 0.0 }

        if ([Math]::Abs($ledgerQty - $remainingQty) -gt 0.0005) {
            [void]$warnings.Add(
                "Reconstructed valued-layer qty ($([Math]::Round($remainingQty,3))) differs from signed BUSY ledger qty ($([Math]::Round($ledgerQty,3))). This normally indicates negative stock or an unresolved parent-return history."
            )
        }

        $layerCount = 0
        Ensure-LayerList -McCode $selectedMcCode
        $selectedLayerList = $layersByMc[[string]$selectedMcCode]

        foreach ($layer in $selectedLayerList) {
            if ([double]$layer.remainingQty -gt $eps) { $layerCount++ }
        }

        if (Get-Command Write-SJDebug -ErrorAction SilentlyContinue) {
            Write-SJDebug (
                "ItemCost Item=$ItemCode MC='$MaterialCentre'($selectedMcCode) AsOf=$($costDate.ToString('yyyy-MM-dd')) " +
                "Qty=$remainingQty Value=$remainingValue Avg=$averageCost LedgerQty=$ledgerQty Layers=$layerCount"
            )
        }

        return @{
            success = $true
            data = @{
                itemCode        = $ItemCode
                itemName        = $itemName.Trim()
                materialCentre  = $MaterialCentre
                materialCentreCode = $selectedMcCode
                asOf            = $costDate.ToString("dd-MM-yyyy")
                remainingQty    = [Math]::Round($remainingQty, 3)
                remainingValue  = [Math]::Round($remainingValue, 2)
                averageCost     = [Math]::Round($averageCost, 6)
                ledgerQty       = [Math]::Round($ledgerQty, 3)
                remainingLayers = $layerCount
                stockValMethod   = $stockValMethod
                calculationMode = "fifo-layer-replay-fallback"
                valuationSource = "fifo-layer-replay-current-average"
                warnings        = @($warnings)
            }
        }
    }
    catch {
        $errorLine = 0
        $errorPosition = ""
        $errorStack = ""

        try { $errorLine = [int]$_.InvocationInfo.ScriptLineNumber } catch {}
        try { $errorPosition = [string]$_.InvocationInfo.PositionMessage } catch {}
        try { $errorStack = [string]$_.ScriptStackTrace } catch {}

        try {
            Write-Host (
                "  [ITEM COST ERROR] Line=$errorLine Message=$($_.Exception.Message)" +
                $(if ($errorStack) { "`n  Stack: $errorStack" } else { "" })
            ) -ForegroundColor Red
        }
        catch {}

        return @{
            success = $false
            error   = "Item costing failed at line ${errorLine}: $($_.Exception.Message)"
            debug   = @{
                line     = $errorLine
                position = $errorPosition
                stack    = $errorStack
            }
        }
    }
    finally {
        Disconnect-BUSY $fi
    }
}

