# 11-returns.ps1
# Extracted from vouchers.ps1. Keep functions behavior-compatible with the monolith.

function Search-OriginalVouchers {
    param(
        [int]$VchType,
        [string]$VchNo = "",
        [string]$Party = "",
        [string]$FromDate = "",
        [string]$ToDate = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $fi = Connect-BUSY `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $fi) {
        return @{
            success = $false
            error   = "BUSY connection failed"
        }
    }

    try {
        # --------------------------------------------------------
        # Validate supported transaction types
        # --------------------------------------------------------
        if ($VchType -notin @(2, 3, 9, 10)) {
            return @{
                success = $false
                error   = "Unsupported pending bill voucher type: $VchType"
            }
        }

        # --------------------------------------------------------
        # Transaction label
        # --------------------------------------------------------
        $transactionType = switch ($VchType) {
            2  { "Purchase" }
            3  { "Sale Return" }
            9  { "Sale" }
            10 { "Purchase Return" }
            default { "Voucher" }
        }

        # --------------------------------------------------------
        # Resolve party master code
        # --------------------------------------------------------
        $partyCode = 0

        if (-not [string]::IsNullOrWhiteSpace($Party)) {
            $safeParty = $Party.Trim() -replace "'", "''"

            $partyRst = $fi.GetRecordset(
                "SELECT Code
                 FROM Master1
                 WHERE Name='$safeParty'"
            )

            if ($partyRst -and -not $partyRst.EOF) {
                try {
                    $partyCode = [int]$partyRst.Fields.Item(
                        "Code"
                    ).Value
                } catch {
                    $partyCode = 0
                }
            }

            if ($partyRst) {
                try { $partyRst.Close() } catch {}
            }

            if ($partyCode -le 0) {
                return @{
                    success = $true
                    count   = 0
                    data    = @()
                }
            }
        }

        # --------------------------------------------------------
        # Build original reference filter
        #
        # Method 1 = original/new bill reference
        # Method 3/4 retained for compatibility with existing data
        # --------------------------------------------------------
        $whereParts = @(
            "T3.RecType=1",
            "T3.Method IN (1,3,4)",
            "T3.VchType=$VchType",
            "H.Cancelled=0",
            "H.VchCancelled=0"
        )

        if ($partyCode -gt 0) {
            $whereParts += "T3.MasterCode1=$partyCode"
        }

        if (-not [string]::IsNullOrWhiteSpace($VchNo)) {
            $safeVchNo = $VchNo.Trim() -replace "'", "''"

            $whereParts += @"
(
    H.VchNo LIKE '*$safeVchNo*'
    OR T3.[No] LIKE '*$safeVchNo*'
)
"@
        }

        if (-not [string]::IsNullOrWhiteSpace($FromDate)) {
            try {
                $fromParsed = [datetime]::ParseExact(
                    $FromDate,
                    "dd-MM-yyyy",
                    [System.Globalization.CultureInfo]::InvariantCulture
                )

                $fromAccess = $fromParsed.ToString("MM/dd/yyyy")
                $whereParts += "T3.[Date] >= #$fromAccess#"
            } catch {}
        }

        if (-not [string]::IsNullOrWhiteSpace($ToDate)) {
            try {
                $toParsed = [datetime]::ParseExact(
                    $ToDate,
                    "dd-MM-yyyy",
                    [System.Globalization.CultureInfo]::InvariantCulture
                )

                $toAccess = $toParsed.ToString("MM/dd/yyyy")
                $whereParts += "T3.[Date] <= #$toAccess#"
            } catch {}
        }

        $whereClause = $whereParts -join " AND "

        # --------------------------------------------------------
        # Load original bill references
        # --------------------------------------------------------
        $originalQuery = @"
SELECT
    T3.RefCode,
    T3.VchCode,
    T3.VchType,
    T3.MasterCode1,
    T3.[No] AS RefNo,
    T3.[Date] AS RefDate,
    T3.DueDate,
    T3.Value1 AS ParentValue,
    H.VchNo AS CurrentVchNo,
    H.AutoVchNo,
    H.VchAmtBaseCur,
    H.OrgVchAmtBaseCur
FROM Tran3 T3
INNER JOIN Tran1 H
    ON H.VchCode=T3.VchCode
WHERE $whereClause
ORDER BY T3.[Date] DESC
"@

        $originalRst = $fi.GetRecordset($originalQuery)

        $results = @()
        $seenReferenceCodes = @{}

        if ($originalRst -and -not $originalRst.EOF) {
            $originalRst.MoveFirst()

            while (-not $originalRst.EOF) {
                $refCode = 0
                $parentVchCode = 0
                $refNo = ""
                $dateString = ""
                $dueDateString = ""
                $parentAmount = 0.0

                try {
                    $refCode = [int]$originalRst.Fields.Item(
                        "RefCode"
                    ).Value
                } catch {}

                try {
                    $parentVchCode = [int]$originalRst.Fields.Item(
                        "VchCode"
                    ).Value
                } catch {}

                # ------------------------------------------------
                # Use current Tran1 voucher number first.
                # Tran3.[No] can contain an old/stale number.
                # ------------------------------------------------
                try {
                    $rawCurrentNo = $originalRst.Fields.Item(
                        "CurrentVchNo"
                    ).Value

                    if (
                        $rawCurrentNo -ne [System.DBNull]::Value -and
                        $null -ne $rawCurrentNo
                    ) {
                        $refNo = $rawCurrentNo.ToString().Trim()
                    }
                } catch {}

                if ([string]::IsNullOrWhiteSpace($refNo)) {
                    try {
                        $rawReferenceNo = $originalRst.Fields.Item(
                            "RefNo"
                        ).Value

                        if (
                            $rawReferenceNo -ne [System.DBNull]::Value -and
                            $null -ne $rawReferenceNo
                        ) {
                            $refNo = $rawReferenceNo.ToString().Trim()
                        }
                    } catch {}
                }

                if ([string]::IsNullOrWhiteSpace($refNo)) {
                    try {
                        $rawAutoNo = $originalRst.Fields.Item(
                            "AutoVchNo"
                        ).Value

                        if (
                            $rawAutoNo -ne [System.DBNull]::Value -and
                            [int]$rawAutoNo -gt 0
                        ) {
                            $refNo = [string]$rawAutoNo
                        }
                    } catch {}
                }

                # ------------------------------------------------
                # Date
                # ------------------------------------------------
                try {
                    $rawDate = $originalRst.Fields.Item(
                        "RefDate"
                    ).Value

                    if (
                        $rawDate -ne [System.DBNull]::Value -and
                        $null -ne $rawDate
                    ) {
                        if ($rawDate -is [datetime]) {
                            $dateString = $rawDate.ToString(
                                "dd-MM-yyyy"
                            )
                        } else {
                            $dateString = [datetime]::Parse(
                                $rawDate.ToString()
                            ).ToString("dd-MM-yyyy")
                        }
                    }
                } catch {}

                # ------------------------------------------------
                # Due date
                # ------------------------------------------------
                $dueDateString = $dateString

                try {
                    $rawDueDate = $originalRst.Fields.Item(
                        "DueDate"
                    ).Value

                    if (
                        $rawDueDate -ne [System.DBNull]::Value -and
                        $null -ne $rawDueDate
                    ) {
                        $parsedDueDate = if (
                            $rawDueDate -is [datetime]
                        ) {
                            $rawDueDate
                        } else {
                            [datetime]::Parse(
                                $rawDueDate.ToString()
                            )
                        }

                        if ($parsedDueDate.Year -gt 1900) {
                            $dueDateString = $parsedDueDate.ToString(
                                "dd-MM-yyyy"
                            )
                        }
                    }
                } catch {}

                # ------------------------------------------------
                # Parent transaction amount
                #
                # Tran3.Value1 is normally the original reference
                # amount. Fall back to Tran1 amount if necessary.
                # ------------------------------------------------
                try {
                    $rawParentValue = $originalRst.Fields.Item(
                        "ParentValue"
                    ).Value

                    if (
                        $rawParentValue -ne [System.DBNull]::Value -and
                        $null -ne $rawParentValue
                    ) {
                        $parentAmount = [Math]::Abs(
                            [double]$rawParentValue
                        )
                    }
                } catch {}

                if ($parentAmount -le 0) {
                    try {
                        $rawVoucherAmount = $originalRst.Fields.Item(
                            "VchAmtBaseCur"
                        ).Value

                        if (
                            $rawVoucherAmount -ne [System.DBNull]::Value -and
                            $null -ne $rawVoucherAmount
                        ) {
                            $parentAmount = [Math]::Abs(
                                [double]$rawVoucherAmount
                            )
                        }
                    } catch {}
                }

                if ($parentAmount -le 0) {
                    try {
                        $rawOriginalAmount = $originalRst.Fields.Item(
                            "OrgVchAmtBaseCur"
                        ).Value

                        if (
                            $rawOriginalAmount -ne [System.DBNull]::Value -and
                            $null -ne $rawOriginalAmount
                        ) {
                            $parentAmount = [Math]::Abs(
                                [double]$rawOriginalAmount
                            )
                        }
                    } catch {}
                }

                $parentAmount = [Math]::Round(
                    $parentAmount,
                    2
                )

                # ------------------------------------------------
                # Calculate linked adjustment categories
                #
                # VchType 3  = Sale Return
                # VchType 10 = Purchase Return
                # VchType 14 = Receipt
                # VchType 19 = Payment
                # ------------------------------------------------
                $saleReturnedAmount = 0.0
                $purchaseReturnedAmount = 0.0
                $alreadyReceivedAmount = 0.0
                $alreadyPaidAmount = 0.0
                $otherAdjustedAmount = 0.0

                if ($refCode -gt 0) {
                    $adjustmentQuery = @"
SELECT
    A.VchType,
    SUM(ABS(A.Value1)) AS AdjustedAmount
FROM Tran3 A
INNER JOIN Tran1 AH
    ON AH.VchCode=A.VchCode
WHERE
    A.RecType=1
    AND A.Method=2
    AND A.RefCode=$refCode
    AND A.VchCode<>$parentVchCode
    AND AH.Cancelled=0
    AND AH.VchCancelled=0
GROUP BY A.VchType
"@

                    $adjustmentRst = $fi.GetRecordset(
                        $adjustmentQuery
                    )

                    if (
                        $adjustmentRst -and
                        -not $adjustmentRst.EOF
                    ) {
                        $adjustmentRst.MoveFirst()

                        while (-not $adjustmentRst.EOF) {
                            $adjustmentVchType = 0
                            $adjustmentAmount = 0.0

                            try {
                                $adjustmentVchType = [int](
                                    $adjustmentRst.Fields.Item(
                                        "VchType"
                                    ).Value
                                )
                            } catch {}

                            try {
                                $rawAdjustedAmount =
                                    $adjustmentRst.Fields.Item(
                                        "AdjustedAmount"
                                    ).Value

                                if (
                                    $rawAdjustedAmount -ne
                                        [System.DBNull]::Value -and
                                    $null -ne $rawAdjustedAmount
                                ) {
                                    $adjustmentAmount = [Math]::Abs(
                                        [double]$rawAdjustedAmount
                                    )
                                }
                            } catch {}

                            switch ($adjustmentVchType) {
                                3 {
                                    $saleReturnedAmount +=
                                        $adjustmentAmount
                                }

                                10 {
                                    $purchaseReturnedAmount +=
                                        $adjustmentAmount
                                }

                                14 {
                                    $alreadyReceivedAmount +=
                                        $adjustmentAmount
                                }

                                19 {
                                    $alreadyPaidAmount +=
                                        $adjustmentAmount
                                }

                                default {
                                    $otherAdjustedAmount +=
                                        $adjustmentAmount
                                }
                            }

                            $adjustmentRst.MoveNext()
                        }
                    }

                    if ($adjustmentRst) {
                        try { $adjustmentRst.Close() } catch {}
                    }
                }

                $saleReturnedAmount = [Math]::Round(
                    $saleReturnedAmount,
                    2
                )

                $purchaseReturnedAmount = [Math]::Round(
                    $purchaseReturnedAmount,
                    2
                )

                $alreadyReceivedAmount = [Math]::Round(
                    $alreadyReceivedAmount,
                    2
                )

                $alreadyPaidAmount = [Math]::Round(
                    $alreadyPaidAmount,
                    2
                )

                $otherAdjustedAmount = [Math]::Round(
                    $otherAdjustedAmount,
                    2
                )

                # ------------------------------------------------
                # Calculate pending amount by parent transaction
                #
                # Sale:
                # Parent - Sale Returns - Receipts
                #
                # Purchase Return:
                # Parent - Receipts
                #
                # Purchase:
                # Parent - Purchase Returns - Payments
                #
                # Sale Return:
                # Parent - Payments
                # ------------------------------------------------
                $pendingAmount = $parentAmount

                switch ($VchType) {
                    9 {
                        $pendingAmount =
                            $parentAmount -
                            $saleReturnedAmount -
                            $alreadyReceivedAmount -
                            $otherAdjustedAmount
                    }

                    10 {
                        $pendingAmount =
                            $parentAmount -
                            $alreadyReceivedAmount -
                            $otherAdjustedAmount
                    }

                    2 {
                        $pendingAmount =
                            $parentAmount -
                            $purchaseReturnedAmount -
                            $alreadyPaidAmount -
                            $otherAdjustedAmount
                    }

                    3 {
                        $pendingAmount =
                            $parentAmount -
                            $alreadyPaidAmount -
                            $otherAdjustedAmount
                    }
                }

                $pendingAmount = [Math]::Max(
                    0,
                    [Math]::Round($pendingAmount, 2)
                )

                # ------------------------------------------------
                # Keep only unique original references
                # ------------------------------------------------
                $uniqueKey = if ($refCode -gt 0) {
                    "REF:$refCode"
                } else {
                    "VCH:$parentVchCode"
                }

                if (
                    $pendingAmount -gt 0.005 -and
                    -not $seenReferenceCodes.ContainsKey($uniqueKey)
                ) {
                    $returnedAmountForCurrentType = switch (
                        $VchType
                    ) {
                        9 { $saleReturnedAmount }
                        2 { $purchaseReturnedAmount }
                        default { 0.0 }
                    }

                    $settledAmountForCurrentType = switch (
                        $VchType
                    ) {
                        9  { $alreadyReceivedAmount }
                        10 { $alreadyReceivedAmount }
                        2  { $alreadyPaidAmount }
                        3  { $alreadyPaidAmount }
                        default { 0.0 }
                    }

                    $results += @{
                        refCode = $refCode

                        vchCode = $parentVchCode
                        vchNo   = $refNo
                        date    = $dateString
                        dueDate = $dueDateString

                        vchType         = $VchType
                        transactionType = $transactionType
                        voucherType     = $transactionType

                        party = $Party

                        # New detailed values
                        parentAmount = $parentAmount

                        saleReturnedAmount =
                            $saleReturnedAmount

                        purchaseReturnedAmount =
                            $purchaseReturnedAmount

                        alreadyReceivedAmount =
                            $alreadyReceivedAmount

                        alreadyPaidAmount =
                            $alreadyPaidAmount

                        pendingAmount = $pendingAmount

                        # Generic frontend convenience fields
                        returnedAmount =
                            $returnedAmountForCurrentType

                        settledAmount =
                            $settledAmountForCurrentType

                        # Backward-compatible fields
                        totalAmt = $parentAmount
                        netAmt   = $pendingAmount
                        returnedAmt =
                            $returnedAmountForCurrentType

                        fullyAdjusted =
                            ($pendingAmount -le 0.005)

                        fullyReturned =
                            ($pendingAmount -le 0.005)
                    }

                    $seenReferenceCodes[$uniqueKey] = $true
                }

                $originalRst.MoveNext()
            }
        }

        if ($originalRst) {
            try { $originalRst.Close() } catch {}
        }

        # --------------------------------------------------------
        # Sort newest first
        # --------------------------------------------------------
        $sortedResults = @(
            $results | Sort-Object {
                try {
                    [datetime]::ParseExact(
                        $_.date,
                        "dd-MM-yyyy",
                        [System.Globalization.CultureInfo]::InvariantCulture
                    )
                } catch {
                    [datetime]::MinValue
                }
            } -Descending
        )

        return @{
            success = $true
            count   = $sortedResults.Count
            data    = $sortedResults
        }
    }
    catch {
        return @{
            success = $false
            error   = $_.Exception.Message
        }
    }
    finally {
        Disconnect-BUSY $fi
    }
}

# ═══════════════════════════════════════════════════════════════
#  GET RETURN HISTORY
# ═══════════════════════════════════════════════════════════════

function Get-ReturnHistory {
    param(
        [int]$OrigVchType,
        [int]$OrigVchCode,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $returnVchType = if ($OrigVchType -eq 9) { 3 } else { 10 }

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) {
        return @{
            success = $false
            error   = "BUSY connection failed"
        }
    }

    try {
        # ------------------------------------------------------------
        # Load master names
        # ------------------------------------------------------------
        $masterMap = @{}

        $mRst = $fi.GetRecordset(
            "SELECT Code, Name FROM Master1"
        )

        if ($null -ne $mRst -and -not $mRst.EOF) {
            $mRst.MoveFirst()

            while (-not $mRst.EOF) {
                $code = ""
                $name = ""

                try {
                    $rawCode = $mRst.Fields.Item("Code").Value
                    if ($rawCode -ne [System.DBNull]::Value) {
                        $code = $rawCode.ToString().Trim()
                    }
                } catch {}

                try {
                    $rawName = $mRst.Fields.Item("Name").Value
                    if ($rawName -ne [System.DBNull]::Value) {
                        $name = $rawName.ToString().Trim()
                    }
                } catch {}

                if (-not [string]::IsNullOrWhiteSpace($code)) {
                    $masterMap[$code] = $name
                }

                $mRst.MoveNext()
            }

            try { $mRst.Close() } catch {}
        }

        # ------------------------------------------------------------
        # Read and accumulate original voucher items
        # Use item code as the key so duplicate item rows are summed.
        # ------------------------------------------------------------
        $originalItems = @{}

        $oRst = $fi.GetRecordset(
            "SELECT MasterCode1, Value1, Value2
             FROM Tran2
             WHERE VchCode=$OrigVchCode
               AND RecType=2"
        )

        if ($null -ne $oRst -and -not $oRst.EOF) {
            $oRst.MoveFirst()

            while (-not $oRst.EOF) {
                $itemCode = ""
                $qty = 0.0
                $altQty = 0.0

                try {
                    $rawCode = $oRst.Fields.Item("MasterCode1").Value
                    if ($rawCode -ne [System.DBNull]::Value) {
                        $itemCode = $rawCode.ToString().Trim()
                    }
                } catch {}

                try {
                    $rawQty = $oRst.Fields.Item("Value1").Value
                    if ($rawQty -ne [System.DBNull]::Value) {
                        $qty = [Math]::Abs([double]$rawQty)
                    }
                } catch {}

                try {
                    $rawAltQty = $oRst.Fields.Item("Value2").Value
                    if ($rawAltQty -ne [System.DBNull]::Value) {
                        $altQty = [Math]::Abs([double]$rawAltQty)
                    }
                } catch {}

                if (
                    -not [string]::IsNullOrWhiteSpace($itemCode) -and
                    $qty -gt 0
                ) {
                    $itemName = if ($masterMap.ContainsKey($itemCode)) {
                        $masterMap[$itemCode]
                    } else {
                        $itemCode
                    }

                    if (-not $originalItems.ContainsKey($itemCode)) {
                        $originalItems[$itemCode] = @{
                            itemCode       = $itemCode
                            itemName       = $itemName
                            originalQty    = 0.0
                            originalAltQty = 0.0
                        }
                    }

                    # Important: accumulate duplicate item rows
                    $originalItems[$itemCode].originalQty += $qty
                    $originalItems[$itemCode].originalAltQty += $altQty
                }

                $oRst.MoveNext()
            }

            try { $oRst.Close() } catch {}
        }

        # ------------------------------------------------------------
        # Find all return vouchers linked to the original voucher
        # ------------------------------------------------------------
        $linkedReturnCodes = @()

        $lRst = $fi.GetRecordset(
            "SELECT DISTINCT T2.VchCode
             FROM Tran3 T1, Tran3 T2
             WHERE T1.RefCode = T2.RefCode
               AND T1.Method = 1
               AND T1.VchCode = $OrigVchCode
               AND T1.VchType = $OrigVchType
               AND T2.VchType = $returnVchType
               AND T2.Method = 2"
        )

        if ($null -ne $lRst -and -not $lRst.EOF) {
            $lRst.MoveFirst()

            while (-not $lRst.EOF) {
                $returnCode = 0

                try {
                    $rawReturnCode = $lRst.Fields.Item("VchCode").Value
                    if ($rawReturnCode -ne [System.DBNull]::Value) {
                        $returnCode = [int]$rawReturnCode
                    }
                } catch {}

                if (
                    $returnCode -gt 0 -and
                    $linkedReturnCodes -notcontains $returnCode
                ) {
                    $linkedReturnCodes += $returnCode
                }

                $lRst.MoveNext()
            }

            try { $lRst.Close() } catch {}
        }

        # ------------------------------------------------------------
        # Read each linked return voucher
        # ------------------------------------------------------------
        $returnVouchers = @()

        foreach ($retCode in $linkedReturnCodes) {
            $retVchNo = ""
            $retDate = ""
            $retItemsByCode = @{}

            $hRst = $fi.GetRecordset(
                "SELECT VchNo, Date
                 FROM Tran1
                 WHERE VchCode=$retCode"
            )

            if ($null -ne $hRst -and -not $hRst.EOF) {
                $hRst.MoveFirst()

                try {
                    $rawVchNo = $hRst.Fields.Item("VchNo").Value
                    if ($rawVchNo -ne [System.DBNull]::Value) {
                        $retVchNo = $rawVchNo.ToString().Trim()
                    }
                } catch {}

                try {
                    $rawDate = $hRst.Fields.Item("Date").Value

                    if ($rawDate -ne [System.DBNull]::Value) {
                        if ($rawDate -is [datetime]) {
                            $retDate = $rawDate.ToString("dd-MM-yyyy")
                        } else {
                            $retDate = [datetime]::Parse(
                                $rawDate.ToString()
                            ).ToString("dd-MM-yyyy")
                        }
                    }
                } catch {}
            }

            if ($null -ne $hRst) {
                try { $hRst.Close() } catch {}
            }

            $riRst = $fi.GetRecordset(
                "SELECT MasterCode1, Value1, Value2
                 FROM Tran2
                 WHERE VchCode=$retCode
                   AND RecType=2"
            )

            if ($null -ne $riRst -and -not $riRst.EOF) {
                $riRst.MoveFirst()

                while (-not $riRst.EOF) {
                    $itemCode = ""
                    $qty = 0.0
                    $altQty = 0.0

                    try {
                        $rawCode = $riRst.Fields.Item("MasterCode1").Value
                        if ($rawCode -ne [System.DBNull]::Value) {
                            $itemCode = $rawCode.ToString().Trim()
                        }
                    } catch {}

                    try {
                        $rawQty = $riRst.Fields.Item("Value1").Value
                        if ($rawQty -ne [System.DBNull]::Value) {
                            $qty = [Math]::Abs([double]$rawQty)
                        }
                    } catch {}

                    try {
                        $rawAltQty = $riRst.Fields.Item("Value2").Value
                        if ($rawAltQty -ne [System.DBNull]::Value) {
                            $altQty = [Math]::Abs([double]$rawAltQty)
                        }
                    } catch {}

                    if (
                        -not [string]::IsNullOrWhiteSpace($itemCode) -and
                        $qty -gt 0
                    ) {
                        $itemName = if ($masterMap.ContainsKey($itemCode)) {
                            $masterMap[$itemCode]
                        } else {
                            $itemCode
                        }

                        if (-not $retItemsByCode.ContainsKey($itemCode)) {
                            $retItemsByCode[$itemCode] = @{
                                itemCode       = $itemCode
                                itemName       = $itemName
                                returnedQty    = 0.0
                                returnedAltQty = 0.0
                            }
                        }

                        # Accumulate duplicate rows in the return voucher
                        $retItemsByCode[$itemCode].returnedQty += $qty
                        $retItemsByCode[$itemCode].returnedAltQty += $altQty
                    }

                    $riRst.MoveNext()
                }

                try { $riRst.Close() } catch {}
            }

            $retItems = @()

            foreach ($itemCode in $retItemsByCode.Keys) {
                $item = $retItemsByCode[$itemCode]

                $retItems += @{
                    itemCode       = $item.itemCode
                    itemName       = $item.itemName
                    returnedQty    = [Math]::Round(
                        [double]$item.returnedQty,
                        3
                    )
                    returnedAltQty = [Math]::Round(
                        [double]$item.returnedAltQty,
                        3
                    )
                }
            }

            $returnVouchers += @{
                vchCode = $retCode
                vchNo   = $retVchNo
                date    = $retDate
                items   = $retItems
            }
        }

        # ------------------------------------------------------------
        # Build final per-item return summary
        # ------------------------------------------------------------
        $itemSummary = @()

        foreach ($itemCode in $originalItems.Keys) {
            $originalItem = $originalItems[$itemCode]

            $originalQty = [Math]::Round(
                [double]$originalItem.originalQty,
                3
            )

            $originalAltQty = [Math]::Round(
                [double]$originalItem.originalAltQty,
                3
            )

            $returnedQty = 0.0
            $returnedAltQty = 0.0

            foreach ($returnVoucher in $returnVouchers) {
                foreach ($returnedItem in @($returnVoucher.items)) {
                    if (
                        [string]$returnedItem.itemCode -eq
                        [string]$itemCode
                    ) {
                        $returnedQty += [double]$returnedItem.returnedQty
                        $returnedAltQty += [double]$returnedItem.returnedAltQty
                    }
                }
            }

            $returnedQty = [Math]::Round($returnedQty, 3)
            $returnedAltQty = [Math]::Round($returnedAltQty, 3)

            $availableQty = [Math]::Max(
                0,
                [Math]::Round(
                    $originalQty - $returnedQty,
                    3
                )
            )

            $availableAltQty = [Math]::Max(
                0,
                [Math]::Round(
                    $originalAltQty - $returnedAltQty,
                    3
                )
            )

            $itemSummary += @{
                itemCode        = $itemCode
                itemName        = $originalItem.itemName
                originalQty     = $originalQty
                returnedQty     = $returnedQty
                availableQty    = $availableQty
                originalAltQty  = $originalAltQty
                returnedAltQty  = $returnedAltQty
                availableAltQty = $availableAltQty
                fullyReturned   = ($availableQty -le 0.000001)
            }
        }

        return @{
            success         = $true
            originalVchCode = $OrigVchCode
            returnVouchers  = $returnVouchers
            itemSummary     = $itemSummary
        }
    }
    catch {
        return @{
            success = $false
            error   = $_.Exception.Message
        }
    }
    finally {
        Disconnect-BUSY $fi
    }
}

# ═══════════════════════════════════════════════════════════════
#  PRODUCTION BOM HELPERS
# ═══════════════════════════════════════════════════════════════

