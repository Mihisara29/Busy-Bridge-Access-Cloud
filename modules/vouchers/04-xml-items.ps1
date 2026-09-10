# 04-xml-items.ps1
# Extracted from vouchers.ps1. Keep functions behavior-compatible with the monolith.

function Build-ItemsXml {
    param(
        $items,
        [string]$defaultMC = "Main Store",
        [string]$salesmanName = ""
    )

    $culture = [System.Globalization.CultureInfo]::InvariantCulture

    function Format-XmlNumber {
        param(
            [double]$Value,
            [string]$Format = "0.####"
        )

        return $Value.ToString($Format, $culture)
    }

    $xml = "<ItemEntries>"

    foreach ($item in @($items)) {
        if (
            -not $item.itemName -or
            [string]::IsNullOrWhiteSpace([string]$item.itemName)
        ) {
            continue
        }

        # ─────────────────────────────────────────────────────
        # Basic item values
        # ─────────────────────────────────────────────────────

        $mc = if (
            $item.mc -and
            -not [string]::IsNullOrWhiteSpace([string]$item.mc)
        ) {
            [string]$item.mc
        }
        else {
            $defaultMC
        }

        $qty = if ($null -ne $item.qty) {
            [double]$item.qty
        }
        else {
            0.0
        }

        $frontendNetPrice = if ($null -ne $item.price) {
            [double]$item.price
        }
        else {
            0.0
        }

        $frontendAmount = if ($null -ne $item.amount) {
            [double]$item.amount
        }
        else {
            0.0
        }

        $listPrice = if (
            $null -ne $item.listPrice -and
            [double]$item.listPrice -gt 0
        ) {
            [double]$item.listPrice
        }
        else {
            $frontendNetPrice
        }

        $discType = if ($item.discType) {
            [string]$item.discType
        }
        else {
            "U"
        }

        $discVal = if ($null -ne $item.discVal) {
            [double]$item.discVal
        }
        else {
            0.0
        }

        # ─────────────────────────────────────────────────────
        # Unit conversion values
        # ─────────────────────────────────────────────────────

        $conFactor = if (
            $null -ne $item.conFactor -and
            [double]$item.conFactor -ne 0
        ) {
            [double]$item.conFactor
        }
        else {
            1.0
        }

        $altQtyConFactor = if (
            $null -ne $item.altQtyConFactor -and
            [double]$item.altQtyConFactor -ne 0
        ) {
            [double]$item.altQtyConFactor
        }
        else {
            $conFactor
        }

        $conType = if ($item.conType) {
            [int]$item.conType
        }
        else {
            1
        }

        $hasAltUnit =
            $item.altUnit -and
            -not [string]::IsNullOrWhiteSpace([string]$item.altUnit)

        $enteredInAltUnit =
            ($item.enteredInAltUnit -eq $true) -or
            (
                $hasAltUnit -and
                [string]$item.unit -eq [string]$item.altUnit
            )

        if ($enteredInAltUnit) {
            $qtyAlt = $qty

            $qtyMain = if ($altQtyConFactor -ne 0) {
                [Math]::Round(
                    $qtyAlt / $altQtyConFactor,
                    4
                )
            }
            else {
                $qty
            }

            $qtyOut = $qtyMain
        }
        else {
            $qtyMain = $qty

            $qtyAlt = if ($hasAltUnit) {
                [Math]::Round(
                    $qtyMain * $altQtyConFactor,
                    4
                )
            }
            else {
                $qtyMain
            }

            $qtyOut = $qtyMain
        }

        # ─────────────────────────────────────────────────────
        # Discount values
        #
        # BUSY Discount = discount amount per unit.
        # ─────────────────────────────────────────────────────

        $discountPerUnit = 0.0
        $discountPercent = 0.0
        $compoundDiscount = 0.0

        if ($discVal -gt 0) {
            if ($discType -eq "%") {
                $discountPercent = [Math]::Round(
                    $discVal,
                    4
                )

                $discountPerUnit = [Math]::Round(
                    $listPrice * $discountPercent / 100,
                    4
                )

                $compoundDiscount = $discountPercent
            }
            else {
                # Fixed discount per unit.
                $discountPerUnit = [Math]::Round(
                    $discVal,
                    4
                )
            }
        }

        # ─────────────────────────────────────────────────────
        # Final net price
        #
        # Prefer the value already calculated by the frontend.
        # Only calculate it here when frontend price is missing.
        # This prevents discount from being applied twice.
        # ─────────────────────────────────────────────────────

        if ($frontendNetPrice -gt 0) {
            $netPrice = [Math]::Round(
                $frontendNetPrice,
                4
            )
        }
        elseif ($discVal -gt 0) {
            $netPrice = [Math]::Round(
                $listPrice - $discountPerUnit,
                4
            )
        }
        else {
            $netPrice = [Math]::Round(
                $listPrice,
                4
            )
        }

        if ($netPrice -lt 0) {
            $netPrice = 0
        }

        # ─────────────────────────────────────────────────────
        # Final net amount
        # ─────────────────────────────────────────────────────

        if ($frontendAmount -gt 0) {
            $netAmount = [Math]::Round(
                $frontendAmount,
                2
            )
        }
        else {
            $netAmount = [Math]::Round(
                $qtyMain * $netPrice,
                2
            )
        }

        # ─────────────────────────────────────────────────────
        # Alternative-unit price
        # ─────────────────────────────────────────────────────

        $frontendAltPrice = if (
            $null -ne $item.altPrice -and
            [double]$item.altPrice -gt 0
        ) {
            [double]$item.altPrice
        }
        else {
            0.0
        }

        if ($hasAltUnit) {
            if ($enteredInAltUnit -and $frontendAltPrice -gt 0) {
                # The frontend altPrice is already the final alt-unit price.
                $priceAltUnit = [Math]::Round(
                    $frontendAltPrice,
                    4
                )
            }
            elseif ($qtyAlt -ne 0) {
                $priceAltUnit = [Math]::Round(
                    $netAmount / $qtyAlt,
                    4
                )
            }
            else {
                $priceAltUnit = $netPrice
            }
        }
        else {
            $priceAltUnit = $netPrice
        }

        # ─────────────────────────────────────────────────────
        # XML-safe formatted values
        # ─────────────────────────────────────────────────────

        $qtyOutText = Format-XmlNumber $qtyOut
        $qtyMainText = Format-XmlNumber $qtyMain
        $qtyAltText = Format-XmlNumber $qtyAlt

        $conFactorText = Format-XmlNumber $conFactor
        $altQtyConFactorText = Format-XmlNumber $altQtyConFactor

        $netPriceText = Format-XmlNumber $netPrice
        $priceAltUnitText = Format-XmlNumber $priceAltUnit
        $listPriceText = Format-XmlNumber $listPrice
        $netAmountText = Format-XmlNumber $netAmount "0.##"

        $discountPerUnitText = Format-XmlNumber $discountPerUnit
        $discountPercentText = Format-XmlNumber $discountPercent
        $compoundDiscountText = Format-XmlNumber $compoundDiscount "0.00"

        # ─────────────────────────────────────────────────────
        # Build ItemDetail XML
        # ─────────────────────────────────────────────────────

        $xml += "<ItemDetail>"

        $xml += "<SrNo>$($item.srNo)</SrNo>"

        $xml += "<ItemName>$(
            [System.Security.SecurityElement]::Escape(
                [string]$item.itemName
            )
        )</ItemName>"

        if (-not [string]::IsNullOrWhiteSpace($salesmanName)) {
            $xml += "<Broker>$(
                [System.Security.SecurityElement]::Escape(
                    [string]$salesmanName
                )
            )</Broker>"
            $xml += "<BrokerInvolved>True</BrokerInvolved>"
        }

        if ($item.itemType) {
            $xml += "<ItemType>$($item.itemType)</ItemType>"
        }

        $xml += "<UnitName>$(
            [System.Security.SecurityElement]::Escape(
                [string]$item.unit
            )
        )</UnitName>"

        if ($hasAltUnit) {
            $xml += "<AltUnitName>$(
                [System.Security.SecurityElement]::Escape(
                    [string]$item.altUnit
                )
            )</AltUnitName>"

            $xml += "<ConFactor>$conFactorText</ConFactor>"
        }

        $xml += "<Qty>$qtyOutText</Qty>"
        $xml += "<QtyMainUnit>$qtyMainText</QtyMainUnit>"
        $xml += "<QtyAltUnit>$qtyAltText</QtyAltUnit>"

        if ($hasAltUnit) {
            $xml += "<AltQtyConFactor>$altQtyConFactorText</AltQtyConFactor>"
            $xml += "<ConFactorType>$conType</ConFactorType>"
        }

        # Manual BUSY XML stores the final net price here.
        $xml += "<Price>$netPriceText</Price>"

        if ($hasAltUnit) {
            $xml += "<PriceAltUnit>$priceAltUnitText</PriceAltUnit>"
        }

        # Original price before discount.
        $xml += "<ListPrice>$listPriceText</ListPrice>"

        # Manual BUSY XML stores the final net amount in both fields.
        $xml += "<Amt>$netAmountText</Amt>"
        $xml += "<NettAmount>$netAmountText</NettAmount>"

        # Native BUSY discount information.
        $xml += "<Discount>$discountPerUnitText</Discount>"

        if (
            $discVal -gt 0 -and
            $discType -eq "%"
        ) {
            $xml += "<DiscountPercent>$discountPercentText</DiscountPercent>"
        }

        $xml += "<CompoundDiscount>$compoundDiscountText</CompoundDiscount>"

        # Original list/MRP price.
        $xml += "<ItemMRP>$listPriceText</ItemMRP>"

        # Final effective net price.
        $xml += "<tmpNettPrice>$netPriceText</tmpNettPrice>"

        $xml += "<MC>$(
            [System.Security.SecurityElement]::Escape(
                [string]$mc
            )
        )</MC>"

        $xml += "<tmpDiscountBasis>1</tmpDiscountBasis>"

        if ($discVal -gt 0) {
            $xml += "<DiscountStructure>Simple Discount % of Amount</DiscountStructure>"
        }

        $xml += "<tmpNettPriceAfterDisc>$netPriceText</tmpNettPriceAfterDisc>"

        $xml += "</ItemDetail>"
    }

    $xml += "</ItemEntries>"

    return $xml
}

function Build-ConsumedItemsXml {
    param(
        $items,
        [string]$defaultMC = "Main Store"
    )

    # IMPORTANT:
    # BUSY Stock Journal consumed section is ItemEntries1, not ConsumedItemEntries.
    $xml = "<ItemEntries1>"

    foreach ($item in @($items)) {
        if (-not $item.itemName -or [string]::IsNullOrWhiteSpace([string]$item.itemName)) {
            continue
        }

        $mc = if ($item.mc -and -not [string]::IsNullOrWhiteSpace([string]$item.mc)) {
            [string]$item.mc
        } else {
            $defaultMC
        }

        $conFactor = 1.0
        try {
            if ($item.conFactor -and [double]$item.conFactor -ne 0) {
                $conFactor = [double]$item.conFactor
            }
        } catch {}

        $altQtyConFactor = $conFactor
        try {
            if ($item.altQtyConFactor -and [double]$item.altQtyConFactor -ne 0) {
                $altQtyConFactor = [double]$item.altQtyConFactor
            }
        } catch {}

        $conType = 1
        try {
            if ($item.conType) {
                $conType = [int]$item.conType
            }
        } catch {}

        $qty = 0.0
        $price = 0.0
        $amt = 0.0
        $altPrice = 0.0

        try { $qty = [double]$item.qty } catch {}
        try { $price = [double]$item.price } catch {}
        try { $amt = [double]$item.amount } catch {}
        try {
            if ($item.altPrice) {
                $altPrice = [double]$item.altPrice
            }
        } catch {}

        $enteredInAltUnit = $false
        try {
            $enteredInAltUnit =
                ($item.enteredInAltUnit -eq $true) -or
                ($item.altUnit -and [string]$item.unit -eq [string]$item.altUnit)
        } catch {}

        if ($enteredInAltUnit) {
            $qtyAlt = $qty
            $qtyMain = if ($altQtyConFactor -ne 0) {
                [Math]::Round($qtyAlt / $altQtyConFactor, 3)
            } else {
                $qty
            }
            $qtyOut = $qtyMain
        } else {
            $qtyMain = $qty
            $qtyAlt = [Math]::Round($qty * $altQtyConFactor, 3)
            $qtyOut = $qtyMain
        }

        $sendAltPrice = $false

        if ($enteredInAltUnit -and $conType -eq 1) {
            if ($altPrice -ne 0) {
                $amt = [Math]::Round($qtyAlt * $altPrice, 2)
                $price = if ($qtyMain -ne 0) {
                    [Math]::Round($amt / $qtyMain, 2)
                } else {
                    0
                }
            }
            $sendAltPrice = $true
        }
        elseif (-not $enteredInAltUnit -and $conType -eq 1) {
            if ($price -ne 0) {
                $amt = [Math]::Round($qtyMain * $price, 2)
            }
            $altPrice = 0
            $sendAltPrice = $false
        }
        elseif ($conType -eq 2) {
            if ($price -ne 0) {
                $amt = [Math]::Round($qtyOut * $price, 2)
            }
            $sendAltPrice = ($altPrice -ne 0)
        }

        $listPrice = $price
        $discType = "U"
        $discVal = 0.0

        try {
            if ($item.listPrice) {
                $listPrice = [double]$item.listPrice
            }
        } catch {}

        try {
            if ($item.discType) {
                $discType = [string]$item.discType
            }
        } catch {}

        try {
            if ($item.discVal) {
                $discVal = [double]$item.discVal
            }
        } catch {}

        $srNo = 1
        try {
            if ($item.srNo) {
                $srNo = [int]$item.srNo
            }
        } catch {}

        $xml += "<ItemDetail>"
        $xml += "<SrNo>$srNo</SrNo>"
        $xml += "<ItemName>$([System.Security.SecurityElement]::Escape([string]$item.itemName))</ItemName>"
        $xml += "<ItemType>2</ItemType>"
        $xml += "<UnitName>$([System.Security.SecurityElement]::Escape([string]$item.unit))</UnitName>"
        $xml += "<Qty>$qtyOut</Qty>"
        $xml += "<QtyMainUnit>$qtyMain</QtyMainUnit>"

        if ($item.altUnit -and [string]$item.altUnit -ne "") {
            $xml += "<AltUnitName>$([System.Security.SecurityElement]::Escape([string]$item.altUnit))</AltUnitName>"
            $xml += "<ConFactor>$conFactor</ConFactor>"
            $xml += "<AltQtyConFactor>$altQtyConFactor</AltQtyConFactor>"
            $xml += "<ConFactorType>$conType</ConFactorType>"
            $xml += "<QtyAltUnit>$qtyAlt</QtyAltUnit>"

            if ($sendAltPrice) {
                $xml += "<PriceAltUnit>$altPrice</PriceAltUnit>"
            }
        } else {
            $xml += "<QtyAltUnit>$qtyOut</QtyAltUnit>"
        }

        $xml += "<Price>$price</Price>"
        $xml += "<Amt>$amt</Amt>"
        $xml += "<MC>$([System.Security.SecurityElement]::Escape($mc))</MC>"

        if ($discVal -gt 0) {
            $xml += "<ListPrice>$listPrice</ListPrice>"

            if ($discType -eq "%") {
                $xml += "<DiscountPercent>$discVal</DiscountPercent>"
                $discAmt = [Math]::Round($listPrice * $discVal / 100, 2)
                $xml += "<Discount>$discAmt</Discount>"
            } else {
                $xml += "<Discount>$discVal</Discount>"
            }

            $xml += "<DiscountStructure>Simple Discount, % of Price</DiscountStructure>"
        }

        $xml += "</ItemDetail>"
    }

    $xml += "</ItemEntries1>"

    return $xml
}

