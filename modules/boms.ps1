# modules/boms.ps1
# BUSY 21 - Bill of Materials Master
#
# MasterType = 15
#
# BUSY BOM XML structure verified from the supplied BUSY export:
#   <BOM>
#     <Name>...</Name>
#     <Alias>...</Alias>
#     <Item>...</Item>
#     <Unit>...</Unit>
#     <DefaultMCName1>...</DefaultMCName1>
#     <DefaultMCName2>...</DefaultMCName2>
#     <ExtraExpensePerUnit>...</ExtraExpensePerUnit>
#     <Qty>...</Qty>
#     <RMDetails>...</RMDetails>
#     <BPDetails>...</BPDetails>
#   </BOM>
#
# MasterSupport mapping verified from the company database:
#   MasterCode = BOM code
#   CM1        = component item code
#   CM2        = unit code
#   D1         = quantity
#   I1 = 2     = Raw Material Consumed
#   I1 = 3     = By-Product Generated
#   I2         = sequence within section

function ConvertTo-BomXmlSafe {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return ""
    }

    return [System.Security.SecurityElement]::Escape(
        [string]$Value
    )
}

function Get-BomNumber {
    param(
        [AllowNull()][object]$Value,
        [double]$Default = 0
    )

    if (
        $null -eq $Value -or
        $Value -eq [System.DBNull]::Value -or
        [string]::IsNullOrWhiteSpace([string]$Value)
    ) {
        return $Default
    }

    $number = 0.0

    if (
        [double]::TryParse(
            [string]$Value,
            [System.Globalization.NumberStyles]::Any,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [ref]$number
        )
    ) {
        return $number
    }

    if ([double]::TryParse([string]$Value, [ref]$number)) {
        return $number
    }

    return $Default
}

function ConvertTo-BomInvariantNumber {
    param(
        [AllowNull()][object]$Value,
        [double]$Default = 0
    )

    $number = Get-BomNumber -Value $Value -Default $Default

    return $number.ToString(
        "0.########",
        [System.Globalization.CultureInfo]::InvariantCulture
    )
}

function Get-BomMasterNameByCode {
    param(
        $Fi,
        [int]$Code
    )

    if ($Code -le 0) {
        return ""
    }

    $rst = $null

    try {
        $rst = $Fi.GetRecordset(
            "SELECT Name FROM Master1 WHERE Code=$Code"
        )

        if ($rst -and -not $rst.EOF) {
            $value = $rst.Fields.Item("Name").Value

            if (
                $null -ne $value -and
                $value -ne [System.DBNull]::Value
            ) {
                return ([string]$value).Trim()
            }
        }
    }
    catch {
        return ""
    }
    finally {
        if ($rst) {
            try { $rst.Close() } catch {}
        }
    }

    return ""
}

function Get-BomMasterCode {
    param(
        $Fi,
        [string]$Name,
        [int]$MasterType
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return 0
    }

    try {
        return [int]$Fi.MasterName2Code(
            $Name.Trim(),
            $MasterType
        )
    }
    catch {
        return 0
    }
}

function Test-BomPayload {
    param(
        $Data,
        $Fi,
        [switch]$IsUpdate
    )

    if (-not $Data) {
        return @{
            success = $false
            error   = "BOM data is required"
        }
    }

    $name = ([string]$Data.name).Trim()
    $alias = ([string]$Data.alias).Trim()
    $itemName = ([string]$Data.item).Trim()
    $unitName = ([string]$Data.unit).Trim()
    $generatedMC = ([string]$Data.generatedMaterialCentre).Trim()
    $consumedMC = ([string]$Data.consumedMaterialCentre).Trim()

    # Also accept API aliases used by the production-voucher side.
    if ([string]::IsNullOrWhiteSpace($itemName)) {
        $itemName = ([string]$Data.mainItemName).Trim()
    }

    if ([string]::IsNullOrWhiteSpace($unitName)) {
        $unitName = ([string]$Data.mainUnit).Trim()
    }

    if ([string]::IsNullOrWhiteSpace($generatedMC)) {
        $generatedMC = ([string]$Data.defaultMCName1).Trim()
    }

    if ([string]::IsNullOrWhiteSpace($consumedMC)) {
        $consumedMC = ([string]$Data.defaultMCName2).Trim()
    }

    if ([string]::IsNullOrWhiteSpace($name)) {
        return @{ success = $false; error = "BOM name is required" }
    }

    if ([string]::IsNullOrWhiteSpace($itemName)) {
        return @{ success = $false; error = "Item to Produce is required" }
    }

    if ([string]::IsNullOrWhiteSpace($unitName)) {
        return @{ success = $false; error = "Unit is required" }
    }

    $qty = Get-BomNumber -Value $Data.qty -Default 0

    if ($qty -le 0) {
        return @{ success = $false; error = "Quantity must be greater than zero" }
    }

    $mainItemCode = Get-BomMasterCode `
        -Fi $Fi `
        -Name $itemName `
        -MasterType 6

    if ($mainItemCode -le 0) {
        return @{
            success = $false
            error   = "Invalid Item to Produce. Select a valid Item Master."
        }
    }

    $unitCode = Get-BomMasterCode `
        -Fi $Fi `
        -Name $unitName `
        -MasterType 8

    if ($unitCode -le 0) {
        return @{
            success = $false
            error   = "Invalid unit. Select a valid Unit Master."
        }
    }

    $generatedMCCode = 0

    if (-not [string]::IsNullOrWhiteSpace($generatedMC)) {
        $generatedMCCode = Get-BomMasterCode `
            -Fi $Fi `
            -Name $generatedMC `
            -MasterType 11

        if ($generatedMCCode -le 0) {
            return @{
                success = $false
                error   = "Invalid Generated Material Centre."
            }
        }
    }

    $consumedMCCode = 0

    if (-not [string]::IsNullOrWhiteSpace($consumedMC)) {
        $consumedMCCode = Get-BomMasterCode `
            -Fi $Fi `
            -Name $consumedMC `
            -MasterType 11

        if ($consumedMCCode -le 0) {
            return @{
                success = $false
                error   = "Invalid Consumed Material Centre."
            }
        }
    }

    $rawMaterials = @()

    if ($null -ne $Data.rawMaterials) {
        $rawMaterials = @($Data.rawMaterials)
    }
    elseif ($null -ne $Data.rmDetails) {
        $rawMaterials = @($Data.rmDetails)
    }

    $byProducts = @()

    if ($null -ne $Data.byProducts) {
        $byProducts = @($Data.byProducts)
    }
    elseif ($null -ne $Data.bpDetails) {
        $byProducts = @($Data.bpDetails)
    }

    $validateRows = {
        param(
            $Rows,
            [string]$SectionLabel
        )

        $cleanRows = @()
        $sequence = 1

        foreach ($row in @($Rows)) {
            $rowItem = ([string]$row.item).Trim()

            if ([string]::IsNullOrWhiteSpace($rowItem)) {
                $rowItem = ([string]$row.itemName).Trim()
            }

            # Completely blank UI rows are ignored.
            if ([string]::IsNullOrWhiteSpace($rowItem)) {
                continue
            }

            $rowQty = Get-BomNumber -Value $row.qty -Default 0

            if ($rowQty -le 0) {
                return @{
                    success = $false
                    error   = "$SectionLabel row $sequence quantity must be greater than zero"
                }
            }

            $rowUnit = ([string]$row.unit).Trim()

            if ([string]::IsNullOrWhiteSpace($rowUnit)) {
                return @{
                    success = $false
                    error   = "$SectionLabel row $sequence unit is required"
                }
            }

            $itemCode = Get-BomMasterCode `
                -Fi $Fi `
                -Name $rowItem `
                -MasterType 6

            if ($itemCode -le 0) {
                return @{
                    success = $false
                    error   = "$SectionLabel row $sequence has an invalid Item Master: $rowItem"
                }
            }

            $rowUnitCode = Get-BomMasterCode `
                -Fi $Fi `
                -Name $rowUnit `
                -MasterType 8

            if ($rowUnitCode -le 0) {
                return @{
                    success = $false
                    error   = "$SectionLabel row $sequence has an invalid Unit Master: $rowUnit"
                }
            }

            $cleanRows += @{
                item     = $rowItem
                itemCode = $itemCode
                qty      = $rowQty
                unit     = $rowUnit
                unitCode = $rowUnitCode
                srNo     = $sequence
            }

            $sequence++
        }

        return @{
            success = $true
            rows    = @($cleanRows)
        }
    }

    $rmValidation = & $validateRows $rawMaterials "Raw Material"

    if ($rmValidation.success -eq $false) {
        return $rmValidation
    }

    $bpValidation = & $validateRows $byProducts "By-Product"

    if ($bpValidation.success -eq $false) {
        return $bpValidation
    }

    return @{
        success = $true

        data = @{
            name                   = $name
            alias                  = $alias
            item                   = $itemName
            itemCode               = $mainItemCode
            unit                   = $unitName
            unitCode               = $unitCode
            qty                    = $qty
            extraExpensePerUnit    = Get-BomNumber -Value $Data.extraExpensePerUnit -Default 0
            generatedMaterialCentre = $generatedMC
            generatedMCCode        = $generatedMCCode
            consumedMaterialCentre = $consumedMC
            consumedMCCode         = $consumedMCCode
            rawMaterials           = @($rmValidation.rows)
            byProducts             = @($bpValidation.rows)
        }
    }
}

function Build-BomXml {
    param(
        $Data,
        [string]$OriginalName = ""
    )

    $nameForSave =
        if (-not [string]::IsNullOrWhiteSpace($OriginalName)) {
            $OriginalName.Trim()
        }
        else {
            ([string]$Data.name).Trim()
        }

    $itemName = ([string]$Data.item).Trim()

    if ([string]::IsNullOrWhiteSpace($itemName)) {
        $itemName = ([string]$Data.mainItemName).Trim()
    }

    $unitName = ([string]$Data.unit).Trim()

    if ([string]::IsNullOrWhiteSpace($unitName)) {
        $unitName = ([string]$Data.mainUnit).Trim()
    }

    $generatedMC = ([string]$Data.generatedMaterialCentre).Trim()

    if ([string]::IsNullOrWhiteSpace($generatedMC)) {
        $generatedMC = ([string]$Data.defaultMCName1).Trim()
    }

    $consumedMC = ([string]$Data.consumedMaterialCentre).Trim()

    if ([string]::IsNullOrWhiteSpace($consumedMC)) {
        $consumedMC = ([string]$Data.defaultMCName2).Trim()
    }

    $rawMaterials =
        if ($null -ne $Data.rawMaterials) {
            @($Data.rawMaterials)
        }
        elseif ($null -ne $Data.rmDetails) {
            @($Data.rmDetails)
        }
        else {
            @()
        }

    $byProducts =
        if ($null -ne $Data.byProducts) {
            @($Data.byProducts)
        }
        elseif ($null -ne $Data.bpDetails) {
            @($Data.bpDetails)
        }
        else {
            @()
        }

    $xml = "<BOM>"

    $xml += "<Name>$(ConvertTo-BomXmlSafe $nameForSave)</Name>"

    if (-not [string]::IsNullOrWhiteSpace([string]$Data.alias)) {
        $xml += "<Alias>$(ConvertTo-BomXmlSafe $Data.alias)</Alias>"
    }

    $xml += "<Item>$(ConvertTo-BomXmlSafe $itemName)</Item>"
    $xml += "<Unit>$(ConvertTo-BomXmlSafe $unitName)</Unit>"

    if (-not [string]::IsNullOrWhiteSpace($generatedMC)) {
        $xml += "<DefaultMCName1>$(ConvertTo-BomXmlSafe $generatedMC)</DefaultMCName1>"
    }

    # BUSY's own export proves this field can be absent.
    if (-not [string]::IsNullOrWhiteSpace($consumedMC)) {
        $xml += "<DefaultMCName2>$(ConvertTo-BomXmlSafe $consumedMC)</DefaultMCName2>"
    }

    $xml += "<ExtraExpensePerUnit>$(ConvertTo-BomInvariantNumber $Data.extraExpensePerUnit)</ExtraExpensePerUnit>"
    $xml += "<Qty>$(ConvertTo-BomInvariantNumber $Data.qty)</Qty>"

    $xml += "<RMDetails>"

    $rmSrNo = 1

    foreach ($row in $rawMaterials) {
        $rowItem = ([string]$row.item).Trim()

        if ([string]::IsNullOrWhiteSpace($rowItem)) {
            $rowItem = ([string]$row.itemName).Trim()
        }

        if ([string]::IsNullOrWhiteSpace($rowItem)) {
            continue
        }

        $xml += "<ItemDetails>"
        $xml += "<Item>$(ConvertTo-BomXmlSafe $rowItem)</Item>"
        $xml += "<ItemType>2</ItemType>"
        $xml += "<Qty>$(ConvertTo-BomInvariantNumber $row.qty)</Qty>"
        $xml += "<Unit>$(ConvertTo-BomXmlSafe $row.unit)</Unit>"
        $xml += "<SrNo>$rmSrNo</SrNo>"
        $xml += "</ItemDetails>"

        $rmSrNo++
    }

    $xml += "</RMDetails>"

    $xml += "<BPDetails>"

    $bpSrNo = 1

    foreach ($row in $byProducts) {
        $rowItem = ([string]$row.item).Trim()

        if ([string]::IsNullOrWhiteSpace($rowItem)) {
            $rowItem = ([string]$row.itemName).Trim()
        }

        if ([string]::IsNullOrWhiteSpace($rowItem)) {
            continue
        }

        $xml += "<ItemDetails>"
        $xml += "<Item>$(ConvertTo-BomXmlSafe $rowItem)</Item>"
        $xml += "<ItemType>3</ItemType>"
        $xml += "<Qty>$(ConvertTo-BomInvariantNumber $row.qty)</Qty>"
        $xml += "<Unit>$(ConvertTo-BomXmlSafe $row.unit)</Unit>"
        $xml += "<SrNo>$bpSrNo</SrNo>"
        $xml += "</ItemDetails>"

        $bpSrNo++
    }

    $xml += "</BPDetails>"

    $xml += "</BOM>"

    return $xml
}

function Create-Bom {
    param(
        $Data,
        [string]$InstanceId  = "",
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
        $validation = Test-BomPayload `
            -Data $Data `
            -Fi $fi

        if ($validation.success -eq $false) {
            return $validation
        }

        $existingCode = Get-BomMasterCode `
            -Fi $fi `
            -Name ([string]$Data.name) `
            -MasterType 15

        if ($existingCode -gt 0) {
            return @{
                success = $false
                error   = "A BOM with this name already exists"
            }
        }

        $xml = Build-BomXml -Data $Data

        $err = ""

        $saved = $fi.SaveMasterFromXML(
            15,
            $xml,
            [ref]$err,
            $false
        )

        if ($saved -eq $true) {
            $createdCode = Get-BomMasterCode `
                -Fi $fi `
                -Name ([string]$Data.name) `
                -MasterType 15

            return @{
                success = $true
                message = "BOM created successfully"
                data    = @{
                    code = $createdCode
                    name = ([string]$Data.name).Trim()
                }
            }
        }

        return @{
            success = $false
            error   = if ($err) { $err } else { "BOM save failed" }
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

function Update-Bom {
    param(
        $Data,
        [string]$InstanceId  = "",
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
        $validation = Test-BomPayload `
            -Data $Data `
            -Fi $fi `
            -IsUpdate

        if ($validation.success -eq $false) {
            return $validation
        }

        $originalName =
            if (
                $Data._originalName -and
                -not [string]::IsNullOrWhiteSpace([string]$Data._originalName)
            ) {
                ([string]$Data._originalName).Trim()
            }
            else {
                ([string]$Data.name).Trim()
            }

        $existingCode = Get-BomMasterCode `
            -Fi $fi `
            -Name $originalName `
            -MasterType 15

        if ($existingCode -le 0) {
            return @{
                success = $false
                error   = "BOM not found"
            }
        }

        # Matches the current Account/Item master update pattern:
        # the original master name is placed in <Name> so BUSY can
        # identify the existing master during SaveMasterFromXML(...,$true).
        $xml = Build-BomXml `
            -Data $Data `
            -OriginalName $originalName

        $err = ""

        $saved = $fi.SaveMasterFromXML(
            15,
            $xml,
            [ref]$err,
            $true
        )

        if ($saved -eq $true) {
            return @{
                success = $true
                message = "BOM updated successfully"
                data    = @{
                    code = $existingCode
                    name = $originalName
                }
            }
        }

        return @{
            success = $false
            error   = if ($err) { $err } else { "BOM update failed" }
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

function Get-BomList {
    param(
        [string]$Search = "",
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
        $safeSearch = ([string]$Search).Trim().Replace("'", "''")

        $where = "MasterType=15"

        if (-not [string]::IsNullOrWhiteSpace($safeSearch)) {
            $where += " AND (Name LIKE '*$safeSearch*' OR Alias LIKE '*$safeSearch*')"
        }

        $sql = @"
SELECT
    Code,
    Name,
    Alias,
    CM1,
    CM2,
    CM3,
    CM4,
    D1,
    D2
FROM Master1
WHERE $where
ORDER BY Name
"@

        $rst = $fi.GetRecordset($sql)
        $data = @()

        if ($rst -and -not $rst.EOF) {
            $rst.MoveFirst()

            while (-not $rst.EOF) {
                $code = [int]$rst.Fields.Item("Code").Value
                $mainItemCode = [int]$rst.Fields.Item("CM1").Value
                $unitCode = [int]$rst.Fields.Item("CM2").Value
                $generatedMCCode = [int]$rst.Fields.Item("CM3").Value
                $consumedMCCode = [int]$rst.Fields.Item("CM4").Value

                $aliasValue = $rst.Fields.Item("Alias").Value

                $data += @{
                    code                  = $code
                    name                  = ([string]$rst.Fields.Item("Name").Value).Trim()
                    alias                 = if (
                        $null -ne $aliasValue -and
                        $aliasValue -ne [System.DBNull]::Value
                    ) {
                        ([string]$aliasValue).Trim()
                    } else { "" }

                    mainItemCode          = $mainItemCode
                    mainItemName          = Get-BomMasterNameByCode -Fi $fi -Code $mainItemCode

                    unitCode              = $unitCode
                    unit                  = Get-BomMasterNameByCode -Fi $fi -Code $unitCode

                    generatedMCCode       = $generatedMCCode
                    generatedMaterialCentre = Get-BomMasterNameByCode -Fi $fi -Code $generatedMCCode

                    consumedMCCode        = $consumedMCCode
                    consumedMaterialCentre = Get-BomMasterNameByCode -Fi $fi -Code $consumedMCCode

                    baseQty               = Get-BomNumber -Value $rst.Fields.Item("D1").Value
                    qty                   = Get-BomNumber -Value $rst.Fields.Item("D1").Value
                    extraExpense          = Get-BomNumber -Value $rst.Fields.Item("D2").Value
                    extraExpensePerUnit   = Get-BomNumber -Value $rst.Fields.Item("D2").Value
                }

                $rst.MoveNext()
            }

            try { $rst.Close() } catch {}
        }

        return @{
            success = $true
            count   = $data.Count
            data    = @($data)
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

function Get-BomDetail {
    param(
        [int]$BomCode = 0,
        [string]$BomName = "",
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
        if ($BomCode -le 0) {
            if ([string]::IsNullOrWhiteSpace($BomName)) {
                return @{
                    success = $false
                    error   = "BOM code or BOM name is required"
                }
            }

            $BomCode = Get-BomMasterCode `
                -Fi $fi `
                -Name $BomName `
                -MasterType 15
        }

        if ($BomCode -le 0) {
            return @{
                success = $false
                error   = "BOM not found"
            }
        }

        $headerRst = $fi.GetRecordset(
            "SELECT * FROM Master1 WHERE Code=$BomCode AND MasterType=15"
        )

        if (-not $headerRst -or $headerRst.EOF) {
            return @{
                success = $false
                error   = "BOM not found"
            }
        }

        $name = ([string]$headerRst.Fields.Item("Name").Value).Trim()

        $aliasValue = $headerRst.Fields.Item("Alias").Value
        $alias = ""

        if (
            $null -ne $aliasValue -and
            $aliasValue -ne [System.DBNull]::Value
        ) {
            $alias = ([string]$aliasValue).Trim()
        }

        $mainItemCode = [int]$headerRst.Fields.Item("CM1").Value
        $unitCode = [int]$headerRst.Fields.Item("CM2").Value
        $generatedMCCode = [int]$headerRst.Fields.Item("CM3").Value
        $consumedMCCode = [int]$headerRst.Fields.Item("CM4").Value

        $qty = Get-BomNumber -Value $headerRst.Fields.Item("D1").Value
        $extraExpense = Get-BomNumber -Value $headerRst.Fields.Item("D2").Value

        try { $headerRst.Close() } catch {}

        $rawMaterials = @()
        $byProducts = @()

        $supportRst = $fi.GetRecordset(
            "SELECT CM1, CM2, D1, I1, I2 FROM MasterSupport WHERE MasterCode=$BomCode AND MasterType=15 ORDER BY I1, I2"
        )

        if ($supportRst -and -not $supportRst.EOF) {
            $supportRst.MoveFirst()

            while (-not $supportRst.EOF) {
                $itemCode = [int]$supportRst.Fields.Item("CM1").Value
                $rowUnitCode = [int]$supportRst.Fields.Item("CM2").Value
                $rowQty = Get-BomNumber -Value $supportRst.Fields.Item("D1").Value
                $itemType = [int]$supportRst.Fields.Item("I1").Value
                $sequence = [int]$supportRst.Fields.Item("I2").Value

                $row = @{
                    code      = $itemCode
                    itemCode  = $itemCode
                    name      = Get-BomMasterNameByCode -Fi $fi -Code $itemCode
                    itemName  = Get-BomMasterNameByCode -Fi $fi -Code $itemCode
                    qty       = $rowQty
                    unitCode  = $rowUnitCode
                    unit      = Get-BomMasterNameByCode -Fi $fi -Code $rowUnitCode
                    sequence  = $sequence
                    srNo      = $sequence
                    itemType  = $itemType
                }

                if ($itemType -eq 2) {
                    $rawMaterials += $row
                }
                elseif ($itemType -eq 3) {
                    $byProducts += $row
                }

                $supportRst.MoveNext()
            }

            try { $supportRst.Close() } catch {}
        }

        $mainItemName = Get-BomMasterNameByCode -Fi $fi -Code $mainItemCode
        $unitName = Get-BomMasterNameByCode -Fi $fi -Code $unitCode
        $generatedMCName = Get-BomMasterNameByCode -Fi $fi -Code $generatedMCCode
        $consumedMCName = Get-BomMasterNameByCode -Fi $fi -Code $consumedMCCode

        return @{
            success = $true

            data = @{
                code = $BomCode
                name = $name
                alias = $alias

                item = $mainItemName
                mainItemName = $mainItemName

                qty = $qty
                baseQty = $qty

                unit = $unitName
                unitCode = $unitCode

                extraExpensePerUnit = $extraExpense

                generatedMaterialCentre = @{
                    code = $generatedMCCode
                    name = $generatedMCName
                }

                consumedMaterialCentre = @{
                    code = $consumedMCCode
                    name = $consumedMCName
                }

                defaultMCName1 = $generatedMCName
                defaultMCName2 = $consumedMCName

                mainItem = @{
                    code = $mainItemCode
                    name = $mainItemName
                    qty  = $qty
                    unitCode = $unitCode
                    unit = $unitName
                }

                rawMaterials = @($rawMaterials)
                byProducts = @($byProducts)
            }
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
