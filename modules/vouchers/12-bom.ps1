# 12-bom.ps1
# Extracted from vouchers.ps1. Keep functions behavior-compatible with the monolith.

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

        $safeSearch = ""
        if ($Search) {
            $safeSearch = $Search.Trim() -replace "'", "''"
        }

        $where = "MasterType = 15"

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

        $result = @()

        if ($rst -and -not $rst.EOF) {

            $rst.MoveFirst()

            while (-not $rst.EOF) {

                $code = 0
                $name = ""
                $alias = ""

                $mainItemCode = 0
                $unitCode = 0
                $generatedMCCode = 0
                $consumedMCCode = 0

                $baseQty = 0.0
                $extraExpense = 0.0

                try {
                    $code = [int]$rst.Fields.Item("Code").Value
                } catch {}

                try {
                    $name = $rst.Fields.Item("Name").Value.ToString().Trim()
                } catch {}

                try {
                    $aliasRaw = $rst.Fields.Item("Alias").Value

                    if (
                        $null -ne $aliasRaw -and
                        $aliasRaw -ne [System.DBNull]::Value
                    ) {
                        $alias = $aliasRaw.ToString().Trim()
                    }
                } catch {}

                try {
                    $mainItemCode = [int]$rst.Fields.Item("CM1").Value
                } catch {}

                try {
                    $unitCode = [int]$rst.Fields.Item("CM2").Value
                } catch {}

                try {
                    $generatedMCCode = [int]$rst.Fields.Item("CM3").Value
                } catch {}

                try {
                    $consumedMCCode = [int]$rst.Fields.Item("CM4").Value
                } catch {}

                try {
                    $baseQty = [double]$rst.Fields.Item("D1").Value
                } catch {}

                try {
                    $extraExpense = [double]$rst.Fields.Item("D2").Value
                } catch {}

                $result += @{
                    code             = $code
                    name             = $name
                    alias            = $alias
                    mainItemCode     = $mainItemCode
                    unitCode         = $unitCode
                    generatedMCCode  = $generatedMCCode
                    consumedMCCode   = $consumedMCCode
                    baseQty          = $baseQty
                    extraExpense     = $extraExpense
                }

                $rst.MoveNext()
            }

            try {
                $rst.Close()
            } catch {}
        }

        return @{
            success = $true
            count   = $result.Count
            data    = @($result)
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

        # -------------------------------------------------------
        # Helper: get one Master1 record by Code
        # -------------------------------------------------------
        function Get-MasterByCode {
            param([int]$Code)

            if ($Code -le 0) {
                return $null
            }

            $rst = $fi.GetRecordset(
                "SELECT Code, MasterType, Name, Alias FROM Master1 WHERE Code=$Code"
            )

            if (-not $rst -or $rst.EOF) {
                return $null
            }

            $name = ""
            $alias = ""
            $masterType = 0

            try {
                $name = $rst.Fields.Item("Name").Value.ToString().Trim()
            } catch {}

            try {
                $masterType = [int]$rst.Fields.Item("MasterType").Value
            } catch {}

            try {
                $aliasRaw = $rst.Fields.Item("Alias").Value

                if (
                    $null -ne $aliasRaw -and
                    $aliasRaw -ne [System.DBNull]::Value
                ) {
                    $alias = $aliasRaw.ToString().Trim()
                }
            } catch {}

            try {
                $rst.Close()
            } catch {}

            return @{
                code       = $Code
                name       = $name
                alias      = $alias
                masterType = $masterType
            }
        }


        # -------------------------------------------------------
        # 1. Resolve BOM header
        # -------------------------------------------------------
        $headerSql = ""

        if ($BomCode -gt 0) {

            $headerSql = @"
SELECT *
FROM Master1
WHERE MasterType=15
AND Code=$BomCode
"@
        }
        else {

            $safeBomName = $BomName.Trim() -replace "'", "''"

            $headerSql = @"
SELECT *
FROM Master1
WHERE MasterType=15
AND Name='$safeBomName'
"@
        }

        $bomRst = $fi.GetRecordset($headerSql)

        if (-not $bomRst -or $bomRst.EOF) {

            return @{
                success = $false
                error   = "BOM not found"
            }
        }


        # -------------------------------------------------------
        # 2. Read BOM header
        # -------------------------------------------------------
        $resolvedBomCode = 0
        $resolvedBomName = ""
        $bomAlias = ""

        $mainItemCode = 0
        $unitCode = 0

        $generatedMCCode = 0
        $consumedMCCode = 0

        $baseQty = 0.0
        $extraExpense = 0.0

        try {
            $resolvedBomCode =
                [int]$bomRst.Fields.Item("Code").Value
        } catch {}

        try {
            $resolvedBomName =
                $bomRst.Fields.Item("Name").Value.ToString().Trim()
        } catch {}

        try {
            $aliasRaw = $bomRst.Fields.Item("Alias").Value

            if (
                $null -ne $aliasRaw -and
                $aliasRaw -ne [System.DBNull]::Value
            ) {
                $bomAlias = $aliasRaw.ToString().Trim()
            }
        } catch {}

        try {
            $mainItemCode =
                [int]$bomRst.Fields.Item("CM1").Value
        } catch {}

        try {
            $unitCode =
                [int]$bomRst.Fields.Item("CM2").Value
        } catch {}

        try {
            $generatedMCCode =
                [int]$bomRst.Fields.Item("CM3").Value
        } catch {}

        try {
            $consumedMCCode =
                [int]$bomRst.Fields.Item("CM4").Value
        } catch {}

        try {
            $baseQty =
                [double]$bomRst.Fields.Item("D1").Value
        } catch {}

        try {
            $extraExpense =
                [double]$bomRst.Fields.Item("D2").Value
        } catch {}

        try {
            $bomRst.Close()
        } catch {}


        # -------------------------------------------------------
        # 3. Resolve header master names
        # -------------------------------------------------------
        $mainItemMaster =
            Get-MasterByCode -Code $mainItemCode

        $unitMaster =
            Get-MasterByCode -Code $unitCode

        $generatedMCMaster =
            Get-MasterByCode -Code $generatedMCCode

        $consumedMCMaster =
            Get-MasterByCode -Code $consumedMCCode


        # -------------------------------------------------------
        # 4. Read BOM components from MasterSupport
        #
        # I1 = 2 -> Raw Material Consumed
        # I1 = 3 -> By-product Generated
        #
        # I2 = sequence within its section
        # CM1 = item code
        # CM2 = unit code
        # D1  = quantity
        # -------------------------------------------------------
        $supportSql = @"
SELECT
    CM1,
    CM2,
    D1,
    I1,
    I2
FROM MasterSupport
WHERE MasterCode=$resolvedBomCode
AND MasterType=15
ORDER BY I1, I2
"@

        $supportRst =
            $fi.GetRecordset($supportSql)

        $rawMaterials = @()
        $byProducts = @()

        if (
            $supportRst -and
            -not $supportRst.EOF
        ) {

            $supportRst.MoveFirst()

            while (-not $supportRst.EOF) {

                $componentItemCode = 0
                $componentUnitCode = 0

                $componentQty = 0.0

                $componentType = 0
                $componentSequence = 0

                try {
                    $componentItemCode =
                        [int]$supportRst.Fields.Item("CM1").Value
                } catch {}

                try {
                    $componentUnitCode =
                        [int]$supportRst.Fields.Item("CM2").Value
                } catch {}

                try {
                    $componentQty =
                        [double]$supportRst.Fields.Item("D1").Value
                } catch {}

                try {
                    $componentType =
                        [int]$supportRst.Fields.Item("I1").Value
                } catch {}

                try {
                    $componentSequence =
                        [int]$supportRst.Fields.Item("I2").Value
                } catch {}

                $itemMaster =
                    Get-MasterByCode -Code $componentItemCode

                $componentUnit =
                    Get-MasterByCode -Code $componentUnitCode

                $component = @{
                    code      = $componentItemCode
                    name      = if ($itemMaster) {
                        $itemMaster.name
                    } else {
                        ""
                    }

                    alias     = if ($itemMaster) {
                        $itemMaster.alias
                    } else {
                        ""
                    }

                    unitCode  = $componentUnitCode

                    unit      = if ($componentUnit) {
                        $componentUnit.name
                    } else {
                        ""
                    }

                    qty       = $componentQty
                    sequence  = $componentSequence
                }


                if ($componentType -eq 2) {

                    $rawMaterials += $component
                }
                elseif ($componentType -eq 3) {

                    $byProducts += $component
                }

                $supportRst.MoveNext()
            }

            try {
                $supportRst.Close()
            } catch {}
        }


        # -------------------------------------------------------
        # 5. Return clean BOM object
        # -------------------------------------------------------
        return @{
            success = $true

            data = @{

                code = $resolvedBomCode
                name = $resolvedBomName
                alias = $bomAlias

                baseQty = $baseQty
                extraExpensePerUnit = $extraExpense

                mainItem = @{
                    code = $mainItemCode

                    name = if ($mainItemMaster) {
                        $mainItemMaster.name
                    } else {
                        ""
                    }

                    alias = if ($mainItemMaster) {
                        $mainItemMaster.alias
                    } else {
                        ""
                    }

                    qty = $baseQty

                    unitCode = $unitCode

                    unit = if ($unitMaster) {
                        $unitMaster.name
                    } else {
                        ""
                    }
                }

                generatedMaterialCentre = @{
                    code = $generatedMCCode

                    name = if ($generatedMCMaster) {
                        $generatedMCMaster.name
                    } else {
                        ""
                    }
                }

                consumedMaterialCentre = @{
                    code = $consumedMCCode

                    name = if ($consumedMCMaster) {
                        $consumedMCMaster.name
                    } else {
                        ""
                    }
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


# ═══════════════════════════════════════════════════════════════
#  MODIFY VOUCHER
# ═══════════════════════════════════════════════════════════════

