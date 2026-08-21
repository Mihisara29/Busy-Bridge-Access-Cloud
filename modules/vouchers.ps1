# modules/vouchers.ps1
# Voucher Management — Generic Engine (Multi-Instance Version)

if ($null -eq $script:Config) {
    . "$PSScriptRoot\connection.ps1"
    . "$PSScriptRoot\utils.ps1"
    . "$PSScriptRoot\config.ps1"
}

# ═══════════════════════════════════════════════════════════════
#  VOUCHER CONFIG TABLE
# ═══════════════════════════════════════════════════════════════
$script:VoucherConfig = @{
    9  = @{ xmlRoot = "Sale";           typeField = "STPTName"; typeDataKey = "saleType";     hasBillNo = $true;  isAccounting = $false; requiredKeys = @("vchSeries","date","party","saleType","items") }
    2  = @{ xmlRoot = "Purchase";       typeField = "STPTName"; typeDataKey = "purchaseType"; hasBillNo = $true;  isAccounting = $false; requiredKeys = @("vchSeries","date","party","purchaseType","items") }
    3  = @{ xmlRoot = "SaleReturn";     typeField = "STPTName"; typeDataKey = "saleType";     hasBillNo = $false; isAccounting = $false; requiredKeys = @("vchSeries","date","party","saleType","items") }
    10 = @{ xmlRoot = "PurchaseReturn"; typeField = "STPTName"; typeDataKey = "purchaseType"; hasBillNo = $false; isAccounting = $false; requiredKeys = @("vchSeries","date","party","purchaseType","items") }
    11 = @{ xmlRoot = "MaterialIssue";   typeField = "STPTName"; typeDataKey = "saleType";     hasBillNo = $false; isAccounting = $false; requiredKeys = @("vchSeries","date","party","saleType","items") }
    4  = @{ xmlRoot = "MaterialReceipt"; typeField = "STPTName"; typeDataKey = "purchaseType"; hasBillNo = $true;  isAccounting = $false; requiredKeys = @("vchSeries","date","party","purchaseType","items") }
    12 = @{ xmlRoot = "SaleOrder";      typeField = "STPTName"; typeDataKey = "saleType";     hasBillNo = $false; isAccounting = $false; requiredKeys = @("vchSeries","date","party","saleType","items") }
    13 = @{ xmlRoot = "PurchaseOrder";  typeField = "STPTName"; typeDataKey = "purchaseType"; hasBillNo = $true;  isAccounting = $false; requiredKeys = @("vchSeries","date","party","purchaseType","items") }
    14 = @{ xmlRoot = "Receipt";        hasBillNo = $false; isAccounting = $true;  requiredKeys = @("vchSeries","date","accounts") }
    15 = @{ xmlRoot = "Contra";         hasBillNo = $false; isAccounting = $true;  requiredKeys = @("vchSeries","date","accounts") }
    16 = @{ xmlRoot = "Journal";        hasBillNo = $false; isAccounting = $true;  requiredKeys = @("vchSeries","date","accounts") }
    19 = @{ xmlRoot = "Payment";        hasBillNo = $false; isAccounting = $true;  requiredKeys = @("vchSeries","date","accounts") }
    5  = @{ xmlRoot = "StockTransfer";  typeField = "VchSeriesName"; typeDataKey = "vchSeries"; hasBillNo = $false; isAccounting = $false; requiredKeys = @("vchSeries","date","party","items") }
    8  = @{ xmlRoot = "StockJournal";   typeField = "VchSeriesName"; typeDataKey = "vchSeries"; hasBillNo = $false; isAccounting = $false; requiredKeys = @("vchSeries","date","party","items") }
    6 = @{xmlRoot = "Production";  typeField = "VchSeriesName";  typeDataKey = "vchSeries";  hasBillNo = $false;  isAccounting = $false;
    requiredKeys = @("vchSeries","date","items")
}
}

function Safe-ParseDate {
    param([string]$dateStr)
    if ([string]::IsNullOrWhiteSpace($dateStr)) { return $null }
    $formats = @("dd-MM-yyyy", "yyyy-MM-dd", "MM/dd/yyyy", "d-M-yyyy", "yyyy/MM/dd")
    foreach ($fmt in $formats) {
        try {
            return [datetime]::ParseExact($dateStr, $fmt, [System.Globalization.CultureInfo]::InvariantCulture)
        } catch {}
    }
    try {
        return [datetime]::Parse($dateStr)
    } catch {}
    return $null
}

function Get-VoucherOptionalFields {
    param(
        [int]$VchType,
        [string]$SeriesName,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) {
        return @{ success = $false; error = "BUSY connection failed" }
    }

    try {
        # 1. Resolve Series Master Code
        $prefixStr = "{0:D2}" -f $VchType
        $prefixedSeriesName = if ($SeriesName.StartsWith($prefixStr)) { $SeriesName } else { "$prefixStr$SeriesName" }
        
        $safeSeriesName = $SeriesName -replace "'", "''"
        $safePrefixedName = $prefixedSeriesName -replace "'", "''"

        $seriesCode = 0
        $sRst = $fi.GetRecordset("SELECT Code FROM Master1 WHERE MasterType=21 AND (Name='$safeSeriesName' OR Name='$safePrefixedName')")
        if ($sRst -and -not $sRst.EOF) {
            $seriesCode = [int]$sRst.Fields.Item("Code").Value
            $sRst.Close()
        }

        if ($seriesCode -eq 0) {
            return @{ success = $true; data = @() }
        }

        # 2. Query Config table
        $qry = "SELECT * FROM Config WHERE RecType=1 AND L1=$seriesCode"
        $rst = $fi.GetRecordset($qry)
        $fields = @()

        if ($rst -and -not $rst.EOF) {
            $rst.MoveFirst()
            
            # Loop up to 20 directly to ensure all configured fields are caught 
            for ($i = 1; $i -le 20; $i++) {
                $colName = "C$i"
                $fName = ""
                try {
                    if ($rst.Fields.Item($colName).Value -ne [System.DBNull]::Value) {
                        $fName = $rst.Fields.Item($colName).Value.ToString().Trim()
                    }
                } catch {}

                if ($fName -ne "") {
                    $typeColOffset = 5 + $i
                    $subColOffset = 25 + $i

                    $rawType = 0
                    $rawSub  = 0

                    try { $rawType = [int]$rst.Fields.Item("I$typeColOffset").Value } catch {}
                    try { $rawSub  = [int]$rst.Fields.Item("I$subColOffset").Value } catch {}

                    # Safe and resilient Type Mapping based on database structures
                    $fieldType = "text"
                    $decimalPlaces = 0
                    $maintainMaster = $false

                    # A. Exhaustive Business Vocabulary matching
                    if ($fName -match "Date|Dated|Expiry|Due|Period|Mfg|Format|Year|Month|Day") {
                        $fieldType = "date"
                    }
                    elseif ($fName -match "Bool|Booleom|Booleon|YesNo|Status|Active|Enabled|Hold|Block|Approved") {
                        $fieldType = "boolean"
                    }
                    elseif ($fName -match "Number|Qty|Amt|Rate|Val|Numeric|Discount|Price|Tax|Gst|Balance|Percent|Charge|Cost|Comm|Commission|Duty|Freight") {
                        $fieldType = "numeric"
                        $decimalPlaces = 3 # Matches FiedNumber default in screenshot
                    }
                    else {
                        $fieldType = "text"
                    }

                    # B. SELF-LEARNING DROPDOWN DETECTION: Query the database directly. If master records 
                    # exist under MasterType = 1000 + i, automatically configure as a dropdown list.
                    if ($fieldType -eq "text") {
                        $targetMasterType = 1000 + $i
                        $chkRst = $fi.GetRecordset("SELECT COUNT(*) AS TotalCount FROM Master1 WHERE MasterType=$targetMasterType")
                        if ($chkRst -and -not $chkRst.EOF) {
                            $cnt = [int]$chkRst.Fields.Item("TotalCount").Value
                            if ($cnt -gt 0) {
                                $maintainMaster = $true
                            }
                            $chkRst.Close()
                        }
                    }

                    $fields += @{
                        fieldKey       = "OptionField$i"
                        fieldName      = $fName
                        fieldType      = $fieldType
                        decimalPlaces  = $decimalPlaces
                        maintainMaster = $maintainMaster
                    }
                }
            }
            $rst.Close()
        }

        return @{ success = $true; count = $fields.Count; data = $fields }
    } catch {
        return @{ success = $false; error = $_.Exception.Message }
    } finally {
        Disconnect-BUSY $fi
    }
}

# ═══════════════════════════════════════════════════════════════
#  2. GET OPTIONAL FIELD MASTER AUTCOMPLETE SUGGESTIONS
# ═══════════════════════════════════════════════════════════════
function Get-OptionalFieldMasterValues {
    param(
        [int]$VchType,
        [string]$SeriesName,
        [int]$FieldNo,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) {
        return @{ success = $false; error = "BUSY connection failed" }
    }

    try {
        # DIRECTLY QUERY THE MAIN MASTER TABLE USING MASTERTYPE LOGIC (1000 + FieldNo)
        $targetMasterType = 1000 + $FieldNo
        $qry = "SELECT Name FROM Master1 WHERE MasterType=$targetMasterType ORDER BY Name"
        $rst = $fi.GetRecordset($qry)

        $values = @()
        if ($rst -and -not $rst.EOF) {
            $rst.MoveFirst()
            while (-not $rst.EOF) {
                $val = $rst.Fields.Item("Name").Value
                if ($val -ne [System.DBNull]::Value -and $val -ne $null) {
                    $values += $val.ToString().Trim()
                }
                $rst.MoveNext()
            }
            $rst.Close()
        }

        return @{ success = $true; data = $values }
    } catch {
        return @{ success = $false; error = $_.Exception.Message }
    } finally {
        Disconnect-BUSY $fi
    }
}

# ═══════════════════════════════════════════════════════════════
#  DIRECT NATIVE HELPER: Get-VchCode-Direct
# ═══════════════════════════════════════════════════════════════
function Get-VchCode-Direct {
    param($conn, [int]$VchType, [string]$VchNo, [string]$VchDate, [bool]$isSql)
    
    $dateFilter = ""
    if ($VchDate -match "\d{2}-\d{2}-\d{4}") {
        $d = [datetime]::ParseExact($VchDate, "dd-MM-yyyy", $null)
        $dateFilter = if ($isSql) { " AND [Date]='" + $d.ToString("yyyy-MM-dd") + "'" } else { " AND [Date]=#" + $d.ToString("MM/dd/yyyy") + "#" }
    } elseif ($VchDate -match "\d{4}-\d{2}-\d{2}") {
        $d = [datetime]::ParseExact($VchDate, "yyyy-MM-dd", $null)
        $dateFilter = if ($isSql) { " AND [Date]='" + $d.ToString("yyyy-MM-dd") + "'" } else { " AND [Date]=#" + $d.ToString("MM/dd/yyyy") + "#" }
    }

    $targetNo = $VchNo.Trim().ToLower()
    $vchCode = 0

    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT [VchCode], [VchNo], [AutoVchNo] FROM [Tran1] WHERE [VchType] = $VchType" + $dateFilter
        $rdr = $cmd.ExecuteReader()
        while ($rdr.Read()) {
            $dbVchNo = ""
            $v = $rdr["VchNo"]
            if ($null -ne $v -and "$v" -ne "") { $dbVchNo = $v.ToString().Trim().ToLower() }
            if ($dbVchNo -eq "") {
                $av = $rdr["AutoVchNo"]
                if ($null -ne $av -and "$v" -ne "") { $dbVchNo = $av.ToString().Trim().ToLower() }
            }
            if ($dbVchNo -eq $targetNo) {
                $vchCode = [int]$rdr["VchCode"]
                break
            }
        }
        $rdr.Close()
    } catch {}

    if ($vchCode -eq 0) {
        try {
            $cmd = $conn.CreateCommand()
            $cmd.CommandText = "SELECT [VchCode], [VchNo], [AutoVchNo] FROM [Tran1] WHERE [VchType] = $VchType"
            $rdr = $cmd.ExecuteReader()
            while ($rdr.Read()) {
                $dbVchNo = ""
                $v = $rdr["VchNo"]
                if ($null -ne $v -and "$v" -ne "") { $dbVchNo = $v.ToString().Trim().ToLower() }
                if ($dbVchNo -eq "") {
                    $av = $rdr["AutoVchNo"]
                    if ($null -ne $av -and "$av" -ne "") { $dbVchNo = $av.ToString().Trim().ToLower() }
                }
                if ($dbVchNo -eq $targetNo) {
                    $vchCode = [int]$rdr["VchCode"]
                    break
                }
            }
            $rdr.Close()
        } catch {}
    }
    return $vchCode
}

# ═══════════════════════════════════════════════════════════════
#  CORE UTILITY HELPERS
# ═══════════════════════════════════════════════════════════════
function Validate-VoucherData {
    param($Data, $Cfg)

    foreach ($key in $Cfg.requiredKeys) {
        $val = $Data.$key
        if ($key -eq "items") {
            if (-not $val -or $val.Count -eq 0) {
                return "At least one item is required"
            }
        } elseif ($key -eq "accounts") {
            if (-not $val -or $val.Count -eq 0) {
                return "At least one account entry is required"
            }
        } else {
            if (-not $val -or [string]$val -eq "") {
                return "$key is required"
            }
        }
    }
    return $null
}

function Pad-VchNo {
    param([string]$VchNo)
    if ($null -eq $VchNo) { return "" }
    if ($VchNo.Length -ge 25) { return $VchNo }
    return $VchNo.PadLeft(25)
}

function Get-VoucherAutoSequence {
    param(
        $fi,
        [int]$VchType,
        [string]$SeriesName,
        [string]$VchNo,
        [string]$VchDateStr
    )

    if ($null -eq $fi) { return 0 }

    try {
        $prefixStr = "{0:D2}" -f $VchType
        $prefixedSeriesName = $SeriesName
        if (-not $SeriesName.StartsWith($prefixStr)) {
            $prefixedSeriesName = "$prefixStr$SeriesName"
        }
        $safeSeriesName = $SeriesName -replace "'", "''"
        $safePrefixedName = $prefixedSeriesName -replace "'", "''"

        $seriesCode = 0
        $sRst = $fi.GetRecordset("SELECT Code FROM Master1 WHERE MasterType=21 AND (Name='$safeSeriesName' OR Name='$safePrefixedName')")
        if ($sRst -and -not $sRst.EOF) {
            $seriesCode = [int]$sRst.Fields.Item("Code").Value
            $sRst.Close()
        }

        if ($seriesCode -eq 0) { return 0 }

        $isAuto = $false
        $startNo = 1
        $frequency = 0
        $cfgRst = $fi.GetRecordset("SELECT I2, L2 FROM Config WHERE RecType=6 AND L1=$seriesCode")
        if ($cfgRst -and -not $cfgRst.EOF) {
            if ($cfgRst.Fields.Item("L2").Value -ne [System.DBNull]::Value) {
                $l2val = [long]$cfgRst.Fields.Item("L2").Value
                $isAuto = ($l2val -ge 1)
                if ($l2val -gt 1) { $startNo = $l2val }
            }
            if ($cfgRst.Fields.Item("I2").Value -ne [System.DBNull]::Value) {
                $frequency = [int]$cfgRst.Fields.Item("I2").Value
            }
            $cfgRst.Close()
        }

        if (-not $isAuto) { return 0 }

        $existingAutoNo = 0
        $safeNo = $VchNo.Trim()
        $lookupNo = Pad-VchNo $safeNo
        $lookupNoExact = $safeNo -replace "'", "''"
        $lookupNoPadded = $lookupNo -replace "'", "''"

        $eRst = $fi.GetRecordset("SELECT AutoVchNo FROM Tran1 WHERE VchType=$VchType AND (VchNo='$lookupNoExact' OR VchNo='$lookupNoPadded')")
        if ($eRst -and -not $eRst.EOF) {
            if ($eRst.Fields.Item("AutoVchNo").Value -ne [System.DBNull]::Value) {
                $existingAutoNo = [int]$eRst.Fields.Item("AutoVchNo").Value
            }
            $eRst.Close()
        }

        if ($existingAutoNo -gt 0) {
            return $existingAutoNo
        }

        $vchDate = [datetime]::Now
        if ($VchDateStr -match "(\d{2})-(\d{2})-(\d{4})") {
            $vchDate = [datetime]::ParseExact($VchDateStr, "dd-MM-yyyy", $null)
        } elseif ($VchDateStr -match "(\d{4})-(\d{2})-(\d{2})") {
            $vchDate = [datetime]::ParseExact($VchDateStr, "yyyy-MM-dd", $null)
        }

        $dateFilter = ""
        if ($frequency -eq 1) {
            $dateFilter = " AND [Date] = #$($vchDate.ToString('MM/dd/yyyy'))#"
        }
        elseif ($frequency -eq 2) {
            $startDate = New-Object DateTime $vchDate.Year, $vchDate.Month, 1
            $endDate = $startDate.AddMonths(1).AddDays(-1)
            $dateFilter = " AND [Date] >= #$($startDate.ToString('MM/dd/yyyy'))# AND [Date] <= #$($endDate.ToString('MM/dd/yyyy'))#"
        }
        elseif ($frequency -eq 3) {
            $fyStartYear = if ($vchDate.Month -ge 4) { $vchDate.Year } else { $vchDate.Year - 1 }
            $startDate = New-Object DateTime $fyStartYear, 4, 1
            $endDate = $startDate.AddYears(1).AddDays(-1)
            $dateFilter = " AND [Date] >= #$($startDate.ToString('MM/dd/yyyy'))# AND [Date] <= #$($endDate.ToString('MM/dd/yyyy'))#"
        }

        $lastSeq = 0
        $tranRst = $fi.GetRecordset("SELECT AutoVchNo FROM Tran1 WHERE VchSeriesCode=$seriesCode AND Cancelled=0 AND VchCancelled=0" + $dateFilter)
        if ($tranRst -and -not $tranRst.EOF) {
            $tranRst.MoveFirst()
            while (-not $tranRst.EOF) {
                $dbAutoNo = 0
                try {
                    if ($tranRst.Fields.Item("AutoVchNo").Value -ne [System.DBNull]::Value) {
                        $dbAutoNo = [int]$tranRst.Fields.Item("AutoVchNo").Value
                    }
                } catch {}
                if ($dbAutoNo -gt $lastSeq) { $lastSeq = $dbAutoNo }
                $tranRst.MoveNext()
            }
            $tranRst.Close()
        }

        $nextSeq = if ($lastSeq -gt 0) { $lastSeq + 1 } else { $startNo }
        return $nextSeq
    } catch {
        return 0
    }
}

function Clear-StockCaches {
    param([string]$InstanceId = "", [string]$CompanyCode = "")
    
    $prefix = "$InstanceId|$CompanyCode|"
    $keysToRemove = @($script:_cache.Keys | Where-Object {
        $_ -like "${prefix}items|*" -or
        $_ -like "${prefix}item|*" -or
        $_ -like "${prefix}parties|*" -or
        $_ -like "${prefix}vouchers|*" -or
        $_ -like "${prefix}voucherseries|*"
    })
    
    foreach ($k in $keysToRemove) {
        $script:_cache.Remove($k)
    }
}

function Update-CheckListCreator {
    param(
        $fi,
        [int]$VchType,
        [string]$VchNo,
        [string]$VchDate,
        [string]$UserName,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )
    if ([string]::IsNullOrEmpty($UserName)) { return }

    try {
        $dbType = 0
        if ($InstanceId -and $CompanyCode) {
            $instance = Get-InstanceConfig -InstanceId $InstanceId
            if ($null -ne $instance -and $null -ne $instance.dbType) {
                $dbType = [int]$instance.dbType
            }
        }
        $wildcard = if ($dbType -eq 1) { "%" } else { "*" }

        $vchCode = 0
        $safeNo = $VchNo.Trim().ToLower()

        $qry = "SELECT VchCode FROM Tran1 WHERE VchType=$VchType AND VchNo LIKE '$wildcard$safeNo$wildcard'"
        $rst = $fi.GetRecordset($qry)

        if ($rst -and -not $rst.EOF) {
            $rst.MoveFirst()
            $vchCode = [int]$rst.Fields.Item("VchCode").Value
            $rst.Close()
        }

        if ($vchCode -gt 0) {
            $safeUser = $UserName -replace "'", "''"
            $fi.ExecuteQuery("UPDATE CheckList SET UserName='$safeUser' WHERE Code=$vchCode")
        }
    } catch {}
}

# ═══════════════════════════════════════════════════════════════
#  XML GENERATOR HELPERS
# ═══════════════════════════════════════════════════════
function Build-ItemsXml {
    param(
        $items,
        [string]$defaultMC = "Main Store"
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

function Build-BillSundriesXml {
    param($billSundries)

    if (-not $billSundries -or $billSundries.Count -eq 0) {
        return ""
    }

    $xml = "<BillSundries>"
    foreach ($bs in $billSundries) {
        $xml += "<BSDetail>"
        $xml += "<SrNo>$($bs.srNo)</SrNo>"
        $xml += "<BSName>$([System.Security.SecurityElement]::Escape($bs.name))</BSName>"
        $xml += "<Amt>$($bs.amount)</Amt>"
        if ($bs.percentVal) {
            $xml += "<PercentVal>$($bs.percentVal)</PercentVal>"
        }
        $xml += "</BSDetail>"
    }
    $xml += "</BillSundries>"

    return $xml
}

function Build-AccEntriesXml {
    param($items, [string]$PartyName, [double]$TotalAmt,[int]$VchType, $Settlements)

    $salesAcc = "Sales"
    $purcAcc = "Purchase"
    
    foreach ($item in $items) {
        if ($item.salesAccount -and $item.salesAccount -ne "") {
            $salesAcc = $item.salesAccount
        }
        if ($item.purchaseAccount -and $item.purchaseAccount -ne "") {
            $purcAcc = $item.purchaseAccount
        }
    }

    $cashAmt = 0; $cardAmt = 0; $giftAmt = 0
    $cashAcc = "Cash"; $cardAcc = ""; $giftAcc = ""
    
    if ($null -ne $Settlements) {
        if ($Data.settlements.cash -and $Data.settlements.cash.amount) {
            $cashAmt = [double]$Data.settlements.cash.amount
            if ($Data.settlements.cash.account) { $cashAcc = $Data.settlements.cash.account }
        }
        if ($Data.settlements.card -and $Data.settlements.card.amount) {
            $cardAmt = [double]$Data.settlements.card.amount
            if ($Data.settlements.card.account) { $cardAcc = $Data.settlements.card.account }
        }
        if ($Data.settlements.gift -and $Data.settlements.gift.amount) {
            $giftAmt = [double]$Data.settlements.gift.amount
            if ($Data.settlements.gift.account) { $giftAcc = $Data.settlements.gift.account }
        }
    }

    $totalSettled = $cashAmt + $cardAmt + $giftAmt
    $partyBalance = $TotalAmt - $totalSettled

    $xml = "<AccEntries>"
    
    if ($VchType -eq 9 -or $VchType -eq 3) {
        
                        # 1. Dynamic assignments to handle Dr/Cr directions cleanly
            if ($VchType -eq 9) {
                # Sale Invoice: Party Dr (1, Negative), Sales Cr (2, Positive)
                $partyAmtType = 1; $partyAmt = -$partyBalance
                $saleAmtType  = 2; $saleAmt  = $TotalAmt
            } else {
                # Sale Return: Party Cr (2, Positive), Sales Dr (1, Negative)
                $partyAmtType = 2; $partyAmt = $partyBalance
                $saleAmtType  = 1; $saleAmt  = -$TotalAmt
            }
        
           # 2. Customer Party AccDetail (updated to use $partyAmtType)
            if ([Math]::Round($partyBalance, 2) -gt 0 -or $totalSettled -eq 0) {
                $xml += "<AccDetail>"
                $xml += "<AccountName>$([System.Security.SecurityElement]::Escape($PartyName))</AccountName>"
                $xml += "<AmountType>$partyAmtType</AmountType>"
                $xml += "<AmtMainCur>$($partyAmt.ToString('0.##',[System.Globalization.CultureInfo]::InvariantCulture))</AmtMainCur>"
                $xml += "</AccDetail>"
            }
        
                $xml += "<AccDetail>"
                $xml += "<AccountName>$([System.Security.SecurityElement]::Escape($salesAcc))</AccountName>"
                $xml += "<AmountType>$saleAmtType</AmountType>"
                $xml += "<AmtMainCur>$($saleAmt.ToString('0.##',[System.Globalization.CultureInfo]::InvariantCulture))</AmtMainCur>"
                $xml += "</AccDetail>"

        if ($cashAmt -gt 0 -and $cashAcc) {
            $cAmt = -$cashAmt
            $xml += "<AccDetail><AccountName>$([System.Security.SecurityElement]::Escape($cashAcc))</AccountName><AmountType>1</AmountType><AmtMainCur>$($cAmt.ToString('0.##',[System.Globalization.CultureInfo]::InvariantCulture))</AmtMainCur></AccDetail>"
        }
        if ($cardAmt -gt 0 -and $cardAcc) {
            $cAmt = -$cardAmt
            $xml += "<AccDetail><AccountName>$([System.Security.SecurityElement]::Escape($cardAcc))</AccountName><AmountType>1</AmountType><AmtMainCur>$($cAmt.ToString('0.##',[System.Globalization.CultureInfo]::InvariantCulture))</AmtMainCur></AccDetail>"
        }
        if ($giftAmt -gt 0 -and $giftAcc) {
            $cAmt = -$giftAmt
            $xml += "<AccDetail><AccountName>$([System.Security.SecurityElement]::Escape($giftAcc))</AccountName><AmountType>1</AmountType><AmtMainCur>$($cAmt.ToString('0.##',[System.Globalization.CultureInfo]::InvariantCulture))</AmtMainCur></AccDetail>"
        }
    } elseif ($VchType -eq 2 -or $VchType -eq 10) {

            # 1. Dynamic assignments to handle Dr/Cr directions cleanly
    if ($VchType -eq 2) {
        # Purchase: Party Cr (2, Positive), Purchase Dr (1, Negative)
        $partyAmtType = 2; $partyAmtVal = $TotalAmt
        $purcAmtType  = 1; $purcAmtVal  = -$TotalAmt
    } else {
        # Purchase Return: Party Dr (1, Negative), Purchase Cr (2, Positive)
        $partyAmtType = 1; $partyAmtVal = -$TotalAmt
        $purcAmtType  = 2; $purcAmtVal  = $TotalAmt
    }
    
        $negAmt = -$TotalAmt
        
        # 2. Supplier Party AccDetail
        $xml += "<AccDetail>"
        $xml += "<AccountName>$([System.Security.SecurityElement]::Escape($PartyName))</AccountName>"
        $xml += "<AmountType>$partyAmtType</AmountType>"
        $xml += "<AmtMainCur>$($partyAmtVal.ToString('0.##',[System.Globalization.CultureInfo]::InvariantCulture))</AmtMainCur>"
        $xml += "</AccDetail>"
        
        # 3. Purchase Account AccDetail
        $xml += "<AccDetail>"
        $xml += "<AccountName>$([System.Security.SecurityElement]::Escape($purcAcc))</AccountName>"
        $xml += "<AmountType>$purcAmtType</AmountType>"
        $xml += "<AmtMainCur>$($purcAmtVal.ToString('0.##',[System.Globalization.CultureInfo]::InvariantCulture))</AmtMainCur>"
        $xml += "</AccDetail>"
    }
    
    $xml += "</AccEntries>"
    return $xml
}

function Build-PendingBillDetailsXml {
    param($refEntries, [string]$PartyName, [string]$VchDate, [int]$VchType,[string]$VchNo)

    if ($null -eq $refEntries) { return "" }
    $refs = @($refEntries)
    if ($refs.Count -eq 0) { return "" }

    $xml = "<BillDetail>"
    $xml += "<MasterName1>$([System.Security.SecurityElement]::Escape($PartyName))</MasterName1>"

    foreach ($ref in $refs) {
        $rNo = if ($ref.refNo -eq "AUTO" -or [string]::IsNullOrWhiteSpace($ref.refNo)) { $VchNo } else { $ref.refNo }
        $dueDate = if ($ref.dueDate) { $ref.dueDate } else { $VchDate }
        
        $val = [double]$ref.amount
        if ($VchType -eq 9 -or $VchType -eq 10) {
            $val = -($val)
        }
        $valStr = $val.ToString("0.00",[System.Globalization.CultureInfo]::InvariantCulture)

        $xml += "<BillRefs>"
        $xml += "<Method>$($ref.refType)</Method>"
        $xml += "<SrNo>$($ref.srNo)</SrNo>"
        $xml += "<RefNo>$([System.Security.SecurityElement]::Escape($rNo))</RefNo>"
        $xml += "<Date>$VchDate</Date>"
        $dueDate = if ($ref.dueDate) { $ref.dueDate } else { $VchDate }
        $xml += "<DueDate>$dueDate</DueDate>"
        $xml += "<Value1>$valStr</Value1>"
        $xml += "<VchType>$VchType</VchType>"
        $xml += "<MfgDate>$VchDate</MfgDate>"
        $xml += "</BillRefs>"
    }

    $xml += "</BillDetail>"
    return $xml
}

function Build-PendingOrdersXml {
    param($items, [string]$PartyName, $linkedOrder, [int]$VchType, [string]$VchDate = "")

    if (-not $linkedOrder -or -not $linkedOrder.vchNo) {
        return ""
    }

    $xml = "<PendingOrders>"
    foreach ($item in $items) {
        $qty = [double]$item.qty
        $altQty = if ($item.altQty) { [double]$item.altQty } else { $qty }
        $price = if ($item.price) { [double]$item.price } else { 0 }
        $altPrice = if ($item.altPrice) { [double]$item.altPrice } else { $price }
        
        $qtyStr = (-$qty).ToString("0.000",[System.Globalization.CultureInfo]::InvariantCulture)
        $altQtyStr = (-$altQty).ToString("0.000",[System.Globalization.CultureInfo]::InvariantCulture)

        $xml += "<OrderDetail>"
        $xml += "<MasterName1>$([System.Security.SecurityElement]::Escape($item.itemName))</MasterName1>"
        $xml += "<MasterName2>$([System.Security.SecurityElement]::Escape($PartyName))</MasterName2>"
        $xml += "<OrderRefs>"
        $xml += "<Method>2</Method>"
        $xml += "<SrNo>1</SrNo>"
        $xml += "<RefNo>$([System.Security.SecurityElement]::Escape($linkedOrder.vchNo))</RefNo>"
        $xml += "<Date>$($linkedOrder.date)</Date>"
        $dueDate = if ($linkedOrder.date) { $linkedOrder.date } else { $VchDate }
        $xml += "<DueDate>$dueDate</DueDate>"
        $xml += "<Value1>$qtyStr</Value1>"
        $xml += "<Value2>$altQtyStr</Value2>"
        $xml += "<MainTranPrice>$price</MainTranPrice>"
        $xml += "<VchType>$VchType</VchType>"
        $xml += "<ItemSrNo>$($item.srNo)</ItemSrNo>"
        $xml += "<AltTranPrice>$altPrice</AltTranPrice>"
        $xml += "<MfgDate>$($linkedOrder.date)</MfgDate>"
        $xml += "</OrderRefs>"
        $xml += "</OrderDetail>"
    }
    
    $xml += "</PendingOrders>"
    return $xml
}

function Build-PendingChallansXml {
    param($items, [string]$PartyName, $linkedChallans, [int]$VchType, [string]$VchNo, [string]$VchDate, $fi)

    if ($VchType -eq 11 -or $VchType -eq 4) {
        $xml = "<PendingChallans>"
        foreach ($item in $items) {
            $qty        = [double]$item.qty
            $conFactor  = if ($item.altQtyConFactor -and [double]$item.altQtyConFactor -ne 0) { [double]$item.altQtyConFactor } else { 1 }
            $altQty     = if ($item.altQty -and [double]$item.altQty -ne 0) { [double]$item.altQty } else { [Math]::Round($qty * $conFactor, 3) }
            $price      = [double]$item.price
            $altPrice   = if ($altQty -ne 0) { [Math]::Round(($price * $qty) / $altQty, 4) } else { $price }

            $qtyStr     = (-$qty).ToString("0.000", [System.Globalization.CultureInfo]::InvariantCulture)
            $altQtyStr  = (-$altQty).ToString("0.000", [System.Globalization.CultureInfo]::InvariantCulture)
            $paddedVchNo = Pad-VchNo $VchNo

            $xml += "<ChallanDetail>"
            $xml += "<MasterName1>$([System.Security.SecurityElement]::Escape($item.itemName))</MasterName1>"
            $xml += "<MasterName2>$([System.Security.SecurityElement]::Escape($PartyName))</MasterName2>"
            $xml += "<ChallanRefs>"
            $xml += "<Method>1</Method>"
            $xml += "<SrNo>1</SrNo>"
            $xml += "<RefNo>$([System.Security.SecurityElement]::Escape($paddedVchNo))</RefNo>"
            $xml += "<Date>$VchDate</Date>"
            $xml += "<DueDate>$VchDate</DueDate>"
            $xml += "<Value1>$qtyStr</Value1>"
            $xml += "<Value2>$altQtyStr</Value2>"
            $xml += "<MainTranPrice>$price</MainTranPrice>"
            $xml += "<VchType>$VchType</VchType>"
            $xml += "<ItemSrNo>$($item.srNo)</ItemSrNo>"
            $xml += "<AltTranPrice>$altPrice</AltTranPrice>"
            $xml += "<MfgDate>$VchDate</MfgDate>"
            $xml += "</ChallanRefs>"
            $xml += "</ChallanDetail>"
        }
        $xml += "</PendingChallans>"
        return $xml
    }

    if (($VchType -eq 9 -or $VchType -eq 2) -and $linkedChallans -and @($linkedChallans).Count -gt 0 -and $null -ne $fi) {
        $xml = "<PendingChallans>"
        
        $partyCode = ""
        $safeParty = $PartyName -replace "'", "''"
        $mRst = $fi.GetRecordset("SELECT Code FROM Master1 WHERE Name = '$safeParty'")
        if ($mRst -and -not $mRst.EOF) {
            $partyCode = $mRst.Fields.Item("Code").Value.ToString()
        }
        if ($mRst) { try { $mRst.Close() } catch {} }

        $masterMap = @{}
        $itemQryRst = $fi.GetRecordset("SELECT Code, Name FROM Master1 WHERE MasterType = 6")
        if ($itemQryRst -and -not $itemQryRst.EOF) {
            while (-not $itemQryRst.EOF) {
                $masterMap[$itemQryRst.Fields.Item("Code").Value.ToString()] = $itemQryRst.Fields.Item("Name").Value.ToString()
                $itemQryRst.MoveNext()
            }
            $itemQryRst.Close()
        }

        $dbType = 0
        if ($script:ActiveConnection -ne $null) {
            $instance = Get-InstanceConfig -InstanceId $script:ActiveInstanceId
            if ($null -ne $instance -and $null -ne $instance.dbType) {
                $dbType = [int]$instance.dbType
            }
        }
        $wildcard = if ($dbType -eq 1) { "%" } else { "*" }

        foreach ($item in $items) {
            $qty      = [double]$item.qty
            $altQty   = if ($item.altQty  -and [double]$item.altQty  -ne 0) { [double]$item.altQty  } else { $qty }
            $altPrice = if ($item.altPrice -and [double]$item.altPrice -ne 0) { [double]$item.altPrice } else { [double]$item.price }
            $itemSrNo = [int]$item.srNo
            
            foreach ($lc in $linkedChallans) {
                $paddedRefNo = Pad-VchNo $lc.vchNo
                
                $resolvedRefCode  = ""
                $resolvedItemCode = ""
                if ($partyCode -ne "") {
                    $safeNo = $lc.vchNo.Trim() -replace "'", "''"
                    $targetRecType = if ($VchType -eq 9) { 3 } else { 4 }
                    
                    $refQry = "SELECT RefCode, MasterCode1 FROM Tran3 WHERE Method=1 AND RecType=$targetRecType AND MasterCode2=$partyCode AND [No] LIKE '$wildcard$safeNo$wildcard' AND ItemSrNo=$itemSrNo"
                    $refRst = $fi.GetRecordset($refQry)
                    if ($refRst -and -not $refRst.EOF) {
                        $resolvedRefCode  = $refRst.Fields.Item("RefCode").Value.ToString()
                        $resolvedItemCode = $refRst.Fields.Item("MasterCode1").Value.ToString()
                    }
                    if ($refRst) { try { $refRst.Close() } catch {} }
                }

                if ($resolvedItemCode -eq "") {
                    $safeItem = $item.itemName -replace "'", "''"
                    $iRst = $fi.GetRecordset("SELECT Code FROM Master1 WHERE Name = '$safeItem'")
                    if ($iRst -and -not $iRst.EOF) {
                        $resolvedItemCode = $iRst.Fields.Item("Code").Value.ToString()
                    }
                    if ($iRst) { try { $iRst.Close() } catch {} }
                }

                $resolvedItemName = if ($null -ne $masterMap -and $masterMap.ContainsKey($resolvedItemCode)) { $masterMap[$resolvedItemCode] } else { $item.itemName }

                $xml += "<ChallanDetail>"
                $xml += "<MasterName1>$([System.Security.SecurityElement]::Escape($resolvedItemName))</MasterName1>"
                $xml += "<MasterName2>$([System.Security.SecurityElement]::Escape($PartyName))</MasterName2>"
                if ($resolvedItemCode -ne "") { $xml += "<tmpMasterCode1>$resolvedItemCode</tmpMasterCode1>" }
                if ($partyCode -ne "")        { $xml += "<tmpMasterCode2>$partyCode</tmpMasterCode2>" }
                
                $xml += "<ChallanRefs>"
                $xml += "<Method>2</Method>"
                $xml += "<SrNo>1</SrNo>"
                $xml += "<RefNo>$([System.Security.SecurityElement]::Escape($paddedRefNo))</RefNo>"
                $xml += "<Date>$VchDate</Date>"
                $dueDate = if ($lc.date) { $lc.date } else { $VchDate }
                $xml += "<DueDate>$dueDate</DueDate>"
                $xml += "<Value1>$qty</Value1>"
                $xml += "<Value2>$altQty</Value2>"
                $xml += "<MainTranPrice>$([double]$item.price)</MainTranPrice>"
                $xml += "<VchType>$VchType</VchType>"
                $xml += "<ItemSrNo>$itemSrNo</ItemSrNo>"
                $xml += "<AltTranPrice>$altPrice</AltTranPrice>"
                $xml += "<MfgDate>$VchDate</MfgDate>"
                
                if ($resolvedRefCode -ne "") {
                    $xml += "<tmpRefCode>$resolvedRefCode</tmpRefCode>"
                    $xml += "<tmpRecType>$targetRecType</tmpRecType>"
                    if ($resolvedItemCode -ne "") { $xml += "<tmpMasterCode1>$resolvedItemCode</tmpMasterCode1>" }
                    if ($partyCode -ne "")        { $xml += "<tmpMasterCode2>$partyCode</tmpMasterCode2>" }
                    $xml += "<TranType>6</TranType>"
                }
                
                $xml += "</ChallanRefs>"
                $xml += "</ChallanDetail>"
            }
        }
        $xml += "</PendingChallans>"
        return $xml
    }
    return ""
}

function Get-BusyAutoVchNoFromVoucherNo {
    param(
        $fi,
        [int]$VchType,
        [string]$SeriesName,
        [string]$VchNo
    )

    if ($null -eq $fi) {
        throw "BUSY connection is unavailable."
    }

    if ([string]::IsNullOrWhiteSpace($VchNo)) {
        throw "Voucher number is empty."
    }

    $cleanSeriesName = ([string]$SeriesName).Trim()
    $cleanVchNo = ([string]$VchNo).Trim()

    $typePrefix = "{0:D2}" -f $VchType
    $prefixedSeriesName = $cleanSeriesName

    if (
        -not $cleanSeriesName.StartsWith(
            $typePrefix,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        $prefixedSeriesName = "$typePrefix$cleanSeriesName"
    }

    $safeSeriesName =
        $cleanSeriesName.Replace("'", "''")

    $safePrefixedSeriesName =
        $prefixedSeriesName.Replace("'", "''")

    $seriesCode = 0
    $seriesRst = $null

    try {
        $seriesQuery = @"
SELECT Code
FROM Master1
WHERE MasterType = 21
  AND (
        Name = '$safeSeriesName'
        OR Name = '$safePrefixedSeriesName'
      )
"@

        $seriesRst = $fi.GetRecordset($seriesQuery)

        if ($seriesRst -and -not $seriesRst.EOF) {
            $value =
                $seriesRst.Fields.Item("Code").Value

            if (
                $null -ne $value -and
                $value -ne [System.DBNull]::Value
            ) {
                $seriesCode = [int]$value
            }
        }
    }
    finally {
        if ($seriesRst) {
            try {
                $seriesRst.Close()
            }
            catch {
            }
        }
    }

    if ($seriesCode -le 0) {
        throw "BUSY series '$cleanSeriesName' was not found."
    }

    $prefix = ""
    $suffix = ""
    $separator = ""
    $dateEnabled = 0
    $frequency = 0
    $datePosition = 0
    $dateFormat = 0

    $cfgRst = $null

    try {
        $configQuery = @"
SELECT C1, C2, C4, I1, I2, I7, I8
FROM Config
WHERE RecType = 6
  AND L1 = $seriesCode
"@

        $cfgRst = $fi.GetRecordset($configQuery)

        if ($cfgRst -and -not $cfgRst.EOF) {
            $readString = {
                param([string]$Name)

                $value =
                    $cfgRst.Fields.Item($Name).Value

                if (
                    $null -eq $value -or
                    $value -eq [System.DBNull]::Value
                ) {
                    return ""
                }

                return ([string]$value).Trim()
            }

            $readInt = {
                param([string]$Name)

                $value =
                    $cfgRst.Fields.Item($Name).Value

                if (
                    $null -eq $value -or
                    $value -eq [System.DBNull]::Value
                ) {
                    return 0
                }

                return [int]$value
            }

            $suffix = & $readString "C1"
            $prefix = & $readString "C2"
            $separator = & $readString "C4"

            $dateEnabled = & $readInt "I1"
            $frequency = & $readInt "I2"
            $datePosition = & $readInt "I7"
            $dateFormat = & $readInt "I8"
        }
    }
    finally {
        if ($cfgRst) {
            try {
                $cfgRst.Close()
            }
            catch {
            }
        }
    }

    $working = $cleanVchNo

    if (
        $prefix -ne "" -and
        $working.StartsWith(
            $prefix,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        $working =
            $working.Substring($prefix.Length)
    }

    if (
        $suffix -ne "" -and
        $working.EndsWith(
            $suffix,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        $working =
            $working.Substring(
                0,
                $working.Length - $suffix.Length
            )
    }

    if ($separator -ne "") {
        $parts = $working.Split(
            @($separator),
            [System.StringSplitOptions]::RemoveEmptyEntries
        )

        $numericParts = @(
            $parts |
            Where-Object {
                $_ -match '^\d+$'
            }
        )

        if ($numericParts.Count -gt 0) {
            # In BUSY formats such as 2026-27-8AZQ,
            # the running number is the final numeric component.
            $candidate =
                $numericParts[$numericParts.Count - 1]

            $parsed = 0L

            if (
                [long]::TryParse(
                    $candidate,
                    [ref]$parsed
                ) -and
                $parsed -gt 0
            ) {
                return $parsed
            }
        }
    }

    # Fallback: use the last numeric block before trailing letters.
    #
    # Example:
    # 2026-27-8AZQ -> 8
    # ABC-001XYZ   -> 001
    $match = [regex]::Match(
        $cleanVchNo,
        '(\d+)(?=[A-Za-z_-]*$)'
    )

    if ($match.Success) {
        $parsed = 0L

        if (
            [long]::TryParse(
                $match.Groups[1].Value,
                [ref]$parsed
            ) -and
            $parsed -gt 0
        ) {
            return $parsed
        }
    }

    throw (
        "Could not extract BUSY AutoVchNo from voucher number " +
        "'$cleanVchNo'."
    )
}

function Build-VoucherXml {
    param($Data, $Cfg, [int]$VchType, [string]$VchNo, [bool]$SkipBBA = $false, $fi = $null)

    # Resolve sequential count to populate AutoVchNo natively in database
    $autoVchNo = Get-BusyAutoVchNoFromVoucherNo `
    -fi $fi `
    -VchType $VchType `
    -SeriesName ([string]$Data.vchSeries) `
    -VchNo $VchNo

    $root = $Cfg.xmlRoot
    $typeField = $Cfg.typeField
    $typeValue = [string]$Data.($Cfg.typeDataKey)
    $matCentre = if ($Data.matCentre) { $Data.matCentre } else { "Main Store" }
    $narration = if ($Data.narration) { $Data.narration } else { "" }
    $totalAmt = ($Data.items | ForEach-Object { [double]$_.amount } | Measure-Object -Sum).Sum
    $inputType = if ($Data.inputType) { [int]$Data.inputType } else { 1 }

    $isChallanType = ($VchType -eq 11 -or $VchType -eq 4)
    $tranType = 0
    if ($isChallanType) {
        $tranType = 3
    } elseif ($inputType -eq 2) {
        $tranType = 6
    }

    $xml = "<$root>"
    $seriesNameToSend = $Data.vchSeries
    $xml += "<VchSeriesName>$([System.Security.SecurityElement]::Escape($seriesNameToSend))</VchSeriesName>"
    $xml += "<Date>$($Data.date)</Date>"
    $xml += "<VchType>$VchType</VchType>"
    $xml += "<TranType>$tranType</TranType>"
    $xml += "<VchNo>$([System.Security.SecurityElement]::Escape($VchNo))</VchNo>"

    # Inject sequence number tag
    if ($autoVchNo -gt 0) {
        $xml += "<AutoVchNo>$autoVchNo</AutoVchNo>"
    }

$xml += "<$typeField>$([System.Security.SecurityElement]::Escape($typeValue))</$typeField>"

# ============================================================
# PRODUCTION VOUCHER - VchType 6
# ============================================================
if ($VchType -eq 6) {

    # Production XML mapping discovered from actual BUSY voucher:
    #
    # MasterName1       = Generated Material Centre
    # MasterName2       = BOM Name
    # ConMCName         = Consumed Material Centre
    # ExtraExpenseInBOM = Extra Expense / Unit

    $generatedMC = if ($Data.matCentre) {
        ([string]$Data.matCentre).Trim()
    } else {
        ""
    }

    $consumedMC = if ($Data.party) {
        ([string]$Data.party).Trim()
    } else {
        ""
    }

    $bomName = if ($Data.bomName) {
        ([string]$Data.bomName).Trim()
    } else {
        ""
    }

    $extraExpenseInBOM = 0.0

    try {
        if ($null -ne $Data.extraExpenseInBOM) {
            $extraExpenseInBOM =
                [double]$Data.extraExpenseInBOM
        }
    }
    catch {
        $extraExpenseInBOM = 0.0
    }

    if ([string]::IsNullOrWhiteSpace($bomName)) {
        throw "Production BOM name is required."
    }

    if ([string]::IsNullOrWhiteSpace($generatedMC)) {
        throw "Generated material centre is required."
    }

    if ([string]::IsNullOrWhiteSpace($consumedMC)) {
        throw "Consumed material centre is required."
    }

    $xml += "<MasterName1>$([System.Security.SecurityElement]::Escape($generatedMC))</MasterName1>"

    $xml += "<MasterName2>$([System.Security.SecurityElement]::Escape($bomName))</MasterName2>"

    $xml += "<ConMCName>$([System.Security.SecurityElement]::Escape($consumedMC))</ConMCName>"

    $extraExpenseText =
        $extraExpenseInBOM.ToString(
            "0.####",
            [System.Globalization.CultureInfo]::InvariantCulture
        )

    $xml += "<ExtraExpenseInBOM>$extraExpenseText</ExtraExpenseInBOM>"

    # Frontend sends the actual selected BOM code.
    # This matches BUSY's native Production XML tmpMasterCode2.
    $bomCode = 0

    try {
        if ($null -ne $Data.bomCode) {
            $bomCode = [int]$Data.bomCode
        }
    }
    catch {
        $bomCode = 0
    }

    if ($bomCode -gt 0) {
        $xml += "<tmpMasterCode2>$bomCode</tmpMasterCode2>"
    }
}
else {

    # ========================================================
    # EXISTING NON-PRODUCTION VOUCHER BEHAVIOUR
    # ========================================================

    $xml += "<MasterName1>$([System.Security.SecurityElement]::Escape([string]$Data.party))</MasterName1>"

    $xml += "<MasterName2>$([System.Security.SecurityElement]::Escape($matCentre))</MasterName2>"

    # Stock Journal consumed material centre
    if ($VchType -eq 8) {
        $xml += "<ConMCName>$([System.Security.SecurityElement]::Escape([string]$Data.party))</ConMCName>"
    }
}

$xml += "<TranCurName>Rs.</TranCurName>"

    $xml += "<InputType>$inputType</InputType>"
    $xml += "<BillingDetails><PartyName>$([System.Security.SecurityElement]::Escape($Data.party))</PartyName></BillingDetails>"
    
    $xml += "<VchOtherInfoDetails>"

    if ($Data.optionalFields) {
        $xml += "<OFInfo>"
        # Populate the XML elements OF1 through OF20 sequentially
        for ($i = 1; $i -le 20; $i++) {
            $key = "OptionField$i"
            $val = ""

            # DUAL-COMPATIBLE PROPERTY DETECTION: Handles both [PSCustomObject] and [Hashtable] safely
            if ($Data.optionalFields.psobject.Properties[$key]) {
                $val = [string]$Data.optionalFields.$key
            } elseif ($Data.optionalFields.GetType().Name -eq "Hashtable" -and $Data.optionalFields.ContainsKey($key)) {
                $val = [string]$Data.optionalFields[$key]
            }

            if ($val -ne "") {
                # AUTOMATIC DATE FORMATTER: Detect and reformat standard HTML5 yyyy-MM-dd values to dd-MM-yyyy
                if ($val -match "^(\d{4})-(\d{2})-(\d{2})$") {
                    $val = "$($Matches[3])-$($Matches[2])-$($Matches[1])"
                }
                $xml += "<OF$i>$([System.Security.SecurityElement]::Escape($val))</OF$i>"
            }
        }
        $xml += "</OFInfo>"
    }

    if ($Cfg.hasBillNo) {
        $billNo = if ($Data.supplierBillNo) { $Data.supplierBillNo } elseif ($Data.purchaseBillNo) { $Data.purchaseBillNo } else { "" }
        $xml += "<PurchaseBillNo>$([System.Security.SecurityElement]::Escape($billNo))</PurchaseBillNo>"
    }
    
    $xml += "<Narration1>$([System.Security.SecurityElement]::Escape($narration))</Narration1>"
    $xml += "</VchOtherInfoDetails>"

if ($VchType -eq 8 -or $VchType -eq 6) {

    # ========================================================
    # STOCK JOURNAL + PRODUCTION
    #
    # ItemEntries  = Items Generated
    # ItemEntries1 = Items Consumed
    # ========================================================

    $genItems = @(
        $Data.items | Where-Object {
            $null -eq $_.itemType -or
            [int]$_.itemType -eq 1
        }
    )

    $conItems = @(
        $Data.items | Where-Object {
            $null -ne $_.itemType -and
            [int]$_.itemType -eq 2
        }
    )

    $debugLabel =
        if ($VchType -eq 6) {
            "Production"
        }
        else {
            "StockJournal"
        }

    Write-Host "[$debugLabel Save DEBUG] Generated items=$($genItems.Count), Consumed items=$($conItems.Count)" -ForegroundColor Cyan
    Write-Host "[$debugLabel Save DEBUG] Generated MC=$matCentre, Consumed MC=$($Data.party)" -ForegroundColor Cyan

    # Generated section
    $xml += Build-ItemsXml `
        -items $genItems `
        -defaultMC $matCentre

    # Consumed section
    $xml += Build-ConsumedItemsXml `
        -items $conItems `
        -defaultMC ([string]$Data.party)
}
else {

    # Existing behaviour for Sales / Purchase / Challan etc.
    $xml += Build-ItemsXml `
        -items $Data.items `
        -defaultMC $matCentre
}
    $xml += Build-BillSundriesXml -billSundries $Data.billSundries

    if (-not $isChallanType) {
        $xml += Build-AccEntriesXml -items $Data.items -PartyName $Data.party -TotalAmt $totalAmt -VchType $VchType -Settlements $Data.settlements
    }

    $cAmt = 0; $cc1Amt = 0; $cc2Amt = 0
    $cAcc = "Cash"; $cc1Acc = ""; $cc2Acc = ""
    $tendered = 0

    if ($Data.settlements) {
        if ($Data.settlements.cash.amount) { $cAmt = [double]$Data.settlements.cash.amount }
        if ($Data.settlements.cash.account) { $cAcc = $Data.settlements.cash.account }

        if ($Data.settlements.card.amount) { $cc1Amt = [double]$Data.settlements.card.amount }
        if ($Data.settlements.card.account) { $cc1Acc = $Data.settlements.card.account }

        if ($Data.settlements.gift.amount) { $cc2Amt = [double]$Data.settlements.gift.amount }
        if ($Data.settlements.gift.account) { $cc2Acc = $Data.settlements.gift.account }

        $tendered = $cAmt
        if ($Data.settlements.cashTendered) { $tendered = [double]$Data.settlements.cashTendered }
        if ($tendered -lt $cAmt) { $tendered = $cAmt }
    }

    $bbaBlocks = ""

    if (-not $SkipBBA -and -not $isChallanType) {
        if (-not $Data.refEntries -or @($Data.refEntries).Count -eq 0) {
            $partyBalance = $totalAmt
            if ($Data.settlements) {
                $partyBalance = $totalAmt - ($cAmt + $cc1Amt + $cc2Amt)
            }
            if ([Math]::Round($partyBalance, 2) -gt 0) {
                $autoRef = @( @{ srNo = 1; refType = 1; refNo = $VchNo; amount = $partyBalance; dueDate = $Data.date } )
                $bbaBlocks += Build-PendingBillDetailsXml -refEntries $autoRef -PartyName $Data.party -VchDate $Data.date -VchType $VchType -VchNo $VchNo
            }
        } else {
            $bbaBlocks += Build-PendingBillDetailsXml -refEntries $Data.refEntries -PartyName $Data.party -VchDate $Data.date -VchType $VchType -VchNo $VchNo
        }

        if ($Data.settlements) {
            if ($cc1Amt -gt 0 -and $cc1Acc -ne "") {
                $autoRefCard = @( @{ srNo = 1; refType = 1; refNo = $VchNo; amount = $cc1Amt; dueDate = $Data.date } )
                $bbaBlocks += Build-PendingBillDetailsXml -refEntries $autoRefCard -PartyName $cc1Acc -VchDate $Data.date -VchType $VchType -VchNo $VchNo
            }
            if ($cc2Amt -gt 0 -and $cc2Acc -ne "") {
                $autoRefGift = @( @{ srNo = 1; refType = 1; refNo = $VchNo; amount = $cc2Amt; dueDate = $Data.date } )
                $bbaBlocks += Build-PendingBillDetailsXml -refEntries $autoRefGift -PartyName $cc2Acc -VchDate $Data.date -VchType $VchType -VchNo $VchNo
            }
        }
    }

    if ($bbaBlocks -ne "") {
        $xml += "<PendingBillDetails>" + $bbaBlocks + "</PendingBillDetails>"
    }

    if ($Data.linkedOrder -and $Data.linkedOrder.vchNo) {
        $xml += Build-PendingOrdersXml -items $Data.items -PartyName $Data.party -linkedOrder $Data.linkedOrder -VchType $VchType
    }

    if (-not [string]::IsNullOrWhiteSpace($VchNo)) {
        $xml += Build-PendingChallansXml -items $Data.items -PartyName $Data.party -linkedChallans $Data.linkedChallans -VchType $VchType -VchNo $VchNo -VchDate $Data.date -fi $fi
    }

    if ($Data.settlements) {
        $xml += "<PosEnabled>True</PosEnabled>"
        $xml += "<POSVchData>"
        $xml += "<CashAmt>$($cAmt.ToString('0.00',[System.Globalization.CultureInfo]::InvariantCulture))</CashAmt>"
        $xml += "<CashRecvdAmt>$($tendered.ToString('0.00',[System.Globalization.CultureInfo]::InvariantCulture))</CashRecvdAmt>"
        $xml += "<CreditCardAmt1>$($cc1Amt.ToString('0.00',[System.Globalization.CultureInfo]::InvariantCulture))</CreditCardAmt1>"
        $xml += "<CreditCardAmt2>$($cc2Amt.ToString('0.00',[System.Globalization.CultureInfo]::InvariantCulture))</CreditCardAmt2>"
        $xml += "<CarrySettlement>True</CarrySettlement>"
        $xml += "<CashAccName>$([System.Security.SecurityElement]::Escape($cAcc))</CashAccName>"
        $xml += "<CCAccName1>$([System.Security.SecurityElement]::Escape($cc1Acc))</CCAccName1>"
        $xml += "<CCAccName2>$([System.Security.SecurityElement]::Escape($cc2Acc))</CCAccName2>"
        $xml += "</POSVchData>"
    }

    $xml += "</$root>"
    return $xml
}

# FIXED: Re-balanced XML mapping of accounting vouchers to support standard ledger balance checks
function Build-AccountingVoucherXml {
    param($Data, $Cfg,[int]$VchType,[string]$VchNo, [bool]$SkipBBA = $false, $fi = $null)

    # Resolve sequential count to populate AutoVchNo natively in database
    $autoVchNo = Get-BusyAutoVchNoFromVoucherNo `
    -fi $fi `
    -VchType $VchType `
    -SeriesName ([string]$Data.vchSeries) `
    -VchNo $VchNo

    $root = $Cfg.xmlRoot
    $narration = if ($Data.narration) { $Data.narration } else { "" }

    # Query if "Double Entry System for Payment & Receipt" is enabled (RecType=1, L1=16, L2=1 means Double Entry)
    $isDoubleEntry = $false
    if ($fi -and ($VchType -eq 14 -or $VchType -eq 19)) {
        try {
            $cfgRst = $fi.GetRecordset("SELECT L2 FROM Config WHERE RecType=1 AND L1=16")
            if ($cfgRst -and -not $cfgRst.EOF) {
                if ($cfgRst.Fields.Item("L2").Value -ne [System.DBNull]::Value) {
                    $isDoubleEntry = ([int]$cfgRst.Fields.Item("L2").Value -eq 1)
                }
                $cfgRst.Close()
            }
        } catch {
            $isDoubleEntry = $false
        }
    }

    $xml = "<$root>"
    $xml += "<VchSeriesName>$([System.Security.SecurityElement]::Escape($Data.vchSeries))</VchSeriesName>"
    $xml += "<Date>$($Data.date)</Date>"
    $xml += "<VchType>$VchType</VchType>"
    $xml += "<VchNo>$([System.Security.SecurityElement]::Escape($VchNo))</VchNo>"

    # Inject sequence number tag
    if ($autoVchNo -gt 0) {
        $xml += "<AutoVchNo>$autoVchNo</AutoVchNo>"
    }

    $xml += "<TranCurName>Rs.</TranCurName>"

    # Single Entry logic: Specifying MasterName1 Mode at the root
    if (-not $isDoubleEntry -and ($VchType -eq 14 -or $VchType -eq 19) -and $Data.accounts.Count -gt 0) {
        $modeAccount = $Data.accounts[0]
        $xml += "<MasterName1>$([System.Security.SecurityElement]::Escape($modeAccount.accountName))</MasterName1>"
    }

    # AccEntries must contain ALL standard Debit and Credit account lines to satisfy double-entry balancing
    $xml += "<AccEntries>"
    foreach ($acc in $Data.accounts) {
        $amtType = if ($acc.dc -eq 'D') { 1 } else { 2 }
        $amt = [double]$acc.amount 
        
        $xml += "<AccDetail>"
        $xml += "<SrNo>$($acc.srNo)</SrNo>"
        $xml += "<AccountName>$([System.Security.SecurityElement]::Escape($acc.accountName))</AccountName>"
        $xml += "<AmountType>$amtType</AmountType>"
        $xml += "<AmtMainCur>$amt</AmtMainCur>"
        if ($acc.shortNarration) {
            $xml += "<ShortNar>$([System.Security.SecurityElement]::Escape($acc.shortNarration))</ShortNar>"
        }
        
        if (-not $SkipBBA -and $acc.bbaEntries -and @($acc.bbaEntries).Count -gt 0) {
            $xml += "<BillRefs>"
            $bbaSrNo = 1
            foreach ($ref in $acc.bbaEntries) {
                $xml += "<BillDetails>"
                $xml += "<Method>$($ref.refType)</Method>"
                $xml += "<SrNo>$bbaSrNo</SrNo>"
                $rNo = if ($ref.refNo -eq "AUTO" -or [string]::IsNullOrWhiteSpace($ref.refNo)) { $VchNo } else { $ref.refNo }
                $xml += "<RefNo>$([System.Security.SecurityElement]::Escape($rNo))</RefNo>"
                $xml += "<Date>$($Data.date)</Date>"
                $dueDate = if ($ref.dueDate) { $ref.dueDate } else { $Data.date }
                $xml += "<DueDate>$dueDate</DueDate>"
                
                $val1 = if ($amtType -eq 1) { -[Math]::Abs([double]$ref.amount) } else { [Math]::Abs([double]$ref.amount) }
                $xml += "<Value1>$val1</Value1>"
                $xml += "<VchType>$VchType</VchType>"
                $xml += "</BillDetails>"
                $bbaSrNo++
            }
            $xml += "</BillRefs>"
        }
        $xml += "</AccDetail>"
    }
    $xml += "</AccEntries>"

    if (-not $SkipBBA) {
        $hasBba = $false
        foreach ($acc in $Data.accounts) {
            if ($acc.bbaEntries -and @($acc.bbaEntries).Count -gt 0) { $hasBba = $true; break }
        }

        if ($hasBba) {
            $xml += "<PendingBillDetails>"
            foreach ($acc in $Data.accounts) {
                if ($acc.bbaEntries -and @($acc.bbaEntries).Count -gt 0) {
                    $amtType = if ($acc.dc -eq 'D') { 1 } else { 2 }
                    $xml += "<BillDetail>"
                    $xml += "<MasterName1>$([System.Security.SecurityElement]::Escape($acc.accountName))</MasterName1>"
                    $bbaSrNo = 1
                    foreach ($ref in $acc.bbaEntries) {
                        $xml += "<BillRefs>"
                        $xml += "<Method>$($ref.refType)</Method>"
                        $xml += "<SrNo>$bbaSrNo</SrNo>"
                        $rNo = if ($ref.refNo -eq "AUTO" -or [string]::IsNullOrWhiteSpace($ref.refNo)) { $VchNo } else { $ref.refNo }
                        $xml += "<RefNo>$([System.Security.SecurityElement]::Escape($rNo))</RefNo>"
                        $xml += "<Date>$($Data.date)</Date>"
                        $dueDate = if ($ref.dueDate) { $ref.dueDate } else { $Data.date }
                        $xml += "<DueDate>$dueDate</DueDate>"
                        $val1 = if ($amtType -eq 1) { -[Math]::Abs([double]$ref.amount) } else { [Math]::Abs([double]$ref.amount) }
                        $xml += "<Value1>$val1</Value1>"
                        $xml += "<VchType>$VchType</VchType>"
                        $xml += "</BillRefs>"
                        $bbaSrNo++
                    }
                    $xml += "</BillDetail>"
                }
            }
            $xml += "</PendingBillDetails>"
        }
    }

    if ($narration) { $xml += "<VchOtherInfoDetails><Narration1>$([System.Security.SecurityElement]::Escape($narration))</Narration1></VchOtherInfoDetails>" }
    $xml += "</$root>"
    
    return $xml
}

# Helper to natively resolve VchCode directly via high-performance direct DB query
function Get-VchCode-Direct {
    param($conn, [int]$VchType, [string]$VchNo, [string]$VchDate, [bool]$isSql)
    
    $dateFilter = ""
    if ($VchDate -match "\d{2}-\d{2}-\d{4}") {
        $d = [datetime]::ParseExact($VchDate, "dd-MM-yyyy", $null)
        $dateFilter = if ($isSql) { " AND [Date]='" + $d.ToString("yyyy-MM-dd") + "'" } else { " AND [Date]=#" + $d.ToString("MM/dd/yyyy") + "#" }
    } elseif ($VchDate -match "\d{4}-\d{2}-\d{2}") {
        $d = [datetime]::ParseExact($VchDate, "yyyy-MM-dd", $null)
        $dateFilter = if ($isSql) { " AND [Date]='" + $d.ToString("yyyy-MM-dd") + "'" } else { " AND [Date]=#" + $d.ToString("MM/dd/yyyy") + "#" }
    }

    $targetNo = $VchNo.Trim().ToLower()
    $vchCode = 0

    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT [VchCode], [VchNo], [AutoVchNo] FROM [Tran1] WHERE [VchType] = $VchType" + $dateFilter
        $rdr = $cmd.ExecuteReader()
        while ($rdr.Read()) {
            $dbVchNo = ""
            $v = $rdr["VchNo"]
            if ($null -ne $v -and "$v" -ne "") { $dbVchNo = $v.ToString().Trim().ToLower() }
            if ($dbVchNo -eq "") {
                $av = $rdr["AutoVchNo"]
                if ($null -ne $av -and "$v" -ne "") { $dbVchNo = $av.ToString().Trim().ToLower() }
            }
            if ($dbVchNo -eq $targetNo) {
                $vchCode = [int]$rdr["VchCode"]
                break
            }
        }
        $rdr.Close()
    } catch {}

    if ($vchCode -eq 0) {
        try {
            $cmd = $conn.CreateCommand()
            $cmd.CommandText = "SELECT [VchCode], [VchNo], [AutoVchNo] FROM [Tran1] WHERE [VchType] = $VchType"
            $rdr = $cmd.ExecuteReader()
            while ($rdr.Read()) {
                $dbVchNo = ""
                $v = $rdr["VchNo"]
                if ($null -ne $v -and "$v" -ne "") { $dbVchNo = $v.ToString().Trim().ToLower() }
                if ($dbVchNo -eq "") {
                    $av = $rdr["AutoVchNo"]
                    if ($null -ne $av -and "$av" -ne "") { $dbVchNo = $av.ToString().Trim().ToLower() }
                }
                if ($dbVchNo -eq $targetNo) {
                    $vchCode = [int]$rdr["VchCode"]
                    break
                }
            }
            $rdr.Close()
        } catch {}
    }
    return $vchCode
}

# ═══════════════════════════════════════════════════════════════
#  PUBLIC API CONTROLLER HANDLERS
# ═══════════════════════════════════════════════════════════════
function Get-VoucherSeries {
    param([int]$VchType = 0,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $cacheKey = "$InstanceId|$CompanyCode|voucherseries|$VchType"
    if ($script:_cache -and $script:_cache.ContainsKey($cacheKey)) {
        $entry = $script:_cache[$cacheKey]
        if ((Get-Date) -lt $entry.Expires) {
            return $entry.Data
        }
    }

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) {
        return @{ success = $false; error = "BUSY connection failed" }
    }

    try {
        $rst = $fi.GetRecordset("SELECT * FROM Master1 WHERE MasterType = 21")
        $allSeries = @()

        if ($null -ne $rst -and -not $rst.EOF) {
            $rst.MoveFirst()
            while (-not $rst.EOF) {
                $code = ""
                $name = ""

                try {
                    $v = $rst.Fields.Item("Code").Value
                    if ($null -ne $v -and $v -ne [System.DBNull]::Value) {
                        $code = $v.ToString().Trim()
                    }
                } catch {}

                try {
                    $v = $rst.Fields.Item("Name").Value
                    if ($null -ne $v -and $v -ne [System.DBNull]::Value) {
                        $name = $v.ToString().Trim()
                    }
                } catch {}

                $seriesVchType = 0

                foreach ($col in @("I1", "ParentGrp", "CM1", "CM2")) {
                    try {
                        $val = $rst.Fields.Item($col).Value
                        if ($null -ne $val -and $val -ne [System.DBNull]::Value) {
                            $parsed = [int]($val.ToString().Trim())
                            if ($parsed -gt 0) {
                                $seriesVchType = $parsed
                                break
                            }
                        }
                    } catch {}
                }

                $cleanName = $name
                if ($seriesVchType -gt 0) {
                    $prefix = "{0:D2}" -f $seriesVchType
                    if ($cleanName.StartsWith($prefix)) {
                        $cleanName = $cleanName.Substring(2)
                    }
                }

                $allSeries += @{
                    code    = $code
                    name    = $cleanName
                    vchType = $seriesVchType
                }
                $rst.MoveNext()
            }
        }

        if ($null -ne $rst) {
            try { $rst.Close() } catch {}
        }

        $filtered = if ($VchType -eq 0) {
            $allSeries
        } else {
            @($allSeries | Where-Object { $_.vchType -eq $VchType })
        }

        if ($VchType -ne 0 -and $filtered.Count -eq 0 -and $allSeries.Count -gt 0) {
            $filtered = $allSeries
        }

        $response = @{ success = $true; data = $filtered }

        if (-not $script:_cache) {
            $script:_cache = @{}
        }
        $cfg = Get-Config
        $ttl = if ($cfg.CACHE_TTL -gt 0) { $cfg.CACHE_TTL } else { 300 }
        $script:_cache[$cacheKey] = @{
            Data    = $response
            Expires = (Get-Date).AddSeconds($ttl)
        }

        return $response
    } catch {
        return @{ success = $false; error = $_.Exception.Message }
    } finally {
        Disconnect-BUSY $fi
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
                # SUM(Value1) works in both Access and SQL Server — no VAL() needed
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

function Get-NumberingConfig {
    param(
        [int]$VchType,
        [string]$SeriesName,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $fi = $null

    try {
        # Detect database type for correct date syntax.
        $targetInst = Get-InstanceConfig -InstanceId $InstanceId
        $dbType = 0

        if (
            $null -ne $targetInst -and
            $null -ne $targetInst.dbType
        ) {
            $dbType = [int]$targetInst.dbType
        }

        $isSql = ($dbType -eq 1)

        $fi = Connect-BUSY `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (-not $fi) {
            return @{
                success = $false
                error   = "BUSY connection failed"
            }
        }

        if ([string]::IsNullOrWhiteSpace($SeriesName)) {
            return @{
                success = $false
                error   = "Series name is required"
            }
        }

        $SeriesName = $SeriesName.Trim()

        # ------------------------------------------------------------
        # Resolve voucher-series code
        # ------------------------------------------------------------

        $prefixStr = "{0:D2}" -f $VchType
        $prefixedSeriesName = $SeriesName

        if (
            -not $SeriesName.StartsWith(
                $prefixStr,
                [System.StringComparison]::OrdinalIgnoreCase
            )
        ) {
            $prefixedSeriesName = "$prefixStr$SeriesName"
        }

        $safeSeriesName =
            $SeriesName.Replace("'", "''")

        $safePrefixedName =
            $prefixedSeriesName.Replace("'", "''")

        $seriesCode = 0
        $sRst = $null

        try {
            $seriesQuery = @"
SELECT Code
FROM Master1
WHERE MasterType = 21
  AND (
        Name = '$safeSeriesName'
        OR Name = '$safePrefixedName'
      )
"@

            $sRst = $fi.GetRecordset($seriesQuery)

            if ($sRst -and -not $sRst.EOF) {
                $codeValue =
                    $sRst.Fields.Item("Code").Value

                if (
                    $null -ne $codeValue -and
                    $codeValue -ne [System.DBNull]::Value
                ) {
                    $seriesCode = [int]$codeValue
                }
            }
        }
        finally {
            if ($sRst) {
                try {
                    $sRst.Close()
                }
                catch {
                }

                try {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $sRst
                    ) | Out-Null
                }
                catch {
                }
            }
        }

        if ($seriesCode -le 0) {
            return @{
                success = $false
                error   = "Series '$SeriesName' not found"
            }
        }

        # ------------------------------------------------------------
        # Default numbering configuration
        # ------------------------------------------------------------

        $c1 = ""
        $c2 = ""
        $c3 = ""
        $c4 = ""

        $i1 = 0
        $i6 = 0
        $i7 = 0
        $i8 = 0

        $prefix = ""
        $suffix = ""
        $paddingLength = 0
        $isAuto = $false
        $frequency = 0
        $endingNo = 0L
        $startNo = 1L
        $gotConfig = $false

        # Helper for safely reading one numbering-config row.
        function Read-NumberingConfigRow {
            param(
                $Recordset
            )

            $result = @{
                c1            = ""
                c2            = ""
                c3            = ""
                c4            = ""
                i1            = 0
                frequency     = 0
                paddingLength = 0
                i6            = 0
                i7            = 0
                i8            = 0
                isAuto        = $false
                startNo       = 1L
                endingNo      = 0L
            }

            foreach ($fieldName in @("C1", "C2", "C3", "C4")) {
                try {
                    $value =
                        $Recordset.Fields.Item($fieldName).Value

                    if (
                        $null -ne $value -and
                        $value -ne [System.DBNull]::Value
                    ) {
                        $result[$fieldName.ToLower()] =
                            ([string]$value).Trim()
                    }
                }
                catch {
                }
            }

            foreach ($fieldName in @("I1", "I2", "I3", "I6", "I7", "I8")) {
                try {
                    $value =
                        $Recordset.Fields.Item($fieldName).Value

                    if (
                        $null -ne $value -and
                        $value -ne [System.DBNull]::Value
                    ) {
                        switch ($fieldName) {
                            "I1" {
                                $result.i1 = [int]$value
                            }

                            "I2" {
                                $result.frequency = [int]$value
                            }

                            "I3" {
                                $result.paddingLength = [int]$value
                            }

                            "I6" {
                                $result.i6 = [int]$value
                            }

                            "I7" {
                                $result.i7 = [int]$value
                            }

                            "I8" {
                                $result.i8 = [int]$value
                            }
                        }
                    }
                }
                catch {
                }
            }

            try {
                $value =
                    $Recordset.Fields.Item("L2").Value

                if (
                    $null -ne $value -and
                    $value -ne [System.DBNull]::Value
                ) {
                    $l2Value = [long]$value

                    $result.isAuto = ($l2Value -ge 1)

                    if ($l2Value -gt 1) {
                        $result.startNo = $l2Value
                    }
                }
            }
            catch {
            }

            try {
                $value =
                    $Recordset.Fields.Item("L3").Value

                if (
                    $null -ne $value -and
                    $value -ne [System.DBNull]::Value
                ) {
                    $result.endingNo = [long]$value
                }
            }
            catch {
            }

            return $result
        }

        # ------------------------------------------------------------
        # Read primary numbering configuration
        # ------------------------------------------------------------

        $cfgRst = $null

        try {
            $configQuery = @"
SELECT
    C1,
    C2,
    C3,
    C4,
    I1,
    I2,
    I3,
    I6,
    I7,
    I8,
    L2,
    L3
FROM Config
WHERE RecType = 6
  AND L1 = $seriesCode
"@

            $cfgRst = $fi.GetRecordset($configQuery)

            if ($cfgRst -and -not $cfgRst.EOF) {
                $gotConfig = $true
                $configValues =
                    Read-NumberingConfigRow -Recordset $cfgRst
            }
        }
        finally {
            if ($cfgRst) {
                try {
                    $cfgRst.Close()
                }
                catch {
                }

                try {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $cfgRst
                    ) | Out-Null
                }
                catch {
                }
            }
        }

        # ------------------------------------------------------------
        # Fallback numbering configuration
        # ------------------------------------------------------------

        if (-not $gotConfig) {
            $cfgRst2 = $null

            try {
                $fallbackQuery = @"
SELECT
    C1,
    C2,
    C3,
    C4,
    I1,
    I2,
    I3,
    I6,
    I7,
    I8,
    L2,
    L3
FROM Config
WHERE L1 = $seriesCode
  AND (I2 = 2 OR I2 = 3)
  AND L2 = 1
"@

                $cfgRst2 = $fi.GetRecordset($fallbackQuery)

                if ($cfgRst2 -and -not $cfgRst2.EOF) {
                    $gotConfig = $true
                    $configValues =
                        Read-NumberingConfigRow -Recordset $cfgRst2
                }
            }
            finally {
                if ($cfgRst2) {
                    try {
                        $cfgRst2.Close()
                    }
                    catch {
                    }

                    try {
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                            $cfgRst2
                        ) | Out-Null
                    }
                    catch {
                    }
                }
            }
        }

        if (-not $gotConfig) {
            return @{
                success = $false
                error   = "Numbering configuration not found for series '$SeriesName'"
            }
        }

        $c1 = [string]$configValues.c1
        $c2 = [string]$configValues.c2
        $c3 = [string]$configValues.c3
        $c4 = [string]$configValues.c4

        $i1 = [int]$configValues.i1
        $frequency = [int]$configValues.frequency
        $paddingLength = [int]$configValues.paddingLength
        $i6 = [int]$configValues.i6
        $i7 = [int]$configValues.i7
        $i8 = [int]$configValues.i8

        $isAuto = [bool]$configValues.isAuto
        $startNo = [long]$configValues.startNo
        $endingNo = [long]$configValues.endingNo

        # ------------------------------------------------------------
        # Resolve visible voucher-number format
        # ------------------------------------------------------------

        $padChar = if (-not [string]::IsNullOrEmpty($c3)) {
            [char]$c3[0]
        }
        else {
            [char]" "
        }

        $sep = $c4
        $prefix = $c2
        $suffix = $c1
        $embedPos = "none"

        if ($i1 -gt 0) {
            switch ($i7) {
                1 {
                    $embedPos = "prefix"
                }

                2 {
                    $embedPos = "suffix"
                }

                0 {
                    $embedPos = "none"
                }

                default {
                    if ($i6 -eq 2) {
                        $embedPos = "prefix"
                    }
                    elseif ($i6 -eq 1) {
                        $embedPos = "suffix"
                    }
                    elseif ($c1 -ne "") {
                        $embedPos = "prefix"
                    }
                    elseif ($c2 -ne "") {
                        $embedPos = "suffix"
                    }
                    else {
                        $embedPos = "suffix"
                    }
                }
            }
        }

        # ------------------------------------------------------------
        # Build date text used inside visible voucher numbers
        # ------------------------------------------------------------

        $now = Get-Date

        $monthAbbr =
            $now.ToString(
                "MMM",
                [System.Globalization.CultureInfo]::InvariantCulture
            ).ToUpperInvariant()

        $financialYear = if ($now.Month -ge 4) {
            $now.Year
        }
        else {
            $now.Year - 1
        }

        $financialYearNext =
            $financialYear + 1

        $fyShort =
            ($financialYear % 100).ToString("D2")

        $fyNextShort =
            ($financialYearNext % 100).ToString("D2")

        $dateStr = ""

        if ($isAuto -and $i1 -gt 0) {
            if ($frequency -eq 3) {
                $dateStr = if ($i8 -eq 1) {
                    "$fyShort-$fyNextShort"
                }
                else {
                    "$financialYear-$fyNextShort"
                }
            }
            elseif ($frequency -eq 2) {
                switch ($i8) {
                    2 {
                        $dateStr = $now.ToString("MMyy")
                    }

                    3 {
                        $dateStr = $now.ToString("MM-yyyy")
                    }

                    4 {
                        $dateStr = $now.ToString("MMyyyy")
                    }

                    11 {
                        $dateStr =
                            "$monthAbbr-$($now.ToString('yyyy'))"
                    }

                    12 {
                        $dateStr =
                            "$monthAbbr-$($now.ToString('yy'))"
                    }

                    13 {
                        $dateStr = $now.ToString("MM-yyyy")
                    }

                    14 {
                        $dateStr = $now.ToString("MM-yy")
                    }

                    15 {
                        $dateStr = $now.ToString("MMyy")
                    }

                    16 {
                        $dateStr = $now.ToString("yyMM")
                    }

                    default {
                        $dateStr = $now.ToString("MMyy")
                    }
                }
            }
        }

        # ------------------------------------------------------------
        # Build date filter for current frequency period
        # ------------------------------------------------------------

        $dateFilter = ""

        if ($isAuto -and $frequency -gt 0) {
            $periodStart = $null
            $periodEndExclusive = $null

            if ($frequency -eq 1) {
                $periodStart = $now.Date
                $periodEndExclusive =
                    $periodStart.AddDays(1)
            }
            elseif ($frequency -eq 2) {
                $periodStart = New-Object DateTime `
                    $now.Year,
                    $now.Month,
                    1

                $periodEndExclusive =
                    $periodStart.AddMonths(1)
            }
            elseif ($frequency -eq 3) {
                $fyStartYear = if ($now.Month -ge 4) {
                    $now.Year
                }
                else {
                    $now.Year - 1
                }

                $periodStart = New-Object DateTime `
                    $fyStartYear,
                    4,
                    1

                $periodEndExclusive =
                    $periodStart.AddYears(1)
            }

            if (
                $null -ne $periodStart -and
                $null -ne $periodEndExclusive
            ) {
                if ($isSql) {
                    $sqlStart =
                        $periodStart.ToString(
                            "yyyy-MM-ddTHH:mm:ss",
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )

                    $sqlEnd =
                        $periodEndExclusive.ToString(
                            "yyyy-MM-ddTHH:mm:ss",
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )

                    $dateFilter = (
                        " AND [Date] >= '$sqlStart'" +
                        " AND [Date] < '$sqlEnd'"
                    )
                }
                else {
                    $accessStart =
                        $periodStart.ToString(
                            "MM/dd/yyyy",
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )

                    $accessEnd =
                        $periodEndExclusive.ToString(
                            "MM/dd/yyyy",
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )

                    $dateFilter = (
                        " AND [Date] >= #$accessStart#" +
                        " AND [Date] < #$accessEnd#"
                    )
                }
            }
        }

        # ------------------------------------------------------------
        # Find highest visible voucher sequence
        # ------------------------------------------------------------

        $lastVisibleSeq = 0L
        $lastAutoSeq = 0L
        $tranRst = $null

        try {
            $tranQuery = (
                "SELECT AutoVchNo, VchNo " +
                "FROM Tran1 " +
                "WHERE VchType = $VchType " +
                "AND VchSeriesCode = $seriesCode " +
                "AND Cancelled = 0 " +
                "AND VchCancelled = 0" +
                $dateFilter
            )

            $tranRst = $fi.GetRecordset($tranQuery)

            if ($tranRst -and -not $tranRst.EOF) {
                $tranRst.MoveFirst()

                while (-not $tranRst.EOF) {
                    $dbAutoNo = 0L
                    $dbVchNo = ""

                    try {
                        $autoValue =
                            $tranRst.Fields.Item("AutoVchNo").Value

                        if (
                            $null -ne $autoValue -and
                            $autoValue -ne [System.DBNull]::Value
                        ) {
                            $dbAutoNo = [long]$autoValue
                        }
                    }
                    catch {
                    }

                    try {
                        $vchNoValue =
                            $tranRst.Fields.Item("VchNo").Value

                        if (
                            $null -ne $vchNoValue -and
                            $vchNoValue -ne [System.DBNull]::Value
                        ) {
                            $dbVchNo =
                                ([string]$vchNoValue).Trim()
                        }
                    }
                    catch {
                    }

                    if ($dbAutoNo -gt $lastAutoSeq) {
                        $lastAutoSeq = $dbAutoNo
                    }

                    if (
                        -not [string]::IsNullOrWhiteSpace($dbVchNo) -and
                        $isAuto
                    ) {
                        $temp = $dbVchNo.Trim()

                        if (
                            $prefix -ne "" -and
                            $temp.StartsWith(
                                $prefix,
                                [System.StringComparison]::OrdinalIgnoreCase
                            )
                        ) {
                            $temp =
                                $temp.Substring($prefix.Length)
                        }

                        if (
                            $suffix -ne "" -and
                            $temp.EndsWith(
                                $suffix,
                                [System.StringComparison]::OrdinalIgnoreCase
                            )
                        ) {
                            $temp =
                                $temp.Substring(
                                    0,
                                    $temp.Length - $suffix.Length
                                )
                        }

                        if (
                            $dateStr -ne "" -and
                            $temp.IndexOf(
                                $dateStr,
                                [System.StringComparison]::OrdinalIgnoreCase
                            ) -ge 0
                        ) {
                            $temp =
                                [regex]::Replace(
                                    $temp,
                                    [regex]::Escape($dateStr),
                                    "",
                                    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
                                )
                        }

                        # Remove the configured separator.
                        if ($sep -ne "") {
                            $temp =
                                $temp.Replace($sep, "")
                        }

                        # Do not remove the padding character.
                        # Leading zeroes are valid and TryParse handles them.
                        $temp = $temp.Trim()

                        $parsedVisible = 0L

                        if (
                            [long]::TryParse(
                                $temp,
                                [ref]$parsedVisible
                            ) -and
                            $parsedVisible -gt $lastVisibleSeq
                        ) {
                            $lastVisibleSeq =
                                $parsedVisible
                        }
                    }
                    elseif (
                        -not [string]::IsNullOrWhiteSpace($dbVchNo) -and
                        -not $isAuto
                    ) {
                        $parsedManual = 0L

                        if (
                            [long]::TryParse(
                                $dbVchNo,
                                [ref]$parsedManual
                            ) -and
                            $parsedManual -gt $lastVisibleSeq
                        ) {
                            $lastVisibleSeq =
                                $parsedManual
                        }
                    }

                    $tranRst.MoveNext()
                }
            }
        }
        finally {
            if ($tranRst) {
                try {
                    $tranRst.Close()
                }
                catch {
                }

                try {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $tranRst
                    ) | Out-Null
                }
                catch {
                }
            }
        }

        # Prefer visible voucher number because AutoVchNo can use a
        # different internal sequence.
        $lastSeq = if ($lastVisibleSeq -gt 0) {
            $lastVisibleSeq
        }
        elseif ($lastAutoSeq -gt 0) {
            $lastAutoSeq
        }
        else {
            0L
        }

        $currentNo = if ($lastSeq -gt 0) {
            $lastSeq + 1
        }
        else {
            $startNo
        }

        if (
            $endingNo -gt 0 -and
            $currentNo -gt $endingNo
        ) {
            return @{
                success = $false
                error   = (
                    "Next voucher number $currentNo exceeds configured " +
                    "ending number $endingNo"
                )
            }
        }

        function Build-LocalVchNo {
            param(
                [long]$Sequence
            )

            $numberText =
                [string]$Sequence

            if ($paddingLength -gt 0) {
                $numberText =
                    $numberText.PadLeft(
                        $paddingLength,
                        $padChar
                    )
            }

            switch ($embedPos) {
                "suffix" {
                    return (
                        "$prefix$numberText$sep$dateStr"
                    ).Trim()
                }

                "prefix" {
                    return (
                        "$dateStr$sep$numberText$suffix"
                    ).Trim()
                }

                default {
                    return (
                        "$prefix$numberText$suffix"
                    ).Trim()
                }
            }
        }

        $nextVchNo = if ($isAuto) {
            Build-LocalVchNo -Sequence $currentNo
        }
        else {
            [string]$currentNo
        }

        $lastVchNo = if ($lastSeq -gt 0) {
            Build-LocalVchNo -Sequence $lastSeq
        }
        else {
            "(none yet)"
        }

        return @{
            success = $true

            data = @{
                vch_type       = $VchType
                series_name    = $SeriesName
                prefix         = $prefix
                suffix         = $suffix
                padding_length = $paddingLength
                current_no     = $currentNo
                is_auto        = $isAuto
                frequency      = $frequency
                ending_no      = $endingNo
                next_vch_no    = $nextVchNo
                last_vch_no    = $lastVchNo
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
        if ($fi) {
            try {
                Disconnect-BUSY $fi
            }
            catch {
            }
        }
    }
}

function Test-VoucherNumberExists {
    param(
        $fi,
        [int]$VchType,
        [string]$SeriesName,
        [string]$VchNo
    )

    if ($null -eq $fi) {
        throw "BUSY connection is not available."
    }

    if ([string]::IsNullOrWhiteSpace($SeriesName)) {
        throw "Voucher series is required for duplicate checking."
    }

    if ([string]::IsNullOrWhiteSpace($VchNo)) {
        throw "Voucher number is required for duplicate checking."
    }

    $cleanSeriesName = $SeriesName.Trim()
    $cleanVchNo = $VchNo.Trim()

    try {
        # BUSY sometimes stores the series master name with the voucher
        # type prefix, for example LP-1518 or 09LP-1518.
        $typePrefix = "{0:D2}" -f $VchType
        $prefixedSeriesName = $cleanSeriesName

        if (
            -not $cleanSeriesName.StartsWith(
                $typePrefix,
                [System.StringComparison]::OrdinalIgnoreCase
            )
        ) {
            $prefixedSeriesName = "$typePrefix$cleanSeriesName"
        }

        $safeSeriesName =
            $cleanSeriesName.Replace("'", "''")

        $safePrefixedSeriesName =
            $prefixedSeriesName.Replace("'", "''")

        # ------------------------------------------------------------
        # Resolve the voucher-series code
        # ------------------------------------------------------------

        $seriesCode = 0
        $seriesRst = $null

        try {
            $seriesQuery = @"
SELECT Code
FROM Master1
WHERE MasterType = 21
  AND (
        Name = '$safeSeriesName'
        OR Name = '$safePrefixedSeriesName'
      )
"@

            $seriesRst = $fi.GetRecordset($seriesQuery)

            if ($seriesRst -and -not $seriesRst.EOF) {
                $codeValue =
                    $seriesRst.Fields.Item("Code").Value

                if (
                    $null -ne $codeValue -and
                    $codeValue -ne [System.DBNull]::Value
                ) {
                    $seriesCode = [int]$codeValue
                }
            }
        }
        finally {
            if ($seriesRst) {
                try {
                    $seriesRst.Close()
                }
                catch {
                }

                try {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $seriesRst
                    ) | Out-Null
                }
                catch {
                }
            }
        }

        if ($seriesCode -le 0) {
            throw "Voucher series '$cleanSeriesName' was not found."
        }

        # BUSY can store VchNo as a fixed-width field with leading spaces.
        # LTrim/RTrim is supported through the BUSY recordset provider for
        # both SQL and Access companies.
        $safeVchNo =
            $cleanVchNo.Replace("'", "''")

        $paddedVchNo = $cleanVchNo

        try {
            $paddedVchNo = Pad-VchNo $cleanVchNo
        }
        catch {
            # Continue with the exact number when Pad-VchNo is unavailable.
            $paddedVchNo = $cleanVchNo
        }

        $safePaddedVchNo =
            ([string]$paddedVchNo).Replace("'", "''")

        # ------------------------------------------------------------
        # Check for the same voucher type, series and visible number
        # ------------------------------------------------------------

        $duplicateQuery = @"
SELECT VchCode, VchNo
FROM Tran1
WHERE VchType = $VchType
  AND VchSeriesCode = $seriesCode
  AND Cancelled = 0
  AND VchCancelled = 0
  AND (
        LTRIM(RTRIM(VchNo)) = '$safeVchNo'
        OR VchNo = '$safePaddedVchNo'
      )
"@

        $duplicateRst = $null

        try {
            $duplicateRst =
                $fi.GetRecordset($duplicateQuery)

            return (
                $duplicateRst -and
                -not $duplicateRst.EOF
            )
        }
        finally {
            if ($duplicateRst) {
                try {
                    $duplicateRst.Close()
                }
                catch {
                }

                try {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $duplicateRst
                    ) | Out-Null
                }
                catch {
                }
            }
        }
    }
    catch {
        throw (
            "Duplicate voucher check failed: " +
            $_.Exception.Message
        )
    }
}

function Create-Voucher {
    param($Data, [string]$InstanceId = "", [string]$CompanyCode = "")

    $vchType = [int]$Data.vchType
    $cfg = $script:VoucherConfig[$vchType]
    if (-not $cfg) { return @{ success=$false; error="Unsupported vchType: $vchType" } }

    $validationError = Validate-VoucherData $Data $cfg
    if ($validationError) { return @{ success=$false; error=$validationError } }

    $isChallanType = ($vchType -eq 11 -or $vchType -eq 4)
    $maxAttempts = 2
    $attempt = 1
    $lastExceptionMsg = ""

    while ($attempt -le $maxAttempts) {
        $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
        if (-not $fi) { return @{ success=$false; error="BUSY connection failed" } }

        try {
            $seriesName = ([string]$Data.vchSeries).Trim()
            $busyDateText = [string]$Data.date
            $voucherDate = [datetime]::Now.Date
            if ($busyDateText -match '^\d{2}-\d{2}-\d{4}$') {
                $voucherDate = [datetime]::ParseExact($busyDateText, 'dd-MM-yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
            } elseif ($busyDateText -match '^\d{4}-\d{2}-\d{2}$') {
                $voucherDate = [datetime]::ParseExact($busyDateText, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
            }

            # Read the administrator-selected source for this type + series.
            $adminConfigResult = Get-WebNumberingConfig `
                -VchType $vchType `
                -SeriesName $seriesName `
                -VoucherDate $voucherDate.ToString('yyyy-MM-dd') `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode

            if (-not $adminConfigResult.success) { throw $adminConfigResult.error }
            $adminConfig = $adminConfigResult.data

            # WEB source is authoritative. BUSY source keeps the number submitted by
            # the existing BUSY numbering flow.
            $vchNo = if ($adminConfig.source -eq 'WEB' -and $adminConfig.is_active) {
                ([string]$adminConfig.next_vch_no).Trim()
            } else {
                ([string]$Data.vchNo).Trim()
            }

            if ([string]::IsNullOrWhiteSpace($vchNo)) { return @{ success=$false; error='Voucher number is required.' } }

            $exists = Test-VoucherNumberExists -fi $fi -VchType $vchType -SeriesName $seriesName -VchNo $vchNo
            if ($exists) {
                return @{ success=$false; error="Voucher number '$vchNo' already exists in series '$seriesName'. Refresh and try again." }
            }

            # Ensure XML builders receive the final backend-selected number.
            $Data.vchNo = $vchNo
            $xml = if ($cfg.isAccounting) {
                Build-AccountingVoucherXml -Data $Data -Cfg $cfg -VchType $vchType -VchNo $vchNo -SkipBBA $false -fi $fi
            } else {
                Build-VoucherXml -Data $Data -Cfg $cfg -VchType $vchType -VchNo $vchNo -SkipBBA $isChallanType -fi $fi
            }

            $errMsg = ""
            $saved = $fi.SaveVchFromXML($vchType, $xml, [ref]$errMsg)
            if ($saved -ne $true) {
                return @{ success=$false; error=if ($errMsg) { $errMsg } else { 'Unknown BUSY error' } }
            }

            if ($Data.bridgeUserName) {
                Update-CheckListCreator -fi $fi -VchType $vchType -VchNo $vchNo -VchDate $Data.date -UserName $Data.bridgeUserName -InstanceId $InstanceId -CompanyCode $CompanyCode
            }

            Clear-StockCaches -InstanceId $InstanceId -CompanyCode $CompanyCode
            return @{
                success=$true
                message="$($cfg.xmlRoot) created successfully"
                data=@{ vchType=$vchType; vchSeries=$seriesName; vchNo=$vchNo; date=$Data.date; party=$Data.party; numberingSource=$adminConfig.source }
            }
        } catch {
            $lastExceptionMsg = $_.Exception.Message
            $script:ActiveConnection = $null
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($fi) | Out-Null } catch {}
            $attempt++
        } finally {
            Disconnect-BUSY $fi
        }
    }

    return @{ success=$false; error="Database error. Connection reset. Details: $lastExceptionMsg" }
}

function Get-Vouchers {
    param([int]$VchType,
        [string]$From = "",
        [string]$To = "",[string]$Party = "",
        [string]$Search = "",
        [string]$InstanceId = "",[string]$CompanyCode = ""
    )

    $cfg = $script:VoucherConfig[$VchType]
    if (-not $cfg) { return @{ success = $false; error = "Unsupported vchType" } }
    
    # Resolve if SQL or Access from instances.json
    $targetInst = Get-InstanceConfig -InstanceId $InstanceId
    $dbType = 0
    if ($null -ne $targetInst -and $null -ne $targetInst.dbType) { $dbType = [int]$targetInst.dbType }
    $isSql = ($dbType -eq 1)

    if ($isSql) {
        # ── SQL PATH: Direct Native SQL Connection (Incredibly stable & fast!) ──
        $dirConn = Get-DirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
        if ($null -eq $dirConn) {
            return @{ success = $false; error = "Could not build direct database connection" }
        }
        $sqlConn = $dirConn.connection
        $vouchers = @()
        try {
            $sqlConn.Open()

            # 1. Fetch masterMap (Code -> Name)
            $masterMap = @{}
            $mCmd = $sqlConn.CreateCommand()
            $mCmd.CommandText = "SELECT [Code], [Name] FROM [Master1]"
            $mRdr = $mCmd.ExecuteReader()
            while ($mRdr.Read()) {
                $c = $mRdr.GetValue(0)
                $n = $mRdr.GetValue(1)
                if ($null -ne $c -and "$c" -ne "") {
                    $masterMap[$c.ToString().Trim()] = if ($null -ne $n -and "$n" -ne "") { $n.ToString().Trim() } else { "" }
                }
            }
            $mRdr.Close()

            # 2. Build Where Filter (Using culture-insensitive Safe-ParseDate helper with single quotes for SQL)
            $where = "VchType = $VchType"
            if ($From -ne "") {
                $dFrom = Safe-ParseDate -dateStr $From
                if ($null -ne $dFrom) {
                    $where += " AND [Date] >= '" + $dFrom.ToString("yyyy-MM-dd") + "'"
                }
            }
            if ($To -ne "") {
                $dTo = Safe-ParseDate -dateStr $To
                if ($null -ne $dTo) {
                    $where += " AND [Date] <= '" + $dTo.ToString("yyyy-MM-dd") + "'"
                }
            }
            if ($Party -ne "") {
                $pCode = ""
                foreach ($k in $masterMap.Keys) {
                    if ($masterMap[$k] -eq $Party) { $pCode = $k; break }
                }
                if ($pCode -ne "") { $where += " AND MasterCode1 = $pCode" }
            }

            # 3. Query Vouchers (SQL direct query is incredibly robust)
            $vCmd = $sqlConn.CreateCommand()
            $stptColExists = $false
            try {
                $colCmd = $sqlConn.CreateCommand()
                $colCmd.CommandText = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'Tran1' AND COLUMN_NAME = 'STPTCode'"
                $stptColExists = ([int]$colCmd.ExecuteScalar() -gt 0)
            } catch {}

            $stptSelect = if ($stptColExists) { "[STPTCode]," } else { "NULL AS [STPTCode]," }

            $vCmd = $sqlConn.CreateCommand()
            $vCmd.CommandText = "SELECT [VchCode], [VchNo], [AutoVchNo], [Date], [MasterCode1], [MasterCode2], [VchSeriesCode], $stptSelect [CM1], [VchAmtBaseCur] FROM [Tran1] WHERE $where ORDER BY COALESCE(TRY_CAST(VchNo AS DECIMAL(18,2)), 0) DESC, VchNo DESC"
            $rdr = $vCmd.ExecuteReader()
            while ($rdr.Read()) {
                $vNo = ""
                $vchCode = 0
                $partyName = ""
                $seriesName = ""
                $typeName = ""
                $amt = 0.0
                $strDate = ""

                try { $vchCode = [int]$rdr["VchCode"] } catch {}

                try {
                    $dbVchNo = $rdr["VchNo"]
                    if ($null -ne $dbVchNo -and "$dbVchNo" -ne "") { $vNo = $dbVchNo.ToString().Trim() }
                    if ($vNo -eq "") {
                        $av = $rdr["AutoVchNo"]
                        if ($null -ne $av -and "$av" -ne "") { $vNo = $av.ToString().Trim() }
                    }
                } catch {}

                try {
                    $rawDate = $rdr["Date"]
                    if ($rawDate -is [datetime]) { $strDate = $rawDate.ToString("dd-MM-yyyy") }
                    elseif ($null -ne $rawDate -and "$rawDate" -ne "") { $strDate = [datetime]::Parse($rawDate.ToString()).ToString("dd-MM-yyyy") }
                } catch {}

                $pCode = ""
                try { $pCode = $rdr["MasterCode1"].ToString().Trim() } catch {}
                $partyName = if ($masterMap.ContainsKey($pCode)) { $masterMap[$pCode] } else { $pCode }

                $p2Code = ""
                try { $p2Code = $rdr["MasterCode2"].ToString().Trim() } catch {}
                $toCenterName = if ($masterMap.ContainsKey($p2Code)) { $masterMap[$p2Code] } else { $p2Code }

                $sCode = ""
                try { $sCode = $rdr["VchSeriesCode"].ToString().Trim() } catch {}
                $seriesName = if ($masterMap.ContainsKey($sCode)) { $masterMap[$sCode] } else { $sCode }

                $tCode = ""
                try { $tCode = $rdr["STPTCode"].ToString().Trim() } catch {
                    try { $tCode = $rdr["CM1"].ToString().Trim() } catch {}
                }
                $typeName = if ($masterMap.ContainsKey($tCode)) { $masterMap[$tCode] } else { $tCode }

                try {
                    $v = $rdr["VchAmtBaseCur"]
                    if ($null -ne $v -and "$v" -ne "") { $amt = [Math]::Abs([Convert]::ToDouble($v.ToString())) }
                } catch {}

                $vouchers += @{
                    vchNo = $vNo
                    vchSeries = $seriesName
                    date = $strDate
                    party = $partyName
                    matCentre = $toCenterName
                    debitAccount = "—"
                    creditAccount = "—"
                    voucherType = $typeName
                    amount = $amt
                    vchType = $VchType
                    vchCode = $vchCode
                }
            }
            $rdr.Close()

            # Debit/Credit accounts if financial voucher
            if ($VchType -in @(14, 15, 16, 19) -and $vouchers.Count -gt 0) {
                for ($idx = 0; $idx -lt $vouchers.Count; $idx++) {
                    $vch = $vouchers[$idx]
                    $vchCode = $vch.vchCode
                    if ($vchCode -gt 0) {
                        try {
                            $accCmd = $sqlConn.CreateCommand()
                            $accCmd.CommandText = "SELECT [MasterCode1], [Value1] FROM [Tran2] WHERE [VchCode] = $vchCode AND [RecType] IN (1,2)"
                            $accRdr = $accCmd.ExecuteReader()
                            $dAccList = @()
                            $cAccList = @()
                            while ($accRdr.Read()) {
                                $mc1 = $accRdr.GetValue(0)
                                $val1 = $accRdr.GetValue(1)
                                if ($null -ne $mc1 -and "$mc1" -ne "" -and $null -ne $val1 -and "$val1" -ne "") {
                                    $amt2 = [double]$val1
                                    $accStr = $mc1.ToString().Trim()
                                    $accName = if ($masterMap.ContainsKey($accStr)) { $masterMap[$accStr] } else { $accStr }
                                    if ($amt2 -gt 0) { $dAccList += $accName } elseif ($amt2 -lt 0) { $cAccList += $accName }
                                }
                            }
                            $accRdr.Close()
                            $vouchers[$idx].debitAccount = if ($dAccList.Count -gt 0) { $dAccList -join ", " } else { "—" }
                            $vouchers[$idx].creditAccount = if ($cAccList.Count -gt 0) { $cAccList -join ", " } else { "—" }
                        } catch {}
                    }
                }
            }

            # Remove temporary vchCode
            $outputVouchers = @()
            foreach ($v in $vouchers) {
                $v.Remove("vchCode")
                $outputVouchers += $v
            }
            return @{ success = $true; count = $outputVouchers.Count; vchType = $VchType; data = $outputVouchers }
        } catch {
            return @{ success = $false; error = $_.Exception.Message }
        } finally {
            if ($null -ne $sqlConn) { try { $sqlConn.Close() } catch {} }
        }
    } else {
        # ── ACCESS PATH: Original stable COM recordset query ──
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
                        debitAccount = if ($debitAcc -ne "") { $debitAcc } else { "—" }
                        creditAccount = if ($creditAcc -ne "") { $creditAcc } else { "—" }
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
        $dirConn = Get-DirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
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

        $matCentre = ""
        $party = ""
        $bomName = ""
        $extraExpenseInBOM = 0.0

        if ($VchType -eq 6) {

            # ─────────────────────────────────────────────
            # PRODUCTION VOUCHER
            #
            # MasterName1 = Generated Material Centre
            # ConMCName   = Consumed Material Centre
            # MasterName2 = BOM Name
            # ─────────────────────────────────────────────

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
                $dirConn = Get-DirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
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
                $dirConn = Get-DirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
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
            $dirConn = Get-DirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
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
        # ── SQL PATH: Resolve VchCode natively using direct SQL connection ──
        $dirConn = Get-DirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
        if ($null -eq $dirConn) { return @{ success = $false; error = "Could not build direct database connection" } }
        $conn = $dirConn.connection
        try {
            $conn.Open()
            $vchCode = Get-VchCode-Direct -conn $conn -VchType $VchType -VchNo $VchNo -VchDate $VchDate -isSql $true
        } finally {
            if ($null -ne $conn) { try { $conn.Close() } catch {} }
        }
    } else {
        # ── ACCESS PATH: Original stable COM recordset query ──
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

# ═══════════════════════════════════════════════════════════════
#  SEARCH ORIGINAL VOUCHERS (Specifically for BBA)
# ═══════════════════════════════════════════════════════════════
# ═══════════════════════════════════════════════════════════════
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
# ═══════════════════════════════════════════════════════════════
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
function Modify-Voucher {
    param($Data, [string]$InstanceId = "",[string]$CompanyCode = "")

    $vchType = [int]$Data.vchType
    $cfg = $script:VoucherConfig[$vchType]
    
    if (-not $cfg) { return @{ success = $false; error = "Unsupported vchType" } }
    if (-not $Data.vchNo -or [string]$Data.vchNo -eq "") { return @{ success = $false; error = "vchNo is required for modify" } }
    
    $maxAttempts = 2
    $attempt = 1
    $lastExceptionMsg = ""

    while ($attempt -le $maxAttempts) {
        $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
        if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
        
        try {
            $xml = if ($cfg.isAccounting) {
                Build-AccountingVoucherXml -Data $Data -Cfg $cfg -VchType $vchType -VchNo ([string]$Data.vchNo) -SkipBBA $false -fi $fi
            } else {
                Build-VoucherXml -Data $Data -Cfg $cfg -VchType $vchType -VchNo ([string]$Data.vchNo) -SkipBBA $false -fi $fi
            }

            $errMsg = ""
            $saved = $fi.SaveVchFromXML($vchType, $xml, [ref]$errMsg, $true)
            
            if ($saved -eq $true) {
                if ($Data.bridgeUserName) {
                    Update-CheckListCreator -fi $fi -VchType $vchType -VchNo ([string]$Data.vchNo) -VchDate $Data.date -UserName $Data.bridgeUserName -InstanceId $InstanceId -CompanyCode $CompanyCode
                }

                Clear-StockCaches -InstanceId $InstanceId -CompanyCode $CompanyCode
                return @{
                    success = $true
                    message = "$($cfg.xmlRoot) modified"
                    data = @{
                        vchType = $vchType
                        vchSeries = $Data.vchSeries
                        vchNo = $Data.vchNo
                        date = $Data.date
                    }
                }
            } else {
                return @{ success = $false; error = if ($errMsg) { $errMsg } else { "Unknown BUSY error" } }
            }
        } catch {
            $lastExceptionMsg = $_.Exception.Message
            $script:ActiveConnection = $null
            try { $script:ActiveConnection.CloseDB() } catch {}
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($fi) | Out-Null } catch {}
            
            if ($attempt -eq 1) {
                Write-Host "2026-06-26 [WARNING] Idle connection timeout detected on Modify. Reconnecting for Attempt 2..." -ForegroundColor Yellow
            }
            $attempt++
        } finally {
            Disconnect-BUSY $fi
        }
    }

    return @{ success = $false; error = "Database error. Connection reset. Details: $lastExceptionMsg" }
}



# ═══════════════════════════════════════════════════════════════
#  DELETE VOUCHER
# ═══════════════════════════════════════════════════════════════
function Delete-Voucher {
    param($Data, [string]$InstanceId = "", [string]$CompanyCode = "")

    $cfg = $script:VoucherConfig[[int]$Data.vchType]
    if (-not $cfg) { return @{ success = $false; error = "Unsupported vchType" } }
    
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
    
    try {
        $errMsg = ""
        $deleted = $fi.DeleteVch($Data.vchType, $Data.vchSeries, $Data.vchDate, $Data.vchNo, [ref]$errMsg)
        
        if ($deleted -eq $true) {
            Clear-StockCaches -InstanceId $InstanceId -CompanyCode $CompanyCode
            return @{ success = $true; message = "Voucher deleted successfully" }
        } else {
            return @{ success = $false; error = if ($errMsg) { $errMsg } else { "Unknown BUSY error" } }
        }
    } catch {
        $script:ActiveConnection = $null
        try { $script:ActiveConnection.CloseDB() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($fi) | Out-Null } catch {}
        return @{ success = $false; error = "Database error. Connection reset. Please try deleting again." }
    } finally {
        Disconnect-BUSY $fi
    }
}