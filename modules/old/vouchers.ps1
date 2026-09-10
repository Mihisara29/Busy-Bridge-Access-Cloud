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
    26 = @{ xmlRoot = "SaleQuotation";     typeField = "STPTName"; typeDataKey = "saleType";     hasBillNo = $true; isAccounting = $false; requiredKeys = @("vchSeries","date","party","saleType","items") }
    27 = @{ xmlRoot = "PurchaseQuotation"; typeField = "STPTName"; typeDataKey = "purchaseType"; hasBillNo = $true; isAccounting = $false; requiredKeys = @("vchSeries","date","party","purchaseType","items") }

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
        [string]$VchSeries = "",
        [string]$UserName,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    # Historical name retained for compatibility. On MODIFY this function must
    # update only the newest CheckList Action=2 row. Updating every row would
    # overwrite Created By (Action=1) and Approved By (Action=3).
    if ([string]::IsNullOrWhiteSpace($UserName) -or [string]::IsNullOrWhiteSpace($VchNo)) { return }

    try {
        $state = Get-VoucherApprovalStateByIdentity `
            -VchType $VchType `
            -VchNo $VchNo `
            -VchDate $VchDate `
            -VchSeries $VchSeries `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode `
            -ExistingFi $fi

        if (-not $state.success -or -not $state.found -or [int]$state.vchCode -le 0) { return }

        $vchCode = [int]$state.vchCode
        $safeUser = $UserName.Trim().Replace("'", "''")
        $fi.ExecuteQuery(@"
UPDATE CheckList
SET UserName='$safeUser'
WHERE Code=$vchCode
  AND Action=2
  AND ActionTime=(SELECT MAX(ActionTime) FROM CheckList WHERE Code=$vchCode AND Action=2)
"@)
    } catch {
        Write-Host "  [WARN] Could not update latest modifier audit row: $($_.Exception.Message)" -ForegroundColor DarkYellow
    }
}

function Get-VoucherApprovalStateByIdentity {
    param(
        [int]$VchType,
        [string]$VchNo,
        [string]$VchDate = "",
        [string]$VchSeries = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        $ExistingFi = $null
    )

    if ($VchType -le 0 -or [string]::IsNullOrWhiteSpace($VchNo)) {
        return @{ success=$false; found=$false; error="vchType and vchNo are required." }
    }

    $fi = $ExistingFi
    $ownsConnection = $false
    if (-not $fi) {
        $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
        $ownsConnection = $true
    }
    if (-not $fi) { return @{ success=$false; found=$false; error="BUSY database connection failed" } }

    try {
        $dbType = 0
        $foundCompany = Get-InstanceForCompany -CompanyCode $CompanyCode -InstanceId $InstanceId
        if ($foundCompany -and $null -ne $foundCompany.instance.dbType) { $dbType = [int]$foundCompany.instance.dbType }
        $wildcard = if ($dbType -eq 1) { "%" } else { "*" }

        $targetNo = $VchNo.Trim()
        $targetSeries = ([string]$VchSeries).Trim()
        $targetDate = $null
        if (-not [string]::IsNullOrWhiteSpace($VchDate)) {
            try { $targetDate = Safe-ParseDate ([string]$VchDate) } catch {}
        }

        $safeNo = $targetNo.Replace("'", "''")
        $rst = $fi.GetRecordset(@"
SELECT VchCode,VchType,VchNo,VchSeriesCode,[Date],ApprovalStatus
FROM Tran1
WHERE VchType=$VchType
  AND VchNo LIKE '$wildcard$safeNo$wildcard'
ORDER BY VchCode DESC
"@)

        if ($rst) {
            while (-not $rst.EOF) {
                $dbNo = ""
                try { $dbNo = ([string]$rst.Fields.Item("VchNo").Value).Trim() } catch {}

                if ($dbNo.Equals($targetNo, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $seriesCode = 0
                    $dateValue = $null
                    try { $seriesCode = [int]$rst.Fields.Item("VchSeriesCode").Value } catch {}
                    try { $dateValue = $rst.Fields.Item("Date").Value } catch {}

                    $dateMatches = $true
                    if ($null -ne $targetDate -and $null -ne $dateValue) {
                        try { $dateMatches = ([datetime]$dateValue).Date -eq ([datetime]$targetDate).Date } catch { $dateMatches = $false }
                    }

                    $seriesMatches = $true
                    $seriesName = ""
                    if (-not [string]::IsNullOrWhiteSpace($targetSeries)) {
                        $seriesMatches = $false
                        if ($seriesCode -gt 0) {
                            $sr = $fi.GetRecordset("SELECT Name FROM Master1 WHERE Code=$seriesCode")
                            if ($sr -and -not $sr.EOF) {
                                try { $seriesName = ([string]$sr.Fields.Item("Name").Value).Trim() } catch {}
                            }
                            if ($sr) { try { $sr.Close() } catch {} }
                        }

                        $prefix = "{0:D2}" -f $VchType
                        if ($seriesName.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                            $seriesName = $seriesName.Substring($prefix.Length)
                        }

                        $seriesMatches = $seriesName.Equals($targetSeries, [System.StringComparison]::OrdinalIgnoreCase)
                    }

                    if ($dateMatches -and $seriesMatches) {
                        $result = @{
                            success = $true
                            found = $true
                            vchCode = [int]$rst.Fields.Item("VchCode").Value
                            vchType = [int]$rst.Fields.Item("VchType").Value
                            vchNo = $dbNo
                            vchSeriesCode = $seriesCode
                            vchSeries = $seriesName
                            date = $dateValue
                            approvalStatus = [int]$rst.Fields.Item("ApprovalStatus").Value
                        }
                        try { $rst.Close() } catch {}
                        return $result
                    }
                }
                $rst.MoveNext()
            }
            try { $rst.Close() } catch {}
        }

        return @{ success=$true; found=$false }
    }
    catch { return @{ success=$false; found=$false; error=$_.Exception.Message } }
    finally { if ($ownsConnection) { Disconnect-BUSY $fi } }
}

function Set-WebCreatedVoucherApprovalState {
    param(
        $fi,
        [int]$VchType,
        [string]$VchNo,
        [string]$VchSeries = "",
        [string]$VchDate = "",
        [string]$UserName,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    # Verified BUSY mapping:
    # 0 Approval Not Required | 1 Approved | 2 To be Approved
    # CheckList: 1 Created | 2 Modified | 3 Approved
    if ([string]::IsNullOrWhiteSpace($VchNo)) {
        return @{ success=$false; error="Voucher number is required to set approval state." }
    }

    try {
        $config = Get-VoucherApprovalConfig -VchType $VchType -InstanceId $InstanceId -CompanyCode $CompanyCode -ExistingFi $fi
        if (-not $config.success) { return @{ success=$false; error="Could not read voucher approval configuration. $($config.error)" } }

        $approvalRequired = [bool]$config.data.approval_required
        $targetStatus = if ($approvalRequired) { 2 } else { 0 }

        $state = Get-VoucherApprovalStateByIdentity -VchType $VchType -VchNo $VchNo -VchSeries $VchSeries -VchDate $VchDate -InstanceId $InstanceId -CompanyCode $CompanyCode -ExistingFi $fi
        if (-not $state.success) { return $state }
        if (-not $state.found) { return @{ success=$false; error="Created voucher was saved, but its VchCode could not be resolved for approval processing." } }

        $vchCode = [int]$state.vchCode
        $fi.ExecuteQuery("UPDATE Tran1 SET ApprovalStatus=$targetStatus WHERE VchCode=$vchCode")

        # SaveVchFromXML can auto-approve at creation. Normalize a newly-created
        # voucher to the BusyCloud policy before any human can act on it.
        $fi.ExecuteQuery("DELETE FROM CheckList WHERE Code=$vchCode AND Action=3")

        if (-not [string]::IsNullOrWhiteSpace($UserName)) {
            $safeUser = $UserName.Trim().Replace("'", "''")
            $fi.ExecuteQuery("UPDATE CheckList SET UserName='$safeUser' WHERE Code=$vchCode AND Action=1")
        }

        $verify = $fi.GetRecordset("SELECT ApprovalStatus FROM Tran1 WHERE VchCode=$vchCode")
        $actual = -1
        if ($verify -and -not $verify.EOF) { $actual = [int]$verify.Fields.Item("ApprovalStatus").Value }
        if ($verify) { try { $verify.Close() } catch {} }
        if ($actual -ne $targetStatus) {
            return @{ success=$false; error="Voucher was created, but BUSY approval status could not be normalized."; vchCode=$vchCode; approvalStatus=$actual; expectedApprovalStatus=$targetStatus }
        }

        $text = if ($targetStatus -eq 2) { "To be Approved" } else { "Approval Not Required" }
        return @{ success=$true; vchCode=$vchCode; approvalRequired=$approvalRequired; approvalStatus=$targetStatus; approvalStatusText=$text }
    }
    catch { return @{ success=$false; error=$_.Exception.Message } }
}

function Set-WebCreatedVoucherPendingApproval {
    param($fi,[int]$VchType,[string]$VchNo,[string]$VchSeries="",[string]$VchDate="",[string]$UserName,[string]$InstanceId="",[string]$CompanyCode="")
    return Set-WebCreatedVoucherApprovalState -fi $fi -VchType $VchType -VchNo $VchNo -VchSeries $VchSeries -VchDate $VchDate -UserName $UserName -InstanceId $InstanceId -CompanyCode $CompanyCode
}

# -----------------------------------------------------------------------------
# BusyCloud approval read/write helpers.
# IMPORTANT: Approval pages/actions must never open BUSY COM. SQL companies can
# take 1-4 minutes to initialize OpenCSDB on some installations. The approval
# feature only needs Tran1 / CheckList / Master1, so use the direct fiscal DB
# connection resolved by vch_setting.ps1 for both SQL Server and Access/BDS.
# -----------------------------------------------------------------------------
function Read-BusyCloudApprovalReaderValue {
    param($Reader, [string]$Field, $Default = $null)
    try {
        $idx = $Reader.GetOrdinal($Field)
        if ($idx -ge 0 -and -not $Reader.IsDBNull($idx)) { return $Reader.GetValue($idx) }
    } catch {}
    return $Default
}

function Convert-BusyCloudApprovalNumberLiteral {
    param($Value)
    try { return ([double]$Value).ToString('0.################', [System.Globalization.CultureInfo]::InvariantCulture) }
    catch { return '0' }
}

function Get-BusyCloudApprovalDateLiteral {
    param([datetime]$Value, [int]$DbType)
    if ($DbType -eq 1) { return "'" + $Value.ToString('yyyy-MM-dd HH:mm:ss.fff') + "'" }
    return '#' + $Value.ToString('MM/dd/yyyy HH:mm:ss') + '#'
}

function Invoke-BusyCloudApprovalNonQuery {
    param($Connection, $Transaction, [string]$Sql)
    $cmd = $Connection.CreateCommand()
    if ($null -ne $Transaction) { $cmd.Transaction = $Transaction }
    try { $cmd.CommandTimeout = 15 } catch {}
    $cmd.CommandText = $Sql
    return $cmd.ExecuteNonQuery()
}

function Get-VoucherApprovalQueue-Direct {
    param(
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$Status = 'pending',
        [int]$VchType = 0,
        [int]$Limit = 200,
        [string]$InstanceId = '',
        [string]$CompanyCode = ''
    )

    if ([string]::IsNullOrWhiteSpace($UserName)) {
        return @{ success=$false; httpStatus=401; error='Authenticated BUSY user is required.' }
    }

    $allowed = Get-VoucherApprovalTypesForUser -UserName $UserName -IsAdmin:$IsAdmin -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $allowed.success) { return $allowed }
    $allowedTypes = @($allowed.data | ForEach-Object { [int]$_ } | Where-Object { $_ -gt 0 } | Sort-Object -Unique)

    if ($VchType -gt 0) {
        if ($allowedTypes -notcontains $VchType) {
            return @{ success=$false; httpStatus=403; error="You are not allowed to approve voucher type $VchType." }
        }
        $effectiveTypes = @($VchType)
    } else {
        $effectiveTypes = @($allowedTypes)
    }

    $normalized = ([string]$Status).Trim().ToLowerInvariant()
    if ($normalized -notin @('pending','approved','history')) {
        return @{ success=$false; httpStatus=400; error='status must be pending, approved, or history.' }
    }

    if ($effectiveTypes.Count -eq 0) {
        return @{ success=$true; data=@{ allowed_vch_types=@(); status=$normalized; items=@() } }
    }

    if ($normalized -eq 'history') {
        $historyUser = if ($IsAdmin) { '' } else { $UserName }
        $history = Get-BusyCloudVoucherApprovalHistory `
            -VchTypes $effectiveTypes `
            -VchType $VchType `
            -ActionBy $historyUser `
            -Limit $Limit `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
        if (-not $history.success) { return $history }
        return @{ success=$true; data=@{ allowed_vch_types=@($allowedTypes); status='history'; items=@($history.data) } }
    }

    $approvalStatus = if ($normalized -eq 'approved') { 1 } else { 2 }
    if ($Limit -lt 1) { $Limit = 1 }
    if ($Limit -gt 500) { $Limit = 500 }

    $ctx = $null
    try {
        $ctx = Get-BusyCloudApprovalDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $conn = $ctx.connection
        $typesSql = ($effectiveTypes -join ',')

        # 1) Read voucher headers in one query.
        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 15 } catch {}
        $cmd.CommandText = @"
SELECT TOP $Limit
    VchCode,VchType,VchNo,[Date],VchSeriesCode,MasterCode1,
    VchAmtBaseCur,VchSalePurcAmt,ApprovalStatus
FROM Tran1
WHERE VchType IN ($typesSql)
  AND ApprovalStatus=$approvalStatus
ORDER BY VchCode DESC
"@
        $rdr = $cmd.ExecuteReader()
        $rows = @()
        $voucherCodes = @()
        $masterCodes = @()
        while ($rdr.Read()) {
            $code = [int](Read-BusyCloudApprovalReaderValue $rdr 'VchCode' 0)
            $type = [int](Read-BusyCloudApprovalReaderValue $rdr 'VchType' 0)
            $seriesCode = [int](Read-BusyCloudApprovalReaderValue $rdr 'VchSeriesCode' 0)
            $partyCode = [int](Read-BusyCloudApprovalReaderValue $rdr 'MasterCode1' 0)
            $rawDate = Read-BusyCloudApprovalReaderValue $rdr 'Date' $null
            $dateText = ''
            if ($null -ne $rawDate) {
                try { $dateText = ([datetime]$rawDate).ToString('yyyy-MM-dd') }
                catch { $dateText = [string]$rawDate }
            }
            $amount = [double](Read-BusyCloudApprovalReaderValue $rdr 'VchAmtBaseCur' 0)
            if ([math]::Abs($amount) -lt 0.000001) {
                $amount = [double](Read-BusyCloudApprovalReaderValue $rdr 'VchSalePurcAmt' 0)
            }

            $rows += @{
                vchCode=$code
                vchType=$type
                vchNo=([string](Read-BusyCloudApprovalReaderValue $rdr 'VchNo' '')).Trim()
                date=$dateText
                seriesCode=$seriesCode
                partyCode=$partyCode
                amount=[math]::Abs($amount)
            }
            if ($code -gt 0 -and $voucherCodes -notcontains $code) { $voucherCodes += $code }
            foreach ($mc in @($seriesCode,$partyCode)) {
                if ($mc -gt 0 -and $masterCodes -notcontains $mc) { $masterCodes += $mc }
            }
        }
        $rdr.Close()

        if ($rows.Count -eq 0) {
            return @{ success=$true; data=@{ allowed_vch_types=@($allowedTypes); status=$normalized; items=@() } }
        }

        # 2) Read all relevant creator/approval audit markers in one query.
        $createdMap = @{}
        $approvedMap = @{}
        $codesSql = ($voucherCodes -join ',')
        if ($codesSql) {
            $cmd = $conn.CreateCommand()
            try { $cmd.CommandTimeout = 15 } catch {}
            $cmd.CommandText = "SELECT Code,Action,ActionTime,UserName,Notes FROM CheckList WHERE Code IN ($codesSql) AND Action IN (1,3) ORDER BY Code,ActionTime,Action"
            $cr = $cmd.ExecuteReader()
            while ($cr.Read()) {
                $code = [int](Read-BusyCloudApprovalReaderValue $cr 'Code' 0)
                $action = [int](Read-BusyCloudApprovalReaderValue $cr 'Action' 0)
                $user = ([string](Read-BusyCloudApprovalReaderValue $cr 'UserName' '')).Trim()
                $notes = [string](Read-BusyCloudApprovalReaderValue $cr 'Notes' '')
                $rawTime = Read-BusyCloudApprovalReaderValue $cr 'ActionTime' $null
                $timeText = ''
                if ($null -ne $rawTime) { try { $timeText = ([datetime]$rawTime).ToString('yyyy-MM-dd HH:mm:ss') } catch {} }

                if ($action -eq 1 -and -not $createdMap.ContainsKey($code)) {
                    $createdMap[$code] = @{ user=$user; time=$timeText }
                } elseif ($action -eq 3) {
                    # Rows are ordered by time ascending; last Action=3 wins.
                    $approvedMap[$code] = @{ user=$user; time=$timeText; notes=$notes }
                }
            }
            $cr.Close()
        }

        # 3) Resolve only the Master1 names needed by these rows.
        $masterMap = @{}
        if ($masterCodes.Count -gt 0) {
            $mastersSql = ($masterCodes -join ',')
            $cmd = $conn.CreateCommand()
            try { $cmd.CommandTimeout = 15 } catch {}
            $cmd.CommandText = "SELECT Code,Name FROM Master1 WHERE Code IN ($mastersSql)"
            $mr = $cmd.ExecuteReader()
            while ($mr.Read()) {
                $mc = [int](Read-BusyCloudApprovalReaderValue $mr 'Code' 0)
                if ($mc -gt 0) { $masterMap[$mc] = [string](Read-BusyCloudApprovalReaderValue $mr 'Name' '') }
            }
            $mr.Close()
        }

        $items = @()
        foreach ($row in $rows) {
            $code = [int]$row.vchCode
            $type = [int]$row.vchType
            $seriesCode = [int]$row.seriesCode
            $partyCode = [int]$row.partyCode

            $seriesName = if ($masterMap.ContainsKey($seriesCode)) { ([string]$masterMap[$seriesCode]).Trim() } else { 'Main' }
            $prefix = '{0:D2}' -f $type
            if ($seriesName.StartsWith($prefix,[System.StringComparison]::OrdinalIgnoreCase)) {
                $seriesName = $seriesName.Substring($prefix.Length)
            }

            $created = if ($createdMap.ContainsKey($code)) { $createdMap[$code] } else { @{user='';time=''} }
            $approved = if ($approvedMap.ContainsKey($code)) { $approvedMap[$code] } else { @{user='';time='';notes=''} }

            $items += @{
                vchCode=$code
                vchType=$type
                vchNo=[string]$row.vchNo
                date=[string]$row.date
                vchSeriesCode=$seriesCode
                vchSeries=$seriesName
                party=if($masterMap.ContainsKey($partyCode)){[string]$masterMap[$partyCode]}else{''}
                amount=[double]$row.amount
                approvalStatus=$approvalStatus
                approvalStatusText=if($approvalStatus -eq 1){'Approved'}else{'To be Approved'}
                createdBy=[string]$created.user
                createdAt=[string]$created.time
                approvedBy=[string]$approved.user
                approvedAt=[string]$approved.time
                approvalRemarks=[string]$approved.notes
            }
        }

        return @{ success=$true; data=@{ allowed_vch_types=@($allowedTypes); status=$normalized; items=@($items) } }
    }
    catch { return @{ success=$false; error=$_.Exception.Message } }
    finally {
        if ($ctx -and $ctx.connection) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}

function Approve-Voucher-Direct {
    param(
        [int]$VchCode,
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$Remarks = '',
        [string]$InstanceId = '',
        [string]$CompanyCode = ''
    )

    if ($VchCode -le 0) { return @{ success=$false; httpStatus=400; error='vchCode is required.' } }
    if ([string]::IsNullOrWhiteSpace($UserName)) { return @{ success=$false; httpStatus=401; error='Authenticated BUSY user is required.' } }

    $ctx = $null
    $txn = $null
    try {
        $ctx = Get-BusyCloudApprovalDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $conn = $ctx.connection
        $dbType = [int]$ctx.dbType

        $cmd = $conn.CreateCommand(); try{$cmd.CommandTimeout=15}catch{}
        $cmd.CommandText = "SELECT TOP 1 VchType,VchNo,VchSeriesCode,[Date],ApprovalStatus FROM Tran1 WHERE VchCode=$VchCode"
        $r = $cmd.ExecuteReader()
        if (-not $r.Read()) { $r.Close(); return @{success=$false;httpStatus=404;error='Voucher not found.'} }
        $vchType=[int](Read-BusyCloudApprovalReaderValue $r 'VchType' 0)
        $vchNo=([string](Read-BusyCloudApprovalReaderValue $r 'VchNo' '')).Trim()
        $seriesCode=[int](Read-BusyCloudApprovalReaderValue $r 'VchSeriesCode' 0)
        $vchDate=Read-BusyCloudApprovalReaderValue $r 'Date' (Get-Date)
        $current=[int](Read-BusyCloudApprovalReaderValue $r 'ApprovalStatus' -1)
        $r.Close()

        $access=Test-VoucherApprover -UserName $UserName -VchType $vchType -IsAdmin:$IsAdmin -InstanceId $InstanceId -CompanyCode $CompanyCode
        if(-not $access.success){return $access}
        if(-not $access.allowed){return @{success=$false;httpStatus=403;error='You are not allowed to approve this voucher type.'}}
        if($current -eq 0){return @{success=$false;httpStatus=409;error='Approval is not required for this voucher.'}}
        if($current -eq 1){return @{success=$true;message='Voucher is already approved.';data=@{vchCode=$VchCode;vchType=$vchType;vchNo=$vchNo;approvalStatus=1}}}
        if($current -ne 2){return @{success=$false;httpStatus=409;error='Voucher is not in a pending approval state.'}}

        $cmd=$conn.CreateCommand();try{$cmd.CommandTimeout=15}catch{}
        $cmd.CommandText="SELECT TOP 1 Type,D1,D2,D3,D4,D5 FROM CheckList WHERE Code=$VchCode AND Action=1 ORDER BY ActionTime"
        $cr=$cmd.ExecuteReader()
        $checkType=2;$d1='0';$d2='0';$d3='0';$d4='0';$d5='0'
        if($cr.Read()){
            $checkType=[int](Read-BusyCloudApprovalReaderValue $cr 'Type' 2)
            $d1=Convert-BusyCloudApprovalNumberLiteral (Read-BusyCloudApprovalReaderValue $cr 'D1' 0)
            $d2=Convert-BusyCloudApprovalNumberLiteral (Read-BusyCloudApprovalReaderValue $cr 'D2' 0)
            $d3=Convert-BusyCloudApprovalNumberLiteral (Read-BusyCloudApprovalReaderValue $cr 'D3' 0)
            $d4=Convert-BusyCloudApprovalNumberLiteral (Read-BusyCloudApprovalReaderValue $cr 'D4' 0)
            $d5=Convert-BusyCloudApprovalNumberLiteral (Read-BusyCloudApprovalReaderValue $cr 'D5' 0)
        }
        $cr.Close()

        $seriesName='Main'
        if($seriesCode -gt 0){
            $cmd=$conn.CreateCommand();try{$cmd.CommandTimeout=15}catch{}
            $cmd.CommandText="SELECT TOP 1 Name FROM Master1 WHERE Code=$seriesCode"
            $raw=$cmd.ExecuteScalar();if($null-ne$raw-and$raw-ne[System.DBNull]::Value){$seriesName=([string]$raw).Trim()}
        }
        $prefix='{0:D2}' -f $vchType
        if($seriesName.StartsWith($prefix,[System.StringComparison]::OrdinalIgnoreCase)){$seriesName=$seriesName.Substring($prefix.Length)}

        $safeUser=$UserName.Trim().Replace("'","''")
        $safeRemarks=([string]$Remarks).Trim().Replace("'","''")
        $computer=if($env:COMPUTERNAME){$env:COMPUTERNAME}else{'BUSYCLOUD'}
        $safeComputer=$computer.Replace("'","''")
        $now=Get-Date
        $dateExpr=Get-BusyCloudApprovalDateLiteral -Value $now -DbType $dbType

        $txn=$conn.BeginTransaction()
        $updated=Invoke-BusyCloudApprovalNonQuery -Connection $conn -Transaction $txn -Sql "UPDATE Tran1 SET ApprovalStatus=1 WHERE VchCode=$VchCode AND ApprovalStatus=2"
        if([int]$updated -le 0){throw 'Voucher approval state changed before this request completed. Refresh and try again.'}
        [void](Invoke-BusyCloudApprovalNonQuery -Connection $conn -Transaction $txn -Sql "DELETE FROM CheckList WHERE Code=$VchCode AND Action=3")
        [void](Invoke-BusyCloudApprovalNonQuery -Connection $conn -Transaction $txn -Sql "INSERT INTO CheckList (Type,Code,Action,ActionTime,UserName,D1,D2,D3,D4,D5,Notes,ComputerName) VALUES ($checkType,$VchCode,3,$dateExpr,'$safeUser',$d1,$d2,$d3,$d4,$d5,'$safeRemarks','$safeComputer')")
        $txn.Commit();$txn=$null

        $audit=Write-BusyCloudVoucherApprovalAudit -VchCode $VchCode -VchType $vchType -VchNo $vchNo -VchSeries $seriesName -VchDate $vchDate -EventType 'APPROVE' -PreviousStatus 2 -NewStatus 1 -ActionBy $UserName -Remarks ([string]$Remarks).Trim() -InstanceId $InstanceId -CompanyCode $CompanyCode
        $result=@{success=$true;message='Voucher approved successfully.';data=@{vchCode=$VchCode;vchType=$vchType;vchNo=$vchNo;approvalStatus=1;approvalStatusText='Approved';approvedBy=$UserName;approvedAt=$now.ToString('yyyy-MM-dd HH:mm:ss');remarks=([string]$Remarks).Trim()}}
        if(-not$audit.success){$result.warning="Voucher was approved, but BusyCloud history could not be written: $($audit.error)"}
        return $result
    }
    catch {
        if($txn){try{$txn.Rollback()}catch{}}
        return @{success=$false;error=$_.Exception.Message}
    }
    finally {
        if($ctx -and $ctx.connection){try{$ctx.connection.Close()}catch{};try{$ctx.connection.Dispose()}catch{}}
    }
}

function Unapprove-Voucher-Direct {
    param(
        [int]$VchCode,
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$Reason = '',
        [string]$InstanceId = '',
        [string]$CompanyCode = ''
    )

    if($VchCode -le 0){return @{success=$false;httpStatus=400;error='vchCode is required.'}}
    $reasonText=([string]$Reason).Trim()
    if([string]::IsNullOrWhiteSpace($reasonText)){return @{success=$false;httpStatus=400;error='An unapprove reason is required.'}}

    $ctx=$null;$txn=$null
    try {
        $ctx=Get-BusyCloudApprovalDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $conn=$ctx.connection

        $cmd=$conn.CreateCommand();try{$cmd.CommandTimeout=15}catch{}
        $cmd.CommandText="SELECT TOP 1 VchType,VchNo,VchSeriesCode,[Date],ApprovalStatus FROM Tran1 WHERE VchCode=$VchCode"
        $r=$cmd.ExecuteReader()
        if(-not$r.Read()){$r.Close();return @{success=$false;httpStatus=404;error='Voucher not found.'}}
        $vchType=[int](Read-BusyCloudApprovalReaderValue $r 'VchType' 0)
        $vchNo=([string](Read-BusyCloudApprovalReaderValue $r 'VchNo' '')).Trim()
        $seriesCode=[int](Read-BusyCloudApprovalReaderValue $r 'VchSeriesCode' 0)
        $vchDate=Read-BusyCloudApprovalReaderValue $r 'Date' (Get-Date)
        $current=[int](Read-BusyCloudApprovalReaderValue $r 'ApprovalStatus' -1)
        $r.Close()

        $access=Test-VoucherApprover -UserName $UserName -VchType $vchType -IsAdmin:$IsAdmin -InstanceId $InstanceId -CompanyCode $CompanyCode
        if(-not$access.success){return $access}
        if(-not$access.allowed){return @{success=$false;httpStatus=403;error='You are not allowed to unapprove this voucher type.'}}
        if($current -eq 0){return @{success=$false;httpStatus=409;error='Approval is not required for this voucher.'}}
        if($current -eq 2){return @{success=$true;message='Voucher is already pending approval.';data=@{vchCode=$VchCode;approvalStatus=2}}}
        if($current -ne 1){return @{success=$false;httpStatus=409;error='Voucher is not currently approved.'}}

        $cmd=$conn.CreateCommand();try{$cmd.CommandTimeout=15}catch{}
        $cmd.CommandText="SELECT TOP 1 Action FROM CheckList WHERE Code=$VchCode AND Action=3 ORDER BY ActionTime DESC"
        $marker=$cmd.ExecuteScalar()
        if($null-eq$marker-or$marker-eq[System.DBNull]::Value){return @{success=$false;httpStatus=409;error='The approved voucher has no CheckList Action=3 marker; refusing to create an inconsistent unapprove state.'}}

        $seriesName='Main'
        if($seriesCode -gt 0){
            $cmd=$conn.CreateCommand();try{$cmd.CommandTimeout=15}catch{}
            $cmd.CommandText="SELECT TOP 1 Name FROM Master1 WHERE Code=$seriesCode"
            $raw=$cmd.ExecuteScalar();if($null-ne$raw-and$raw-ne[System.DBNull]::Value){$seriesName=([string]$raw).Trim()}
        }
        $prefix='{0:D2}' -f $vchType
        if($seriesName.StartsWith($prefix,[System.StringComparison]::OrdinalIgnoreCase)){$seriesName=$seriesName.Substring($prefix.Length)}

        $txn=$conn.BeginTransaction()
        $updated=Invoke-BusyCloudApprovalNonQuery -Connection $conn -Transaction $txn -Sql "UPDATE Tran1 SET ApprovalStatus=2 WHERE VchCode=$VchCode AND ApprovalStatus=1"
        if([int]$updated -le 0){throw 'Voucher approval state changed before this request completed. Refresh and try again.'}
        [void](Invoke-BusyCloudApprovalNonQuery -Connection $conn -Transaction $txn -Sql "DELETE FROM CheckList WHERE Code=$VchCode AND Action=3")
        $txn.Commit();$txn=$null

        $now=Get-Date
        $audit=Write-BusyCloudVoucherApprovalAudit -VchCode $VchCode -VchType $vchType -VchNo $vchNo -VchSeries $seriesName -VchDate $vchDate -EventType 'UNAPPROVE' -PreviousStatus 1 -NewStatus 2 -ActionBy $UserName -Remarks $reasonText -InstanceId $InstanceId -CompanyCode $CompanyCode
        $result=@{success=$true;message='Voucher moved back to To be Approved.';data=@{vchCode=$VchCode;vchType=$vchType;vchNo=$vchNo;approvalStatus=2;approvalStatusText='To be Approved';unapprovedBy=$UserName;unapprovedAt=$now.ToString('yyyy-MM-dd HH:mm:ss');reason=$reasonText}}
        if(-not$audit.success){$result.warning="Voucher was unapproved, but BusyCloud history could not be written: $($audit.error)"}
        return $result
    }
    catch {
        if($txn){try{$txn.Rollback()}catch{}}
        return @{success=$false;error=$_.Exception.Message}
    }
    finally {
        if($ctx -and $ctx.connection){try{$ctx.connection.Close()}catch{};try{$ctx.connection.Dispose()}catch{}}
    }
}


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

    $usesSalesman = @(26, 12, 9, 3) -contains $VchType
    $salesmanName = ""
    $salesmanCode = 0

    if ($usesSalesman) {
        try { $salesmanName = ([string]$Data.salesmanName).Trim() } catch {}
        try {
            if ($null -ne $Data.salesmanCode) {
                $salesmanCode = [int]$Data.salesmanCode
            }
        }
        catch { $salesmanCode = 0 }
    }

    # ------------------------------------------------------------
    # Challan types require BUSY TranType = 3.
    #
    # SkipBBA itself is decided by Create-Voucher / Modify-Voucher
    # and passed into this function explicitly.
    # ------------------------------------------------------------
    $isChallanType = ($vchType -eq 11 -or $vchType -eq 4)

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

    if ($usesSalesman -and -not [string]::IsNullOrWhiteSpace($salesmanName)) {
        $xml += "<BrokerInvolved>True</BrokerInvolved>"
        $xml += "<BrokerName>$([System.Security.SecurityElement]::Escape($salesmanName))</BrokerName>"
    }

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
        -defaultMC $matCentre `
        -salesmanName $salesmanName
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

    if ($usesSalesman -and $salesmanCode -gt 0) {
        $xml += "<tmpBrokerCode>$salesmanCode</tmpBrokerCode>"
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

    # ------------------------------------------------------------
    # VOUCHERS THAT MUST NOT GENERATE BBA
    #
    # 11 = Delivery Order / Material Issue
    # 4  = Goods Received Note / Material Receipt
    # 26 = Sales Quotation
    # 27 = Purchase Quotation
    #
    # Quotations are non-accounting documents, so they must not
    # create Pending Bill / Bill-by-Bill Adjustment references.
    # ------------------------------------------------------------
    $isChallanType = ($vchType -eq 11 -or $vchType -eq 4)
    $isQuotationType = ($vchType -eq 26 -or $vchType -eq 27)

    $skipBBA = [bool]($isChallanType -or $isQuotationType)

    $maxAttempts = 2
    $attempt = 1
    $lastExceptionMsg = ""
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
                Build-VoucherXml -Data $Data -Cfg $cfg -VchType $vchType -VchNo $vchNo -SkipBBA $skipBBA -fi $fi
            }

            $errMsg = ""
            $saved = $fi.SaveVchFromXML($vchType, $xml, [ref]$errMsg)
            if ($saved -ne $true) {
                return @{ success=$false; error=if ($errMsg) { $errMsg } else { 'Unknown BUSY error' } }
            }

            # Apply the administrator's BusyCloud approval policy immediately
            # after BUSY creates the voucher:
            #   policy OFF -> ApprovalStatus 0 (Approval Not Required)
            #   policy ON  -> ApprovalStatus 2 (To be Approved)
            $approvalStateResult = Set-WebCreatedVoucherApprovalState `
                -fi $fi `
                -VchType $vchType `
                -VchNo $vchNo `
                -VchSeries ([string]$Data.vchSeries) `
                -VchDate ([string]$Data.date) `
                -UserName ([string]$Data.bridgeUserName) `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode

            if (-not $approvalStateResult.success) {
                # The voucher already exists at this point. Explicitly return a
                # partial-success response so the client never retries creation.
                return @{
                    success = $false
                    created = $true
                    error = "Voucher was created in BUSY, but BusyCloud could not apply its configured approval state. Do not create it again. Details: $($approvalStateResult.error)"
                    data = @{
                        vchType = $vchType
                        vchSeries = $seriesName
                        vchNo = $vchNo
                        date = $Data.date
                        vchCode = $approvalStateResult.vchCode
                    }
                }
            }

            Clear-StockCaches -InstanceId $InstanceId -CompanyCode $CompanyCode
            return @{
                success=$true
                message="$($cfg.xmlRoot) created successfully"
                data=@{
                    vchType=$vchType
                    vchSeries=$seriesName
                    vchNo=$vchNo
                    date=$Data.date
                    party=$Data.party
                    numberingSource=$adminConfig.source
                    vchCode=$approvalStateResult.vchCode
                    approvalRequired=$approvalStateResult.approvalRequired
                    approvalStatus=$approvalStateResult.approvalStatus
                    approvalStatusText=$approvalStateResult.approvalStatusText
                }
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

# ═══════════════════════════════════════════════════════════════
#  MATERIAL-CENTRE-SPECIFIC CURRENT AVERAGE ITEM COST
# ═══════════════════════════════════════════════════════════════
# Replays BUSY stock layers from opening stock + inventory vouchers.
# FIFO is used to maintain the remaining layers; the UI cost is the
# weighted average of those remaining layers, not the next FIFO rate.
# ═══════════════════════════════════════════════════════════════
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

        # ------------------------------------------------------------
    # Keep BBA behaviour consistent between Create and Modify.
    # ------------------------------------------------------------
    $isChallanType = ($vchType -eq 11 -or $vchType -eq 4)
    $isQuotationType = ($vchType -eq 26 -or $vchType -eq 27)

    $skipBBA = [bool]($isChallanType -or $isQuotationType)
    
    $maxAttempts = 2
    $attempt = 1
    $lastExceptionMsg = ""

    while ($attempt -le $maxAttempts) {
        $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
        if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
        
        try {
            # routes.ps1 resolves the existing voucher before modify and places
            # its current approval state here. Preserve 0/1/2 exactly across
            # SaveVchFromXML; otherwise BUSY may auto-change approval state.
            $approvalStatusToPreserve = -1
            try {
                if ($null -ne $Data.preserveApprovalStatus) {
                    $candidateStatus = [int]$Data.preserveApprovalStatus
                    if ($candidateStatus -in @(0,1,2)) { $approvalStatusToPreserve = $candidateStatus }
                }
            } catch {}

            $approvalVchCode = 0
            $approvalMarkerSnapshot = $null

            if ($approvalStatusToPreserve -in @(0,1,2)) {
                $beforeState = Get-VoucherApprovalStateByIdentity `
                    -VchType $vchType `
                    -VchNo ([string]$Data.vchNo) `
                    -VchDate ([string]$Data.date) `
                    -VchSeries ([string]$Data.vchSeries) `
                    -InstanceId $InstanceId `
                    -CompanyCode $CompanyCode `
                    -ExistingFi $fi

                if ($beforeState.success -and $beforeState.found) {
                    $approvalVchCode = [int]$beforeState.vchCode
                }

                # Approved vouchers must retain the original Action=3 marker so
                # BUSY continues to show the original Approved By / Approved At.
                if ($approvalStatusToPreserve -eq 1 -and $approvalVchCode -gt 0) {
                    $approvalRs = $fi.GetRecordset("SELECT TOP 1 Type,ActionTime,UserName,D1,D2,D3,D4,D5,Notes,ComputerName FROM CheckList WHERE Code=$approvalVchCode AND Action=3 ORDER BY ActionTime DESC")
                    if ($approvalRs -and -not $approvalRs.EOF) {
                        function Read-ModifyApprovalField { param($R,[string]$F,$D); try { $v=$R.Fields.Item($F).Value; if ($null -ne $v -and $v -ne [System.DBNull]::Value) { return $v } } catch {}; return $D }
                        $approvalMarkerSnapshot = @{
                            Type = [int](Read-ModifyApprovalField $approvalRs "Type" 2)
                            ActionTime = Read-ModifyApprovalField $approvalRs "ActionTime" (Get-Date)
                            UserName = [string](Read-ModifyApprovalField $approvalRs "UserName" "")
                            D1 = Read-ModifyApprovalField $approvalRs "D1" 0
                            D2 = Read-ModifyApprovalField $approvalRs "D2" 0
                            D3 = Read-ModifyApprovalField $approvalRs "D3" 0
                            D4 = Read-ModifyApprovalField $approvalRs "D4" 0
                            D5 = Read-ModifyApprovalField $approvalRs "D5" 0
                            Notes = [string](Read-ModifyApprovalField $approvalRs "Notes" "")
                            ComputerName = [string](Read-ModifyApprovalField $approvalRs "ComputerName" "BUSYCLOUD")
                        }
                    }
                    if ($approvalRs) { try { $approvalRs.Close() } catch {} }
                }
            }

           $xml = if ($cfg.isAccounting) {
                Build-AccountingVoucherXml `
                    -Data $Data `
                    -Cfg $cfg `
                    -VchType $vchType `
                    -VchNo ([string]$Data.vchNo) `
                    -SkipBBA $false `
                    -fi $fi
            } else {
                Build-VoucherXml `
                    -Data $Data `
                    -Cfg $cfg `
                    -VchType $vchType `
                    -VchNo ([string]$Data.vchNo) `
                    -SkipBBA $skipBBA `
                    -fi $fi
            }

            $errMsg = ""
            $saved = $fi.SaveVchFromXML($vchType, $xml, [ref]$errMsg, $true)
            
            if ($saved -eq $true) {
                # Restore the exact approval state that existed before modify.
                # 0 = Approval Not Required, 1 = Approved, 2 = To be Approved.
                if ($approvalStatusToPreserve -in @(0,1,2)) {
                    if ($approvalVchCode -le 0) {
                        $afterState = Get-VoucherApprovalStateByIdentity `
                            -VchType $vchType `
                            -VchNo ([string]$Data.vchNo) `
                            -VchDate ([string]$Data.date) `
                            -VchSeries ([string]$Data.vchSeries) `
                            -InstanceId $InstanceId `
                            -CompanyCode $CompanyCode `
                            -ExistingFi $fi
                        if ($afterState.success -and $afterState.found) { $approvalVchCode = [int]$afterState.vchCode }
                    }

                    if ($approvalVchCode -gt 0) {
                        $fi.ExecuteQuery("UPDATE Tran1 SET ApprovalStatus=$approvalStatusToPreserve WHERE VchCode=$approvalVchCode")

                        if ($approvalStatusToPreserve -ne 1) {
                            # Pending / Not Required must not carry an approval marker.
                            $fi.ExecuteQuery("DELETE FROM CheckList WHERE Code=$approvalVchCode AND Action=3")
                        }
                        else {
                            # Keep exactly the original approval marker. If BUSY
                            # generates/removes Action=3 during SaveVchFromXML,
                            # normalize it back to the pre-modify Approved By / At.
                            if ($null -ne $approvalMarkerSnapshot) {
                                function To-ModifyNumberLiteral { param($Value); try { return ([double]$Value).ToString("0.################", [System.Globalization.CultureInfo]::InvariantCulture) } catch { return "0" } }
                                $found = Get-InstanceForCompany -CompanyCode $CompanyCode -InstanceId $InstanceId
                                $dbType = if ($found -and $null -ne $found.instance.dbType) { [int]$found.instance.dbType } else { 0 }
                                $dt = [datetime]$approvalMarkerSnapshot.ActionTime
                                $dateExpr = if ($dbType -eq 1) { "'" + $dt.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'" } else { "#" + $dt.ToString("MM/dd/yyyy HH:mm:ss") + "#" }
                                $safeApprovedBy = ([string]$approvalMarkerSnapshot.UserName).Replace("'", "''")
                                $safeApprovalNotes = ([string]$approvalMarkerSnapshot.Notes).Replace("'", "''")
                                $safeApprovalComputer = ([string]$approvalMarkerSnapshot.ComputerName).Replace("'", "''")
                                $d1 = To-ModifyNumberLiteral $approvalMarkerSnapshot.D1
                                $d2 = To-ModifyNumberLiteral $approvalMarkerSnapshot.D2
                                $d3 = To-ModifyNumberLiteral $approvalMarkerSnapshot.D3
                                $d4 = To-ModifyNumberLiteral $approvalMarkerSnapshot.D4
                                $d5 = To-ModifyNumberLiteral $approvalMarkerSnapshot.D5
                                $markerType = [int]$approvalMarkerSnapshot.Type
                                $fi.ExecuteQuery("DELETE FROM CheckList WHERE Code=$approvalVchCode AND Action=3")
                                $fi.ExecuteQuery("INSERT INTO CheckList (Type,Code,Action,ActionTime,UserName,D1,D2,D3,D4,D5,Notes,ComputerName) VALUES ($markerType,$approvalVchCode,3,$dateExpr,'$safeApprovedBy',$d1,$d2,$d3,$d4,$d5,'$safeApprovalNotes','$safeApprovalComputer')")
                            }
                        }
                    }
                }

                if ($Data.bridgeUserName) {
                    Update-CheckListCreator -fi $fi -VchType $vchType -VchNo ([string]$Data.vchNo) -VchDate $Data.date -VchSeries ([string]$Data.vchSeries) -UserName $Data.bridgeUserName -InstanceId $InstanceId -CompanyCode $CompanyCode
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

function Get-VoucherApprovalQueue-AccessCom {
    param(
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$Status = "pending",
        [int]$VchType = 0,
        [int]$Limit = 200,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ([string]::IsNullOrWhiteSpace($UserName)) { return @{ success=$false; httpStatus=401; error="Authenticated BUSY user is required." } }

    $allowed = Get-VoucherApprovalTypesForUser -UserName $UserName -IsAdmin:$IsAdmin -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $allowed.success) { return $allowed }
    $allowedTypes = @($allowed.data | ForEach-Object { [int]$_ } | Where-Object { $_ -gt 0 } | Sort-Object -Unique)

    if ($VchType -gt 0) {
        if ($allowedTypes -notcontains $VchType) { return @{ success=$false; httpStatus=403; error="You are not allowed to approve voucher type $VchType." } }
        $effectiveTypes = @($VchType)
    } else { $effectiveTypes = @($allowedTypes) }

    $normalized = ([string]$Status).Trim().ToLowerInvariant()
    if ($normalized -notin @("pending","approved","history")) { return @{ success=$false; httpStatus=400; error="status must be pending, approved, or history." } }

    if ($effectiveTypes.Count -eq 0) {
        return @{ success=$true; data=@{ allowed_vch_types=@(); status=$normalized; items=@() } }
    }

    if ($normalized -eq "history") {
        $historyUser = if ($IsAdmin) { "" } else { $UserName }
        $history = Get-BusyCloudVoucherApprovalHistory `
            -VchTypes $effectiveTypes `
            -VchType $VchType `
            -ActionBy $historyUser `
            -Limit $Limit `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
        if (-not $history.success) { return $history }
        return @{ success=$true; data=@{ allowed_vch_types=@($allowedTypes); status="history"; items=@($history.data) } }
    }

    $approvalStatus = if ($normalized -eq "approved") { 1 } else { 2 }
    if ($Limit -lt 1) { $Limit=1 }; if ($Limit -gt 500) { $Limit=500 }

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success=$false; error="BUSY database connection failed" } }

    try {
        function Read-ApprovalField { param($R,[string]$F,$D); try { $v=$R.Fields.Item($F).Value; if ($null -ne $v -and $v -ne [System.DBNull]::Value) { return $v } } catch {}; return $D }

        $masterMap = @{}
        $m = $fi.GetRecordset("SELECT Code,Name FROM Master1")
        if ($m) {
            while (-not $m.EOF) { try { $c=[int]$m.Fields.Item("Code").Value; if ($c -gt 0) { $masterMap[$c]=[string]$m.Fields.Item("Name").Value } } catch {}; $m.MoveNext() }
            try { $m.Close() } catch {}
        }

        $typesSql = ($effectiveTypes -join ",")
        $rst = $fi.GetRecordset(@"
SELECT TOP $Limit VchCode,VchType,VchNo,[Date],VchSeriesCode,MasterCode1,VchAmtBaseCur,VchSalePurcAmt,ApprovalStatus
FROM Tran1
WHERE VchType IN ($typesSql) AND ApprovalStatus=$approvalStatus
ORDER BY VchCode DESC
"@)

        $items = @()
        if ($rst) {
            while (-not $rst.EOF) {
                $code=[int](Read-ApprovalField $rst "VchCode" 0); $type=[int](Read-ApprovalField $rst "VchType" 0)
                $createdBy=""; $createdAt=""; $approvedBy=""; $approvedAt=""; $approvalRemarks=""

                $createdRs = $fi.GetRecordset("SELECT TOP 1 UserName,ActionTime FROM CheckList WHERE Code=$code AND Action=1 ORDER BY ActionTime")
                if ($createdRs -and -not $createdRs.EOF) {
                    $createdBy=([string](Read-ApprovalField $createdRs "UserName" "")).Trim(); $raw=Read-ApprovalField $createdRs "ActionTime" $null
                    if ($null -ne $raw) { try { $createdAt=([datetime]$raw).ToString("yyyy-MM-dd HH:mm:ss") } catch {} }
                }
                if ($createdRs) { try { $createdRs.Close() } catch {} }

                if ($approvalStatus -eq 1) {
                    $a = $fi.GetRecordset("SELECT TOP 1 UserName,ActionTime,Notes FROM CheckList WHERE Code=$code AND Action=3 ORDER BY ActionTime DESC")
                    if ($a -and -not $a.EOF) {
                        $approvedBy=([string](Read-ApprovalField $a "UserName" "")).Trim(); $approvalRemarks=[string](Read-ApprovalField $a "Notes" ""); $raw=Read-ApprovalField $a "ActionTime" $null
                        if ($null -ne $raw) { try { $approvedAt=([datetime]$raw).ToString("yyyy-MM-dd HH:mm:ss") } catch {} }
                    }
                    if ($a) { try { $a.Close() } catch {} }
                }

                # Any currently assigned approver can Unapprove an approved
                # voucher in an allowed voucher type. Therefore the Approved
                # list is type-scoped, not "approved by me" scoped.
                $include = $true

                if ($include) {
                    $seriesCode=[int](Read-ApprovalField $rst "VchSeriesCode" 0); $masterCode1=[int](Read-ApprovalField $rst "MasterCode1" 0)
                    $seriesName=if($masterMap.ContainsKey($seriesCode)){[string]$masterMap[$seriesCode]}else{"Main"}; $prefix="{0:D2}" -f $type
                    if($seriesName.StartsWith($prefix,[System.StringComparison]::OrdinalIgnoreCase)){$seriesName=$seriesName.Substring($prefix.Length)}
                    $rawDate=Read-ApprovalField $rst "Date" $null; $dateText=""; if($null-ne $rawDate){try{$dateText=([datetime]$rawDate).ToString("yyyy-MM-dd")}catch{$dateText=[string]$rawDate}}
                    $amount=[double](Read-ApprovalField $rst "VchAmtBaseCur" 0); if([math]::Abs($amount)-lt 0.000001){$amount=[double](Read-ApprovalField $rst "VchSalePurcAmt" 0)}
                    $items += @{
                        vchCode=$code; vchType=$type; vchNo=[string](Read-ApprovalField $rst "VchNo" ""); date=$dateText; vchSeriesCode=$seriesCode; vchSeries=$seriesName;
                        party=if($masterMap.ContainsKey($masterCode1)){[string]$masterMap[$masterCode1]}else{""}; amount=[math]::Abs($amount);
                        approvalStatus=$approvalStatus; approvalStatusText=if($approvalStatus-eq1){"Approved"}else{"To be Approved"};
                        createdBy=$createdBy; createdAt=$createdAt; approvedBy=$approvedBy; approvedAt=$approvedAt; approvalRemarks=$approvalRemarks
                    }
                }
                $rst.MoveNext()
            }
            try { $rst.Close() } catch {}
        }

        return @{ success=$true; data=@{ allowed_vch_types=@($allowedTypes); status=$normalized; items=@($items) } }
    }
    catch { return @{ success=$false; error=$_.Exception.Message } }
    finally { Disconnect-BUSY $fi }
}

function Approve-Voucher-AccessCom {
    param(
        [int]$VchCode,
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$Remarks = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ($VchCode -le 0) { return @{ success=$false; httpStatus=400; error="vchCode is required." } }
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success=$false; error="BUSY database connection failed" } }

    try {
        $rst=$fi.GetRecordset("SELECT VchType,VchNo,VchSeriesCode,[Date],ApprovalStatus FROM Tran1 WHERE VchCode=$VchCode")
        if(-not $rst -or $rst.EOF){return @{success=$false;httpStatus=404;error="Voucher not found."}}
        $vchType=[int]$rst.Fields.Item("VchType").Value; $vchNo=([string]$rst.Fields.Item("VchNo").Value).Trim(); $seriesCode=[int]$rst.Fields.Item("VchSeriesCode").Value; $vchDate=$rst.Fields.Item("Date").Value; $current=[int]$rst.Fields.Item("ApprovalStatus").Value
        try{$rst.Close()}catch{}

        $access=Test-VoucherApprover -UserName $UserName -VchType $vchType -IsAdmin:$IsAdmin -InstanceId $InstanceId -CompanyCode $CompanyCode -ExistingFi $fi
        if(-not $access.success){return $access}; if(-not $access.allowed){return @{success=$false;httpStatus=403;error="You are not allowed to approve this voucher type."}}
        if($current-eq0){return @{success=$false;httpStatus=409;error="Approval is not required for this voucher."}}
        if($current-eq1){return @{success=$true;message="Voucher is already approved.";data=@{vchCode=$VchCode;vchType=$vchType;vchNo=$vchNo;approvalStatus=1}}}
        if($current-ne2){return @{success=$false;httpStatus=409;error="Voucher is not in a pending approval state."}}

        $storage=Ensure-BusyCloudVoucherApprovalAuditTable -InstanceId $InstanceId -CompanyCode $CompanyCode
        if(-not $storage.success){return @{success=$false;error="Approval audit storage is unavailable. $($storage.error)"}}

        function RAF { param($R,[string]$F,$D);try{$v=$R.Fields.Item($F).Value;if($null-ne$v-and$v-ne[System.DBNull]::Value){return$v}}catch{};return$D }
        function NumLit { param($V);try{return([double]$V).ToString("0.################",[System.Globalization.CultureInfo]::InvariantCulture)}catch{return"0"} }

        $src=$fi.GetRecordset("SELECT TOP 1 Type,D1,D2,D3,D4,D5 FROM CheckList WHERE Code=$VchCode AND Action=1 ORDER BY ActionTime")
        $checkType=2;$d1="0";$d2="0";$d3="0";$d4="0";$d5="0"
        if($src-and-not$src.EOF){$checkType=[int](RAF $src "Type" 2);$d1=NumLit(RAF $src "D1" 0);$d2=NumLit(RAF $src "D2" 0);$d3=NumLit(RAF $src "D3" 0);$d4=NumLit(RAF $src "D4" 0);$d5=NumLit(RAF $src "D5" 0)}
        if($src){try{$src.Close()}catch{}}

        $seriesName="Main";$sr=$fi.GetRecordset("SELECT Name FROM Master1 WHERE Code=$seriesCode");if($sr-and-not$sr.EOF){$seriesName=([string]$sr.Fields.Item("Name").Value).Trim()};if($sr){try{$sr.Close()}catch{}};$prefix="{0:D2}"-f$vchType;if($seriesName.StartsWith($prefix,[System.StringComparison]::OrdinalIgnoreCase)){$seriesName=$seriesName.Substring($prefix.Length)}

        $safeUser=$UserName.Trim().Replace("'","''");$safeRemarks=([string]$Remarks).Trim().Replace("'","''");$computer=if($env:COMPUTERNAME){$env:COMPUTERNAME}else{"BUSYCLOUD"};$safeComputer=$computer.Replace("'","''")
        $found=Get-InstanceForCompany -CompanyCode $CompanyCode -InstanceId $InstanceId;$dbType=if($found-and$null-ne$found.instance.dbType){[int]$found.instance.dbType}else{0};$nowExpr=if($dbType-eq1){"GETDATE()"}else{"Now()"}

        $fi.ExecuteQuery("UPDATE Tran1 SET ApprovalStatus=1 WHERE VchCode=$VchCode AND ApprovalStatus=2")
        try {
            $fi.ExecuteQuery("DELETE FROM CheckList WHERE Code=$VchCode AND Action=3")
            $fi.ExecuteQuery("INSERT INTO CheckList (Type,Code,Action,ActionTime,UserName,D1,D2,D3,D4,D5,Notes,ComputerName) VALUES ($checkType,$VchCode,3,$nowExpr,'$safeUser',$d1,$d2,$d3,$d4,$d5,'$safeRemarks','$safeComputer')")
            $audit=Write-BusyCloudVoucherApprovalAudit -VchCode $VchCode -VchType $vchType -VchNo $vchNo -VchSeries $seriesName -VchDate $vchDate -EventType "APPROVE" -PreviousStatus 2 -NewStatus 1 -ActionBy $UserName -Remarks ([string]$Remarks).Trim() -InstanceId $InstanceId -CompanyCode $CompanyCode
            if(-not$audit.success){throw "Could not write BusyCloud approval history. $($audit.error)"}
        } catch {
            try{$fi.ExecuteQuery("DELETE FROM CheckList WHERE Code=$VchCode AND Action=3")}catch{};try{$fi.ExecuteQuery("UPDATE Tran1 SET ApprovalStatus=2 WHERE VchCode=$VchCode")}catch{};throw
        }

        return @{success=$true;message="Voucher approved successfully.";data=@{vchCode=$VchCode;vchType=$vchType;vchNo=$vchNo;approvalStatus=1;approvalStatusText="Approved";approvedBy=$UserName;approvedAt=$audit.actionTime;remarks=([string]$Remarks).Trim()}}
    }
    catch{return @{success=$false;error=$_.Exception.Message}}
    finally{Disconnect-BUSY $fi}
}

function Unapprove-Voucher-AccessCom {
    param(
        [int]$VchCode,
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$Reason = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if($VchCode-le0){return @{success=$false;httpStatus=400;error="vchCode is required."}}
    $reasonText=([string]$Reason).Trim();if([string]::IsNullOrWhiteSpace($reasonText)){return @{success=$false;httpStatus=400;error="An unapprove reason is required."}}
    $fi=Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode;if(-not$fi){return @{success=$false;error="BUSY database connection failed"}}

    try {
        function URF { param($R,[string]$F,$D);try{$v=$R.Fields.Item($F).Value;if($null-ne$v-and$v-ne[System.DBNull]::Value){return$v}}catch{};return$D }
        function UNum { param($V);try{return([double]$V).ToString("0.################",[System.Globalization.CultureInfo]::InvariantCulture)}catch{return"0"} }

        $rst=$fi.GetRecordset("SELECT VchType,VchNo,VchSeriesCode,[Date],ApprovalStatus FROM Tran1 WHERE VchCode=$VchCode");if(-not$rst-or$rst.EOF){return @{success=$false;httpStatus=404;error="Voucher not found."}}
        $vchType=[int]$rst.Fields.Item("VchType").Value;$vchNo=([string]$rst.Fields.Item("VchNo").Value).Trim();$seriesCode=[int]$rst.Fields.Item("VchSeriesCode").Value;$vchDate=$rst.Fields.Item("Date").Value;$current=[int]$rst.Fields.Item("ApprovalStatus").Value;try{$rst.Close()}catch{}

        $access=Test-VoucherApprover -UserName $UserName -VchType $vchType -IsAdmin:$IsAdmin -InstanceId $InstanceId -CompanyCode $CompanyCode -ExistingFi $fi;if(-not$access.success){return$access};if(-not$access.allowed){return @{success=$false;httpStatus=403;error="You are not allowed to unapprove this voucher type."}}
        if($current-eq0){return @{success=$false;httpStatus=409;error="Approval is not required for this voucher."}};if($current-eq2){return @{success=$true;message="Voucher is already pending approval.";data=@{vchCode=$VchCode;approvalStatus=2}}};if($current-ne1){return @{success=$false;httpStatus=409;error="Voucher is not currently approved."}}

        $approval=$fi.GetRecordset("SELECT TOP 1 Type,ActionTime,UserName,D1,D2,D3,D4,D5,Notes,ComputerName FROM CheckList WHERE Code=$VchCode AND Action=3 ORDER BY ActionTime DESC")
        if(-not$approval-or$approval.EOF){return @{success=$false;httpStatus=409;error="The approved voucher has no CheckList Action=3 marker; refusing to create an inconsistent unapprove state."}}
        $aType=[int](URF $approval "Type" 2);$aTime=URF $approval "ActionTime" (Get-Date);$aUser=[string](URF $approval "UserName" "");$ad1=UNum(URF $approval "D1" 0);$ad2=UNum(URF $approval "D2" 0);$ad3=UNum(URF $approval "D3" 0);$ad4=UNum(URF $approval "D4" 0);$ad5=UNum(URF $approval "D5" 0);$aNotes=[string](URF $approval "Notes" "");$aComputer=[string](URF $approval "ComputerName" "BUSYCLOUD");try{$approval.Close()}catch{}

        $seriesName="Main";$sr=$fi.GetRecordset("SELECT Name FROM Master1 WHERE Code=$seriesCode");if($sr-and-not$sr.EOF){$seriesName=([string]$sr.Fields.Item("Name").Value).Trim()};if($sr){try{$sr.Close()}catch{}};$prefix="{0:D2}"-f$vchType;if($seriesName.StartsWith($prefix,[System.StringComparison]::OrdinalIgnoreCase)){$seriesName=$seriesName.Substring($prefix.Length)}
        $storage=Ensure-BusyCloudVoucherApprovalAuditTable -InstanceId $InstanceId -CompanyCode $CompanyCode;if(-not$storage.success){return @{success=$false;error="Approval audit storage is unavailable. $($storage.error)"}}

        $fi.ExecuteQuery("UPDATE Tran1 SET ApprovalStatus=2 WHERE VchCode=$VchCode AND ApprovalStatus=1")
        try {
            $fi.ExecuteQuery("DELETE FROM CheckList WHERE Code=$VchCode AND Action=3")
            $audit=Write-BusyCloudVoucherApprovalAudit -VchCode $VchCode -VchType $vchType -VchNo $vchNo -VchSeries $seriesName -VchDate $vchDate -EventType "UNAPPROVE" -PreviousStatus 1 -NewStatus 2 -ActionBy $UserName -Remarks $reasonText -InstanceId $InstanceId -CompanyCode $CompanyCode
            if(-not$audit.success){throw "Could not write BusyCloud unapprove history. $($audit.error)"}
        } catch {
            try{$fi.ExecuteQuery("UPDATE Tran1 SET ApprovalStatus=1 WHERE VchCode=$VchCode")}catch{}
            try {
                $found=Get-InstanceForCompany -CompanyCode $CompanyCode -InstanceId $InstanceId;$dbType=if($found-and$null-ne$found.instance.dbType){[int]$found.instance.dbType}else{0};$dt=[datetime]$aTime;$dateExpr=if($dbType-eq1){"'"+$dt.ToString("yyyy-MM-dd HH:mm:ss.fff")+"'"}else{"#"+$dt.ToString("MM/dd/yyyy HH:mm:ss")+"#"}
                $su=$aUser.Replace("'","''");$sn=$aNotes.Replace("'","''");$sc=$aComputer.Replace("'","''")
                $fi.ExecuteQuery("INSERT INTO CheckList (Type,Code,Action,ActionTime,UserName,D1,D2,D3,D4,D5,Notes,ComputerName) VALUES ($aType,$VchCode,3,$dateExpr,'$su',$ad1,$ad2,$ad3,$ad4,$ad5,'$sn','$sc')")
            } catch {}
            throw
        }

        return @{success=$true;message="Voucher moved back to To be Approved.";data=@{vchCode=$VchCode;vchType=$vchType;vchNo=$vchNo;approvalStatus=2;approvalStatusText="To be Approved";unapprovedBy=$UserName;unapprovedAt=$audit.actionTime;reason=$reasonText}}
    }
    catch{return @{success=$false;error=$_.Exception.Message}}
    finally{Disconnect-BUSY $fi}
}




# -----------------------------------------------------------------------------
# HOTFIX V5: SQL approval operations stay on the fast direct DB path.
# Access/BDS approval operations use BUSY COM because the direct db.bds used for
# authentication does not expose the fiscal Tran1 / CheckList / Master1 / Config
# tables consistently. Access COM initialization is local and fast in the
# observed installation, while SQL COM remains deliberately avoided.
# -----------------------------------------------------------------------------

function Test-BusyCloudApprovalAccessRuntime {
    param(
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    try {
        $found = Get-InstanceForCompany -CompanyCode $CompanyCode -InstanceId $InstanceId
        if (-not $found) { return $false }
        $dbType = if ($null -ne $found.instance.dbType) { [int]$found.instance.dbType } else { 0 }
        return ($dbType -eq 0)
    }
    catch {
        return $false
    }
}

function Get-VoucherApprovalQueue {
    param(
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$Status = 'pending',
        [int]$VchType = 0,
        [int]$Limit = 200,
        [string]$InstanceId = '',
        [string]$CompanyCode = ''
    )

    if (Test-BusyCloudApprovalAccessRuntime -InstanceId $InstanceId -CompanyCode $CompanyCode) {
        return Get-VoucherApprovalQueue-AccessCom `
            -UserName $UserName `
            -IsAdmin:$IsAdmin `
            -Status $Status `
            -VchType $VchType `
            -Limit $Limit `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }

    return Get-VoucherApprovalQueue-Direct `
        -UserName $UserName `
        -IsAdmin:$IsAdmin `
        -Status $Status `
        -VchType $VchType `
        -Limit $Limit `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode
}

function Approve-Voucher {
    param(
        [int]$VchCode,
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$Remarks = '',
        [string]$InstanceId = '',
        [string]$CompanyCode = ''
    )

    if (Test-BusyCloudApprovalAccessRuntime -InstanceId $InstanceId -CompanyCode $CompanyCode) {
        return Approve-Voucher-AccessCom `
            -VchCode $VchCode `
            -UserName $UserName `
            -IsAdmin:$IsAdmin `
            -Remarks $Remarks `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }

    return Approve-Voucher-Direct `
        -VchCode $VchCode `
        -UserName $UserName `
        -IsAdmin:$IsAdmin `
        -Remarks $Remarks `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode
}

function Unapprove-Voucher {
    param(
        [int]$VchCode,
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$Reason = '',
        [string]$InstanceId = '',
        [string]$CompanyCode = ''
    )

    if (Test-BusyCloudApprovalAccessRuntime -InstanceId $InstanceId -CompanyCode $CompanyCode) {
        return Unapprove-Voucher-AccessCom `
            -VchCode $VchCode `
            -UserName $UserName `
            -IsAdmin:$IsAdmin `
            -Reason $Reason `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }

    return Unapprove-Voucher-Direct `
        -VchCode $VchCode `
        -UserName $UserName `
        -IsAdmin:$IsAdmin `
        -Reason $Reason `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode
}


# =============================================================================
# BusyCloud Voucher Approval v6 list/search/pagination
# =============================================================================

function Convert-BusyCloudApprovalFilterDate {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $formats = @('yyyy-MM-dd','dd-MM-yyyy','MM/dd/yyyy','yyyy/MM/dd')
    foreach ($fmt in $formats) {
        $dt = [datetime]::MinValue
        if ([datetime]::TryParseExact($Value.Trim(), $fmt, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
            return $dt.Date
        }
    }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($Value, [ref]$parsed)) { return $parsed.Date }
    return $null
}

function New-BusyCloudApprovalPagination {
    param([int]$Page, [int]$PageSize, [int]$Total)
    if ($Page -lt 1) { $Page = 1 }
    if ($PageSize -le 0) {
        return @{
            page=1; page_size='all'; total=$Total; total_pages=1;
            has_previous=$false; has_next=$false
        }
    }
    $totalPages = if ($Total -le 0) { 1 } else { [int][math]::Ceiling($Total / [double]$PageSize) }
    if ($Page -gt $totalPages) { $Page = $totalPages }
    return @{
        page=$Page; page_size=$PageSize; total=$Total; total_pages=$totalPages;
        has_previous=($Page -gt 1); has_next=($Page -lt $totalPages)
    }
}

function Get-BusyCloudVoucherApprovalHistoryV6 {
    param(
        [int[]]$VchTypes=@(),
        [int]$VchType=0,
        [string]$ActionBy='',
        [string]$Search='',
        [string]$FromDate='',
        [string]$ToDate='',
        [int]$Page=1,
        [int]$PageSize=50,
        [string]$InstanceId='',
        [string]$CompanyCode=''
    )
    $ensure = Ensure-BusyCloudVoucherApprovalAuditTable -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $ensure.success) { return $ensure }

    $effectiveTypes = @($VchTypes | ForEach-Object { [int]$_ } | Where-Object { Test-IsBusyCloudApprovalVoucherType -VchType $_ } | Sort-Object -Unique)
    if ($VchType -gt 0) { $effectiveTypes = @($VchType) }
    if ($effectiveTypes.Count -eq 0) {
        return @{ success=$true; data=@(); pagination=(New-BusyCloudApprovalPagination -Page 1 -PageSize $PageSize -Total 0) }
    }

    if ($Page -lt 1) { $Page=1 }
    if ($PageSize -notin @(0,50,100,200,300,500)) { $PageSize=50 }
    $from = Convert-BusyCloudApprovalFilterDate $FromDate
    $to = Convert-BusyCloudApprovalFilterDate $ToDate
    $needle = ([string]$Search).Trim().ToLowerInvariant()

    try {
        $items = @()
        foreach ($line in @(Get-Content -Path $ensure.path -ErrorAction SilentlyContinue)) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            try { $r = $line | ConvertFrom-Json } catch { continue }
            $t=0; try { $t=[int]$r.vchType } catch {}
            if ($effectiveTypes -notcontains $t) { continue }
            if ($ActionBy -and -not ([string]$r.actionBy).Equals($ActionBy,[System.StringComparison]::OrdinalIgnoreCase)) { continue }

            $actionDt=$null
            try { $actionDt=[datetime]$r.actionTime } catch {}
            if ($from -and $actionDt -and $actionDt.Date -lt $from.Date) { continue }
            if ($to -and $actionDt -and $actionDt.Date -gt $to.Date) { continue }

            if ($needle) {
                $hay = @(
                    [string]$r.vchNo, [string]$r.vchSeries, [string]$r.eventType,
                    [string]$r.actionBy, [string]$r.remarks, [string]$r.vchCode
                ) -join ' '
                if (-not $hay.ToLowerInvariant().Contains($needle)) { continue }
            }
            $items += $r
        }

        $items = @($items | Sort-Object @{Expression={ try{[datetime]$_.actionTime}catch{[datetime]::MinValue} };Descending=$true}, @{Expression={ try{[long]$_.id}catch{0} };Descending=$true})
        $total = $items.Count
        if ($PageSize -le 0) {
            $paged = @($items)
            $page = 1
        } else {
            $totalPages = if ($total -le 0) { 1 } else { [int][math]::Ceiling($total/[double]$PageSize) }
            if ($Page -gt $totalPages) { $Page=$totalPages }
            $skip = ($Page-1)*$PageSize
            $paged = @($items | Select-Object -Skip $skip -First $PageSize)
        }
        return @{ success=$true; data=@($paged); pagination=(New-BusyCloudApprovalPagination -Page $Page -PageSize $PageSize -Total $total) }
    }
    catch { return @{ success=$false; error=$_.Exception.Message } }
}

function Get-VoucherApprovalQueue-DirectV6 {
    param(
        [string]$UserName,
        [bool]$IsAdmin=$false,
        [string]$Status='pending',
        [int]$VchType=0,
        [string]$Search='',
        [string]$FromDate='',
        [string]$ToDate='',
        [int]$Page=1,
        [int]$PageSize=50,
        [string]$InstanceId='',
        [string]$CompanyCode=''
    )
    if ([string]::IsNullOrWhiteSpace($UserName)) { return @{success=$false;httpStatus=401;error='Authenticated BUSY user is required.'} }
    $allowed = Get-VoucherApprovalTypesForUser -UserName $UserName -IsAdmin:$IsAdmin -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $allowed.success) { return $allowed }
    $allowedTypes=@($allowed.data | ForEach-Object{[int]$_} | Where-Object{$_ -gt 0} | Sort-Object -Unique)
    if ($VchType -le 0) {
        return @{success=$true;data=@{allowed_vch_types=@($allowedTypes);status=([string]$Status).Trim().ToLowerInvariant();items=@();pagination=(New-BusyCloudApprovalPagination -Page 1 -PageSize $PageSize -Total 0)}}
    }
    if ($allowedTypes -notcontains $VchType) { return @{success=$false;httpStatus=403;error="Approval processing is not enabled for voucher type $VchType, or you are not an approver for it."} }
    $normalized=([string]$Status).Trim().ToLowerInvariant()
    if($normalized -notin @('pending','approved','history')){return @{success=$false;httpStatus=400;error='status must be pending, approved, or history.'}}
    if($Page -lt 1){$Page=1}; if($PageSize -notin @(0,50,100,200,300,500)){$PageSize=50}

    if($normalized -eq 'history'){
        $historyUser=if($IsAdmin){''}else{$UserName}
        $h=Get-BusyCloudVoucherApprovalHistoryV6 -VchTypes @($VchType) -VchType $VchType -ActionBy $historyUser -Search $Search -FromDate $FromDate -ToDate $ToDate -Page $Page -PageSize $PageSize -InstanceId $InstanceId -CompanyCode $CompanyCode
        if(-not $h.success){return$h}
        return @{success=$true;data=@{allowed_vch_types=@($allowedTypes);status='history';items=@($h.data);pagination=$h.pagination}}
    }

    $approvalStatus=if($normalized -eq 'approved'){1}else{2}
    $ctx=$null
    try{
        $ctx=Get-BusyCloudApprovalDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $conn=$ctx.connection
        $where=@("T.VchType=$VchType","T.ApprovalStatus=$approvalStatus")
        $from=Convert-BusyCloudApprovalFilterDate $FromDate
        $to=Convert-BusyCloudApprovalFilterDate $ToDate
        if($from){$where += "T.[Date] >= '"+$from.ToString('yyyy-MM-dd')+"'"}
        if($to){$where += "T.[Date] < '"+$to.AddDays(1).ToString('yyyy-MM-dd')+"'"}
        $needle=([string]$Search).Trim()
        if($needle){
            $safe=$needle.Replace("'","''")
            $like="%$safe%"
            $where += "(T.VchNo LIKE '$like' OR CAST(T.VchCode AS VARCHAR(30)) LIKE '$like' OR ISNULL(P.Name,'') LIKE '$like' OR ISNULL(S.Name,'') LIKE '$like' OR EXISTS (SELECT 1 FROM CheckList C WHERE C.Code=T.VchCode AND C.Action IN (1,3) AND (ISNULL(C.UserName,'') LIKE '$like' OR ISNULL(C.Notes,'') LIKE '$like')))"
        }
        $whereSql=$where -join ' AND '

        $countCmd=$conn.CreateCommand();try{$countCmd.CommandTimeout=20}catch{}
        $countCmd.CommandText="SELECT COUNT(*) FROM Tran1 T LEFT JOIN Master1 P ON P.Code=T.MasterCode1 LEFT JOIN Master1 S ON S.Code=T.VchSeriesCode WHERE $whereSql"
        $total=[int]$countCmd.ExecuteScalar()
        $pagination=New-BusyCloudApprovalPagination -Page $Page -PageSize $PageSize -Total $total
        $Page=[int]$pagination.page

        $sql=@"
SELECT T.VchCode,T.VchType,T.VchNo,T.[Date],T.VchSeriesCode,T.MasterCode1,
       T.VchAmtBaseCur,T.VchSalePurcAmt,T.ApprovalStatus,
       ISNULL(P.Name,'') AS PartyName, ISNULL(S.Name,'') AS SeriesName
FROM Tran1 T
LEFT JOIN Master1 P ON P.Code=T.MasterCode1
LEFT JOIN Master1 S ON S.Code=T.VchSeriesCode
WHERE $whereSql
ORDER BY T.VchCode DESC
"@
        if($PageSize -gt 0){
            $offset=($Page-1)*$PageSize
            $sql += " OFFSET $offset ROWS FETCH NEXT $PageSize ROWS ONLY"
        }
        $cmd=$conn.CreateCommand();try{$cmd.CommandTimeout=20}catch{};$cmd.CommandText=$sql
        $rdr=$cmd.ExecuteReader();$rows=@();$codes=@()
        while($rdr.Read()){
            $code=[int](Read-BusyCloudApprovalReaderValue $rdr 'VchCode' 0)
            $rawDate=Read-BusyCloudApprovalReaderValue $rdr 'Date' $null;$dateText='';if($rawDate){try{$dateText=([datetime]$rawDate).ToString('yyyy-MM-dd')}catch{$dateText=[string]$rawDate}}
            $amount=[double](Read-BusyCloudApprovalReaderValue $rdr 'VchAmtBaseCur' 0);if([math]::Abs($amount) -lt .000001){$amount=[double](Read-BusyCloudApprovalReaderValue $rdr 'VchSalePurcAmt' 0)}
            $seriesName=([string](Read-BusyCloudApprovalReaderValue $rdr 'SeriesName' 'Main')).Trim();if(-not $seriesName){$seriesName='Main'};$prefix='{0:D2}' -f $VchType;if($seriesName.StartsWith($prefix,[System.StringComparison]::OrdinalIgnoreCase)){$seriesName=$seriesName.Substring($prefix.Length)}
            $rows += @{vchCode=$code;vchType=$VchType;vchNo=([string](Read-BusyCloudApprovalReaderValue $rdr 'VchNo' '')).Trim();date=$dateText;vchSeriesCode=[int](Read-BusyCloudApprovalReaderValue $rdr 'VchSeriesCode' 0);vchSeries=$seriesName;party=([string](Read-BusyCloudApprovalReaderValue $rdr 'PartyName' '')).Trim();amount=[math]::Abs($amount)}
            if($code -gt 0){$codes += $code}
        };$rdr.Close()

        $createdMap=@{};$approvedMap=@{}
        if($codes.Count -gt 0){
            $codesSql=($codes -join ',');$ac=$conn.CreateCommand();try{$ac.CommandTimeout=20}catch{};$ac.CommandText="SELECT Code,Action,ActionTime,UserName,Notes FROM CheckList WHERE Code IN ($codesSql) AND Action IN (1,3) ORDER BY Code,ActionTime,Action"
            $ar=$ac.ExecuteReader();while($ar.Read()){
                $code=[int](Read-BusyCloudApprovalReaderValue $ar 'Code' 0);$action=[int](Read-BusyCloudApprovalReaderValue $ar 'Action' 0);$user=([string](Read-BusyCloudApprovalReaderValue $ar 'UserName' '')).Trim();$notes=[string](Read-BusyCloudApprovalReaderValue $ar 'Notes' '');$raw=Read-BusyCloudApprovalReaderValue $ar 'ActionTime' $null;$time='';if($raw){try{$time=([datetime]$raw).ToString('yyyy-MM-dd HH:mm:ss')}catch{}}
                if($action -eq 1 -and -not $createdMap.ContainsKey($code)){$createdMap[$code]=@{user=$user;time=$time}}elseif($action -eq 3){$approvedMap[$code]=@{user=$user;time=$time;notes=$notes}}
            };$ar.Close()
        }
        $items=@();foreach($row in $rows){$c=[int]$row.vchCode;$cr=if($createdMap.ContainsKey($c)){$createdMap[$c]}else{@{user='';time=''}};$ap=if($approvedMap.ContainsKey($c)){$approvedMap[$c]}else{@{user='';time='';notes=''}};$items+=@{vchCode=$c;vchType=$row.vchType;vchNo=$row.vchNo;date=$row.date;vchSeriesCode=$row.vchSeriesCode;vchSeries=$row.vchSeries;party=$row.party;amount=$row.amount;approvalStatus=$approvalStatus;approvalStatusText=if($approvalStatus -eq 1){'Approved'}else{'To be Approved'};createdBy=$cr.user;createdAt=$cr.time;approvedBy=$ap.user;approvedAt=$ap.time;approvalRemarks=$ap.notes}}
        return @{success=$true;data=@{allowed_vch_types=@($allowedTypes);status=$normalized;items=@($items);pagination=$pagination}}
    }catch{return @{success=$false;error=$_.Exception.Message}}finally{if($ctx -and $ctx.connection){try{$ctx.connection.Close()}catch{};try{$ctx.connection.Dispose()}catch{}}}
}

function Get-VoucherApprovalQueue-AccessComV6 {
    param(
        [string]$UserName,[bool]$IsAdmin=$false,[string]$Status='pending',[int]$VchType=0,
        [string]$Search='',[string]$FromDate='',[string]$ToDate='',[int]$Page=1,[int]$PageSize=50,
        [string]$InstanceId='',[string]$CompanyCode=''
    )
    if([string]::IsNullOrWhiteSpace($UserName)){return @{success=$false;httpStatus=401;error='Authenticated BUSY user is required.'}}
    $allowed=Get-VoucherApprovalTypesForUser -UserName $UserName -IsAdmin:$IsAdmin -InstanceId $InstanceId -CompanyCode $CompanyCode
    if(-not $allowed.success){return$allowed};$allowedTypes=@($allowed.data|ForEach-Object{[int]$_}|Where-Object{$_ -gt 0}|Sort-Object -Unique)
    if($VchType -le 0){return @{success=$true;data=@{allowed_vch_types=@($allowedTypes);status=([string]$Status).Trim().ToLowerInvariant();items=@();pagination=(New-BusyCloudApprovalPagination -Page 1 -PageSize $PageSize -Total 0)}}}
    if($allowedTypes -notcontains $VchType){return @{success=$false;httpStatus=403;error="Approval processing is not enabled for voucher type $VchType, or you are not an approver for it."}}
    $normalized=([string]$Status).Trim().ToLowerInvariant();if($normalized -notin @('pending','approved','history')){return @{success=$false;httpStatus=400;error='status must be pending, approved, or history.'}}
    if($Page -lt 1){$Page=1};if($PageSize -notin @(0,50,100,200,300,500)){$PageSize=50}
    if($normalized -eq 'history'){$historyUser=if($IsAdmin){''}else{$UserName};$h=Get-BusyCloudVoucherApprovalHistoryV6 -VchTypes @($VchType) -VchType $VchType -ActionBy $historyUser -Search $Search -FromDate $FromDate -ToDate $ToDate -Page $Page -PageSize $PageSize -InstanceId $InstanceId -CompanyCode $CompanyCode;if(-not $h.success){return$h};return @{success=$true;data=@{allowed_vch_types=@($allowedTypes);status='history';items=@($h.data);pagination=$h.pagination}}}

    $approvalStatus=if($normalized -eq 'approved'){1}else{2};$fi=Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode;if(-not $fi){return @{success=$false;error='BUSY database connection failed'}}
    try{
        function Read-ApprovalFieldV6{param($R,[string]$F,$D);try{$v=$R.Fields.Item($F).Value;if($null -ne $v -and $v -ne [System.DBNull]::Value){return$v}}catch{};return$D}
        $masterMap=@{};$m=$fi.GetRecordset('SELECT Code,Name FROM Master1');if($m){while(-not $m.EOF){try{$c=[int]$m.Fields.Item('Code').Value;if($c -gt 0){$masterMap[$c]=[string]$m.Fields.Item('Name').Value}}catch{};$m.MoveNext()};try{$m.Close()}catch{}}
        $from=Convert-BusyCloudApprovalFilterDate $FromDate;$to=Convert-BusyCloudApprovalFilterDate $ToDate;$dateSql='';if($from){$dateSql += " AND [Date]>=#"+$from.ToString('MM/dd/yyyy')+'#'};if($to){$dateSql += " AND [Date]<#"+$to.AddDays(1).ToString('MM/dd/yyyy')+'#'}
        $rst=$fi.GetRecordset("SELECT VchCode,VchType,VchNo,[Date],VchSeriesCode,MasterCode1,VchAmtBaseCur,VchSalePurcAmt,ApprovalStatus FROM Tran1 WHERE VchType=$VchType AND ApprovalStatus=$approvalStatus$dateSql ORDER BY VchCode DESC")
        $headers=@();$codes=@();if($rst){while(-not $rst.EOF){$code=[int](Read-ApprovalFieldV6 $rst 'VchCode' 0);$seriesCode=[int](Read-ApprovalFieldV6 $rst 'VchSeriesCode' 0);$partyCode=[int](Read-ApprovalFieldV6 $rst 'MasterCode1' 0);$rawDate=Read-ApprovalFieldV6 $rst 'Date' $null;$dateText='';if($rawDate){try{$dateText=([datetime]$rawDate).ToString('yyyy-MM-dd')}catch{$dateText=[string]$rawDate}};$amt=[double](Read-ApprovalFieldV6 $rst 'VchAmtBaseCur' 0);if([math]::Abs($amt) -lt .000001){$amt=[double](Read-ApprovalFieldV6 $rst 'VchSalePurcAmt' 0)};$headers+=@{vchCode=$code;vchNo=([string](Read-ApprovalFieldV6 $rst 'VchNo' '')).Trim();date=$dateText;seriesCode=$seriesCode;partyCode=$partyCode;amount=[math]::Abs($amt)};if($code -gt 0){$codes+=$code};$rst.MoveNext()};try{$rst.Close()}catch{}}
        $createdMap=@{};$approvedMap=@{};for($i=0;$i -lt $codes.Count;$i+=150){$end=[math]::Min($i+149,$codes.Count-1);$batch=@($codes[$i..$end]);if($batch.Count -eq 0){continue};$cs=$batch-join',';$a=$fi.GetRecordset("SELECT Code,Action,ActionTime,UserName,Notes FROM CheckList WHERE Code IN ($cs) AND Action IN (1,3) ORDER BY Code,ActionTime,Action");if($a){while(-not $a.EOF){$c=[int](Read-ApprovalFieldV6 $a 'Code' 0);$act=[int](Read-ApprovalFieldV6 $a 'Action' 0);$u=([string](Read-ApprovalFieldV6 $a 'UserName' '')).Trim();$n=[string](Read-ApprovalFieldV6 $a 'Notes' '');$raw=Read-ApprovalFieldV6 $a 'ActionTime' $null;$tm='';if($raw){try{$tm=([datetime]$raw).ToString('yyyy-MM-dd HH:mm:ss')}catch{}};if($act -eq 1 -and -not $createdMap.ContainsKey($c)){$createdMap[$c]=@{user=$u;time=$tm}}elseif($act -eq 3){$approvedMap[$c]=@{user=$u;time=$tm;notes=$n}};$a.MoveNext()};try{$a.Close()}catch{}}}
        $needle=([string]$Search).Trim().ToLowerInvariant();$all=@();foreach($h in $headers){$c=[int]$h.vchCode;$series=if($masterMap.ContainsKey([int]$h.seriesCode)){([string]$masterMap[([int]$h.seriesCode)]).Trim()}else{'Main'};$prefix='{0:D2}'-f$VchType;if($series.StartsWith($prefix,[System.StringComparison]::OrdinalIgnoreCase)){$series=$series.Substring($prefix.Length)};$party=if($masterMap.ContainsKey([int]$h.partyCode)){([string]$masterMap[([int]$h.partyCode)]).Trim()}else{''};$cr=if($createdMap.ContainsKey($c)){$createdMap[$c]}else{@{user='';time=''}};$ap=if($approvedMap.ContainsKey($c)){$approvedMap[$c]}else{@{user='';time='';notes=''}};if($needle){$hay=(@($h.vchNo,$series,$party,$cr.user,$ap.user,$ap.notes,$c)-join' ').ToLowerInvariant();if(-not $hay.Contains($needle)){continue}};$all+=@{vchCode=$c;vchType=$VchType;vchNo=$h.vchNo;date=$h.date;vchSeriesCode=$h.seriesCode;vchSeries=$series;party=$party;amount=$h.amount;approvalStatus=$approvalStatus;approvalStatusText=if($approvalStatus -eq 1){'Approved'}else{'To be Approved'};createdBy=$cr.user;createdAt=$cr.time;approvedBy=$ap.user;approvedAt=$ap.time;approvalRemarks=$ap.notes}}
        $total=$all.Count;if($PageSize -le 0){$paged=@($all);$Page=1}else{$tp=if($total -le 0){1}else{[int][math]::Ceiling($total/[double]$PageSize)};if($Page -gt $tp){$Page=$tp};$skip=($Page-1)*$PageSize;$paged=@($all|Select-Object -Skip $skip -First $PageSize)}
        return @{success=$true;data=@{allowed_vch_types=@($allowedTypes);status=$normalized;items=@($paged);pagination=(New-BusyCloudApprovalPagination -Page $Page -PageSize $PageSize -Total $total)}}
    }catch{return @{success=$false;error=$_.Exception.Message}}finally{Disconnect-BUSY $fi}
}

function Get-VoucherApprovalQueue {
    param(
        [string]$UserName,[bool]$IsAdmin=$false,[string]$Status='pending',[int]$VchType=0,
        [string]$Search='',[string]$FromDate='',[string]$ToDate='',[int]$Page=1,[int]$PageSize=50,
        [int]$Limit=0,[string]$InstanceId='',[string]$CompanyCode=''
    )
    # Backward compatibility: old callers used Limit only.
    if($PageSize -eq 50 -and $Limit -gt 0){if($Limit -in @(50,100,200,300,500)){$PageSize=$Limit}}
    if(Test-BusyCloudApprovalAccessRuntime -InstanceId $InstanceId -CompanyCode $CompanyCode){
        return Get-VoucherApprovalQueue-AccessComV6 -UserName $UserName -IsAdmin:$IsAdmin -Status $Status -VchType $VchType -Search $Search -FromDate $FromDate -ToDate $ToDate -Page $Page -PageSize $PageSize -InstanceId $InstanceId -CompanyCode $CompanyCode
    }
    return Get-VoucherApprovalQueue-DirectV6 -UserName $UserName -IsAdmin:$IsAdmin -Status $Status -VchType $VchType -Search $Search -FromDate $FromDate -ToDate $ToDate -Page $Page -PageSize $PageSize -InstanceId $InstanceId -CompanyCode $CompanyCode
}

# -----------------------------------------------------------------------------
# HOTFIX V7: SQL approval-list query hardening
# - keeps default server-side paging at 50
# - removes unnecessary joins from COUNT queries
# - pages Tran1 before joining Master1
# - batches CheckList lookups so "All" does not create one huge IN clause
# - uses read-only NOLOCK reads for list/search screens to avoid blocking BUSY
# - preserves Access/BDS COM path from V6
# -----------------------------------------------------------------------------
function Get-VoucherApprovalQueue-DirectV7 {
    param(
        [string]$UserName,
        [bool]$IsAdmin=$false,
        [string]$Status='pending',
        [int]$VchType=0,
        [string]$Search='',
        [string]$FromDate='',
        [string]$ToDate='',
        [int]$Page=1,
        [int]$PageSize=50,
        [string]$InstanceId='',
        [string]$CompanyCode=''
    )

    if ([string]::IsNullOrWhiteSpace($UserName)) {
        return @{ success=$false; httpStatus=401; error='Authenticated BUSY user is required.' }
    }

    $allowed = Get-VoucherApprovalTypesForUser `
        -UserName $UserName `
        -IsAdmin:$IsAdmin `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $allowed.success) { return $allowed }

    $allowedTypes = @(
        $allowed.data |
            ForEach-Object { [int]$_ } |
            Where-Object { $_ -gt 0 } |
            Sort-Object -Unique
    )

    $normalized = ([string]$Status).Trim().ToLowerInvariant()
    if ($normalized -notin @('pending','approved','history')) {
        return @{ success=$false; httpStatus=400; error='status must be pending, approved, or history.' }
    }

    if ($Page -lt 1) { $Page = 1 }
    if ($PageSize -notin @(0,50,100,200,300,500)) { $PageSize = 50 }

    # Metadata-only request used by the frontend to discover enabled/allowed types.
    if ($VchType -le 0) {
        return @{
            success=$true
            data=@{
                allowed_vch_types=@($allowedTypes)
                status=$normalized
                items=@()
                pagination=(New-BusyCloudApprovalPagination -Page 1 -PageSize $PageSize -Total 0)
            }
        }
    }

    if ($allowedTypes -notcontains $VchType) {
        return @{
            success=$false
            httpStatus=403
            error="Approval processing is not enabled for voucher type $VchType, or you are not an approver for it."
        }
    }

    if ($normalized -eq 'history') {
        $historyUser = if ($IsAdmin) { '' } else { $UserName }
        $history = Get-BusyCloudVoucherApprovalHistoryV6 `
            -VchTypes @($VchType) `
            -VchType $VchType `
            -ActionBy $historyUser `
            -Search $Search `
            -FromDate $FromDate `
            -ToDate $ToDate `
            -Page $Page `
            -PageSize $PageSize `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (-not $history.success) { return $history }

        return @{
            success=$true
            data=@{
                allowed_vch_types=@($allowedTypes)
                status='history'
                items=@($history.data)
                pagination=$history.pagination
            }
        }
    }

    $approvalStatus = if ($normalized -eq 'approved') { 1 } else { 2 }
    $ctx = $null

    try {
        $ctx = Get-BusyCloudApprovalDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $conn = $ctx.connection

        $whereParts = New-Object System.Collections.Generic.List[string]
        [void]$whereParts.Add('T.VchType = @VchType')
        [void]$whereParts.Add('T.ApprovalStatus = @ApprovalStatus')

        $from = Convert-BusyCloudApprovalFilterDate $FromDate
        $to = Convert-BusyCloudApprovalFilterDate $ToDate
        if ($from) { [void]$whereParts.Add('T.[Date] >= @FromDate') }
        if ($to) { [void]$whereParts.Add('T.[Date] < @ToDateExclusive') }

        $needle = ([string]$Search).Trim()
        if (-not [string]::IsNullOrWhiteSpace($needle)) {
            [void]$whereParts.Add(@"
(
    T.VchNo LIKE @SearchLike
    OR CONVERT(VARCHAR(30), T.VchCode) LIKE @SearchLike
    OR EXISTS (
        SELECT 1 FROM Master1 PM WITH (NOLOCK)
        WHERE PM.Code = T.MasterCode1 AND PM.Name LIKE @SearchLike
    )
    OR EXISTS (
        SELECT 1 FROM Master1 SM WITH (NOLOCK)
        WHERE SM.Code = T.VchSeriesCode AND SM.Name LIKE @SearchLike
    )
    OR EXISTS (
        SELECT 1 FROM CheckList CL WITH (NOLOCK)
        WHERE CL.Code = T.VchCode
          AND CL.Action IN (1,3)
          AND (ISNULL(CL.UserName,'') LIKE @SearchLike OR ISNULL(CL.Notes,'') LIKE @SearchLike)
    )
)
"@)
        }

        $whereSql = $whereParts -join ' AND '
        $timeoutSeconds = if ($PageSize -le 0) { 120 } else { 30 }

        # COUNT only Tran1. Do not join Master1 unless search explicitly needs it.
        $countCmd = $conn.CreateCommand()
        $countCmd.CommandTimeout = $timeoutSeconds
        $countCmd.CommandText = "SELECT COUNT_BIG(1) FROM Tran1 T WITH (NOLOCK) WHERE $whereSql"
        [void]$countCmd.Parameters.AddWithValue('@VchType', $VchType)
        [void]$countCmd.Parameters.AddWithValue('@ApprovalStatus', $approvalStatus)
        if ($from) { [void]$countCmd.Parameters.AddWithValue('@FromDate', $from) }
        if ($to) { [void]$countCmd.Parameters.AddWithValue('@ToDateExclusive', $to.AddDays(1)) }
        if (-not [string]::IsNullOrWhiteSpace($needle)) {
            [void]$countCmd.Parameters.AddWithValue('@SearchLike', ('%' + $needle + '%'))
        }

        $total64 = [long]$countCmd.ExecuteScalar()
        $total = if ($total64 -gt [int]::MaxValue) { [int]::MaxValue } else { [int]$total64 }
        $pagination = New-BusyCloudApprovalPagination -Page $Page -PageSize $PageSize -Total $total
        $Page = [int]$pagination.page

        # First page Tran1, then join Master1 only for those rows.
        $pagingSql = 'OFFSET 0 ROWS'
        if ($PageSize -gt 0) {
            $offset = ($Page - 1) * $PageSize
            $pagingSql = "OFFSET $offset ROWS FETCH NEXT $PageSize ROWS ONLY"
        }

        $dataSql = @"
;WITH PageRows AS (
    SELECT
        T.VchCode,
        T.VchType,
        T.VchNo,
        T.[Date],
        T.VchSeriesCode,
        T.MasterCode1,
        T.VchAmtBaseCur,
        T.VchSalePurcAmt,
        T.ApprovalStatus
    FROM Tran1 T WITH (NOLOCK)
    WHERE $whereSql
    ORDER BY T.VchCode DESC
    $pagingSql
)
SELECT
    R.VchCode,
    R.VchType,
    R.VchNo,
    R.[Date],
    R.VchSeriesCode,
    R.MasterCode1,
    R.VchAmtBaseCur,
    R.VchSalePurcAmt,
    R.ApprovalStatus,
    ISNULL(P.Name,'') AS PartyName,
    ISNULL(S.Name,'') AS SeriesName
FROM PageRows R
LEFT JOIN Master1 P WITH (NOLOCK) ON P.Code = R.MasterCode1
LEFT JOIN Master1 S WITH (NOLOCK) ON S.Code = R.VchSeriesCode
ORDER BY R.VchCode DESC
"@

        $cmd = $conn.CreateCommand()
        $cmd.CommandTimeout = $timeoutSeconds
        $cmd.CommandText = $dataSql
        [void]$cmd.Parameters.AddWithValue('@VchType', $VchType)
        [void]$cmd.Parameters.AddWithValue('@ApprovalStatus', $approvalStatus)
        if ($from) { [void]$cmd.Parameters.AddWithValue('@FromDate', $from) }
        if ($to) { [void]$cmd.Parameters.AddWithValue('@ToDateExclusive', $to.AddDays(1)) }
        if (-not [string]::IsNullOrWhiteSpace($needle)) {
            [void]$cmd.Parameters.AddWithValue('@SearchLike', ('%' + $needle + '%'))
        }

        $rows = @()
        $codes = @()
        $rdr = $cmd.ExecuteReader()
        try {
            while ($rdr.Read()) {
                $code = [int](Read-BusyCloudApprovalReaderValue $rdr 'VchCode' 0)
                $rawDate = Read-BusyCloudApprovalReaderValue $rdr 'Date' $null
                $dateText = ''
                if ($rawDate) {
                    try { $dateText = ([datetime]$rawDate).ToString('yyyy-MM-dd') }
                    catch { $dateText = [string]$rawDate }
                }

                $amount = [double](Read-BusyCloudApprovalReaderValue $rdr 'VchAmtBaseCur' 0)
                if ([math]::Abs($amount) -lt 0.000001) {
                    $amount = [double](Read-BusyCloudApprovalReaderValue $rdr 'VchSalePurcAmt' 0)
                }

                $seriesName = ([string](Read-BusyCloudApprovalReaderValue $rdr 'SeriesName' 'Main')).Trim()
                if (-not $seriesName) { $seriesName = 'Main' }
                $prefix = '{0:D2}' -f $VchType
                if ($seriesName.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $seriesName = $seriesName.Substring($prefix.Length)
                }

                $rows += @{
                    vchCode=$code
                    vchType=$VchType
                    vchNo=([string](Read-BusyCloudApprovalReaderValue $rdr 'VchNo' '')).Trim()
                    date=$dateText
                    vchSeriesCode=[int](Read-BusyCloudApprovalReaderValue $rdr 'VchSeriesCode' 0)
                    vchSeries=$seriesName
                    party=([string](Read-BusyCloudApprovalReaderValue $rdr 'PartyName' '')).Trim()
                    amount=[math]::Abs($amount)
                }
                if ($code -gt 0) { $codes += $code }
            }
        }
        finally {
            try { $rdr.Close() } catch {}
        }

        # Read checklist metadata in bounded batches. This is especially important
        # when the user explicitly selects Rows=All.
        $createdMap = @{}
        $approvedMap = @{}
        $batchSize = 250

        for ($start = 0; $start -lt $codes.Count; $start += $batchSize) {
            $end = [math]::Min($start + $batchSize - 1, $codes.Count - 1)
            $batch = @($codes[$start..$end])
            if ($batch.Count -eq 0) { continue }

            $codesSql = $batch -join ','
            $auditCmd = $conn.CreateCommand()
            $auditCmd.CommandTimeout = $timeoutSeconds
            $auditCmd.CommandText = "SELECT Code,Action,ActionTime,UserName,Notes FROM CheckList WITH (NOLOCK) WHERE Code IN ($codesSql) AND Action IN (1,3) ORDER BY Code,ActionTime,Action"

            $auditReader = $auditCmd.ExecuteReader()
            try {
                while ($auditReader.Read()) {
                    $code = [int](Read-BusyCloudApprovalReaderValue $auditReader 'Code' 0)
                    $action = [int](Read-BusyCloudApprovalReaderValue $auditReader 'Action' 0)
                    $user = ([string](Read-BusyCloudApprovalReaderValue $auditReader 'UserName' '')).Trim()
                    $notes = [string](Read-BusyCloudApprovalReaderValue $auditReader 'Notes' '')
                    $rawTime = Read-BusyCloudApprovalReaderValue $auditReader 'ActionTime' $null
                    $time = ''
                    if ($rawTime) {
                        try { $time = ([datetime]$rawTime).ToString('yyyy-MM-dd HH:mm:ss') } catch {}
                    }

                    if ($action -eq 1 -and -not $createdMap.ContainsKey($code)) {
                        $createdMap[$code] = @{ user=$user; time=$time }
                    }
                    elseif ($action -eq 3) {
                        $approvedMap[$code] = @{ user=$user; time=$time; notes=$notes }
                    }
                }
            }
            finally {
                try { $auditReader.Close() } catch {}
            }
        }

        $items = @()
        foreach ($row in $rows) {
            $code = [int]$row.vchCode
            $created = if ($createdMap.ContainsKey($code)) { $createdMap[$code] } else { @{ user=''; time='' } }
            $approved = if ($approvedMap.ContainsKey($code)) { $approvedMap[$code] } else { @{ user=''; time=''; notes='' } }

            $items += @{
                vchCode=$code
                vchType=$row.vchType
                vchNo=$row.vchNo
                date=$row.date
                vchSeriesCode=$row.vchSeriesCode
                vchSeries=$row.vchSeries
                party=$row.party
                amount=$row.amount
                approvalStatus=$approvalStatus
                approvalStatusText=if ($approvalStatus -eq 1) { 'Approved' } else { 'To be Approved' }
                createdBy=$created.user
                createdAt=$created.time
                approvedBy=$approved.user
                approvedAt=$approved.time
                approvalRemarks=$approved.notes
            }
        }

        return @{
            success=$true
            data=@{
                allowed_vch_types=@($allowedTypes)
                status=$normalized
                items=@($items)
                pagination=$pagination
            }
        }
    }
    catch {
        return @{ success=$false; error=$_.Exception.Message }
    }
    finally {
        if ($ctx -and $ctx.connection) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}

# Final runtime dispatcher override. Access/BDS keeps the V6 COM implementation;
# SQL Server uses the optimized V7 direct implementation above.
function Get-VoucherApprovalQueue {
    param(
        [string]$UserName,
        [bool]$IsAdmin=$false,
        [string]$Status='pending',
        [int]$VchType=0,
        [string]$Search='',
        [string]$FromDate='',
        [string]$ToDate='',
        [int]$Page=1,
        [int]$PageSize=50,
        [int]$Limit=0,
        [string]$InstanceId='',
        [string]$CompanyCode=''
    )

    if ($PageSize -eq 50 -and $Limit -gt 0 -and $Limit -in @(50,100,200,300,500)) {
        $PageSize = $Limit
    }

    if (Test-BusyCloudApprovalAccessRuntime -InstanceId $InstanceId -CompanyCode $CompanyCode) {
        return Get-VoucherApprovalQueue-AccessComV6 `
            -UserName $UserName `
            -IsAdmin:$IsAdmin `
            -Status $Status `
            -VchType $VchType `
            -Search $Search `
            -FromDate $FromDate `
            -ToDate $ToDate `
            -Page $Page `
            -PageSize $PageSize `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }

    return Get-VoucherApprovalQueue-DirectV7 `
        -UserName $UserName `
        -IsAdmin:$IsAdmin `
        -Status $Status `
        -VchType $VchType `
        -Search $Search `
        -FromDate $FromDate `
        -ToDate $ToDate `
        -Page $Page `
        -PageSize $PageSize `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode
}
