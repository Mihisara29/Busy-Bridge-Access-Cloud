# 02-optional-fields.ps1
# Extracted from vouchers.ps1. Keep functions behavior-compatible with the monolith.

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

