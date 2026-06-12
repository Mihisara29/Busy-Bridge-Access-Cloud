# modules/vch_setting.ps1
# Voucher Column & POS Settings — Direct Database Persistence Engine

if ($null -eq $script:Config) {
    . "$PSScriptRoot\connection.ps1"
    . "$PSScriptRoot\utils.ps1"
}

# ═══════════════════════════════════════════════════════════════
#  GET COLUMN CONFIG FROM CONFIG TABLE (Type=VchType, D15=DeviceType)
# ═══════════════════════════════════════════════════════════════
function Get-ColumnConfig {
    param(
        [int]$VchType,
        [int]$DeviceType, # 0 = Desktop, 1 = Mobile/Tab
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) {
        return @{ success = $false; error = "BUSY database connection failed" }
    }

    try {
        # RecType=201, [Type] stores the VchType, and D15 stores the DeviceType (0 or 1)
        $qry = "SELECT * FROM Config WHERE RecType=201 AND [Type]=$VchType AND D15=$DeviceType"
        $rst = $fi.GetRecordset($qry)
        
        function Map-IntToBehavior {
            param([int]$val)
            if ($val -eq 2) { return "semi_variable" }
            if ($val -eq 3) { return "fixed" }
            return "variable"
        }

        $config = $null
        if ($rst -and -not $rst.EOF) {
            $rst.MoveFirst()
            
            function Get-SafeVal {
                param($rst, [string]$field, $default)
                try {
                    $item = $rst.Fields.Item($field).Value
                    if ($item -ne [System.DBNull]::Value -and $null -ne $item) {
                        return $item
                    }
                } catch {}
                return $default
            }

            $i1  = [int](Get-SafeVal -rst $rst -field "I1" -default 1)
            $i2  = [int](Get-SafeVal -rst $rst -field "I2" -default 1)
            $i3  = [int](Get-SafeVal -rst $rst -field "I3" -default 1)
            $i4  = [int](Get-SafeVal -rst $rst -field "I4" -default 1)
            $i5  = [int](Get-SafeVal -rst $rst -field "I5" -default 1)
            $i6  = [int](Get-SafeVal -rst $rst -field "I6" -default 1)
            $i7  = [int](Get-SafeVal -rst $rst -field "I7" -default 1)
            $i8  = [int](Get-SafeVal -rst $rst -field "I8" -default 1)
            $i9  = [int](Get-SafeVal -rst $rst -field "I9" -default 1)
            $i10 = [int](Get-SafeVal -rst $rst -field "I10" -default 1)
            $i11 = [int](Get-SafeVal -rst $rst -field "I11" -default 0)
            $c1  = [string](Get-SafeVal -rst $rst -field "C1" -default "")
            $c2  = [string](Get-SafeVal -rst $rst -field "C2" -default "")

            $config = @{
                vch_type             = $VchType
                device_type          = $DeviceType
                enable_item_discount = ($i1 -eq 1)
                enable_alt_units     = ($i2 -eq 1)
                col_qty              = Map-IntToBehavior $i3
                col_unit             = Map-IntToBehavior $i4
                col_price            = Map-IntToBehavior $i5
                col_amount           = Map-IntToBehavior $i6
                col_discount         = Map-IntToBehavior $i7
                col_cfact            = Map-IntToBehavior $i8
                col_alt_qty          = Map-IntToBehavior $i9
                col_alt_price        = Map-IntToBehavior $i10
                enable_pos           = ($i11 -eq 1)
                def_card_acc         = $c1
                def_gift_acc         = $c2
            }
            $rst.Close()
        }

        # Safe defaults if no record exists yet
        if ($null -eq $config) {
            $config = @{
                vch_type             = $VchType
                device_type          = $DeviceType
                enable_item_discount = $true
                enable_alt_units     = $true
                col_qty              = "variable"
                col_unit             = "variable"
                col_price            = "variable"
                col_amount           = "variable"
                col_discount         = "variable"
                col_cfact            = "variable"
                col_alt_qty          = "variable"
                col_alt_price        = "variable"
                enable_pos           = $false
                def_card_acc         = ""
                def_gift_acc         = ""
            }
        }

        return @{ success = $true; data = $config }
    } catch {
        return @{ success = $false; error = $_.Exception.Message }
    } finally {
        Disconnect-BUSY $fi
    }
}

# ═══════════════════════════════════════════════════════════════
#  SAVE COLUMN CONFIG TO CONFIG TABLE (Type=VchType, D15=DeviceType)
# ═══════════════════════════════════════════════════════════════
function Save-ColumnConfig {
    param(
        $Data,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $vchType    = [int]$Data.vch_type
    $deviceType = if ($null -ne $Data.device_type) { [int]$Data.device_type } else { 0 } # Default: 0 (Desktop)

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) {
        return @{ success = $false; error = "BUSY database connection failed" }
    }

    try {
        function Map-BehaviorToInt {
            param([string]$val)
            if ($val -eq "semi_variable") { return 2 }
            if ($val -eq "fixed") { return 3 }
            return 1
        }

        $i1  = if ($Data.enable_item_discount -eq $true -or $Data.enable_item_discount -eq "true") { 1 } else { 0 }
        $i2  = if ($Data.enable_alt_units -eq $true -or $Data.enable_alt_units -eq "true") { 1 } else { 0 }
        $i3  = Map-BehaviorToInt $Data.col_qty
        $i4  = Map-BehaviorToInt $Data.col_unit
        $i5  = Map-BehaviorToInt $Data.col_price
        $i6  = Map-BehaviorToInt $Data.col_amount
        $i7  = Map-BehaviorToInt $Data.col_discount
        $i8  = Map-BehaviorToInt $Data.col_cfact
        $i9  = Map-BehaviorToInt $Data.col_alt_qty
        $i10 = Map-BehaviorToInt $Data.col_alt_price
        $i11 = if ($Data.enable_pos -eq $true -or $Data.enable_pos -eq "true") { 1 } else { 0 }
        
        $c1  = if ($Data.def_card_acc) { $Data.def_card_acc -replace "'", "''" } else { "" }
        $c2  = if ($Data.def_gift_acc) { $Data.def_gift_acc -replace "'", "''" } else { "" }

        # Check if record exists
        $exists = $false
        $rst = $fi.GetRecordset("SELECT RecType FROM Config WHERE RecType=201 AND [Type]=$vchType AND D15=$deviceType")
        if ($rst -and -not $rst.EOF) {
            $exists = $true
            $rst.Close()
        }

        if ($exists) {
            $sql = "UPDATE Config SET I1=$i1, I2=$i2, I3=$i3, I4=$i4, I5=$i5, I6=$i6, I7=$i7, I8=$i8, I9=$i9, I10=$i10, I11=$i11, C1='$c1', C2='$c2' WHERE RecType=201 AND [Type]=$vchType AND D15=$deviceType"
        } else {
            $sql = "INSERT INTO Config (RecType, [Type], D15, I1, I2, I3, I4, I5, I6, I7, I8, I9, I10, I11, C1, C2) VALUES (201, $vchType, $deviceType, $i1, $i2, $i3, $i4, $i5, $i6, $i7, $i8, $i9, $i10, $i11, '$c1', '$c2')"
        }

        $fi.ExecuteQuery($sql)

        return @{ success = $true; message = "Voucher configuration updated successfully" }
    } catch {
        return @{ success = $false; error = $_.Exception.Message }
    } finally {
        Disconnect-BUSY $fi
    }
}