# modules/vch_setting.ps1
# Voucher UI settings + Web/Busy voucher numbering configuration.

if ($null -eq $script:Config) {
    . "$PSScriptRoot\connection.ps1"
    . "$PSScriptRoot\utils.ps1"
}

# -----------------------------------------------------------------------------
# Existing column/POS configuration
# -----------------------------------------------------------------------------
Write-Host "  [FAST-NUMBERING-V4] Direct numbering read path loaded." -ForegroundColor DarkCyan

function Get-ColumnConfig {
    param(
        [int]$VchType,
        [int]$DeviceType,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY database connection failed" } }

    try {
        $qry = "SELECT * FROM Config WHERE RecType=201 AND [Type]=$VchType AND D15=$DeviceType"
        $rst = $fi.GetRecordset($qry)

        function Map-IntToBehavior {
            param([int]$val)
            if ($val -eq 2) { return "semi_variable" }
            if ($val -eq 3) { return "fixed" }
            return "variable"
        }

        function Get-SafeVal {
            param($Recordset, [string]$Field, $Default)
            try {
                $value = $Recordset.Fields.Item($Field).Value
                if ($null -ne $value -and $value -ne [System.DBNull]::Value) { return $value }
            } catch {}
            return $Default
        }

        # Account voucher forms use the same RecType=201 record identity
        # (voucher type + device type) but interpret I1..I10 as account-form settings.
        # Receipt (14) and Payment (19) support Quick/Double.
        # Contra (15) and Journal (16) are Double-entry only.
        $isAccountVoucher = @(14, 15, 16, 19) -contains $VchType
        $supportsQuickMode = @(14, 19) -contains $VchType

        $config = $null
        if ($rst -and -not $rst.EOF) {
            if ($isAccountVoucher) {
                $storedMode = [int](Get-SafeVal $rst "I1" 1)
                $defaultMode = if ($supportsQuickMode -and $storedMode -eq 1) { "single" } else { "double" }

                $config = @{
                    vch_type                         = $VchType
                    device_type                      = $DeviceType
                    account_default_mode             = $defaultMode
                    account_allow_mode_switch        = if ($supportsQuickMode) { ([int](Get-SafeVal $rst "I2" 1) -eq 1) } else { $false }

                    # Quick-entry field behavior
                    acc_quick_col_account            = Map-IntToBehavior ([int](Get-SafeVal $rst "I3" 1))
                    acc_quick_col_amount             = Map-IntToBehavior ([int](Get-SafeVal $rst "I4" 1))
                    acc_quick_col_short_narration    = Map-IntToBehavior ([int](Get-SafeVal $rst "I5" 1))

                    # Double-entry field behavior
                    acc_double_col_dc                = Map-IntToBehavior ([int](Get-SafeVal $rst "I6" 1))
                    acc_double_col_account           = Map-IntToBehavior ([int](Get-SafeVal $rst "I7" 1))
                    acc_double_col_debit             = Map-IntToBehavior ([int](Get-SafeVal $rst "I8" 1))
                    acc_double_col_credit            = Map-IntToBehavior ([int](Get-SafeVal $rst "I9" 1))
                    acc_double_col_short_narration   = Map-IntToBehavior ([int](Get-SafeVal $rst "I10" 1))
                }
            } else {
                $config = @{
                    vch_type             = $VchType
                    device_type          = $DeviceType
                    enable_item_discount = ([int](Get-SafeVal $rst "I1" 1) -eq 1)
                    enable_alt_units     = ([int](Get-SafeVal $rst "I2" 1) -eq 1)
                    col_qty              = Map-IntToBehavior ([int](Get-SafeVal $rst "I3" 1))
                    col_unit             = Map-IntToBehavior ([int](Get-SafeVal $rst "I4" 1))
                    col_price            = Map-IntToBehavior ([int](Get-SafeVal $rst "I5" 1))
                    col_amount           = Map-IntToBehavior ([int](Get-SafeVal $rst "I6" 1))
                    col_discount         = Map-IntToBehavior ([int](Get-SafeVal $rst "I7" 1))
                    col_cfact            = Map-IntToBehavior ([int](Get-SafeVal $rst "I8" 1))
                    col_alt_qty          = Map-IntToBehavior ([int](Get-SafeVal $rst "I9" 1))
                    col_alt_price        = Map-IntToBehavior ([int](Get-SafeVal $rst "I10" 1))
                    enable_pos           = ([int](Get-SafeVal $rst "I11" 0) -eq 1)

                    # C3 is free for RecType 201 inventory-voucher rows and stores whether
                    # material-centre stock balances are shown in item search UIs.
                    # Default to enabled so existing companies keep the current behavior.
                    show_stock_balance    = ([string](Get-SafeVal $rst "C3" "1") -ne "0")
                    def_card_acc          = [string](Get-SafeVal $rst "C1" "")
                    def_gift_acc          = [string](Get-SafeVal $rst "C2" "")
                }
            }
        }

        if ($rst) { try { $rst.Close() } catch {} }

        if ($null -eq $config) {
            if ($isAccountVoucher) {
                $config = @{
                    vch_type                         = $VchType
                    device_type                      = $DeviceType
                    account_default_mode             = if ($supportsQuickMode) { "single" } else { "double" }
                    account_allow_mode_switch        = if ($supportsQuickMode) { $true } else { $false }
                    acc_quick_col_account            = "variable"
                    acc_quick_col_amount             = "variable"
                    acc_quick_col_short_narration    = "variable"
                    acc_double_col_dc                = "variable"
                    acc_double_col_account           = "variable"
                    acc_double_col_debit             = "variable"
                    acc_double_col_credit            = "variable"
                    acc_double_col_short_narration   = "variable"
                }
            } else {
                $config = @{
                    vch_type=$VchType; device_type=$DeviceType
                    enable_item_discount=$true; enable_alt_units=$true
                    col_qty="variable"; col_unit="variable"; col_price="variable"; col_amount="variable"
                    col_discount="variable"; col_cfact="variable"; col_alt_qty="variable"; col_alt_price="variable"
                    enable_pos=$false; show_stock_balance=$true; def_card_acc=""; def_gift_acc=""
                }
            }
        }

        return @{ success = $true; data = $config }
    } catch {
        return @{ success = $false; error = $_.Exception.Message }
    } finally {
        Disconnect-BUSY $fi
    }
}

function Save-ColumnConfig {
    param($Data, [string]$InstanceId = "", [string]$CompanyCode = "")

    $vchType = [int]$Data.vch_type
    $deviceType = if ($null -ne $Data.device_type) { [int]$Data.device_type } else { 0 }
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY database connection failed" } }

    try {
        function Map-BehaviorToInt {
            param([string]$val)
            if ($val -eq "semi_variable") { return 2 }
            if ($val -eq "fixed") { return 3 }
            return 1
        }

        $isAccountVoucher = @(14, 15, 16, 19) -contains $vchType
        $supportsQuickMode = @(14, 19) -contains $vchType

        $exists = $false
        $rst = $fi.GetRecordset("SELECT RecType FROM Config WHERE RecType=201 AND [Type]=$vchType AND D15=$deviceType")
        if ($rst -and -not $rst.EOF) { $exists = $true }
        if ($rst) { try { $rst.Close() } catch {} }

        if ($isAccountVoucher) {
            # Account voucher RecType=201 mapping:
            # I1  = default entry mode: 1 Quick, 2 Double
            # I2  = allow Quick/Double switching: 0/1
            # I3  = Quick Account behavior
            # I4  = Quick Amount behavior
            # I5  = Quick Short Narration behavior
            # I6  = Double D/C behavior
            # I7  = Double Account behavior
            # I8  = Double Debit behavior
            # I9  = Double Credit behavior
            # I10 = Double Short Narration behavior
            $requestedMode = ([string]$Data.account_default_mode).ToLowerInvariant()
            $i1 = if ($supportsQuickMode -and $requestedMode -eq "single") { 1 } else { 2 }
            $i2 = if ($supportsQuickMode -and ($Data.account_allow_mode_switch -eq $true -or $Data.account_allow_mode_switch -eq "true")) { 1 } else { 0 }

            $i3  = Map-BehaviorToInt $Data.acc_quick_col_account
            $i4  = Map-BehaviorToInt $Data.acc_quick_col_amount
            $i5  = Map-BehaviorToInt $Data.acc_quick_col_short_narration
            $i6  = Map-BehaviorToInt $Data.acc_double_col_dc
            $i7  = Map-BehaviorToInt $Data.acc_double_col_account
            $i8  = Map-BehaviorToInt $Data.acc_double_col_debit
            $i9  = Map-BehaviorToInt $Data.acc_double_col_credit
            $i10 = Map-BehaviorToInt $Data.acc_double_col_short_narration

            if ($exists) {
                $sql = "UPDATE Config SET I1=$i1,I2=$i2,I3=$i3,I4=$i4,I5=$i5,I6=$i6,I7=$i7,I8=$i8,I9=$i9,I10=$i10 WHERE RecType=201 AND [Type]=$vchType AND D15=$deviceType"
            } else {
                $sql = "INSERT INTO Config (RecType,[Type],D15,I1,I2,I3,I4,I5,I6,I7,I8,I9,I10) VALUES (201,$vchType,$deviceType,$i1,$i2,$i3,$i4,$i5,$i6,$i7,$i8,$i9,$i10)"
            }
        } else {
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

            # Keep stock visibility ON when older clients do not send this field.
            # C3 is scoped to RecType 201 inventory-voucher rows, so it does not conflict
            # with C3 used by RecType 202 numbering rows.
            $c3 = if ($null -eq $Data.show_stock_balance) { "1" } elseif ($Data.show_stock_balance -eq $true -or $Data.show_stock_balance -eq "true") { "1" } else { "0" }
            $c1 = if ($Data.def_card_acc) { ([string]$Data.def_card_acc).Replace("'", "''") } else { "" }
            $c2 = if ($Data.def_gift_acc) { ([string]$Data.def_gift_acc).Replace("'", "''") } else { "" }

            if ($exists) {
                $sql = "UPDATE Config SET I1=$i1,I2=$i2,I3=$i3,I4=$i4,I5=$i5,I6=$i6,I7=$i7,I8=$i8,I9=$i9,I10=$i10,I11=$i11,C1='$c1',C2='$c2',C3='$c3' WHERE RecType=201 AND [Type]=$vchType AND D15=$deviceType"
            } else {
                $sql = "INSERT INTO Config (RecType,[Type],D15,I1,I2,I3,I4,I5,I6,I7,I8,I9,I10,I11,C1,C2,C3) VALUES (201,$vchType,$deviceType,$i1,$i2,$i3,$i4,$i5,$i6,$i7,$i8,$i9,$i10,$i11,'$c1','$c2','$c3')"
            }
        }

        $fi.ExecuteQuery($sql)
        return @{ success = $true; message = "Voucher configuration updated successfully" }
    } catch {
        return @{ success = $false; error = $_.Exception.Message }
    } finally {
        Disconnect-BUSY $fi
    }
}

# -----------------------------------------------------------------------------
# BusyCloud voucher approval configuration + custom approval audit storage
# -----------------------------------------------------------------------------
# Verified BUSY mappings used by this feature:
#   Tran1.ApprovalStatus: 0 = Approval Not Required, 1 = Approved, 2 = Pending
#   CheckList.Action:      1 = Created, 2 = Modified, 3 = Approved
#
# BusyCloud custom Config mappings:
#   RecType = 203 -> one policy row per voucher type
#       Type = voucher type
#       I1   = approval processing required for NEW vouchers (0/1)
#   RecType = 204 -> approver assignments
#       Type = voucher type
#       I1   = active assignment (0/1)
#       C1   = BUSY username
#
# Approval policy is intentionally company + voucher-type specific and is not
# device-specific. Approver assignments are retained even when processing is
# switched OFF so existing pending vouchers remain manageable and the policy can
# later be re-enabled without losing the assignment list.
# -----------------------------------------------------------------------------

$script:BusyCloudApprovalVoucherTypes = @(9,26,12,11,3,2,27,13,4,10,14,19,15,16,5,8,6)
$script:BusyCloudApprovalAuditTable = "BusyCloudVoucherApprovalAudit"

function Test-IsBusyCloudApprovalVoucherType {
    param([int]$VchType)
    return @($script:BusyCloudApprovalVoucherTypes) -contains $VchType
}

function Get-BusyCloudApprovalSupportedVoucherTypes {
    return @($script:BusyCloudApprovalVoucherTypes)
}

function Get-VoucherApprovalConfig-Direct {
    param(
        [int]$VchType,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if (-not (Test-IsBusyCloudApprovalVoucherType -VchType $VchType)) {
        return @{ success=$false; error="Voucher type $VchType is not supported by BusyCloud approval processing." }
    }

    $ctx = $null
    try {
        $ctx = Get-BusyCloudApprovalDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $conn = $ctx.connection
        $approvalRequired = $false
        $approvers = @()

        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT TOP 1 I1 FROM Config WHERE RecType=203 AND [Type]=$VchType"
        $raw = $cmd.ExecuteScalar()
        if ($null -ne $raw -and $raw -ne [System.DBNull]::Value) { $approvalRequired = ([int]$raw -eq 1) }

        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT C1 FROM Config WHERE RecType=204 AND [Type]=$VchType AND I1=1"
        $rdr = $cmd.ExecuteReader()
        while ($rdr.Read()) {
            $name = ""
            if (-not $rdr.IsDBNull(0)) { $name = ([string]$rdr.GetValue(0)).Trim() }
            if ($name -and $approvers -notcontains $name) { $approvers += $name }
        }
        $rdr.Close()

        return @{ success=$true; data=@{ vch_type=$VchType; approval_required=[bool]$approvalRequired; approvers=@($approvers | Sort-Object) } }
    }
    catch { return @{ success=$false; error=$_.Exception.Message } }
    finally {
        if ($ctx -and $ctx.connection) { try { $ctx.connection.Close() } catch {}; try { $ctx.connection.Dispose() } catch {} }
    }
}

function Save-VoucherApprovalConfig-Direct {
    param(
        $Data,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $vchType = 0
    try { $vchType = [int]$Data.vch_type } catch {}
    if (-not (Test-IsBusyCloudApprovalVoucherType -VchType $vchType)) {
        return @{ success=$false; error="Unsupported or missing vch_type." }
    }

    $rawRequired = ([string]$Data.approval_required).Trim().ToLowerInvariant()
    $approvalRequired = ($Data.approval_required -eq $true -or $rawRequired -eq "1" -or $rawRequired -eq "true")
    $requestedApprovers = @(@($Data.approvers) | ForEach-Object { ([string]$_).Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)

    $usersResult = Get-CompanyUsers -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $usersResult.success) {
        return @{ success=$false; error="Could not validate approvers against BUSY users. $($usersResult.error)" }
    }

    $canonicalByLower = @{}
    foreach ($u in @($usersResult.data)) {
        $n = ([string]$u).Trim()
        if ($n) { $canonicalByLower[$n.ToLowerInvariant()] = $n }
    }

    $approvers = @(); $unknown = @()
    foreach ($requested in $requestedApprovers) {
        $key = $requested.ToLowerInvariant()
        if ($canonicalByLower.ContainsKey($key)) {
            $canonical = [string]$canonicalByLower[$key]
            if ($approvers -notcontains $canonical) { $approvers += $canonical }
        } else { $unknown += $requested }
    }
    if ($unknown.Count -gt 0) { return @{ success=$false; error=("Unknown BUSY user(s): " + ($unknown -join ", ")) } }

    $ctx = $null
    try {
        $ctx = Get-BusyCloudApprovalDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $conn = $ctx.connection
        $requiredInt = if ($approvalRequired) { 1 } else { 0 }

        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT COUNT(*) FROM Config WHERE RecType=203 AND [Type]=$vchType"
        $exists = ([int]$cmd.ExecuteScalar() -gt 0)

        $cmd = $conn.CreateCommand()
        if ($exists) { $cmd.CommandText = "UPDATE Config SET I1=$requiredInt WHERE RecType=203 AND [Type]=$vchType" }
        else { $cmd.CommandText = "INSERT INTO Config (RecType,[Type],I1) VALUES (203,$vchType,$requiredInt)" }
        [void]$cmd.ExecuteNonQuery()

        $cmd = $conn.CreateCommand(); $cmd.CommandText = "DELETE FROM Config WHERE RecType=204 AND [Type]=$vchType"; [void]$cmd.ExecuteNonQuery()
        foreach ($name in $approvers) {
            $safe = $name.Replace("'", "''")
            $cmd = $conn.CreateCommand(); $cmd.CommandText = "INSERT INTO Config (RecType,[Type],I1,C1) VALUES (204,$vchType,1,'$safe')"; [void]$cmd.ExecuteNonQuery()
        }

        return @{ success=$true; message="Voucher approval configuration updated successfully"; data=@{ vch_type=$vchType; approval_required=[bool]$approvalRequired; approvers=@($approvers) } }
    }
    catch { return @{ success=$false; error=$_.Exception.Message } }
    finally {
        if ($ctx -and $ctx.connection) { try { $ctx.connection.Close() } catch {}; try { $ctx.connection.Dispose() } catch {} }
    }
}

function Get-VoucherApprovalTypesForUser-Direct {
    param(
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ($IsAdmin) { return @{ success=$true; data=@($script:BusyCloudApprovalVoucherTypes) } }
    if ([string]::IsNullOrWhiteSpace($UserName)) { return @{ success=$true; data=@() } }

    $ctx = $null
    try {
        $ctx = Get-BusyCloudApprovalDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $conn = $ctx.connection
        $safeUser = $UserName.Trim().Replace("'", "''")
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT [Type] FROM Config WHERE RecType=204 AND I1=1 AND C1='$safeUser'"
        $rdr = $cmd.ExecuteReader()
        $types = @()
        while ($rdr.Read()) {
            $t = [int]$rdr.GetValue(0)
            if ((Test-IsBusyCloudApprovalVoucherType -VchType $t) -and $types -notcontains $t) { $types += $t }
        }
        $rdr.Close()
        return @{ success=$true; data=@($types | Sort-Object) }
    }
    catch { return @{ success=$false; error=$_.Exception.Message } }
    finally {
        if ($ctx -and $ctx.connection) { try { $ctx.connection.Close() } catch {}; try { $ctx.connection.Dispose() } catch {} }
    }
}

function Test-VoucherApprover {
    param(
        [string]$UserName,
        [int]$VchType,
        [bool]$IsAdmin = $false,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if (-not (Test-IsBusyCloudApprovalVoucherType -VchType $VchType)) {
        return @{ success=$true; allowed=$false; allowed_vch_types=@() }
    }

    if ($IsAdmin) {
        return @{ success=$true; allowed=$true; allowed_vch_types=@($script:BusyCloudApprovalVoucherTypes) }
    }

    $types = Get-VoucherApprovalTypesForUser -UserName $UserName -IsAdmin:$false -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $types.success) { return $types }

    return @{
        success = $true
        allowed = (@($types.data) -contains $VchType)
        allowed_vch_types = @($types.data)
    }
}

function Test-VoucherModifyPermissionForUser {
    param(
        [string]$UserName,
        [int]$VchType,
        [bool]$IsAdmin = $false,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ($IsAdmin) { return @{ success=$true; allowed=$true } }
    if ([string]::IsNullOrWhiteSpace($UserName)) { return @{ success=$true; allowed=$false } }

    $map = @{
        12='C2'; 9='C4'; 3='C6'; 11='C8'; 13='C10';
        2='I2'; 10='I4'; 4='I6'; 14='I8'; 19='I10'; 16='I12'; 15='I14';
        5='I16'; 8='I18'; 26='I20'; 27='I22'; 6='B34'
    }
    if (-not $map.ContainsKey($VchType)) { return @{ success=$true; allowed=$false } }

    # MobileUserPreference belongs to the authentication/permission database,
    # not necessarily the fiscal transaction database that contains Tran1/Config.
    # Reuse the already established permission reader instead of querying the
    # approval transaction DB directly.
    try {
        $all = Get-UserPermissions -InstanceId $InstanceId -CompanyCode $CompanyCode
        if (-not $all.success) {
            return @{ success=$false; allowed=$false; error=$all.error }
        }

        $row = @($all.data | Where-Object {
            ([string]$_.name).Equals($UserName.Trim(), [System.StringComparison]::OrdinalIgnoreCase)
        } | Select-Object -First 1)

        $key = [string]$map[$VchType]
        if (-not $row) { return @{ success=$true; allowed=$false; permission_key=$key } }

        $value = 0
        try { $value = [int]$row[0].$key } catch { $value = 0 }
        return @{ success=$true; allowed=($value -ne 0); permission_key=$key }
    }
    catch { return @{ success=$false; allowed=$false; error=$_.Exception.Message } }
}

if ($null -eq $script:BusyCloudApprovalSqlDbCache) {
    $script:BusyCloudApprovalSqlDbCache = @{}
}

function Test-BusyCloudSqlTableExists {
    param(
        $Instance,
        [string]$Database,
        [string]$TableName
    )

    $conn = $null
    try {
        $conn = Open-SqlConnection `
            -SqlServer $Instance.sqlServer `
            -Database $Database `
            -SqlUser $Instance.sqlUser `
            -SqlPassword $Instance.sqlPassword

        $cmd = $conn.CreateCommand()
        $cmd.CommandTimeout = 5
        $cmd.CommandText = "SELECT CASE WHEN OBJECT_ID(@ObjectName, N'U') IS NULL THEN 0 ELSE 1 END"
        [void]$cmd.Parameters.AddWithValue("@ObjectName", "dbo.$TableName")
        return ([int]$cmd.ExecuteScalar() -eq 1)
    }
    catch { return $false }
    finally {
        if ($conn) { try { $conn.Close() } catch {}; try { $conn.Dispose() } catch {} }
    }
}

function Resolve-BusyCloudFiscalSqlDatabaseName {
    param(
        $Instance,
        [string]$CompanyCode,
        [string]$InstanceId
    )

    $cacheKey = "${InstanceId}|${CompanyCode}".ToLowerInvariant()
    if ($script:BusyCloudApprovalSqlDbCache.ContainsKey($cacheKey)) {
        return [string]$script:BusyCloudApprovalSqlDbCache[$cacheKey]
    }

    # BUSY SQL authentication/master data can live in the base database while
    # transaction/configuration tables live in a fiscal-year database such as
    # BusyComp0003_db12026. Get-SqlDatabaseName returns the base name, so discover
    # the active fiscal DB directly from SQL Server without touching BUSY COM.
    $baseDb = Get-SqlDatabaseName -CompanyCode $CompanyCode -InstanceId $InstanceId
    if ([string]::IsNullOrWhiteSpace($baseDb)) {
        throw "Could not resolve the SQL base database for $InstanceId/$CompanyCode."
    }

    $now = Get-Date
    $fyStartYear = if ($now.Month -ge 4) { $now.Year } else { $now.Year - 1 }
    $preferred = "${baseDb}1${fyStartYear}"

    $master = $null
    try {
        $master = Open-SqlConnection `
            -SqlServer $Instance.sqlServer `
            -Database "master" `
            -SqlUser $Instance.sqlUser `
            -SqlPassword $Instance.sqlPassword

        $cmd = $master.CreateCommand()
        $cmd.CommandTimeout = 5
        $cmd.CommandText = @"
SELECT [name]
FROM sys.databases
WHERE state = 0
  AND ([name] = @BaseDb OR [name] LIKE @Prefix)
ORDER BY [name] DESC
"@
        [void]$cmd.Parameters.AddWithValue("@BaseDb", $baseDb)
        [void]$cmd.Parameters.AddWithValue("@Prefix", ($baseDb + "%"))

        $rdr = $cmd.ExecuteReader()
        $candidates = New-Object System.Collections.Generic.List[string]
        while ($rdr.Read()) {
            if (-not $rdr.IsDBNull(0)) { [void]$candidates.Add(([string]$rdr.GetString(0)).Trim()) }
        }
        $rdr.Close()

        # Prefer the current FY naming convention observed in this BUSY setup,
        # then fall back to any matching DB that actually contains dbo.Config.
        $ordered = New-Object System.Collections.Generic.List[string]
        if ($candidates -contains $preferred) { [void]$ordered.Add($preferred) }
        foreach ($name in $candidates) {
            if ($name -ne $preferred -and $name -ne $baseDb) { [void]$ordered.Add($name) }
        }
        if ($candidates -contains $baseDb) { [void]$ordered.Add($baseDb) }

        foreach ($name in $ordered) {
            if (Test-BusyCloudSqlTableExists -Instance $Instance -Database $name -TableName "Config") {
                $script:BusyCloudApprovalSqlDbCache[$cacheKey] = $name
                Write-Host "  [APPROVAL-DB] $InstanceId/$CompanyCode -> $name" -ForegroundColor DarkCyan
                return $name
            }
        }
    }
    finally {
        if ($master) { try { $master.Close() } catch {}; try { $master.Dispose() } catch {} }
    }

    throw "Could not find a fiscal BUSY SQL database containing dbo.Config for $InstanceId/$CompanyCode (base '$baseDb')."
}

function Get-BusyCloudApprovalDbContext {
    param([string]$InstanceId="", [string]$CompanyCode="")

    $found = Get-InstanceForCompany -CompanyCode $CompanyCode -InstanceId $InstanceId
    if (-not $found) { throw "Company not found in instances.json" }

    $inst = $found.instance
    $comp = $found.company
    $dbType = if ($null -ne $inst.dbType) { [int]$inst.dbType } else { 0 }

    if ($dbType -eq 1) {
        $dbName = Resolve-BusyCloudFiscalSqlDatabaseName `
            -Instance $inst `
            -CompanyCode $CompanyCode `
            -InstanceId ([string]$inst.id)

        $conn = Open-SqlConnection `
            -SqlServer $inst.sqlServer `
            -Database $dbName `
            -SqlUser $inst.sqlUser `
            -SqlPassword $inst.sqlPassword

        return @{ dbType=1; connection=$conn; instance=$inst; company=$comp; database=$dbName }
    }

    $companyFolder = Join-Path ([string]$inst.dataPath) ([string]$comp.code)
    $dbFile = Join-Path $companyFolder "db.bds"
    if (-not (Test-Path $dbFile)) { throw "Access/BDS database not found: $dbFile" }
    $conn = Open-BdsConnection -DbFile $dbFile
    return @{ dbType=0; connection=$conn; instance=$inst; company=$comp; database=$dbFile }
}

function Get-BusyCloudApprovalAuditFilePath {
    param([string]$InstanceId="", [string]$CompanyCode="")
    $root = Join-Path (Split-Path -Parent $PSScriptRoot) "data\approval_audit"
    $safeInstance = ([string]$InstanceId) -replace '[^A-Za-z0-9_.-]', '_'
    $safeCompany = ([string]$CompanyCode) -replace '[^A-Za-z0-9_.-]', '_'
    if ([string]::IsNullOrWhiteSpace($safeInstance)) { $safeInstance = "default_instance" }
    if ([string]::IsNullOrWhiteSpace($safeCompany)) { $safeCompany = "default_company" }
    return @{ root=$root; file=(Join-Path $root ("{0}__{1}.jsonl" -f $safeInstance,$safeCompany)) }
}

# Historical function name retained so existing approval code does not need to
# change. Storage is intentionally OUTSIDE the BUSY database. Creating custom
# tables inside a live BUSY SQL database caused schema-lock timeouts and slowed
# COM/OpenCSDB. The sidecar JSONL audit is database-agnostic and safe for both
# SQL Server and Access/BDS companies.
function Ensure-BusyCloudVoucherApprovalAuditTable {
    param([string]$InstanceId="", [string]$CompanyCode="")
    try {
        $p = Get-BusyCloudApprovalAuditFilePath -InstanceId $InstanceId -CompanyCode $CompanyCode
        if (-not (Test-Path $p.root)) { [void](New-Item -ItemType Directory -Path $p.root -Force) }
        if (-not (Test-Path $p.file)) {
            $fs = [System.IO.File]::Create($p.file); $fs.Close(); $fs.Dispose()
        }
        return @{ success=$true; table="BusyCloudVoucherApprovalAudit(sidecar)"; path=$p.file; dbType=-1 }
    }
    catch { return @{ success=$false; error=$_.Exception.Message } }
}

function Initialize-BusyCloudApprovalStorage {
    param([bool]$VerboseOutput=$true)
    $instancesPath = Join-Path $PSScriptRoot "..\instances.json"
    if (-not (Test-Path $instancesPath)) { return @{ success=$false; error="instances.json not found at $instancesPath"; initialized=0; failed=0 } }
    try { $cfg = Get-Content $instancesPath -Raw | ConvertFrom-Json }
    catch { return @{ success=$false; error=$_.Exception.Message; initialized=0; failed=0 } }

    $initialized=0; $failed=0; $errors=@()
    foreach ($inst in @($cfg.instances)) {
        foreach ($comp in @($inst.companies)) {
            $instanceId=[string]$inst.id; $companyCode=[string]$comp.code
            $res=Ensure-BusyCloudVoucherApprovalAuditTable -InstanceId $instanceId -CompanyCode $companyCode
            if ($res.success) {
                $initialized++
                if ($VerboseOutput) { Write-Host "  Approval audit storage ready: $instanceId/$companyCode" -ForegroundColor DarkGreen }
            } else {
                $failed++; $errors += "${instanceId}/${companyCode}: $($res.error)"
                if ($VerboseOutput) { Write-Host "  [WARN] Approval audit storage: $instanceId/$companyCode - $($res.error)" -ForegroundColor DarkYellow }
            }
        }
    }
    return @{ success=($failed-eq0); initialized=$initialized; failed=$failed; errors=@($errors) }
}

function Write-BusyCloudVoucherApprovalAudit {
    param(
        [int]$VchCode,
        [int]$VchType,
        [string]$VchNo="",
        [string]$VchSeries="",
        $VchDate=$null,
        [string]$EventType,
        [int]$PreviousStatus,
        [int]$NewStatus,
        [string]$ActionBy,
        [string]$Remarks="",
        [string]$InstanceId="",
        [string]$CompanyCode=""
    )

    $ensure=Ensure-BusyCloudVoucherApprovalAuditTable -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $ensure.success) { return $ensure }
    $event=([string]$EventType).Trim().ToUpperInvariant()
    if ($event -notin @("APPROVE","UNAPPROVE")) { return @{ success=$false; error="Unsupported approval audit event '$EventType'." } }

    try {
        $now=Get-Date
        $dateText=""
        if ($null-ne$VchDate -and "$VchDate"-ne"") { try { $dateText=([datetime]$VchDate).ToString("yyyy-MM-dd") } catch { $dateText=[string]$VchDate } }
        $record=[ordered]@{
            id=[long][DateTime]::UtcNow.Ticks
            vchCode=$VchCode; vchType=$VchType; vchNo=[string]$VchNo; vchSeries=[string]$VchSeries; date=$dateText
            eventType=$event; previousStatus=$PreviousStatus; newStatus=$NewStatus
            actionBy=[string]$ActionBy; actionTime=$now.ToString("yyyy-MM-dd HH:mm:ss"); remarks=[string]$Remarks
        }
        $line=($record | ConvertTo-Json -Compress -Depth 5) + [Environment]::NewLine
        [System.IO.File]::AppendAllText([string]$ensure.path,$line,[System.Text.Encoding]::UTF8)
        return @{ success=$true; actionTime=$record.actionTime }
    }
    catch { return @{ success=$false; error=$_.Exception.Message } }
}

function Get-BusyCloudVoucherApprovalHistory {
    param(
        [int[]]$VchTypes=@(),
        [int]$VchType=0,
        [string]$ActionBy="",
        [int]$Limit=300,
        [string]$InstanceId="",
        [string]$CompanyCode=""
    )

    $ensure=Ensure-BusyCloudVoucherApprovalAuditTable -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $ensure.success) { return $ensure }
    $effectiveTypes=@($VchTypes | ForEach-Object{[int]$_} | Where-Object{Test-IsBusyCloudApprovalVoucherType -VchType $_} | Sort-Object -Unique)
    if ($VchType -gt 0) { $effectiveTypes=@($VchType) }
    if ($effectiveTypes.Count -eq 0) { return @{ success=$true; data=@() } }
    if ($Limit -lt 1) { $Limit=1 }; if ($Limit -gt 1000) { $Limit=1000 }

    try {
        $items=@()
        foreach ($line in @(Get-Content -Path $ensure.path -ErrorAction SilentlyContinue)) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            try { $r=$line | ConvertFrom-Json } catch { continue }
            $t=0; try{$t=[int]$r.vchType}catch{}
            if ($effectiveTypes -notcontains $t) { continue }
            if ($ActionBy -and -not ([string]$r.actionBy).Equals($ActionBy,[System.StringComparison]::OrdinalIgnoreCase)) { continue }
            $items += $r
        }
        $items=@($items | Sort-Object @{Expression={ try{[datetime]$_.actionTime}catch{[datetime]::MinValue} };Descending=$true}, @{Expression={ try{[long]$_.id}catch{0} };Descending=$true} | Select-Object -First $Limit)
        return @{ success=$true; data=@($items) }
    }
    catch { return @{ success=$false; error=$_.Exception.Message } }
}

# -----------------------------------------------------------------------------
# Web voucher numbering configuration
# Uses BUSY Config table with custom RecType 202.
# One row per voucher type + voucher series.
#
# Mapping:
#   RecType = 202
#   Type    = voucher type
#   D15     = voucher series Master1.Code
#   I1      = source: 0 BUSY, 1 WEB
#   I2      = reset: 0 NEVER, 1 DAILY, 2 MONTHLY, 3 YEARLY, 4 FINANCIAL_YEAR
#   I3      = padding length
#   I4      = year format: 0 NONE, 1 YY, 2 YYYY
#   I5      = month format: 0 NONE, 1 MMM, 2 MM
#   I6      = active: 0/1
#   I7      = numbering date basis: 0 VOUCHER_DATE, 1 REAL_TIME
#   L2      = starting number
#   L3      = ending number (0 = unlimited)
#   C1      = organisation code
#   C2      = static prefix
#   C3      = static suffix
#   C4      = separator
# -----------------------------------------------------------------------------

function Resolve-VoucherSeriesCode {
    param($fi, [int]$VchType, [string]$SeriesName)

    if ([string]::IsNullOrWhiteSpace($SeriesName)) { throw "Voucher series is required." }
    $clean = $SeriesName.Trim()
    $typePrefix = "{0:D2}" -f $VchType
    $prefixed = if ($clean.StartsWith($typePrefix, [System.StringComparison]::OrdinalIgnoreCase)) { $clean } else { "$typePrefix$clean" }
    $safe1 = $clean.Replace("'", "''")
    $safe2 = $prefixed.Replace("'", "''")

    $rst = $fi.GetRecordset("SELECT Code FROM Master1 WHERE MasterType=21 AND (Name='$safe1' OR Name='$safe2')")
    try {
        if ($rst -and -not $rst.EOF) { return [int]$rst.Fields.Item("Code").Value }
    } finally {
        if ($rst) { try { $rst.Close() } catch {} }
    }
    throw "Voucher series '$clean' was not found for voucher type $VchType."
}

function Get-ConfigSafeValue {
    param($Recordset, [string]$Field, $Default)
    try {
        $value = $Recordset.Fields.Item($Field).Value
        if ($null -ne $value -and $value -ne [System.DBNull]::Value) { return $value }
    } catch {}
    return $Default
}

function Convert-ResetNameToInt {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { $Value = "NEVER" }
    switch ($Value.ToUpperInvariant()) {
        "DAILY" { return 1 }
        "MONTHLY" { return 2 }
        "YEARLY" { return 3 }
        "FINANCIAL_YEAR" { return 4 }
        default { return 0 }
    }
}

function Convert-ResetIntToName {
    param([int]$Value)
    switch ($Value) {
        1 { return "DAILY" }
        2 { return "MONTHLY" }
        3 { return "YEARLY" }
        4 { return "FINANCIAL_YEAR" }
        default { return "NEVER" }
    }
}

function Convert-YearFormatToInt {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { $Value = "YY" }
    switch ($Value.ToUpperInvariant()) { "YYYY" {2}; "NONE" {0}; default {1} }
}

function Convert-YearFormatIntToName {
    param([int]$Value)
    switch ($Value) { 2 {"YYYY"}; 0 {"NONE"}; default {"YY"} }
}

function Convert-MonthFormatToInt {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { $Value = "MMM" }
    switch ($Value.ToUpperInvariant()) { "MM" {2}; "NONE" {0}; default {1} }
}

function Convert-MonthFormatIntToName {
    param([int]$Value)
    switch ($Value) { 2 {"MM"}; 0 {"NONE"}; default {"MMM"} }
}

function Convert-DateBasisToInt {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        $Value = "VOUCHER_DATE"
    }

    switch ($Value.Trim().ToUpperInvariant()) {
        "REAL_TIME" { return 1 }
        default { return 0 }
    }
}

function Convert-DateBasisIntToName {
    param([int]$Value)

    if ($Value -eq 1) {
        return "REAL_TIME"
    }

    return "VOUCHER_DATE"
}

function Resolve-NumberingDate {
    param(
        [string]$DateBasis,
        [datetime]$VoucherDate
    )

    if (
        -not [string]::IsNullOrWhiteSpace($DateBasis) -and
        $DateBasis.Trim().ToUpperInvariant() -eq "REAL_TIME"
    ) {
        return (Get-Date).Date
    }

    return $VoucherDate.Date
}

function Build-WebVoucherNumber {
    param($Config, [datetime]$VoucherDate, [long]$Sequence)

    $parts = New-Object System.Collections.Generic.List[string]
    $separator = [string]$Config.separator
    if ($null -eq $separator) { $separator = "_" }

    $prefix = ([string]$Config.prefix).Trim()
    $suffix = ([string]$Config.suffix).Trim()
    $orgCode = (([string]$Config.organisation_code) -replace '\s+', '').ToUpperInvariant()

    $datePart = ""
    if ($Config.year_format -eq "YY") { $datePart += $VoucherDate.ToString("yy") }
    elseif ($Config.year_format -eq "YYYY") { $datePart += $VoucherDate.ToString("yyyy") }

    if ($Config.month_format -eq "MMM") {
        $datePart += $VoucherDate.ToString("MMM", [System.Globalization.CultureInfo]::InvariantCulture).ToUpperInvariant()
    } elseif ($Config.month_format -eq "MM") {
        $datePart += $VoucherDate.ToString("MM")
    }

    if ($prefix) { $parts.Add($prefix) }
    if ($datePart) { $parts.Add($datePart) }
    if ($orgCode) { $parts.Add($orgCode) }

    $padding = [Math]::Max(1, [int]$Config.padding_length)
    $parts.Add($Sequence.ToString().PadLeft($padding, '0'))
    if ($suffix) { $parts.Add($suffix) }

    $number = [string]::Join($separator, $parts)
    if ($number.Length -gt 40) { throw "Generated voucher number '$number' exceeds 40 characters." }
    if ($number -match '\s') { throw "Generated voucher number cannot contain spaces." }
    if ($number -notmatch '\d') { throw "Generated voucher number must contain a numeric sequence." }
    return $number
}

function Get-NumberingPeriod {
    param([datetime]$VoucherDate, [string]$ResetFrequency)

    switch ($ResetFrequency) {
        "DAILY" { return @{ Start=$VoucherDate.Date; End=$VoucherDate.Date.AddDays(1) } }
        "MONTHLY" {
            $s = New-Object DateTime $VoucherDate.Year, $VoucherDate.Month, 1
            return @{ Start=$s; End=$s.AddMonths(1) }
        }
        "YEARLY" {
            $s = New-Object DateTime $VoucherDate.Year, 1, 1
            return @{ Start=$s; End=$s.AddYears(1) }
        }
        "FINANCIAL_YEAR" {
            $year = if ($VoucherDate.Month -ge 4) { $VoucherDate.Year } else { $VoucherDate.Year - 1 }
            $s = New-Object DateTime $year, 4, 1
            return @{ Start=$s; End=$s.AddYears(1) }
        }
        default { return $null }
    }
}

function Get-WebNumberRegex {
    param(
        $Config,
        [datetime]$VoucherDate
    )

    $sequenceMarker = 987654321L

    $effectiveDate = Resolve-NumberingDate `
        -DateBasis ([string]$Config.date_basis) `
        -VoucherDate $VoucherDate

    $sampleNumber = Build-WebVoucherNumber `
        -Config $Config `
        -VoucherDate $effectiveDate `
        -Sequence $sequenceMarker

    $escapedSample = [regex]::Escape(
        ([string]$sampleNumber).Trim()
    )

    $escapedMarker = [regex]::Escape(
        $sequenceMarker.ToString()
    )

    $pattern = '^' + $escapedSample.Replace(
        $escapedMarker,
        '(\d+)'
    ) + '$'

    return New-Object System.Text.RegularExpressions.Regex(
        $pattern,
        (
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase `
            -bor
            [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
        )
    )
}

function Get-WebConfigBySeriesCode {
    param(
        $fi,
        [int]$VchType,
        [int]$SeriesCode,
        [string]$SeriesName = ""
    )

    $rst = $null
    try {
        $rst = $fi.GetRecordset(
            "SELECT * FROM Config WHERE RecType=202 AND [Type]=$VchType AND D15=$SeriesCode"
        )

        if (-not $rst -or $rst.EOF) { return $null }

        return @{
            vch_type          = $VchType
            series_name       = $SeriesName
            series_code       = $SeriesCode
            source            = if ([int](Get-ConfigSafeValue $rst "I1" 0) -eq 1) { "WEB" } else { "BUSY" }
            reset_frequency   = Convert-ResetIntToName ([int](Get-ConfigSafeValue $rst "I2" 3))
            padding_length    = [int](Get-ConfigSafeValue $rst "I3" 1)
            year_format       = Convert-YearFormatIntToName ([int](Get-ConfigSafeValue $rst "I4" 1))
            month_format      = Convert-MonthFormatIntToName ([int](Get-ConfigSafeValue $rst "I5" 1))
            is_active         = ([int](Get-ConfigSafeValue $rst "I6" 1) -eq 1)
            date_basis       = Convert-DateBasisIntToName ([int](Get-ConfigSafeValue $rst "I7" 0))
            starting_number   = [long](Get-ConfigSafeValue $rst "L2" 1)
            ending_number     = [long](Get-ConfigSafeValue $rst "L3" 0)
            organisation_code = ([string](Get-ConfigSafeValue $rst "C1" "")).Trim()
            prefix            = ([string](Get-ConfigSafeValue $rst "C2" "")).Trim()
            suffix            = ([string](Get-ConfigSafeValue $rst "C3" "")).Trim()
            separator         = [string](Get-ConfigSafeValue $rst "C4" "_")
        }
    }
    finally {
        if ($rst) { try { $rst.Close() } catch {} }
    }
}

function Get-ExistingWebVoucherMigrationPlan {
    param(
        $fi,
        $OldConfig,
        $NewConfig,
        [int]$VchType,
        [int]$SeriesCode
    )

    $sourceRows = @()
    $rst = $null

    try {
        $qry = @"
SELECT
    VchCode,
    VchNo,
    [Date]

FROM Tran1

WHERE
    VchType=$VchType
    AND VchSeriesCode=$SeriesCode
    AND Cancelled=0
    AND VchCancelled=0

ORDER BY
    VchCode
"@

        $rst = $fi.GetRecordset($qry)

        if ($rst -and -not $rst.EOF) {
            $rst.MoveFirst()

            while (-not $rst.EOF) {
                $vchCode = 0
                $oldVchNo = ""
                $voucherDate = Get-Date

                try {
                    $rawCode =
                        $rst.Fields.Item(
                            "VchCode"
                        ).Value

                    if (
                        $null -ne $rawCode -and
                        $rawCode -ne
                            [System.DBNull]::Value
                    ) {
                        $vchCode =
                            [int]$rawCode
                    }
                }
                catch {}

                try {
                    $rawNo =
                        $rst.Fields.Item(
                            "VchNo"
                        ).Value

                    if (
                        $null -ne $rawNo -and
                        $rawNo -ne
                            [System.DBNull]::Value
                    ) {
                        $oldVchNo =
                            ([string]$rawNo).Trim()
                    }
                }
                catch {}

                try {
                    $rawDate =
                        $rst.Fields.Item(
                            "Date"
                        ).Value

                    if (
                        $null -ne $rawDate -and
                        $rawDate -ne
                            [System.DBNull]::Value
                    ) {
                        $voucherDate =
                            [datetime]$rawDate
                    }
                }
                catch {}

                if (
                    $vchCode -gt 0 -and
                    -not [string]::IsNullOrWhiteSpace(
                        $oldVchNo
                    )
                ) {
                    $oldNumberingDate =
                        Resolve-NumberingDate `
                            -DateBasis (
                                [string]$OldConfig.date_basis
                            ) `
                            -VoucherDate $voucherDate

                    $oldRegex =
                        Get-WebNumberRegex `
                            -Config $OldConfig `
                            -VoucherDate $oldNumberingDate

                    $match =
                        $oldRegex.Match(
                            $oldVchNo
                        )

                    if ($match.Success) {
                        $sourceRows +=
                            [pscustomobject]@{
                                VchCode =
                                    $vchCode

                                OldVchNo =
                                    $oldVchNo

                                VoucherDate =
                                    $voucherDate
                            }
                    }
                }

                $rst.MoveNext()
            }
        }
    }
    finally {
        if ($rst) {
            try {
                $rst.Close()
            }
            catch {}
        }
    }

    # Use one fixed server date for the whole migration.
    # Otherwise midnight changing during migration could split the results.
    $migrationRealTimeDate =
        (Get-Date).Date

    $startingNumber =
        [Math]::Max(
            1L,
            [long]$NewConfig.starting_number
        )

    $endingNumber =
        [Math]::Max(
            0L,
            [long]$NewConfig.ending_number
        )

    $nextSequenceByPeriod = @{}
    $rows = @()

    foreach (
        $sourceRow in (
            $sourceRows |
            Sort-Object VchCode
        )
    ) {
        $effectiveDate = if (
            (
                [string]$NewConfig.date_basis
            ).Trim().ToUpperInvariant() -eq
                "REAL_TIME"
        ) {
            $migrationRealTimeDate
        }
        else {
            ([datetime]$sourceRow.VoucherDate).Date
        }

        $periodKey = switch (
            (
                [string]$NewConfig.reset_frequency
            ).Trim().ToUpperInvariant()
        ) {
            "DAILY" {
                $effectiveDate.ToString(
                    "yyyyMMdd",
                    [System.Globalization.CultureInfo]::InvariantCulture
                )
            }

            "MONTHLY" {
                $effectiveDate.ToString(
                    "yyyyMM",
                    [System.Globalization.CultureInfo]::InvariantCulture
                )
            }

            "YEARLY" {
                $effectiveDate.ToString(
                    "yyyy",
                    [System.Globalization.CultureInfo]::InvariantCulture
                )
            }

            "FINANCIAL_YEAR" {
                $financialYearStart = if (
                    $effectiveDate.Month -ge 4
                ) {
                    $effectiveDate.Year
                }
                else {
                    $effectiveDate.Year - 1
                }

                "FY$financialYearStart"
            }

            default {
                "NEVER"
            }
        }

        if (
            -not $nextSequenceByPeriod.ContainsKey(
                $periodKey
            )
        ) {
            $nextSequenceByPeriod[$periodKey] =
                $startingNumber
        }

        $sequence =
            [long]$nextSequenceByPeriod[
                $periodKey
            ]

        if (
            $endingNumber -gt 0 -and
            $sequence -gt $endingNumber
        ) {
            throw (
                "Voucher sequence has reached the configured " +
                "ending number $endingNumber for period " +
                "'$periodKey'. No changes were applied."
            )
        }

        $newVchNo =
            Build-WebVoucherNumber `
                -Config $NewConfig `
                -VoucherDate $effectiveDate `
                -Sequence $sequence

        $rows +=
            [pscustomobject]@{
                VchCode =
                    [int]$sourceRow.VchCode

                OldVchNo =
                    [string]$sourceRow.OldVchNo

                NewVchNo =
                    $newVchNo

                Sequence =
                    $sequence

                VoucherDate =
                    [datetime]$sourceRow.VoucherDate

                NumberingDate =
                    $effectiveDate

                PeriodKey =
                    $periodKey
            }

        $nextSequenceByPeriod[$periodKey] =
            $sequence + 1
    }

    $duplicates = @(
        $rows |
        Group-Object NewVchNo |
        Where-Object {
            $_.Count -gt 1
        }
    )

    if ($duplicates.Count -gt 0) {
        $duplicateNumbers =
            $duplicates |
            ForEach-Object {
                $_.Name
            }

        throw (
            "The new configuration would generate duplicate " +
            "voucher numbers: " +
            (
                $duplicateNumbers -join ", "
            ) +
            ". No changes were applied."
        )
    }

    foreach ($row in $rows) {
        $safeNew =
            ([string]$row.NewVchNo).Replace(
                "'",
                "''"
            )

        $dupRst = $null

        try {
            $dupRst =
                $fi.GetRecordset(@"
SELECT VchCode

FROM Tran1

WHERE
    VchType=$VchType
    AND VchSeriesCode=$SeriesCode
    AND LTRIM(RTRIM(VchNo))='$safeNew'
    AND VchCode<>$($row.VchCode)
"@)

            if (
                $dupRst -and
                -not $dupRst.EOF
            ) {
                $conflictingCode =
                    [int]$dupRst.Fields.Item(
                        "VchCode"
                    ).Value

                $includedInMigration =
                    @(
                        $rows |
                        Where-Object {
                            [int]$_.VchCode -eq
                                $conflictingCode
                        }
                    ).Count -gt 0

                if (-not $includedInMigration) {
                    throw (
                        "Voucher number '$($row.NewVchNo)' " +
                        "already exists outside the migration set. " +
                        "No changes were applied."
                    )
                }
            }
        }
        finally {
            if ($dupRst) {
                try {
                    $dupRst.Close()
                }
                catch {}
            }
        }
    }

    return ,$rows
}

function Get-DbVoucherNumberValue {
    param(
        $fi,
        [string]$TableName,
        [string]$ColumnName,
        [int]$VchCode,
        [int]$VchType
    )

    $rst = $null

    try {
        $query = @"
SELECT $ColumnName
FROM $TableName
WHERE VchCode=$VchCode
  AND VchType=$VchType
"@

        $rst = $fi.GetRecordset($query)

        if (-not $rst -or $rst.EOF) {
            return $null
        }

        $value = $rst.Fields.Item($ColumnName).Value

        if (
            $null -eq $value -or
            $value -eq [System.DBNull]::Value
        ) {
            return ""
        }

        return [string]$value
    }
    finally {
        if ($rst) {
            try { $rst.Close() } catch {}
        }
    }
}

function Get-SqlBusyPaddedVoucherNo {
    param(
        [string]$ExistingValue,
        [string]$NewVoucherNo
    )

    if ($null -eq $ExistingValue) {
        return $NewVoucherNo
    }

    # SQL BUSY may return fixed-width CHAR values with leading spaces.
    # Keep the total stored width, not the old number of spaces.
    #
    # Example:
    # Old field : "              26JUL_QQQ_1"
    # New value : "26JUL_QQQP_1"
    # Result    : one fewer leading space, same total field width.
    $fieldWidth = ([string]$ExistingValue).Length
    $cleanNew = ([string]$NewVoucherNo).Trim()

    if ($fieldWidth -le 0) {
        return $cleanNew
    }

    if ($cleanNew.Length -gt $fieldWidth) {
        throw (
            "The new voucher number '$cleanNew' has $($cleanNew.Length) " +
            "characters, but the BUSY SQL field currently allows only " +
            "$fieldWidth characters. Shorten the numbering format."
        )
    }

    return $cleanNew.PadLeft($fieldWidth, [char]' ')
}

function Set-BusyVoucherNumberInTable {
    param(
        $fi,
        [string]$TableName,
        [string]$ColumnName,
        [int]$VchCode,
        [int]$VchType,
        [string]$NewVoucherNo,
        [bool]$IsSql
    )

    $currentValue = Get-DbVoucherNumberValue `
        -fi $fi `
        -TableName $TableName `
        -ColumnName $ColumnName `
        -VchCode $VchCode `
        -VchType $VchType

    # Some voucher types may not create a row in every table.
    if ($null -eq $currentValue) {
        Write-Host (
            "  [WEB-NUMBER MIGRATION] " +
            "$TableName.$ColumnName has no row for " +
            "VchCode=$VchCode, VchType=$VchType. Skipped."
        ) -ForegroundColor DarkYellow

        return $false
    }

    $storedValue = if ($IsSql) {
        Get-SqlBusyPaddedVoucherNo `
            -ExistingValue $currentValue `
            -NewVoucherNo $NewVoucherNo
    }
    else {
        $NewVoucherNo
    }

    $safeStoredValue = $storedValue.Replace("'", "''")

    $updateQuery = @"
UPDATE $TableName
SET $ColumnName='$safeStoredValue'
WHERE VchCode=$VchCode
  AND VchType=$VchType
"@

    $currentValueForLog = ([string]$currentValue) -replace "[\r\n]", ""
    $storedValueForLog = ([string]$storedValue) -replace "[\r\n]", ""

    Write-Host (
        "  [WEB-NUMBER MIGRATION] Updating " +
        "$TableName.$ColumnName | VchCode=$VchCode | " +
        "Old='$currentValueForLog' | " +
        "New='$storedValueForLog'"
    ) -ForegroundColor Cyan

    $fi.ExecuteQuery($updateQuery)

    $savedValue = Get-DbVoucherNumberValue `
        -fi $fi `
        -TableName $TableName `
        -ColumnName $ColumnName `
        -VchCode $VchCode `
        -VchType $VchType

    if ($null -eq $savedValue) {
        throw (
            "Could not verify $TableName.$ColumnName after updating " +
            "voucher code $VchCode."
        )
    }

    if (
        -not $savedValue.Trim().Equals(
            $NewVoucherNo.Trim(),
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        throw (
            "$TableName.$ColumnName was not updated correctly for " +
            "voucher code $VchCode. Expected '$NewVoucherNo', " +
            "database returned '$savedValue'."
        )
    }

    return $true
}

function Update-ExistingWebVoucherNumbers {
    param(
        $fi,
        $MigrationPlan,
        [int]$VchType,
        [int]$SeriesCode,
        [bool]$IsSql = $false
    )

    $updatedVoucherCount = 0

    foreach ($row in @($MigrationPlan)) {
        $vchCode = [int]$row.VchCode
        $newVchNo = [string]$row.NewVchNo
        $oldVchNo = [string]$row.OldVchNo

        Write-Host (
            "[WEB-NUMBER MIGRATION] Voucher " +
            "VchCode=$vchCode | '$oldVchNo' -> '$newVchNo'"
        ) -ForegroundColor Green

        # BUSY stores the visible voucher number in multiple transaction tables.
        #
        # Tran1 : Voucher header
        # Tran2 : Voucher transaction/detail rows
        # Tran3 : Bill/reference rows; the voucher number field is named [No]
        #
        # AutoVchNo is intentionally NOT updated for WEB-numbered vouchers.
        $tran1Updated = Set-BusyVoucherNumberInTable `
            -fi $fi `
            -TableName "Tran1" `
            -ColumnName "VchNo" `
            -VchCode $vchCode `
            -VchType $VchType `
            -NewVoucherNo $newVchNo `
            -IsSql $IsSql

        $tran2Updated = Set-BusyVoucherNumberInTable `
            -fi $fi `
            -TableName "Tran2" `
            -ColumnName "VchNo" `
            -VchCode $vchCode `
            -VchType $VchType `
            -NewVoucherNo $newVchNo `
            -IsSql $IsSql

        $tran3Updated = Set-BusyVoucherNumberInTable `
            -fi $fi `
            -TableName "Tran3" `
            -ColumnName "[No]" `
            -VchCode $vchCode `
            -VchType $VchType `
            -NewVoucherNo $newVchNo `
            -IsSql $IsSql

        if (-not $tran1Updated) {
            throw (
                "Tran1 voucher header was not found for VchCode=$vchCode. " +
                "Migration was stopped."
            )
        }

        Write-Host (
            "[WEB-NUMBER MIGRATION] Verified VchCode=$vchCode | " +
            "Tran1=$tran1Updated, Tran2=$tran2Updated, Tran3=$tran3Updated"
        ) -ForegroundColor Green

        $updatedVoucherCount++
    }

    return $updatedVoucherCount
}

function Get-WebNumberingConfig {
    param(
        [int]$VchType,
        [string]$SeriesName,
        [string]$VoucherDate = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )
    Write-Host (
        "  [WEB-NUMBER-FAST START] {0}/{1} type={2} series='{3}'" -f
        $InstanceId,
        $CompanyCode,
        $VchType,
        $SeriesName
    ) -ForegroundColor DarkCyan


    # -------------------------------------------------------------------------
    # FAST READ-ONLY WEB NUMBERING CONFIG
    #
    # IMPORTANT:
    # The previous implementation called Connect-BUSY before checking RecType
    # 202. That meant EVERY /busy/numbering-config request could spend tens of
    # seconds inside BUSY COM/OpenCSDB even when the voucher actually used
    # normal BUSY numbering.
    #
    # This implementation reads Master1 / Config / Tran1 directly from the
    # active fiscal database and never initializes BUSY COM.
    # -------------------------------------------------------------------------

    $startedAt = Get-Date
    $ctx = $null
    $reader = $null

    try {
        if ($VchType -le 0) {
            return @{
                success = $false
                error = "Voucher type is required."
            }
        }

        if ([string]::IsNullOrWhiteSpace($SeriesName)) {
            return @{
                success = $false
                error = "Voucher series is required."
            }
        }

        $resolver = Get-Command `
            Get-BusyCloudFastConfigDbContext `
            -ErrorAction SilentlyContinue

        if ($null -eq $resolver) {
            throw (
                "Fast fiscal database resolver is unavailable. " +
                "Install the optimized vch_setting.ps1."
            )
        }

        $ctx = Get-BusyCloudFastConfigDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (
            $null -eq $ctx -or
            $null -eq $ctx.connection
        ) {
            throw "Direct fiscal database connection is unavailable."
        }

        $conn = $ctx.connection
        $isSql = ([int]$ctx.dbType -eq 1)

        # ---------------------------------------------------------------------
        # Resolve voucher series directly from Master1.
        # ---------------------------------------------------------------------

        $cleanSeries = $SeriesName.Trim()
        $typePrefix = "{0:D2}" -f $VchType

        $prefixedSeries = if (
            $cleanSeries.StartsWith(
                $typePrefix,
                [System.StringComparison]::OrdinalIgnoreCase
            )
        ) {
            $cleanSeries
        }
        else {
            "$typePrefix$cleanSeries"
        }

        $safeSeries = $cleanSeries.Replace("'", "''")
        $safePrefixedSeries = $prefixedSeries.Replace("'", "''")

        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}

        $cmd.CommandText = @"
SELECT TOP 1 Code
FROM Master1
WHERE MasterType = 21
  AND (
        Name = '$safeSeries'
        OR Name = '$safePrefixedSeries'
      )
"@

        $seriesValue = $cmd.ExecuteScalar()

        if (
            $null -eq $seriesValue -or
            $seriesValue -eq [System.DBNull]::Value
        ) {
            return @{
                success = $false
                error = "Voucher series '$cleanSeries' was not found for voucher type $VchType."
            }
        }

        $seriesCode = [int]$seriesValue

        # ---------------------------------------------------------------------
        # Defaults mean normal BUSY numbering.
        # ---------------------------------------------------------------------

        $cfg = @{
            vch_type          = $VchType
            series_name       = $cleanSeries
            series_code       = $seriesCode
            source            = "BUSY"
            reset_frequency   = "YEARLY"
            padding_length    = 1
            year_format       = "YY"
            month_format      = "MMM"
            is_active         = $true
            date_basis        = "VOUCHER_DATE"
            starting_number   = 1L
            ending_number     = 0L
            organisation_code = ""
            prefix            = ""
            suffix            = ""
            separator         = "_"
        }

        # ---------------------------------------------------------------------
        # Read optional BusyCloud WEB numbering override from Config/RecType 202.
        # ---------------------------------------------------------------------

        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}

        $cmd.CommandText = @"
SELECT TOP 1
    I1,
    I2,
    I3,
    I4,
    I5,
    I6,
    I7,
    L2,
    L3,
    C1,
    C2,
    C3,
    C4
FROM Config
WHERE RecType = 202
  AND [Type] = $VchType
  AND D15 = $seriesCode
"@

        $reader = $cmd.ExecuteReader()

        try {
            if ($reader.Read()) {
                $getValue = {
                    param(
                        [string]$Field,
                        $Default
                    )

                    try {
                        $ordinal = $reader.GetOrdinal($Field)

                        if (
                            $ordinal -ge 0 -and
                            -not $reader.IsDBNull($ordinal)
                        ) {
                            return $reader.GetValue($ordinal)
                        }
                    }
                    catch {
                    }

                    return $Default
                }

                $cfg.source = if (
                    [int](& $getValue "I1" 0) -eq 1
                ) {
                    "WEB"
                }
                else {
                    "BUSY"
                }

                $cfg.reset_frequency =
                    Convert-ResetIntToName `
                        ([int](& $getValue "I2" 3))

                $cfg.padding_length =
                    [int](& $getValue "I3" 1)

                $cfg.year_format =
                    Convert-YearFormatIntToName `
                        ([int](& $getValue "I4" 1))

                $cfg.month_format =
                    Convert-MonthFormatIntToName `
                        ([int](& $getValue "I5" 1))

                $cfg.is_active =
                    ([int](& $getValue "I6" 1) -eq 1)

                $cfg.date_basis =
                    Convert-DateBasisIntToName `
                        ([int](& $getValue "I7" 0))

                $cfg.starting_number =
                    [long](& $getValue "L2" 1)

                $cfg.ending_number =
                    [long](& $getValue "L3" 0)

                $cfg.organisation_code =
                    ([string](& $getValue "C1" "")).Trim()

                $cfg.prefix =
                    ([string](& $getValue "C2" "")).Trim()

                $cfg.suffix =
                    ([string](& $getValue "C3" "")).Trim()

                $cfg.separator =
                    [string](& $getValue "C4" "_")
            }
        }
        finally {
            if ($reader) {
                try { $reader.Close() } catch {}
                try { $reader.Dispose() } catch {}
                $reader = $null
            }
        }

        # ---------------------------------------------------------------------
        # Resolve numbering date.
        # ---------------------------------------------------------------------

        $date = if (
            -not [string]::IsNullOrWhiteSpace($VoucherDate)
        ) {
            [datetime]::Parse(
                $VoucherDate,
                [System.Globalization.CultureInfo]::InvariantCulture
            )
        }
        else {
            Get-Date
        }

        $effectiveDate = Resolve-NumberingDate `
            -DateBasis ([string]$cfg.date_basis) `
            -VoucherDate $date

        # ---------------------------------------------------------------------
        # WEB numbering preview.
        #
        # Only WEB+active needs a Tran1 scan. BUSY numbering skips this entirely
        # and immediately returns to Get-EffectiveNumberingConfig, which then
        # calls the optimized Get-NumberingConfig.
        # ---------------------------------------------------------------------

        if (
            $cfg.source -eq "WEB" -and
            $cfg.is_active
        ) {
            $dateFilter = ""

            if (
                ([string]$cfg.date_basis).Trim().ToUpperInvariant() -eq
                "VOUCHER_DATE"
            ) {
                $period = Get-NumberingPeriod `
                    -VoucherDate $effectiveDate `
                    -ResetFrequency ([string]$cfg.reset_frequency)

                if ($null -ne $period) {
                    if ($isSql) {
                        $startText = $period.Start.ToString(
                            "yyyy-MM-ddTHH:mm:ss",
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )

                        $endText = $period.End.ToString(
                            "yyyy-MM-ddTHH:mm:ss",
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )

                        $dateFilter = (
                            " AND [Date] >= '$startText'" +
                            " AND [Date] < '$endText'"
                        )
                    }
                    else {
                        $startText = $period.Start.ToString(
                            "MM/dd/yyyy",
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )

                        $endText = $period.End.ToString(
                            "MM/dd/yyyy",
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )

                        $dateFilter = (
                            " AND [Date] >= #$startText#" +
                            " AND [Date] < #$endText#"
                        )
                    }
                }
            }

            $webNumberRegex = Get-WebNumberRegex `
                -Config $cfg `
                -VoucherDate $effectiveDate

            $cmd = $conn.CreateCommand()
            try { $cmd.CommandTimeout = 5 } catch {}

            $cmd.CommandText = (
                "SELECT VchNo FROM Tran1 " +
                "WHERE VchType=$VchType " +
                "AND VchSeriesCode=$seriesCode " +
                "AND Cancelled=0 " +
                "AND VchCancelled=0" +
                $dateFilter
            )

            $reader = $cmd.ExecuteReader()

            $highest = 0L

            try {
                while ($reader.Read()) {
                    $raw = ""

                    try {
                        if (-not $reader.IsDBNull(0)) {
                            $raw = (
                                [string]$reader.GetValue(0)
                            ).Trim()
                        }
                    }
                    catch {
                        $raw = ""
                    }

                    if ([string]::IsNullOrWhiteSpace($raw)) {
                        continue
                    }

                    $match = $webNumberRegex.Match($raw)

                    if ($match.Success) {
                        $sequence = 0L

                        if (
                            [long]::TryParse(
                                $match.Groups[1].Value,
                                [ref]$sequence
                            ) -and
                            $sequence -gt $highest
                        ) {
                            $highest = $sequence
                        }
                    }
                }
            }
            finally {
                if ($reader) {
                    try { $reader.Close() } catch {}
                    try { $reader.Dispose() } catch {}
                    $reader = $null
                }
            }

            $startingNumber =
                [Math]::Max(
                    1L,
                    [long]$cfg.starting_number
                )

            $nextSequence = if ($highest -gt 0) {
                $highest + 1
            }
            else {
                $startingNumber
            }

            $endingNumber =
                [Math]::Max(
                    0L,
                    [long]$cfg.ending_number
                )

            if (
                $endingNumber -gt 0 -and
                $nextSequence -gt $endingNumber
            ) {
                throw (
                    "Voucher sequence has reached the configured " +
                    "ending number $endingNumber."
                )
            }

            $cfg.last_sequence = $highest
            $cfg.next_sequence = $nextSequence
            $cfg.next_vch_no = Build-WebVoucherNumber `
                -Config $cfg `
                -VoucherDate $effectiveDate `
                -Sequence $nextSequence
        }
        else {
            $cfg.last_sequence = 0L
            $cfg.next_sequence = 0L
            $cfg.next_vch_no = ""
        }

        $elapsedMs = [int](
            ((Get-Date) - $startedAt).TotalMilliseconds
        )

        Write-Host (
            "  [WEB-NUMBER-FAST] {0}/{1} type={2} series='{3}' source={4} db={5} elapsedMs={6}" -f
            $InstanceId,
            $CompanyCode,
            $VchType,
            $cleanSeries,
            [string]$cfg.source,
            [string]$ctx.database,
            $elapsedMs
        ) -ForegroundColor DarkCyan

        return @{
            success = $true
            data = $cfg
        }
    }
    catch {
        $elapsedMs = [int](
            ((Get-Date) - $startedAt).TotalMilliseconds
        )

        Write-Host (
            "  [WEB-NUMBER-FAST FAIL] {0}/{1} type={2} series='{3}' elapsedMs={4} error={5}" -f
            $InstanceId,
            $CompanyCode,
            $VchType,
            $SeriesName,
            $elapsedMs,
            $_.Exception.Message
        ) -ForegroundColor Red

        # Do not fall back to Connect-BUSY here. This is a read-only page load,
        # and a slow COM fallback would block the entire API process again.
        return @{
            success = $false
            error = $_.Exception.Message
        }
    }
    finally {
        if ($reader) {
            try { $reader.Close() } catch {}
            try { $reader.Dispose() } catch {}
        }

        if (
            $ctx -and
            $ctx.connection
        ) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}

function Save-WebNumberingConfig {
    param(
        $Data,
        [string]$UpdatedBy = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) {
        return @{ success = $false; error = "BUSY database connection failed" }
    }

    try {
        $vchType = [int]$Data.vch_type
        $seriesName = ([string]$Data.series_name).Trim()

        if ($vchType -le 0) { throw "Voucher type is required." }
        if ([string]::IsNullOrWhiteSpace($seriesName)) { throw "Voucher series is required." }

        $seriesCode = Resolve-VoucherSeriesCode -fi $fi -VchType $vchType -SeriesName $seriesName
        $oldConfig = Get-WebConfigBySeriesCode -fi $fi -VchType $vchType -SeriesCode $seriesCode -SeriesName $seriesName

        $source = ([string]$Data.source).Trim().ToUpperInvariant()
        if ($source -notin @("BUSY", "WEB")) { throw "Source must be BUSY or WEB." }

        $org = (([string]$Data.organisation_code) -replace '\s+', '').ToUpperInvariant()
        if ($source -eq "WEB" -and [string]::IsNullOrWhiteSpace($org)) {
            throw "Organisation code is required for WEB numbering."
        }
        if ($org -and $org -notmatch '^[A-Z0-9_-]+$') {
            throw "Organisation code may contain only letters, numbers, underscore and hyphen."
        }

        $prefixRaw = (([string]$Data.prefix) -replace '\s+', '')
        $suffixRaw = (([string]$Data.suffix) -replace '\s+', '')
        $separatorRaw = [string]$Data.separator
        if ($null -eq $separatorRaw) { $separatorRaw = "_" }
        if ($separatorRaw.Length -gt 1) { throw "Separator must be one character or empty." }

        $sourceInt = if ($source -eq "WEB") { 1 } else { 0 }
        $resetInt = Convert-ResetNameToInt ([string]$Data.reset_frequency)
        $yearInt = Convert-YearFormatToInt ([string]$Data.year_format)
        $monthInt = Convert-MonthFormatToInt ([string]$Data.month_format)
        $dateBasisInt = Convert-DateBasisToInt ([string]$Data.date_basis)
        $dateBasis = Convert-DateBasisIntToName $dateBasisInt
        $padding = [Math]::Min(15, [Math]::Max(1, [int]$Data.padding_length))

        $active = 1
        $isActiveProperty = $Data.PSObject.Properties["is_active"]
        if ($null -ne $isActiveProperty) {
            $rawActive = $isActiveProperty.Value
            if ($rawActive -is [bool]) {
                $active = if ([bool]$rawActive) { 1 } else { 0 }
            }
            else {
                $activeText = ([string]$rawActive).Trim().ToLowerInvariant()
                $active = if ($activeText -in @("true", "1", "yes", "on")) { 1 } else { 0 }
            }
        }

        $updateExisting = $false
        $updateProperty = $Data.PSObject.Properties["update_existing_web_vouchers"]
        if ($null -ne $updateProperty) {
            $rawUpdate = $updateProperty.Value
            if ($rawUpdate -is [bool]) {
                $updateExisting = [bool]$rawUpdate
            }
            else {
                $updateExisting = (([string]$rawUpdate).Trim().ToLowerInvariant() -in @("true", "1", "yes", "on"))
            }
        }

        $startNo = [Math]::Max(1L, [long]$Data.starting_number)
        $endNo = [Math]::Max(0L, [long]$Data.ending_number)
        if ($endNo -gt 0 -and $endNo -lt $startNo) {
            throw "Ending number cannot be lower than starting number."
        }

        # REAL_TIME numbering cannot safely use voucher transaction dates to
        # separate historical periods. Require a visible date token so each
        # real-time period can be identified from VchNo itself.
        if ($dateBasis -eq "REAL_TIME") {
            $resetName = Convert-ResetIntToName $resetInt
            $yearName = Convert-YearFormatIntToName $yearInt
            $monthName = Convert-MonthFormatIntToName $monthInt

            if ($resetName -eq "DAILY") {
                throw "Real-Time Date does not support Daily reset because the format has no day token."
            }

            if ($resetName -eq "MONTHLY" -and (
                $yearName -eq "NONE" -or
                $monthName -eq "NONE"
            )) {
                throw "Real-Time Date with Monthly reset requires both year and month in the voucher number."
            }

            if ($resetName -in @("YEARLY", "FINANCIAL_YEAR") -and $yearName -eq "NONE") {
                throw "Real-Time Date with yearly reset requires the year in the voucher number."
            }

            if ($resetName -eq "FINANCIAL_YEAR") {
                throw "Real-Time Date does not support Financial Year reset safely. Use Voucher Form Date for financial-year numbering."
            }
        }

        $newConfig = @{
            vch_type          = $vchType
            series_name       = $seriesName
            series_code       = $seriesCode
            source            = $source
            reset_frequency   = Convert-ResetIntToName $resetInt
            padding_length    = $padding
            year_format       = Convert-YearFormatIntToName $yearInt
            month_format      = Convert-MonthFormatIntToName $monthInt
            is_active         = ($active -eq 1)
            date_basis       = $dateBasis
            starting_number   = $startNo
            ending_number     = $endNo
            organisation_code = $org
            prefix            = $prefixRaw
            suffix            = $suffixRaw
            separator         = $separatorRaw
        }

        $validationDate = Resolve-NumberingDate `
            -DateBasis $dateBasis `
            -VoucherDate (Get-Date)

        [void](Build-WebVoucherNumber `
            -Config $newConfig `
            -VoucherDate $validationDate `
            -Sequence $startNo)

        # Build and validate the complete migration plan BEFORE changing anything.
        $migrationPlan = @()
        if (
            $updateExisting -and
            $null -ne $oldConfig -and
            $oldConfig.source -eq "WEB" -and
            $source -eq "WEB"
        ) {
            $migrationPlan = Get-ExistingWebVoucherMigrationPlan `
                -fi $fi `
                -OldConfig $oldConfig `
                -NewConfig $newConfig `
                -VchType $vchType `
                -SeriesCode $seriesCode
        }

        $safeOrg = $org.Replace("'", "''")
        $safePrefix = $prefixRaw.Replace("'", "''")
        $safeSuffix = $suffixRaw.Replace("'", "''")
        $safeSeparator = $separatorRaw.Replace("'", "''")

        $exists = ($null -ne $oldConfig)
        if ($exists) {
            $sql = "UPDATE Config SET I1=$sourceInt,I2=$resetInt,I3=$padding,I4=$yearInt,I5=$monthInt,I6=$active,I7=$dateBasisInt,L2=$startNo,L3=$endNo,C1='$safeOrg',C2='$safePrefix',C3='$safeSuffix',C4='$safeSeparator' WHERE RecType=202 AND [Type]=$vchType AND D15=$seriesCode"
        }
        else {
            $sql = "INSERT INTO Config (RecType,[Type],D15,I1,I2,I3,I4,I5,I6,I7,L2,L3,C1,C2,C3,C4) VALUES (202,$vchType,$seriesCode,$sourceInt,$resetInt,$padding,$yearInt,$monthInt,$active,$dateBasisInt,$startNo,$endNo,'$safeOrg','$safePrefix','$safeSuffix','$safeSeparator')"
        }
        $fi.ExecuteQuery($sql)

        $updatedCount = 0
        if ($migrationPlan.Count -gt 0) {
            $targetInstance = Get-InstanceConfig `
                -InstanceId $InstanceId

            $dbType = 0

            if (
                $null -ne $targetInstance -and
                $null -ne $targetInstance.dbType
            ) {
                $dbType = [int]$targetInstance.dbType
            }

            $isSql = ($dbType -eq 1)

            Write-Host (
                "[WEB-NUMBER MIGRATION] Database=" +
                $(if ($isSql) { "SQL Server" } else { "Access" }) +
                " | VoucherType=$vchType | SeriesCode=$seriesCode | " +
                "Count=$($migrationPlan.Count)"
            ) -ForegroundColor Magenta

            $updatedCount = Update-ExistingWebVoucherNumbers `
                -fi $fi `
                -MigrationPlan $migrationPlan `
                -VchType $vchType `
                -SeriesCode $seriesCode `
                -IsSql $isSql
        }

        $verifyRst = $null
        $savedActive = -1
        try {
            $verifyRst = $fi.GetRecordset("SELECT I6 FROM Config WHERE RecType=202 AND [Type]=$vchType AND D15=$seriesCode")
            if ($verifyRst -and -not $verifyRst.EOF) {
                $savedValue = $verifyRst.Fields.Item("I6").Value
                if ($null -ne $savedValue -and $savedValue -ne [System.DBNull]::Value) {
                    $savedActive = [int]$savedValue
                }
            }
        }
        finally {
            if ($verifyRst) { try { $verifyRst.Close() } catch {} }
        }

        if ($savedActive -ne $active) {
            throw "The Active setting was not saved correctly. Expected I6=$active but database returned I6=$savedActive."
        }

        $result = Get-WebNumberingConfig `
            -VchType $vchType `
            -SeriesName $seriesName `
            -VoucherDate ([string]$Data.voucher_date) `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if ($result.success) {
            $result.data.updated_existing_count = $updatedCount
        }
        return $result
    }
    catch {
        return @{ success = $false; error = $_.Exception.Message }
    }
    finally {
        if ($fi) { Disconnect-BUSY $fi }
    }
}

function Get-NextWebVoucherNumberInternal {
    param(
        $fi,
        $Config,
        [datetime]$VoucherDate,
        [bool]$IsSql = $false
    )

    if ($null -eq $fi) {
        throw "BUSY connection is unavailable."
    }

    # VoucherDate passed to this function is already the effective numbering
    # date resolved from either VOUCHER_DATE or REAL_TIME.
    $effectiveDate = $VoucherDate.Date
    $dateBasis = ([string]$Config.date_basis).Trim().ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($dateBasis)) {
        $dateBasis = "VOUCHER_DATE"
    }

    # Voucher-date mode can safely filter Tran1.[Date], because that column is
    # the voucher form date. Real-time mode must not filter Tran1.[Date]:
    # backdated vouchers may have a different form date. In real-time mode the
    # exact year/month token in VchNo identifies the active period.
    $dateFilter = ""

    if ($dateBasis -eq "VOUCHER_DATE") {
        $period = Get-NumberingPeriod `
            -VoucherDate $effectiveDate `
            -ResetFrequency ([string]$Config.reset_frequency)

        if ($null -ne $period) {
            if ($IsSql) {
                $startText = $period.Start.ToString(
                    "yyyy-MM-ddTHH:mm:ss",
                    [System.Globalization.CultureInfo]::InvariantCulture
                )
                $endText = $period.End.ToString(
                    "yyyy-MM-ddTHH:mm:ss",
                    [System.Globalization.CultureInfo]::InvariantCulture
                )
                $dateFilter = " AND [Date] >= '$startText' AND [Date] < '$endText'"
            }
            else {
                $startText = $period.Start.ToString(
                    "MM/dd/yyyy",
                    [System.Globalization.CultureInfo]::InvariantCulture
                )
                $endText = $period.End.ToString(
                    "MM/dd/yyyy",
                    [System.Globalization.CultureInfo]::InvariantCulture
                )
                $dateFilter = " AND [Date] >= #$startText# AND [Date] < #$endText#"
            }
        }
    }

    # Build an exact regex for the active date token and WEB format.
    # This ignores old BUSY-generated numbers and numbers from other periods.
    $webNumberRegex = Get-WebNumberRegex `
        -Config $Config `
        -VoucherDate $effectiveDate

    $qry = (
        "SELECT VchNo FROM Tran1 " +
        "WHERE VchType=$([int]$Config.vch_type) " +
        "AND VchSeriesCode=$([int]$Config.series_code) " +
        "AND Cancelled=0 AND VchCancelled=0" +
        $dateFilter
    )

    $rst = $null
    $highest = 0L

    try {
        $rst = $fi.GetRecordset($qry)

        if ($rst -and -not $rst.EOF) {
            $rst.MoveFirst()

            while (-not $rst.EOF) {
                $raw = ([string](Get-ConfigSafeValue $rst "VchNo" "")).Trim()

                if (-not [string]::IsNullOrWhiteSpace($raw)) {
                    $match = $webNumberRegex.Match($raw)

                    if ($match.Success) {
                        $seq = 0L

                        if (
                            [long]::TryParse(
                                $match.Groups[1].Value,
                                [ref]$seq
                            ) -and
                            $seq -gt $highest
                        ) {
                            $highest = $seq
                        }
                    }
                }

                $rst.MoveNext()
            }
        }
    }
    finally {
        if ($rst) {
            try { $rst.Close() } catch {}

            try {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $rst
                ) | Out-Null
            }
            catch {}
        }
    }

    $startingNumber =
        [Math]::Max(1L, [long]$Config.starting_number)

    $next = if ($highest -gt 0) {
        $highest + 1
    }
    else {
        $startingNumber
    }

    $endingNumber =
        [Math]::Max(0L, [long]$Config.ending_number)

    if (
        $endingNumber -gt 0 -and
        $next -gt $endingNumber
    ) {
        throw (
            "Voucher sequence has reached the configured ending " +
            "number $endingNumber."
        )
    }

    return @{
        last_sequence = $highest
        next_sequence = $next
        next_vch_no   = Build-WebVoucherNumber `
            -Config $Config `
            -VoucherDate $effectiveDate `
            -Sequence $next
    }
}

function Get-EffectiveVoucherNumber {
    param(
        $fi,
        [int]$VchType,
        [string]$SeriesName,
        [datetime]$VoucherDate,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $cfgResult = Get-WebNumberingConfig -VchType $VchType -SeriesName $SeriesName -VoucherDate $VoucherDate.ToString("yyyy-MM-dd") -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $cfgResult.success) { throw $cfgResult.error }
    $cfg = $cfgResult.data

    if ($cfg.source -eq "WEB" -and $cfg.is_active) {
        return @{ source="WEB"; vch_no=$cfg.next_vch_no; sequence=$cfg.next_sequence }
    }

    return @{ source="BUSY"; vch_no=""; sequence=0 }
}

# Returns the same response shape already consumed by useNumberingConfig.
# WEB configuration overrides BUSY only when source=WEB and active=true.
function Get-EffectiveNumberingConfig {
    param(
        [int]$VchType,
        [string]$SeriesName,
        [string]$VoucherDate = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $webResult = Get-WebNumberingConfig -VchType $VchType -SeriesName $SeriesName -VoucherDate $VoucherDate -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $webResult.success) { return $webResult }

    $web = $webResult.data
    if ($web.source -eq "WEB" -and $web.is_active) {
        return @{
            success = $true
            data = @{
                vch_type          = $VchType
                series_name       = $SeriesName
                source            = "WEB"
                date_basis        = $web.date_basis
                prefix            = $web.prefix
                suffix            = $web.suffix
                separator         = $web.separator
                organisation_code = $web.organisation_code
                year_format       = $web.year_format
                month_format      = $web.month_format
                reset_frequency   = $web.reset_frequency
                starting_number   = $web.starting_number
                padding_length    = $web.padding_length
                current_no        = $web.next_sequence
                is_auto           = $true
                frequency         = switch ($web.reset_frequency) { "DAILY" {1}; "MONTHLY" {2}; "YEARLY" {3}; "FINANCIAL_YEAR" {3}; default {0} }
                ending_no         = $web.ending_number
                next_vch_no       = $web.next_vch_no
                last_vch_no    = if ($web.last_sequence -gt 0) {
                    $rawVoucherDate = if ($VoucherDate) {
                        [datetime]::Parse(
                            $VoucherDate,
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )
                    }
                    else {
                        Get-Date
                    }

                    $effectiveLastDate = Resolve-NumberingDate `
                        -DateBasis ([string]$web.date_basis) `
                        -VoucherDate $rawVoucherDate

                    Build-WebVoucherNumber `
                        -Config $web `
                        -VoucherDate $effectiveLastDate `
                        -Sequence $web.last_sequence
                }
                else {
                    "(none yet)"
                }
            }
        }
    }

    # Existing BUSY function from vouchers.ps1.
    return Get-NumberingConfig -VchType $VchType -SeriesName $SeriesName -InstanceId $InstanceId -CompanyCode $CompanyCode
}

# -----------------------------------------------------------------------------
# HOTFIX V5: Access/BDS approval configuration must use the BUSY COM recordset.
#
# The Access instance's <company>\db.bds file is the authentication/preferences
# database and does not necessarily contain the fiscal Config table. Existing
# voucher settings already prove that Config is available through Connect-BUSY.
#
# SQL Server keeps the fast direct fiscal-DB path introduced in v3/v4.
# -----------------------------------------------------------------------------

function Test-BusyCloudApprovalUsesAccessCom {
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

function Get-VoucherApprovalConfig-AccessCom {
    param(
        [int]$VchType,
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        $ExistingFi = $null
    )

    $fi = $ExistingFi
    $ownsConnection = $false

    if (-not $fi) {
        $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
        $ownsConnection = $true
    }

    if (-not $fi) {
        return @{ success=$false; error="BUSY database connection failed" }
    }

    try {
        $approvalRequired = $false
        $approvers = @()

        $rst = $fi.GetRecordset("SELECT I1 FROM Config WHERE RecType=203 AND [Type]=$VchType")
        if ($rst -and -not $rst.EOF) {
            try {
                $raw = $rst.Fields.Item("I1").Value
                if ($null -ne $raw -and $raw -ne [System.DBNull]::Value) {
                    $approvalRequired = ([int]$raw -eq 1)
                }
            } catch {}
        }
        if ($rst) { try { $rst.Close() } catch {} }

        $rst = $fi.GetRecordset("SELECT C1 FROM Config WHERE RecType=204 AND [Type]=$VchType AND I1=1")
        if ($rst) {
            while (-not $rst.EOF) {
                $name = ""
                try {
                    $rawName = $rst.Fields.Item("C1").Value
                    if ($null -ne $rawName -and $rawName -ne [System.DBNull]::Value) {
                        $name = ([string]$rawName).Trim()
                    }
                } catch {}

                if ($name -and $approvers -notcontains $name) {
                    $approvers += $name
                }
                $rst.MoveNext()
            }
            try { $rst.Close() } catch {}
        }

        return @{
            success = $true
            data = @{
                vch_type = $VchType
                approval_required = [bool]$approvalRequired
                approvers = @($approvers | Sort-Object)
            }
        }
    }
    catch {
        return @{ success=$false; error=$_.Exception.Message }
    }
    finally {
        if ($ownsConnection -and $fi) {
            Disconnect-BUSY $fi
        }
    }
}

function Save-VoucherApprovalConfig-AccessCom {
    param(
        $Data,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $vchType = 0
    try { $vchType = [int]$Data.vch_type } catch {}

    if (-not (Test-IsBusyCloudApprovalVoucherType -VchType $vchType)) {
        return @{ success=$false; error="Unsupported or missing vch_type." }
    }

    $rawRequired = ([string]$Data.approval_required).Trim().ToLowerInvariant()
    $approvalRequired = (
        $Data.approval_required -eq $true -or
        $rawRequired -eq "1" -or
        $rawRequired -eq "true"
    )

    $requestedApprovers = @(
        @($Data.approvers) |
        ForEach-Object { ([string]$_).Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique
    )

    $usersResult = Get-CompanyUsers -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $usersResult.success) {
        return @{
            success=$false
            error="Could not validate approvers against BUSY users. $($usersResult.error)"
        }
    }

    $canonicalByLower = @{}
    foreach ($u in @($usersResult.data)) {
        $n = ([string]$u).Trim()
        if ($n) { $canonicalByLower[$n.ToLowerInvariant()] = $n }
    }

    $approvers = @()
    $unknown = @()

    foreach ($requested in $requestedApprovers) {
        $key = $requested.ToLowerInvariant()
        if ($canonicalByLower.ContainsKey($key)) {
            $canonical = [string]$canonicalByLower[$key]
            if ($approvers -notcontains $canonical) { $approvers += $canonical }
        }
        else {
            $unknown += $requested
        }
    }

    if ($unknown.Count -gt 0) {
        return @{
            success=$false
            error=("Unknown BUSY user(s): " + ($unknown -join ", "))
        }
    }

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) {
        return @{ success=$false; error="BUSY database connection failed" }
    }

    try {
        $requiredInt = if ($approvalRequired) { 1 } else { 0 }

        $exists = $false
        $rst = $fi.GetRecordset("SELECT RecType FROM Config WHERE RecType=203 AND [Type]=$vchType")
        if ($rst -and -not $rst.EOF) { $exists = $true }
        if ($rst) { try { $rst.Close() } catch {} }

        if ($exists) {
            $fi.ExecuteQuery("UPDATE Config SET I1=$requiredInt WHERE RecType=203 AND [Type]=$vchType")
        }
        else {
            $fi.ExecuteQuery("INSERT INTO Config (RecType,[Type],I1) VALUES (203,$vchType,$requiredInt)")
        }

        $fi.ExecuteQuery("DELETE FROM Config WHERE RecType=204 AND [Type]=$vchType")

        foreach ($name in $approvers) {
            $safe = $name.Replace("'", "''")
            $fi.ExecuteQuery("INSERT INTO Config (RecType,[Type],I1,C1) VALUES (204,$vchType,1,'$safe')")
        }

        return @{
            success=$true
            message="Voucher approval configuration updated successfully"
            data=@{
                vch_type=$vchType
                approval_required=[bool]$approvalRequired
                approvers=@($approvers)
            }
        }
    }
    catch {
        return @{ success=$false; error=$_.Exception.Message }
    }
    finally {
        Disconnect-BUSY $fi
    }
}

function Get-VoucherApprovalTypesForUser-AccessCom {
    param(
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        $ExistingFi = $null
    )

    if ($IsAdmin) {
        return @{ success=$true; data=@($script:BusyCloudApprovalVoucherTypes) }
    }

    if ([string]::IsNullOrWhiteSpace($UserName)) {
        return @{ success=$true; data=@() }
    }

    $fi = $ExistingFi
    $ownsConnection = $false

    if (-not $fi) {
        $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
        $ownsConnection = $true
    }

    if (-not $fi) {
        return @{ success=$false; error="BUSY database connection failed" }
    }

    try {
        $safeUser = $UserName.Trim().Replace("'", "''")
        $rst = $fi.GetRecordset("SELECT [Type] FROM Config WHERE RecType=204 AND I1=1 AND C1='$safeUser'")

        $types = @()
        if ($rst) {
            while (-not $rst.EOF) {
                $t = 0
                try { $t = [int]$rst.Fields.Item("Type").Value } catch {}

                if (
                    (Test-IsBusyCloudApprovalVoucherType -VchType $t) -and
                    $types -notcontains $t
                ) {
                    $types += $t
                }
                $rst.MoveNext()
            }
            try { $rst.Close() } catch {}
        }

        return @{ success=$true; data=@($types | Sort-Object) }
    }
    catch {
        return @{ success=$false; error=$_.Exception.Message }
    }
    finally {
        if ($ownsConnection -and $fi) {
            Disconnect-BUSY $fi
        }
    }
}

# Public dispatchers. SQL keeps the direct fiscal DB path; Access/BDS uses COM.
function Get-VoucherApprovalConfig {
    param(
        [int]$VchType,
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        $ExistingFi = $null
    )

    if (-not (Test-IsBusyCloudApprovalVoucherType -VchType $VchType)) {
        return @{
            success=$false
            error="Voucher type $VchType is not supported by BusyCloud approval processing."
        }
    }

    if (Test-BusyCloudApprovalUsesAccessCom -InstanceId $InstanceId -CompanyCode $CompanyCode) {
        return Get-VoucherApprovalConfig-AccessCom `
            -VchType $VchType `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode `
            -ExistingFi $ExistingFi
    }

    return Get-VoucherApprovalConfig-Direct `
        -VchType $VchType `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode
}

function Save-VoucherApprovalConfig {
    param(
        $Data,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if (Test-BusyCloudApprovalUsesAccessCom -InstanceId $InstanceId -CompanyCode $CompanyCode) {
        return Save-VoucherApprovalConfig-AccessCom `
            -Data $Data `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }

    return Save-VoucherApprovalConfig-Direct `
        -Data $Data `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode
}

function Get-VoucherApprovalTypesForUser {
    param(
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        $ExistingFi = $null
    )

    if ($IsAdmin) {
        return @{ success=$true; data=@($script:BusyCloudApprovalVoucherTypes) }
    }

    if (Test-BusyCloudApprovalUsesAccessCom -InstanceId $InstanceId -CompanyCode $CompanyCode) {
        return Get-VoucherApprovalTypesForUser-AccessCom `
            -UserName $UserName `
            -IsAdmin:$IsAdmin `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode `
            -ExistingFi $ExistingFi
    }

    return Get-VoucherApprovalTypesForUser-Direct `
        -UserName $UserName `
        -IsAdmin:$IsAdmin `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode
}

# Override with ExistingFi support so Access approve/unapprove can reuse the
# active COM session without opening/closing BUSY again.
function Test-VoucherApprover {
    param(
        [string]$UserName,
        [int]$VchType,
        [bool]$IsAdmin = $false,
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        $ExistingFi = $null
    )

    if (-not (Test-IsBusyCloudApprovalVoucherType -VchType $VchType)) {
        return @{ success=$true; allowed=$false; allowed_vch_types=@() }
    }

    if ($IsAdmin) {
        return @{
            success=$true
            allowed=$true
            allowed_vch_types=@($script:BusyCloudApprovalVoucherTypes)
        }
    }

    $types = Get-VoucherApprovalTypesForUser `
        -UserName $UserName `
        -IsAdmin:$false `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode `
        -ExistingFi $ExistingFi

    if (-not $types.success) { return $types }

    return @{
        success=$true
        allowed=(@($types.data) -contains $VchType)
        allowed_vch_types=@($types.data)
    }
}


# =============================================================================
# BusyCloud Voucher Approval v6 policy hardening
# - Approval OFF means NO approver access for that voucher type.
# - Explicit approver assignments are cleared when approval is disabled.
# - Admins remain automatic approvers only for voucher types whose approval
#   processing is currently enabled.
# =============================================================================

function Get-EnabledVoucherApprovalTypes-Direct {
    param([string]$InstanceId="", [string]$CompanyCode="")
    $ctx = $null
    try {
        $ctx = Get-BusyCloudApprovalDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $conn = $ctx.connection
        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 15 } catch {}
        $cmd.CommandText = "SELECT DISTINCT [Type] FROM Config WHERE RecType=203 AND I1 IN (1,3)"
        $rdr = $cmd.ExecuteReader()
        $types = @()
        while ($rdr.Read()) {
            $t = 0
            try { $t = [int]$rdr.GetValue(0) } catch {}
            if ((Test-IsBusyCloudApprovalVoucherType -VchType $t) -and $types -notcontains $t) { $types += $t }
        }
        $rdr.Close()
        return @{ success=$true; data=@($types | Sort-Object) }
    }
    catch { return @{ success=$false; error=$_.Exception.Message } }
    finally {
        if ($ctx -and $ctx.connection) { try { $ctx.connection.Close() } catch {}; try { $ctx.connection.Dispose() } catch {} }
    }
}

function Get-EnabledVoucherApprovalTypes-AccessCom {
    param([string]$InstanceId="", [string]$CompanyCode="", $ExistingFi=$null)
    $fi = $ExistingFi
    $ownsConnection = $false
    if (-not $fi) {
        $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
        $ownsConnection = $true
    }
    if (-not $fi) { return @{ success=$false; error="BUSY database connection failed" } }
    try {
        $rst = $fi.GetRecordset("SELECT [Type] FROM Config WHERE RecType=203 AND I1 IN (1,3)")
        $types = @()
        if ($rst) {
            while (-not $rst.EOF) {
                $t = 0
                try { $t = [int]$rst.Fields.Item("Type").Value } catch {}
                if ((Test-IsBusyCloudApprovalVoucherType -VchType $t) -and $types -notcontains $t) { $types += $t }
                $rst.MoveNext()
            }
            try { $rst.Close() } catch {}
        }
        return @{ success=$true; data=@($types | Sort-Object) }
    }
    catch { return @{ success=$false; error=$_.Exception.Message } }
    finally {
        if ($ownsConnection -and $fi) { Disconnect-BUSY $fi }
    }
}

function Get-EnabledVoucherApprovalTypes {
    param([string]$InstanceId="", [string]$CompanyCode="", $ExistingFi=$null)
    if (Test-BusyCloudApprovalUsesAccessCom -InstanceId $InstanceId -CompanyCode $CompanyCode) {
        return Get-EnabledVoucherApprovalTypes-AccessCom -InstanceId $InstanceId -CompanyCode $CompanyCode -ExistingFi $ExistingFi
    }
    return Get-EnabledVoucherApprovalTypes-Direct -InstanceId $InstanceId -CompanyCode $CompanyCode
}

function Get-VoucherApprovalConfig {
    param(
        [int]$VchType,
        [string]$InstanceId="",
        [string]$CompanyCode="",
        $ExistingFi=$null
    )
    if (-not (Test-IsBusyCloudApprovalVoucherType -VchType $VchType)) {
        return @{ success=$false; error="Voucher type $VchType is not supported by BusyCloud approval processing." }
    }
    $res = if (Test-BusyCloudApprovalUsesAccessCom -InstanceId $InstanceId -CompanyCode $CompanyCode) {
        Get-VoucherApprovalConfig-AccessCom -VchType $VchType -InstanceId $InstanceId -CompanyCode $CompanyCode -ExistingFi $ExistingFi
    } else {
        Get-VoucherApprovalConfig-Direct -VchType $VchType -InstanceId $InstanceId -CompanyCode $CompanyCode
    }
    if ($res.success -and -not [bool]$res.data.approval_required) {
        # Never expose stale assignments while the policy is OFF.
        $res.data.approvers = @()
    }
    return $res
}

function Save-VoucherApprovalConfig {
    param($Data, [string]$InstanceId="", [string]$CompanyCode="")
    $rawRequired = ([string]$Data.approval_required).Trim().ToLowerInvariant()
    $enabled = ($Data.approval_required -eq $true -or $rawRequired -eq "1" -or $rawRequired -eq "true")
    if (-not $enabled) {
        try { $Data.approvers = @() } catch {}
    }
    if (Test-BusyCloudApprovalUsesAccessCom -InstanceId $InstanceId -CompanyCode $CompanyCode) {
        return Save-VoucherApprovalConfig-AccessCom -Data $Data -InstanceId $InstanceId -CompanyCode $CompanyCode
    }
    return Save-VoucherApprovalConfig-Direct -Data $Data -InstanceId $InstanceId -CompanyCode $CompanyCode
}

function Get-VoucherApprovalTypesForUser {
    param(
        [string]$UserName,
        [bool]$IsAdmin=$false,
        [string]$InstanceId="",
        [string]$CompanyCode="",
        $ExistingFi=$null
    )

    $enabledResult = Get-EnabledVoucherApprovalTypes -InstanceId $InstanceId -CompanyCode $CompanyCode -ExistingFi $ExistingFi
    if (-not $enabledResult.success) { return $enabledResult }
    $enabledTypes = @($enabledResult.data | ForEach-Object { [int]$_ } | Sort-Object -Unique)

    if ($IsAdmin) {
        # Admins are automatic approvers, but ONLY while approval processing is ON.
        return @{ success=$true; data=@($enabledTypes) }
    }
    if ([string]::IsNullOrWhiteSpace($UserName) -or $enabledTypes.Count -eq 0) {
        return @{ success=$true; data=@() }
    }

    $assignedResult = if (Test-BusyCloudApprovalUsesAccessCom -InstanceId $InstanceId -CompanyCode $CompanyCode) {
        Get-VoucherApprovalTypesForUser-AccessCom -UserName $UserName -IsAdmin:$false -InstanceId $InstanceId -CompanyCode $CompanyCode -ExistingFi $ExistingFi
    } else {
        Get-VoucherApprovalTypesForUser-Direct -UserName $UserName -IsAdmin:$false -InstanceId $InstanceId -CompanyCode $CompanyCode
    }
    if (-not $assignedResult.success) { return $assignedResult }

    $types = @($assignedResult.data | ForEach-Object { [int]$_ } | Where-Object { $enabledTypes -contains $_ } | Sort-Object -Unique)
    return @{ success=$true; data=@($types) }
}

function Test-VoucherApprover {
    param(
        [string]$UserName,
        [int]$VchType,
        [bool]$IsAdmin=$false,
        [string]$InstanceId="",
        [string]$CompanyCode="",
        $ExistingFi=$null
    )
    if (-not (Test-IsBusyCloudApprovalVoucherType -VchType $VchType)) {
        return @{ success=$true; allowed=$false; allowed_vch_types=@() }
    }
    $types = Get-VoucherApprovalTypesForUser -UserName $UserName -IsAdmin:$IsAdmin -InstanceId $InstanceId -CompanyCode $CompanyCode -ExistingFi $ExistingFi
    if (-not $types.success) { return $types }
    return @{
        success=$true
        allowed=(@($types.data) -contains $VchType)
        allowed_vch_types=@($types.data)
    }
}
# =============================================================================
# FAST VOUCHER SETTINGS DIRECT-DB OVERRIDES
# =============================================================================
# Goal:
#   - Never start BUSY COM just to read/write RecType 201 voucher UI settings.
#   - SQL: resolve the fiscal SQL database and use SqlConnection directly.
#   - Access/BDS: discover the fiscal .bds containing Config and use OleDb
#     directly, the same style used by the fast permission subsystem.
#   - Preserve the old COM implementation as a safety fallback.
#
# These definitions intentionally appear at the END of this module so they
# override the older COM-first functions above without changing routes.ps1.
# =============================================================================

if ($null -eq $script:BusyCloudFastConfigAccessDbCache) {
    $script:BusyCloudFastConfigAccessDbCache = @{}
}

if ($null -eq $script:BusyCloudFastColumnConfigCache) {
    $script:BusyCloudFastColumnConfigCache = @{}
}

# Capture the currently loaded implementations before overriding them.
if ($null -eq $script:BusyCloudLegacyGetColumnConfig) {
    $script:BusyCloudLegacyGetColumnConfig = ${function:Get-ColumnConfig}
}
if ($null -eq $script:BusyCloudLegacySaveColumnConfig) {
    $script:BusyCloudLegacySaveColumnConfig = ${function:Save-ColumnConfig}
}
if ($null -eq $script:BusyCloudLegacyGetVoucherApprovalConfigAccessCom) {
    $script:BusyCloudLegacyGetVoucherApprovalConfigAccessCom = ${function:Get-VoucherApprovalConfig-AccessCom}
}
if ($null -eq $script:BusyCloudLegacySaveVoucherApprovalConfigAccessCom) {
    $script:BusyCloudLegacySaveVoucherApprovalConfigAccessCom = ${function:Save-VoucherApprovalConfig-AccessCom}
}
if ($null -eq $script:BusyCloudLegacyGetEnabledVoucherApprovalTypesAccessCom) {
    $script:BusyCloudLegacyGetEnabledVoucherApprovalTypesAccessCom = ${function:Get-EnabledVoucherApprovalTypes-AccessCom}
}
if ($null -eq $script:BusyCloudLegacyGetVoucherApprovalTypesForUserAccessCom) {
    $script:BusyCloudLegacyGetVoucherApprovalTypesForUserAccessCom = ${function:Get-VoucherApprovalTypesForUser-AccessCom}
}

function Test-BusyCloudAccessConfigDatabase {
    param([string]$DbFile)

    if ([string]::IsNullOrWhiteSpace($DbFile) -or -not (Test-Path -LiteralPath $DbFile)) {
        return $false
    }

    $conn = $null
    $rdr = $null
    try {
        $conn = Open-BdsConnection -DbFile $DbFile
        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 3 } catch {}
        $cmd.CommandText = "SELECT TOP 1 [RecType] FROM [Config]"
        $rdr = $cmd.ExecuteReader()
        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($rdr) {
            try { $rdr.Close() } catch {}
            try { $rdr.Dispose() } catch {}
        }
        if ($conn) {
            try { $conn.Close() } catch {}
            try { $conn.Dispose() } catch {}
        }
    }
}

function Resolve-BusyCloudFiscalAccessDatabasePath {
    param(
        $Instance,
        $Company,
        [string]$InstanceId,
        [string]$CompanyCode
    )

    $cacheKey = ("{0}|{1}" -f $InstanceId, $CompanyCode).ToLowerInvariant()

    if ($script:BusyCloudFastConfigAccessDbCache.ContainsKey($cacheKey)) {
        $cachedPath = [string]$script:BusyCloudFastConfigAccessDbCache[$cacheKey]
        if (Test-Path -LiteralPath $cachedPath) {
            return $cachedPath
        }
        $script:BusyCloudFastConfigAccessDbCache.Remove($cacheKey)
    }

    $companyFolder = Join-Path ([string]$Instance.dataPath) ([string]$Company.code)

    if (-not (Test-Path -LiteralPath $companyFolder)) {
        throw "Access company folder not found: $companyFolder"
    }

    # BUSY keeps the authentication/preferences database in db.bds. Fiscal
    # transaction/config data may be in another .bds file. Discover it once,
    # then cache the exact path for the lifetime of the API process.
    $files = @(
        Get-ChildItem -LiteralPath $companyFolder -Filter "*.bds" -File -ErrorAction SilentlyContinue
    )

    if ($files.Count -eq 0) {
        $files = @(
            Get-ChildItem -LiteralPath $companyFolder -Filter "*.bds" -File -Recurse -ErrorAction SilentlyContinue |
            Select-Object -First 50
        )
    }

    if ($files.Count -eq 0) {
        throw "No .bds files were found under $companyFolder"
    }

    $now = Get-Date
    $fyStartYear = if ($now.Month -ge 4) { $now.Year } else { $now.Year - 1 }
    $fyText = [string]$fyStartYear

    # Prefer an obvious current-financial-year database, then non-main BDS
    # files by recent write time, and only then db.bds.
    $ranked = @(
        $files |
        ForEach-Object {
            $name = ([string]$_.Name).ToLowerInvariant()
            $score = 20

            if ($name -match [regex]::Escape($fyText)) {
                $score = 0
            }
            elseif ($name -ne "db.bds") {
                $score = 10
            }
            else {
                $score = 30
            }

            [pscustomobject]@{
                File = $_
                Score = $score
                LastWriteTime = $_.LastWriteTime
            }
        } |
        Sort-Object Score, @{ Expression = "LastWriteTime"; Descending = $true }
    )

    foreach ($candidate in $ranked) {
        $path = [string]$candidate.File.FullName

        if (Test-BusyCloudAccessConfigDatabase -DbFile $path) {
            $script:BusyCloudFastConfigAccessDbCache[$cacheKey] = $path
            Write-Host "  [FAST-CONFIG-DB] $InstanceId/$CompanyCode -> $path" -ForegroundColor DarkCyan
            return $path
        }
    }

    $names = @($ranked | ForEach-Object { $_.File.Name }) -join ", "
    throw "Could not find an Access/BDS fiscal database containing Config for $InstanceId/$CompanyCode. Checked: $names"
}

function Get-BusyCloudFastConfigDbContext {
    param(
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $found = Get-InstanceForCompany -CompanyCode $CompanyCode -InstanceId $InstanceId
    if (-not $found) {
        throw "Company not found in instances.json"
    }

    $inst = $found.instance
    $comp = $found.company
    $dbType = if ($null -ne $inst.dbType) { [int]$inst.dbType } else { 0 }

    if ($dbType -eq 1) {
        $dbName = Resolve-BusyCloudFiscalSqlDatabaseName `
            -Instance $inst `
            -CompanyCode $CompanyCode `
            -InstanceId ([string]$inst.id)

        $conn = Open-SqlConnection `
            -SqlServer $inst.sqlServer `
            -Database $dbName `
            -SqlUser $inst.sqlUser `
            -SqlPassword $inst.sqlPassword

        return @{
            dbType = 1
            connection = $conn
            instance = $inst
            company = $comp
            database = $dbName
        }
    }

    $dbFile = Resolve-BusyCloudFiscalAccessDatabasePath `
        -Instance $inst `
        -Company $comp `
        -InstanceId ([string]$inst.id) `
        -CompanyCode ([string]$comp.code)

    $conn = Open-BdsConnection -DbFile $dbFile

    return @{
        dbType = 0
        connection = $conn
        instance = $inst
        company = $comp
        database = $dbFile
    }
}

function Get-BusyCloudReaderValue {
    param(
        $Reader,
        [string]$Field,
        $Default = $null
    )

    try {
        $ordinal = $Reader.GetOrdinal($Field)
        if ($ordinal -ge 0 -and -not $Reader.IsDBNull($ordinal)) {
            return $Reader.GetValue($ordinal)
        }
    }
    catch {}

    return $Default
}

function Convert-BusyCloudBehaviorIntToName {
    param([int]$Value)

    if ($Value -eq 2) { return "semi_variable" }
    if ($Value -eq 3) { return "fixed" }
    return "variable"
}

function Convert-BusyCloudBehaviorNameToInt {
    param([string]$Value)

    if ($Value -eq "semi_variable") { return 2 }
    if ($Value -eq "fixed") { return 3 }
    return 1
}

function Get-BusyCloudColumnConfigDefault {
    param(
        [int]$VchType,
        [int]$DeviceType
    )

    $isAccountVoucher = @(14, 15, 16, 19) -contains $VchType
    $supportsQuickMode = @(14, 19) -contains $VchType

    if ($isAccountVoucher) {
        return @{
            vch_type = $VchType
            device_type = $DeviceType
            account_default_mode = if ($supportsQuickMode) { "single" } else { "double" }
            account_allow_mode_switch = if ($supportsQuickMode) { $true } else { $false }
            acc_quick_col_account = "variable"
            acc_quick_col_amount = "variable"
            acc_quick_col_short_narration = "variable"
            acc_double_col_dc = "variable"
            acc_double_col_account = "variable"
            acc_double_col_debit = "variable"
            acc_double_col_credit = "variable"
            acc_double_col_short_narration = "variable"
        }
    }

    return @{
        vch_type = $VchType
        device_type = $DeviceType
        enable_item_discount = $true
        enable_alt_units = $true
        col_qty = "variable"
        col_unit = "variable"
        col_price = "variable"
        col_amount = "variable"
        col_discount = "variable"
        col_cfact = "variable"
        col_alt_qty = "variable"
        col_alt_price = "variable"
        enable_pos = $false
        show_stock_balance = $true
        def_card_acc = ""
        def_gift_acc = ""
    }
}

function Get-ColumnConfig {
    param(
        [int]$VchType,
        [int]$DeviceType,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $cacheKey = ("{0}|{1}|{2}|{3}" -f $InstanceId, $CompanyCode, $VchType, $DeviceType).ToLowerInvariant()

    try {
        if ($script:BusyCloudFastColumnConfigCache.ContainsKey($cacheKey)) {
            $cached = $script:BusyCloudFastColumnConfigCache[$cacheKey]
            if ($cached -and (Get-Date) -lt $cached.expires) {
                return @{ success = $true; data = $cached.data }
            }
            $script:BusyCloudFastColumnConfigCache.Remove($cacheKey)
        }
    }
    catch {}

    $ctx = $null
    $rdr = $null

    try {
        $ctx = Get-BusyCloudFastConfigDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $conn = $ctx.connection

        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}
        $cmd.CommandText = "SELECT TOP 1 * FROM Config WHERE RecType=201 AND [Type]=$VchType AND D15=$DeviceType"

        $rdr = $cmd.ExecuteReader()

        $config = $null
        if ($rdr.Read()) {
            $isAccountVoucher = @(14, 15, 16, 19) -contains $VchType
            $supportsQuickMode = @(14, 19) -contains $VchType

            if ($isAccountVoucher) {
                $storedMode = [int](Get-BusyCloudReaderValue -Reader $rdr -Field "I1" -Default 1)

                $config = @{
                    vch_type = $VchType
                    device_type = $DeviceType
                    account_default_mode = if ($supportsQuickMode -and $storedMode -eq 1) { "single" } else { "double" }
                    account_allow_mode_switch = if ($supportsQuickMode) {
                        ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I2" -Default 1) -eq 1)
                    } else {
                        $false
                    }
                    acc_quick_col_account = Convert-BusyCloudBehaviorIntToName ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I3" -Default 1))
                    acc_quick_col_amount = Convert-BusyCloudBehaviorIntToName ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I4" -Default 1))
                    acc_quick_col_short_narration = Convert-BusyCloudBehaviorIntToName ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I5" -Default 1))
                    acc_double_col_dc = Convert-BusyCloudBehaviorIntToName ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I6" -Default 1))
                    acc_double_col_account = Convert-BusyCloudBehaviorIntToName ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I7" -Default 1))
                    acc_double_col_debit = Convert-BusyCloudBehaviorIntToName ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I8" -Default 1))
                    acc_double_col_credit = Convert-BusyCloudBehaviorIntToName ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I9" -Default 1))
                    acc_double_col_short_narration = Convert-BusyCloudBehaviorIntToName ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I10" -Default 1))
                }
            }
            else {
                $config = @{
                    vch_type = $VchType
                    device_type = $DeviceType
                    enable_item_discount = ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I1" -Default 1) -eq 1)
                    enable_alt_units = ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I2" -Default 1) -eq 1)
                    col_qty = Convert-BusyCloudBehaviorIntToName ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I3" -Default 1))
                    col_unit = Convert-BusyCloudBehaviorIntToName ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I4" -Default 1))
                    col_price = Convert-BusyCloudBehaviorIntToName ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I5" -Default 1))
                    col_amount = Convert-BusyCloudBehaviorIntToName ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I6" -Default 1))
                    col_discount = Convert-BusyCloudBehaviorIntToName ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I7" -Default 1))
                    col_cfact = Convert-BusyCloudBehaviorIntToName ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I8" -Default 1))
                    col_alt_qty = Convert-BusyCloudBehaviorIntToName ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I9" -Default 1))
                    col_alt_price = Convert-BusyCloudBehaviorIntToName ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I10" -Default 1))
                    enable_pos = ([int](Get-BusyCloudReaderValue -Reader $rdr -Field "I11" -Default 0) -eq 1)
                    show_stock_balance = ([string](Get-BusyCloudReaderValue -Reader $rdr -Field "C3" -Default "1") -ne "0")
                    def_card_acc = [string](Get-BusyCloudReaderValue -Reader $rdr -Field "C1" -Default "")
                    def_gift_acc = [string](Get-BusyCloudReaderValue -Reader $rdr -Field "C2" -Default "")
                }
            }
        }

        if ($rdr) {
            try { $rdr.Close() } catch {}
            try { $rdr.Dispose() } catch {}
            $rdr = $null
        }

        if ($null -eq $config) {
            $config = Get-BusyCloudColumnConfigDefault -VchType $VchType -DeviceType $DeviceType
        }

        $script:BusyCloudFastColumnConfigCache[$cacheKey] = @{
            expires = (Get-Date).AddSeconds(30)
            data = $config
        }

        return @{
            success = $true
            data = $config
            storage = if ($ctx.dbType -eq 1) { "SQL-DIRECT" } else { "ACCESS-DIRECT" }
        }
    }
    catch {
        Write-Host "  [FAST-CONFIG WARN] Direct column-config read failed for $InstanceId/$CompanyCode; using COM fallback. $($_.Exception.Message)" -ForegroundColor DarkYellow

        return & $script:BusyCloudLegacyGetColumnConfig `
            -VchType $VchType `
            -DeviceType $DeviceType `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }
    finally {
        if ($rdr) {
            try { $rdr.Close() } catch {}
            try { $rdr.Dispose() } catch {}
        }
        if ($ctx -and $ctx.connection) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}

function Save-ColumnConfig {
    param(
        $Data,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $vchType = [int]$Data.vch_type
    $deviceType = if ($null -ne $Data.device_type) { [int]$Data.device_type } else { 0 }

    $ctx = $null

    try {
        $ctx = Get-BusyCloudFastConfigDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $conn = $ctx.connection

        $isAccountVoucher = @(14, 15, 16, 19) -contains $vchType
        $supportsQuickMode = @(14, 19) -contains $vchType

        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}
        $cmd.CommandText = "SELECT COUNT(*) FROM Config WHERE RecType=201 AND [Type]=$vchType AND D15=$deviceType"
        $exists = ([int]$cmd.ExecuteScalar() -gt 0)

        if ($isAccountVoucher) {
            $requestedMode = ([string]$Data.account_default_mode).ToLowerInvariant()
            $i1 = if ($supportsQuickMode -and $requestedMode -eq "single") { 1 } else { 2 }
            $i2 = if ($supportsQuickMode -and ($Data.account_allow_mode_switch -eq $true -or $Data.account_allow_mode_switch -eq "true")) { 1 } else { 0 }
            $i3 = Convert-BusyCloudBehaviorNameToInt ([string]$Data.acc_quick_col_account)
            $i4 = Convert-BusyCloudBehaviorNameToInt ([string]$Data.acc_quick_col_amount)
            $i5 = Convert-BusyCloudBehaviorNameToInt ([string]$Data.acc_quick_col_short_narration)
            $i6 = Convert-BusyCloudBehaviorNameToInt ([string]$Data.acc_double_col_dc)
            $i7 = Convert-BusyCloudBehaviorNameToInt ([string]$Data.acc_double_col_account)
            $i8 = Convert-BusyCloudBehaviorNameToInt ([string]$Data.acc_double_col_debit)
            $i9 = Convert-BusyCloudBehaviorNameToInt ([string]$Data.acc_double_col_credit)
            $i10 = Convert-BusyCloudBehaviorNameToInt ([string]$Data.acc_double_col_short_narration)

            if ($exists) {
                $sql = "UPDATE Config SET I1=$i1,I2=$i2,I3=$i3,I4=$i4,I5=$i5,I6=$i6,I7=$i7,I8=$i8,I9=$i9,I10=$i10 WHERE RecType=201 AND [Type]=$vchType AND D15=$deviceType"
            }
            else {
                $sql = "INSERT INTO Config (RecType,[Type],D15,I1,I2,I3,I4,I5,I6,I7,I8,I9,I10) VALUES (201,$vchType,$deviceType,$i1,$i2,$i3,$i4,$i5,$i6,$i7,$i8,$i9,$i10)"
            }
        }
        else {
            $i1 = if ($Data.enable_item_discount -eq $true -or $Data.enable_item_discount -eq "true") { 1 } else { 0 }
            $i2 = if ($Data.enable_alt_units -eq $true -or $Data.enable_alt_units -eq "true") { 1 } else { 0 }
            $i3 = Convert-BusyCloudBehaviorNameToInt ([string]$Data.col_qty)
            $i4 = Convert-BusyCloudBehaviorNameToInt ([string]$Data.col_unit)
            $i5 = Convert-BusyCloudBehaviorNameToInt ([string]$Data.col_price)
            $i6 = Convert-BusyCloudBehaviorNameToInt ([string]$Data.col_amount)
            $i7 = Convert-BusyCloudBehaviorNameToInt ([string]$Data.col_discount)
            $i8 = Convert-BusyCloudBehaviorNameToInt ([string]$Data.col_cfact)
            $i9 = Convert-BusyCloudBehaviorNameToInt ([string]$Data.col_alt_qty)
            $i10 = Convert-BusyCloudBehaviorNameToInt ([string]$Data.col_alt_price)
            $i11 = if ($Data.enable_pos -eq $true -or $Data.enable_pos -eq "true") { 1 } else { 0 }

            $c3 = if ($null -eq $Data.show_stock_balance) {
                "1"
            }
            elseif ($Data.show_stock_balance -eq $true -or $Data.show_stock_balance -eq "true") {
                "1"
            }
            else {
                "0"
            }

            $c1 = if ($Data.def_card_acc) { ([string]$Data.def_card_acc).Replace("'", "''") } else { "" }
            $c2 = if ($Data.def_gift_acc) { ([string]$Data.def_gift_acc).Replace("'", "''") } else { "" }

            if ($exists) {
                $sql = "UPDATE Config SET I1=$i1,I2=$i2,I3=$i3,I4=$i4,I5=$i5,I6=$i6,I7=$i7,I8=$i8,I9=$i9,I10=$i10,I11=$i11,C1='$c1',C2='$c2',C3='$c3' WHERE RecType=201 AND [Type]=$vchType AND D15=$deviceType"
            }
            else {
                $sql = "INSERT INTO Config (RecType,[Type],D15,I1,I2,I3,I4,I5,I6,I7,I8,I9,I10,I11,C1,C2,C3) VALUES (201,$vchType,$deviceType,$i1,$i2,$i3,$i4,$i5,$i6,$i7,$i8,$i9,$i10,$i11,'$c1','$c2','$c3')"
            }
        }

        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}
        $cmd.CommandText = $sql
        [void]$cmd.ExecuteNonQuery()

        $cacheKey = ("{0}|{1}|{2}|{3}" -f $InstanceId, $CompanyCode, $vchType, $deviceType).ToLowerInvariant()
        if ($script:BusyCloudFastColumnConfigCache.ContainsKey($cacheKey)) {
            $script:BusyCloudFastColumnConfigCache.Remove($cacheKey)
        }

        return @{
            success = $true
            message = "Voucher configuration updated successfully"
            storage = if ($ctx.dbType -eq 1) { "SQL-DIRECT" } else { "ACCESS-DIRECT" }
        }
    }
    catch {
        Write-Host "  [FAST-CONFIG WARN] Direct column-config save failed for $InstanceId/$CompanyCode; using COM fallback. $($_.Exception.Message)" -ForegroundColor DarkYellow

        return & $script:BusyCloudLegacySaveColumnConfig `
            -Data $Data `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }
    finally {
        if ($ctx -and $ctx.connection) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}

# -----------------------------------------------------------------------------
# Access approval-config fast path
# -----------------------------------------------------------------------------
# The approval queue/actions may still use BUSY COM on Access when they already
# have an ExistingFi. Settings/navigation calls do not need COM, so when no
# ExistingFi is supplied these functions read the fiscal Config BDS directly.

function Get-VoucherApprovalConfig-AccessCom {
    param(
        [int]$VchType,
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        $ExistingFi = $null
    )

    if ($ExistingFi) {
        return & $script:BusyCloudLegacyGetVoucherApprovalConfigAccessCom `
            -VchType $VchType `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode `
            -ExistingFi $ExistingFi
    }

    $ctx = $null
    $rdr = $null

    try {
        $ctx = Get-BusyCloudFastConfigDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $conn = $ctx.connection

        $approvalRequired = $false
        $approvers = @()

        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}
        $cmd.CommandText = "SELECT TOP 1 I1 FROM Config WHERE RecType=203 AND [Type]=$VchType"
        $raw = $cmd.ExecuteScalar()
        if ($null -ne $raw -and $raw -ne [System.DBNull]::Value) {
            $approvalRequired = ([int]$raw -eq 1)
        }

        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}
        $cmd.CommandText = "SELECT C1 FROM Config WHERE RecType=204 AND [Type]=$VchType AND I1=1"
        $rdr = $cmd.ExecuteReader()

        while ($rdr.Read()) {
            $name = ""
            if (-not $rdr.IsDBNull(0)) {
                $name = ([string]$rdr.GetValue(0)).Trim()
            }
            if ($name -and $approvers -notcontains $name) {
                $approvers += $name
            }
        }

        return @{
            success = $true
            data = @{
                vch_type = $VchType
                approval_required = [bool]$approvalRequired
                approvers = @($approvers | Sort-Object)
            }
        }
    }
    catch {
        Write-Host "  [FAST-CONFIG WARN] Direct Access approval-config read failed; using COM fallback. $($_.Exception.Message)" -ForegroundColor DarkYellow

        return & $script:BusyCloudLegacyGetVoucherApprovalConfigAccessCom `
            -VchType $VchType `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }
    finally {
        if ($rdr) {
            try { $rdr.Close() } catch {}
            try { $rdr.Dispose() } catch {}
        }
        if ($ctx -and $ctx.connection) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}

function Save-VoucherApprovalConfig-AccessCom {
    param(
        $Data,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $vchType = 0
    try { $vchType = [int]$Data.vch_type } catch {}

    if (-not (Test-IsBusyCloudApprovalVoucherType -VchType $vchType)) {
        return @{ success = $false; error = "Unsupported or missing vch_type." }
    }

    $rawRequired = ([string]$Data.approval_required).Trim().ToLowerInvariant()
    $approvalRequired = (
        $Data.approval_required -eq $true -or
        $rawRequired -eq "1" -or
        $rawRequired -eq "true"
    )

    $requestedApprovers = @(
        @($Data.approvers) |
        ForEach-Object { ([string]$_).Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique
    )

    $usersResult = Get-CompanyUsers -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $usersResult.success) {
        return @{
            success = $false
            error = "Could not validate approvers against BUSY users. $($usersResult.error)"
        }
    }

    $canonicalByLower = @{}
    foreach ($u in @($usersResult.data)) {
        $n = ([string]$u).Trim()
        if ($n) {
            $canonicalByLower[$n.ToLowerInvariant()] = $n
        }
    }

    $approvers = @()
    $unknown = @()

    foreach ($requested in $requestedApprovers) {
        $key = $requested.ToLowerInvariant()
        if ($canonicalByLower.ContainsKey($key)) {
            $canonical = [string]$canonicalByLower[$key]
            if ($approvers -notcontains $canonical) {
                $approvers += $canonical
            }
        }
        else {
            $unknown += $requested
        }
    }

    if ($unknown.Count -gt 0) {
        return @{
            success = $false
            error = ("Unknown BUSY user(s): " + ($unknown -join ", "))
        }
    }

    $ctx = $null

    try {
        $ctx = Get-BusyCloudFastConfigDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $conn = $ctx.connection
        $requiredInt = if ($approvalRequired) { 1 } else { 0 }

        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}
        $cmd.CommandText = "SELECT COUNT(*) FROM Config WHERE RecType=203 AND [Type]=$vchType"
        $exists = ([int]$cmd.ExecuteScalar() -gt 0)

        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}

        if ($exists) {
            $cmd.CommandText = "UPDATE Config SET I1=$requiredInt WHERE RecType=203 AND [Type]=$vchType"
        }
        else {
            $cmd.CommandText = "INSERT INTO Config (RecType,[Type],I1) VALUES (203,$vchType,$requiredInt)"
        }
        [void]$cmd.ExecuteNonQuery()

        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}
        $cmd.CommandText = "DELETE FROM Config WHERE RecType=204 AND [Type]=$vchType"
        [void]$cmd.ExecuteNonQuery()

        foreach ($name in $approvers) {
            $safe = $name.Replace("'", "''")
            $cmd = $conn.CreateCommand()
            try { $cmd.CommandTimeout = 5 } catch {}
            $cmd.CommandText = "INSERT INTO Config (RecType,[Type],I1,C1) VALUES (204,$vchType,1,'$safe')"
            [void]$cmd.ExecuteNonQuery()
        }

        return @{
            success = $true
            message = "Voucher approval configuration updated successfully"
            data = @{
                vch_type = $vchType
                approval_required = [bool]$approvalRequired
                approvers = @($approvers)
            }
        }
    }
    catch {
        Write-Host "  [FAST-CONFIG WARN] Direct Access approval-config save failed; using COM fallback. $($_.Exception.Message)" -ForegroundColor DarkYellow

        return & $script:BusyCloudLegacySaveVoucherApprovalConfigAccessCom `
            -Data $Data `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }
    finally {
        if ($ctx -and $ctx.connection) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}

function Get-EnabledVoucherApprovalTypes-AccessCom {
    param(
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        $ExistingFi = $null
    )

    if ($ExistingFi) {
        return & $script:BusyCloudLegacyGetEnabledVoucherApprovalTypesAccessCom `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode `
            -ExistingFi $ExistingFi
    }

    $ctx = $null
    $rdr = $null

    try {
        $ctx = Get-BusyCloudFastConfigDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $cmd = $ctx.connection.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}
        $cmd.CommandText = "SELECT DISTINCT [Type] FROM Config WHERE RecType=203 AND I1 IN (1,3)"
        $rdr = $cmd.ExecuteReader()

        $types = @()
        while ($rdr.Read()) {
            $t = 0
            try { $t = [int]$rdr.GetValue(0) } catch {}

            if ((Test-IsBusyCloudApprovalVoucherType -VchType $t) -and $types -notcontains $t) {
                $types += $t
            }
        }

        return @{ success = $true; data = @($types | Sort-Object) }
    }
    catch {
        Write-Host "  [FAST-CONFIG WARN] Direct Access enabled-approval read failed; using COM fallback. $($_.Exception.Message)" -ForegroundColor DarkYellow

        return & $script:BusyCloudLegacyGetEnabledVoucherApprovalTypesAccessCom `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }
    finally {
        if ($rdr) {
            try { $rdr.Close() } catch {}
            try { $rdr.Dispose() } catch {}
        }
        if ($ctx -and $ctx.connection) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}

function Get-VoucherApprovalTypesForUser-AccessCom {
    param(
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        $ExistingFi = $null
    )

    if ($IsAdmin) {
        return @{ success = $true; data = @($script:BusyCloudApprovalVoucherTypes) }
    }

    if ([string]::IsNullOrWhiteSpace($UserName)) {
        return @{ success = $true; data = @() }
    }

    if ($ExistingFi) {
        return & $script:BusyCloudLegacyGetVoucherApprovalTypesForUserAccessCom `
            -UserName $UserName `
            -IsAdmin:$IsAdmin `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode `
            -ExistingFi $ExistingFi
    }

    $ctx = $null
    $rdr = $null

    try {
        $safeUser = $UserName.Trim().Replace("'", "''")
        $ctx = Get-BusyCloudFastConfigDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $cmd = $ctx.connection.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}
        $cmd.CommandText = "SELECT [Type] FROM Config WHERE RecType=204 AND I1=1 AND C1='$safeUser'"
        $rdr = $cmd.ExecuteReader()

        $types = @()
        while ($rdr.Read()) {
            $t = 0
            try { $t = [int]$rdr.GetValue(0) } catch {}

            if ((Test-IsBusyCloudApprovalVoucherType -VchType $t) -and $types -notcontains $t) {
                $types += $t
            }
        }

        return @{ success = $true; data = @($types | Sort-Object) }
    }
    catch {
        Write-Host "  [FAST-CONFIG WARN] Direct Access approver-type read failed; using COM fallback. $($_.Exception.Message)" -ForegroundColor DarkYellow

        return & $script:BusyCloudLegacyGetVoucherApprovalTypesForUserAccessCom `
            -UserName $UserName `
            -IsAdmin:$IsAdmin `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }
    finally {
        if ($rdr) {
            try { $rdr.Close() } catch {}
            try { $rdr.Dispose() } catch {}
        }
        if ($ctx -and $ctx.connection) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}

# =============================================================================
# END FAST VOUCHER SETTINGS DIRECT-DB OVERRIDES
# =============================================================================

# =============================================================================
# FAST OPTIONAL FIELD READ PATH
# =============================================================================
# Overrides the COM-based implementations originally loaded from vouchers.ps1.
# vch_setting.ps1 is loaded after vouchers.ps1, so these definitions win.
#
# Read-only endpoints covered:
#   GET /busy/voucher/optional-fields-config
#   GET /busy/voucher/optional-fields-values
#
# No Connect-BUSY / OpenCSDB / COM is used here.
# =============================================================================

Write-Host "  [FAST-OPTIONAL-FIELDS-V1] Direct optional-field read path loaded." -ForegroundColor DarkCyan

if ($null -eq $script:BusyCloudOptionalFieldsConfigCache) {
    $script:BusyCloudOptionalFieldsConfigCache = @{}
}

if ($null -eq $script:BusyCloudOptionalFieldValuesCache) {
    $script:BusyCloudOptionalFieldValuesCache = @{}
}

function Get-VoucherOptionalFields {
    param(
        [int]$VchType,
        [string]$SeriesName,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $startedAt = [System.Diagnostics.Stopwatch]::StartNew()

    if ($VchType -le 0 -or [string]::IsNullOrWhiteSpace($SeriesName)) {
        return @{
            success = $true
            count = 0
            data = @()
        }
    }

    $cleanSeriesName = $SeriesName.Trim()

    $cacheKey = (
        "{0}|{1}|{2}|{3}" -f
        $InstanceId,
        $CompanyCode,
        $VchType,
        $cleanSeriesName
    ).ToLowerInvariant()

    if ($script:BusyCloudOptionalFieldsConfigCache.ContainsKey($cacheKey)) {
        $entry = $script:BusyCloudOptionalFieldsConfigCache[$cacheKey]

        if (
            $entry -and
            $entry.expiresAt -gt [DateTime]::UtcNow
        ) {
            Write-Host (
                "  [OPTIONAL-FIELDS-FAST] cache HIT {0}/{1} type={2} series='{3}' rows={4}" -f
                $InstanceId,
                $CompanyCode,
                $VchType,
                $cleanSeriesName,
                @($entry.result.data).Count
            ) -ForegroundColor DarkCyan

            return $entry.result
        }

        try {
            $script:BusyCloudOptionalFieldsConfigCache.Remove($cacheKey)
        }
        catch {
        }
    }

    $ctx = $null
    $reader = $null
    $cmd = $null

    try {
        $ctx = Get-BusyCloudFastConfigDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (
            $null -eq $ctx -or
            $null -eq $ctx.connection
        ) {
            throw "Direct fiscal database connection is unavailable."
        }

        $conn = $ctx.connection

        # ---------------------------------------------------------------------
        # 1. Resolve the voucher-series master code.
        # ---------------------------------------------------------------------

        $prefixStr = "{0:D2}" -f $VchType

        $prefixedSeriesName = if (
            $cleanSeriesName.StartsWith(
                $prefixStr,
                [System.StringComparison]::OrdinalIgnoreCase
            )
        ) {
            $cleanSeriesName
        }
        else {
            "$prefixStr$cleanSeriesName"
        }

        $safeSeriesName =
            $cleanSeriesName.Replace("'", "''")

        $safePrefixedName =
            $prefixedSeriesName.Replace("'", "''")

        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}

        $cmd.CommandText = @"
SELECT TOP 1 Code
FROM Master1
WHERE MasterType = 21
  AND (
        Name = '$safeSeriesName'
        OR Name = '$safePrefixedName'
      )
"@

        $seriesValue = $cmd.ExecuteScalar()

        try { $cmd.Dispose() } catch {}
        $cmd = $null

        if (
            $null -eq $seriesValue -or
            $seriesValue -eq [System.DBNull]::Value
        ) {
            $result = @{
                success = $true
                count = 0
                data = @()
            }

            $script:BusyCloudOptionalFieldsConfigCache[$cacheKey] = @{
                expiresAt = [DateTime]::UtcNow.AddMinutes(5)
                result = $result
            }

            return $result
        }

        $seriesCode = [int]$seriesValue

        # ---------------------------------------------------------------------
        # 2. Read the single RecType=1 configuration row.
        #
        # The original logic only uses C1..C20 to build the returned field
        # definitions. I* values are read there but do not affect the result.
        # ---------------------------------------------------------------------

        $columns = @(
            1..20 |
            ForEach-Object { "C$_" }
        ) -join ", "

        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}

        $cmd.CommandText = @"
SELECT TOP 1
    $columns
FROM Config
WHERE RecType = 1
  AND L1 = $seriesCode
"@

        $reader = $cmd.ExecuteReader()

        $fieldNames = @{}

        if ($reader.Read()) {
            for ($i = 1; $i -le 20; $i++) {
                $fieldName = ""

                try {
                    $ordinal =
                        $reader.GetOrdinal("C$i")

                    if (-not $reader.IsDBNull($ordinal)) {
                        $fieldName = (
                            [string]$reader.GetValue($ordinal)
                        ).Trim()
                    }
                }
                catch {
                    $fieldName = ""
                }

                if (-not [string]::IsNullOrWhiteSpace($fieldName)) {
                    $fieldNames[$i] = $fieldName
                }
            }
        }

        try { $reader.Close() } catch {}
        try { $reader.Dispose() } catch {}
        $reader = $null

        try { $cmd.Dispose() } catch {}
        $cmd = $null

        if ($fieldNames.Count -eq 0) {
            $result = @{
                success = $true
                count = 0
                data = @()
            }

            $script:BusyCloudOptionalFieldsConfigCache[$cacheKey] = @{
                expiresAt = [DateTime]::UtcNow.AddMinutes(5)
                result = $result
            }

            return $result
        }

        # ---------------------------------------------------------------------
        # 3. Resolve dropdown/master-backed optional fields in ONE query.
        #
        # Original code could execute a separate COUNT(*) query for every text
        # field. We preserve the same 1000 + field number rule but collapse all
        # checks into a single DISTINCT query.
        # ---------------------------------------------------------------------

        $masterTypesPresent = @{}

        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}

        $cmd.CommandText = @"
SELECT DISTINCT MasterType
FROM Master1
WHERE MasterType >= 1001
  AND MasterType <= 1020
"@

        $reader = $cmd.ExecuteReader()

        while ($reader.Read()) {
            try {
                if (-not $reader.IsDBNull(0)) {
                    $masterTypesPresent[[int]$reader.GetValue(0)] = $true
                }
            }
            catch {
            }
        }

        try { $reader.Close() } catch {}
        try { $reader.Dispose() } catch {}
        $reader = $null

        try { $cmd.Dispose() } catch {}
        $cmd = $null

        # ---------------------------------------------------------------------
        # 4. Build the EXACT existing response shape and vocabulary mapping.
        # ---------------------------------------------------------------------

        $fields = @()

        foreach ($i in ($fieldNames.Keys | Sort-Object)) {
            $fName = [string]$fieldNames[$i]

            $fieldType = "text"
            $decimalPlaces = 0
            $maintainMaster = $false

            if (
                $fName -match
                "Date|Dated|Expiry|Due|Period|Mfg|Format|Year|Month|Day"
            ) {
                $fieldType = "date"
            }
            elseif (
                $fName -match
                "Bool|Booleom|Booleon|YesNo|Status|Active|Enabled|Hold|Block|Approved"
            ) {
                $fieldType = "boolean"
            }
            elseif (
                $fName -match
                "Number|Qty|Amt|Rate|Val|Numeric|Discount|Price|Tax|Gst|Balance|Percent|Charge|Cost|Comm|Commission|Duty|Freight"
            ) {
                $fieldType = "numeric"
                $decimalPlaces = 3
            }
            else {
                $targetMasterType = 1000 + [int]$i

                if ($masterTypesPresent.ContainsKey($targetMasterType)) {
                    $maintainMaster = $true
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

        $result = @{
            success = $true
            count = $fields.Count
            data = @($fields)
        }

        $script:BusyCloudOptionalFieldsConfigCache[$cacheKey] = @{
            expiresAt = [DateTime]::UtcNow.AddMinutes(5)
            result = $result
        }

        $startedAt.Stop()

        Write-Host (
            "  [OPTIONAL-FIELDS-FAST] {0}/{1} type={2} series='{3}' db={4} rows={5} elapsedMs={6}" -f
            $InstanceId,
            $CompanyCode,
            $VchType,
            $cleanSeriesName,
            [string]$ctx.database,
            $fields.Count,
            [int]$startedAt.ElapsedMilliseconds
        ) -ForegroundColor DarkCyan

        return $result
    }
    catch {
        if ($startedAt.IsRunning) {
            $startedAt.Stop()
        }

        Write-Host (
            "  [OPTIONAL-FIELDS-FAST FAIL] {0}/{1} type={2} series='{3}' elapsedMs={4} error={5}" -f
            $InstanceId,
            $CompanyCode,
            $VchType,
            $cleanSeriesName,
            [int]$startedAt.ElapsedMilliseconds,
            $_.Exception.Message
        ) -ForegroundColor Red

        # Deliberately fail fast. A read-only page-load endpoint should not
        # freeze the entire API process by falling back to BUSY COM.
        return @{
            success = $false
            error = $_.Exception.Message
            data = @()
        }
    }
    finally {
        if ($reader) {
            try { $reader.Close() } catch {}
            try { $reader.Dispose() } catch {}
        }

        if ($cmd) {
            try { $cmd.Dispose() } catch {}
        }

        if (
            $ctx -and
            $ctx.connection
        ) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}

function Get-OptionalFieldMasterValues {
    param(
        [int]$VchType,
        [string]$SeriesName,
        [int]$FieldNo,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ($FieldNo -le 0) {
        return @{
            success = $true
            data = @()
        }
    }

    $cacheKey = (
        "{0}|{1}|{2}" -f
        $InstanceId,
        $CompanyCode,
        $FieldNo
    ).ToLowerInvariant()

    if ($script:BusyCloudOptionalFieldValuesCache.ContainsKey($cacheKey)) {
        $entry = $script:BusyCloudOptionalFieldValuesCache[$cacheKey]

        if (
            $entry -and
            $entry.expiresAt -gt [DateTime]::UtcNow
        ) {
            return $entry.result
        }

        try {
            $script:BusyCloudOptionalFieldValuesCache.Remove($cacheKey)
        }
        catch {
        }
    }

    $startedAt = [System.Diagnostics.Stopwatch]::StartNew()
    $ctx = $null
    $reader = $null
    $cmd = $null

    try {
        $ctx = Get-BusyCloudFastConfigDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (
            $null -eq $ctx -or
            $null -eq $ctx.connection
        ) {
            throw "Direct fiscal database connection is unavailable."
        }

        $targetMasterType = 1000 + $FieldNo

        $cmd = $ctx.connection.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}

        $cmd.CommandText = @"
SELECT Name
FROM Master1
WHERE MasterType = $targetMasterType
ORDER BY Name
"@

        $reader = $cmd.ExecuteReader()

        $values = @()

        while ($reader.Read()) {
            try {
                if (-not $reader.IsDBNull(0)) {
                    $value = (
                        [string]$reader.GetValue(0)
                    ).Trim()

                    if (-not [string]::IsNullOrWhiteSpace($value)) {
                        $values += $value
                    }
                }
            }
            catch {
            }
        }

        $result = @{
            success = $true
            data = @($values)
        }

        $script:BusyCloudOptionalFieldValuesCache[$cacheKey] = @{
            expiresAt = [DateTime]::UtcNow.AddMinutes(5)
            result = $result
        }

        $startedAt.Stop()

        Write-Host (
            "  [OPTIONAL-VALUES-FAST] {0}/{1} field={2} db={3} rows={4} elapsedMs={5}" -f
            $InstanceId,
            $CompanyCode,
            $FieldNo,
            [string]$ctx.database,
            $values.Count,
            [int]$startedAt.ElapsedMilliseconds
        ) -ForegroundColor DarkCyan

        return $result
    }
    catch {
        if ($startedAt.IsRunning) {
            $startedAt.Stop()
        }

        Write-Host (
            "  [OPTIONAL-VALUES-FAST FAIL] {0}/{1} field={2} elapsedMs={3} error={4}" -f
            $InstanceId,
            $CompanyCode,
            $FieldNo,
            [int]$startedAt.ElapsedMilliseconds,
            $_.Exception.Message
        ) -ForegroundColor Red

        return @{
            success = $false
            error = $_.Exception.Message
            data = @()
        }
    }
    finally {
        if ($reader) {
            try { $reader.Close() } catch {}
            try { $reader.Dispose() } catch {}
        }

        if ($cmd) {
            try { $cmd.Dispose() } catch {}
        }

        if (
            $ctx -and
            $ctx.connection
        ) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}

# =============================================================================
# END FAST OPTIONAL FIELD READ PATH
# =============================================================================


# =============================================================================
# BUSYCLOUD APPROVAL POLICY V7.2 - INDEPENDENT BUSY + WEB FLAGS
# =============================================================================
# IMPORTANT:
#   RecType=203 / I1 is kept strictly BUSY-native-compatible:
#       I1 = 0 -> BUSY approval OFF
#       I1 = 1 -> BUSY approval ON
#
#   Web Approval is stored separately in BusyCloudWebApprovalPolicy:
#       WebApprovalEnabled = 0/1
#
#   The API still exposes the convenient four-state value:
#       0 NONE, 1 BUSY, 2 WEB, 3 BOTH
#   but value 2/3 is NEVER written into Config.RecType=203.I1.
#
# This separation is required because BUSY-created pending vouchers rely on the
# native approval flag remaining in the value BUSY understands.
# =============================================================================

$script:BusyCloudWebApprovalPolicyTable = "BusyCloudWebApprovalPolicy"

function ConvertTo-VoucherApprovalModeValue {
    param($Value)
    if ($null -eq $Value) { return -1 }
    $text = ([string]$Value).Trim().ToUpperInvariant()
    switch ($text) {
        "0" { return 0 }; "NONE" { return 0 }; "OFF" { return 0 }
        "1" { return 1 }; "BUSY" { return 1 }; "BUSY_APPROVAL" { return 1 }
        "2" { return 2 }; "WEB" { return 2 }; "WEB_APPROVAL" { return 2 }
        "3" { return 3 }; "BOTH" { return 3 }; "WEB_AND_BUSY" { return 3 }; "WEB_BUSY" { return 3 }
        default { return -1 }
    }
}

function Get-VoucherApprovalModeName {
    param([int]$Mode)
    switch ($Mode) {
        1 { return "BUSY" }
        2 { return "WEB" }
        3 { return "BOTH" }
        default { return "NONE" }
    }
}

function Ensure-BusyCloudWebApprovalPolicyTable {
    param($Context)
    if ($null -eq $Context -or $null -eq $Context.connection) {
        throw "Approval policy database context is unavailable."
    }

    $conn = $Context.connection
    $dbType = [int]$Context.dbType

    if ($dbType -eq 1) {
        $cmd = $conn.CreateCommand()
        try {
            $cmd.CommandText = @"
IF OBJECT_ID(N'dbo.BusyCloudWebApprovalPolicy', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.BusyCloudWebApprovalPolicy (
        VoucherType INT NOT NULL PRIMARY KEY,
        WebApprovalEnabled INT NOT NULL CONSTRAINT DF_BCWAP_WebApprovalEnabled DEFAULT(0),
        UpdatedAt DATETIME2 NULL
    )
END
"@
            [void]$cmd.ExecuteNonQuery()
        }
        finally { try { $cmd.Dispose() } catch {} }
        return
    }

    $exists = $true
    $probe = $conn.CreateCommand()
    try {
        $probe.CommandText = "SELECT TOP 1 VoucherType FROM BusyCloudWebApprovalPolicy"
        [void]$probe.ExecuteScalar()
    }
    catch { $exists = $false }
    finally { try { $probe.Dispose() } catch {} }

    if (-not $exists) {
        $cmd = $conn.CreateCommand()
        try {
            $cmd.CommandText = @"
CREATE TABLE [BusyCloudWebApprovalPolicy] (
    [VoucherType] LONG NOT NULL,
    [WebApprovalEnabled] LONG NOT NULL,
    [UpdatedAt] DATETIME,
    CONSTRAINT [PK_BusyCloudWebApprovalPolicy] PRIMARY KEY ([VoucherType])
)
"@
            [void]$cmd.ExecuteNonQuery()
        }
        finally { try { $cmd.Dispose() } catch {} }
    }
}

function Get-BusyCloudWebApprovalEnabledFromContext {
    param($Context, [int]$VchType)
    Ensure-BusyCloudWebApprovalPolicyTable -Context $Context
    $table = if ([int]$Context.dbType -eq 1) { "dbo.BusyCloudWebApprovalPolicy" } else { "BusyCloudWebApprovalPolicy" }
    $cmd = $Context.connection.CreateCommand()
    try {
        $cmd.CommandText = "SELECT TOP 1 WebApprovalEnabled FROM $table WHERE VoucherType=$VchType"
        $raw = $cmd.ExecuteScalar()
        if ($null -eq $raw -or $raw -eq [System.DBNull]::Value) { return $false }
        try { return ([int]$raw -eq 1) } catch { return $false }
    }
    finally { try { $cmd.Dispose() } catch {} }
}

function Set-BusyCloudWebApprovalEnabledInContext {
    param($Context, $Transaction, [int]$VchType, [bool]$Enabled)
    # The caller ensures the table exists before starting a transaction.
    # SQL Server commands created while a transaction is active must be
    # attached to that transaction, so do not run DDL/probe commands here.
    if ($null -eq $Transaction) {
        Ensure-BusyCloudWebApprovalPolicyTable -Context $Context
    }
    $table = if ([int]$Context.dbType -eq 1) { "dbo.BusyCloudWebApprovalPolicy" } else { "BusyCloudWebApprovalPolicy" }
    $enabledInt = if ($Enabled) { 1 } else { 0 }
    $cmd = $Context.connection.CreateCommand()
    if ($null -ne $Transaction) { $cmd.Transaction = $Transaction }
    try {
        $cmd.CommandText = "SELECT COUNT(*) FROM $table WHERE VoucherType=$VchType"
        $exists = ([int]$cmd.ExecuteScalar() -gt 0)
    }
    finally { try { $cmd.Dispose() } catch {} }

    $cmd = $Context.connection.CreateCommand()
    if ($null -ne $Transaction) { $cmd.Transaction = $Transaction }
    try {
        if ([int]$Context.dbType -eq 1) {
            if ($exists) {
                $cmd.CommandText = "UPDATE $table SET WebApprovalEnabled=$enabledInt, UpdatedAt=SYSUTCDATETIME() WHERE VoucherType=$VchType"
            } else {
                $cmd.CommandText = "INSERT INTO $table (VoucherType,WebApprovalEnabled,UpdatedAt) VALUES ($VchType,$enabledInt,SYSUTCDATETIME())"
            }
        } else {
            if ($exists) {
                $cmd.CommandText = "UPDATE $table SET WebApprovalEnabled=$enabledInt, UpdatedAt=Now() WHERE VoucherType=$VchType"
            } else {
                $cmd.CommandText = "INSERT INTO $table (VoucherType,WebApprovalEnabled,UpdatedAt) VALUES ($VchType,$enabledInt,Now())"
            }
        }
        [void]$cmd.ExecuteNonQuery()
    }
    finally { try { $cmd.Dispose() } catch {} }
}

function Get-VoucherApprovalConfig {
    param(
        [int]$VchType,
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        $ExistingFi = $null
    )

    if (-not (Test-IsBusyCloudApprovalVoucherType -VchType $VchType)) {
        return @{ success=$false; error="Voucher type $VchType is not supported by BusyCloud approval processing." }
    }

    $ctx = $null
    $rdr = $null
    try {
        $ctx = Get-BusyCloudFastConfigDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        Ensure-BusyCloudWebApprovalPolicyTable -Context $ctx
        $conn = $ctx.connection

        $nativeBusyValue = 0
        $cmd = $conn.CreateCommand()
        try {
            $cmd.CommandText = "SELECT TOP 1 I1 FROM Config WHERE RecType=203 AND [Type]=$VchType"
            $raw = $cmd.ExecuteScalar()
            if ($null -ne $raw -and $raw -ne [System.DBNull]::Value) {
                try { $nativeBusyValue = [int]$raw } catch { $nativeBusyValue = 0 }
            }
        }
        finally { try { $cmd.Dispose() } catch {} }

        # One-time migration from the earlier four-state experiment.
        # 2 -> WEB only: native BUSY flag becomes 0, web flag becomes 1.
        # 3 -> BOTH:     native BUSY flag becomes 1, web flag becomes 1.
        if ($nativeBusyValue -in @(2,3)) {
            $legacyMode = $nativeBusyValue
            $busyValue = if ($legacyMode -eq 3) { 1 } else { 0 }
            $webEnabled = $true
            $tx = $conn.BeginTransaction()
            try {
                $cmd = $conn.CreateCommand(); $cmd.Transaction = $tx
                $cmd.CommandText = "UPDATE Config SET I1=$busyValue WHERE RecType=203 AND [Type]=$VchType"
                [void]$cmd.ExecuteNonQuery(); try { $cmd.Dispose() } catch {}
                Set-BusyCloudWebApprovalEnabledInContext -Context $ctx -Transaction $tx -VchType $VchType -Enabled:$webEnabled
                $tx.Commit(); $tx = $null
                $nativeBusyValue = $busyValue
                Write-Host "  [APPROVAL-POLICY MIGRATION] type=$VchType legacyMode=$legacyMode -> busy=$busyValue web=1" -ForegroundColor Yellow
            }
            catch {
                if ($tx) { try { $tx.Rollback() } catch {} }
                throw
            }
        }

        $busyEnabled = ($nativeBusyValue -eq 1)
        $webEnabled = Get-BusyCloudWebApprovalEnabledFromContext -Context $ctx -VchType $VchType
        $mode = if ($busyEnabled -and $webEnabled) { 3 } elseif ($webEnabled) { 2 } elseif ($busyEnabled) { 1 } else { 0 }

        $approvers = @()
        if ($busyEnabled) {
            $cmd = $conn.CreateCommand()
            try {
                $cmd.CommandText = "SELECT C1 FROM Config WHERE RecType=204 AND [Type]=$VchType AND I1=1"
                $rdr = $cmd.ExecuteReader()
                while ($rdr.Read()) {
                    $name = ""
                    if (-not $rdr.IsDBNull(0)) { $name = ([string]$rdr.GetValue(0)).Trim() }
                    if ($name -and $approvers -notcontains $name) { $approvers += $name }
                }
                try { $rdr.Close() } catch {}; $rdr = $null
            }
            finally { try { $cmd.Dispose() } catch {} }
        }

        return @{ success=$true; data=@{
            vch_type=$VchType
            approval_mode_value=$mode
            approval_mode=(Get-VoucherApprovalModeName -Mode $mode)
            approval_required=[bool]$busyEnabled
            web_approval_required=[bool]$webEnabled
            approvers=@($approvers | Sort-Object)
        }}
    }
    catch { return @{ success=$false; error=$_.Exception.Message } }
    finally {
        if ($rdr) { try { $rdr.Close() } catch {}; try { $rdr.Dispose() } catch {} }
        if ($ctx -and $ctx.connection) { try { $ctx.connection.Close() } catch {}; try { $ctx.connection.Dispose() } catch {} }
    }
}

function Save-VoucherApprovalConfig {
    param($Data,[string]$InstanceId="",[string]$CompanyCode="")
    $vchType = 0; try { $vchType=[int]$Data.vch_type } catch {}
    if (-not (Test-IsBusyCloudApprovalVoucherType -VchType $vchType)) {
        return @{ success=$false; error="Unsupported or missing vch_type." }
    }

    $mode = -1
    if ($null -ne $Data.approval_mode_value) { $mode=ConvertTo-VoucherApprovalModeValue $Data.approval_mode_value }
    if ($mode -lt 0 -and $null -ne $Data.approval_mode) { $mode=ConvertTo-VoucherApprovalModeValue $Data.approval_mode }
    if ($mode -lt 0) {
        $raw=([string]$Data.approval_required).Trim().ToLowerInvariant()
        $mode=if($Data.approval_required -eq $true -or $raw -in @('1','true')){1}else{0}
    }
    if ($mode -notin @(0,1,2,3)) { return @{success=$false;error="approval_mode must be NONE, BUSY, WEB or BOTH."} }

    $busyEnabled = ($mode -in @(1,3))
    $webEnabled = ($mode -in @(2,3))
    if ($webEnabled -and (Get-Command Test-WebApprovalSupportedVoucherType -ErrorAction SilentlyContinue) -and -not (Test-WebApprovalSupportedVoucherType -VchType $vchType)) {
        return @{success=$false;error="Web Approval is supported only for Sales Quotation, Sale Order, Sale, Sale Return, Receipt and Delivery Order."}
    }

    $approvers=@()
    if ($busyEnabled) {
        $requested=@(@($Data.approvers)|ForEach-Object{([string]$_).Trim()}|Where-Object{$_}|Sort-Object -Unique)
        $users=Get-CompanyUsers -InstanceId $InstanceId -CompanyCode $CompanyCode
        if(-not $users.success){return @{success=$false;error="Could not validate approvers against BUSY users. $($users.error)"}}
        $map=@{}; foreach($u in @($users.data)){ $n=([string]$u).Trim(); if($n){$map[$n.ToLowerInvariant()]=$n} }
        $unknown=@(); foreach($r in $requested){$k=$r.ToLowerInvariant();if($map.ContainsKey($k)){$n=[string]$map[$k];if($approvers -notcontains $n){$approvers+=$n}}else{$unknown+=$r}}
        if($unknown.Count -gt 0){return @{success=$false;error=("Unknown BUSY user(s): "+($unknown -join ', '))}}
    }

    $ctx=$null;$tx=$null
    try {
        $ctx=Get-BusyCloudFastConfigDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        Ensure-BusyCloudWebApprovalPolicyTable -Context $ctx
        $conn=$ctx.connection;$tx=$conn.BeginTransaction()

        # CRITICAL: BUSY-native flag remains 0/1 only.
        $busyInt=if($busyEnabled){1}else{0}
        $cmd=$conn.CreateCommand();$cmd.Transaction=$tx
        $cmd.CommandText="SELECT COUNT(*) FROM Config WHERE RecType=203 AND [Type]=$vchType"
        $exists=([int]$cmd.ExecuteScalar() -gt 0);try{$cmd.Dispose()}catch{}
        $cmd=$conn.CreateCommand();$cmd.Transaction=$tx
        $cmd.CommandText=if($exists){"UPDATE Config SET I1=$busyInt WHERE RecType=203 AND [Type]=$vchType"}else{"INSERT INTO Config (RecType,[Type],I1) VALUES (203,$vchType,$busyInt)"}
        [void]$cmd.ExecuteNonQuery();try{$cmd.Dispose()}catch{}

        Set-BusyCloudWebApprovalEnabledInContext -Context $ctx -Transaction $tx -VchType $vchType -Enabled:$webEnabled

        $cmd=$conn.CreateCommand();$cmd.Transaction=$tx;$cmd.CommandText="DELETE FROM Config WHERE RecType=204 AND [Type]=$vchType";[void]$cmd.ExecuteNonQuery();try{$cmd.Dispose()}catch{}
        if($busyEnabled){foreach($name in $approvers){$safe=$name.Replace("'","''");$cmd=$conn.CreateCommand();$cmd.Transaction=$tx;$cmd.CommandText="INSERT INTO Config (RecType,[Type],I1,C1) VALUES (204,$vchType,1,'$safe')";[void]$cmd.ExecuteNonQuery();try{$cmd.Dispose()}catch{}}}

        $tx.Commit();$tx=$null
        return @{success=$true;message="Voucher approval settings saved successfully.";data=@{
            vch_type=$vchType;approval_mode_value=$mode;approval_mode=(Get-VoucherApprovalModeName $mode);approval_required=[bool]$busyEnabled;web_approval_required=[bool]$webEnabled;approvers=@($approvers)
        }}
    }
    catch { if($tx){try{$tx.Rollback()}catch{}};return @{success=$false;error=$_.Exception.Message} }
    finally { if($ctx -and $ctx.connection){try{$ctx.connection.Close()}catch{};try{$ctx.connection.Dispose()}catch{}} }
}

# Final overrides: native BUSY approval-enabled type discovery must use only the
# BUSY-native flag. Web-only types must never appear in the BUSY approval inbox.
function Get-EnabledVoucherApprovalTypes-Direct {
    param([string]$InstanceId="",[string]$CompanyCode="")
    $ctx=$null
    try{$ctx=Get-BusyCloudApprovalDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode;$cmd=$ctx.connection.CreateCommand();$cmd.CommandText="SELECT DISTINCT [Type] FROM Config WHERE RecType=203 AND I1=1";$r=$cmd.ExecuteReader();$types=@();while($r.Read()){$t=0;try{$t=[int]$r.GetValue(0)}catch{};if((Test-IsBusyCloudApprovalVoucherType $t)-and $types -notcontains $t){$types+=$t}};try{$r.Close()}catch{};return @{success=$true;data=@($types|Sort-Object)}}catch{return @{success=$false;error=$_.Exception.Message}}finally{if($ctx -and $ctx.connection){try{$ctx.connection.Close()}catch{};try{$ctx.connection.Dispose()}catch{}}}
}
function Get-EnabledVoucherApprovalTypes-AccessCom {
    param([string]$InstanceId="",[string]$CompanyCode="",$ExistingFi=$null)
    $fi=$ExistingFi;$owns=$false;if(-not $fi){$fi=Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode;$owns=$true};if(-not $fi){return @{success=$false;error="BUSY database connection failed"}}
    try{$rst=$fi.GetRecordset("SELECT [Type] FROM Config WHERE RecType=203 AND I1=1");$types=@();if($rst){while(-not $rst.EOF){$t=0;try{$t=[int]$rst.Fields.Item('Type').Value}catch{};if((Test-IsBusyCloudApprovalVoucherType $t)-and $types -notcontains $t){$types+=$t};$rst.MoveNext()};try{$rst.Close()}catch{}};return @{success=$true;data=@($types|Sort-Object)}}catch{return @{success=$false;error=$_.Exception.Message}}finally{if($owns -and $fi){Disconnect-BUSY $fi}}
}
function Get-EnabledVoucherApprovalTypes {
    param([string]$InstanceId="",[string]$CompanyCode="",$ExistingFi=$null)
    if(Test-BusyCloudApprovalUsesAccessCom -InstanceId $InstanceId -CompanyCode $CompanyCode){return Get-EnabledVoucherApprovalTypes-AccessCom -InstanceId $InstanceId -CompanyCode $CompanyCode -ExistingFi $ExistingFi}
    return Get-EnabledVoucherApprovalTypes-Direct -InstanceId $InstanceId -CompanyCode $CompanyCode
}

Write-Host "  [APPROVAL-POLICY] Independent BUSY/Web flags v7.2 loaded." -ForegroundColor DarkCyan
