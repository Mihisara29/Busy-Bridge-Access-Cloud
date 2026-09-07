# modules/routes.ps1
# HTTP Route Handlers — Targeted Native Auth & Audit Trail Integration

function Read-RequestBody {
    param($request)
    $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
    return $reader.ReadToEnd()
}

function Send-Response {
    param($response, $data, $statusCode = $null)
    try {
        if ($null -ne $statusCode) { $response.StatusCode = $statusCode }
        $json   = $data | ConvertTo-Json -Depth 10
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
        $response.OutputStream.Write($buffer, 0, $buffer.Length)
        $response.OutputStream.Close()
    } catch { }
}

function Get-QueryStringValue {
    param($queryString, [string]$key, [string]$defaultValue = "")
    $value = $queryString[$key]
    if ($value -and $value -ne "") { return $value }
    return $defaultValue
}

function Get-ServerId {
    return "busy-server-native-unified"
}

# ===============================================================
# VOUCHER PARTY ACCOUNT ACCESS PERMISSIONS
# ===============================================================

function Test-IsPartyAccessVoucherType {
    param([int]$VchType)

    return @(
        9, 26, 12, 11, 3, 14,
        2, 27, 13, 4, 10, 19
    ) -contains $VchType
}

function Test-IsPermissionAdminUser {
    param($User)

    if ($null -eq $User) {
        return $false
    }

    $role = ([string]$User.role).Trim().ToLowerInvariant()
    return ($role -eq "superadmin" -or $role -eq "companyadmin")
}

function Get-PartyGroupAccessForAuthUser {
    param(
        $AuthResult,
        [int]$VchType,
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        [bool]$RequireAuth = $true
    )

    # Trusted bridge-secret/internal calls keep the existing unrestricted path.
    if (-not $RequireAuth) {
        return @{ enforce = $false; groupCodes = @(); allEligibleRoots = $false }
    }

    if (-not (Test-IsPartyAccessVoucherType -VchType $VchType)) {
        return @{ enforce = $false; groupCodes = @(); allEligibleRoots = $false }
    }

    if (
        $null -ne $AuthResult -and
        $null -ne $AuthResult.user -and
        (Test-IsPermissionAdminUser -User $AuthResult.user)
    ) {
        # Admins bypass user branch assignments, but the Party field itself is
        # still limited to the three eligible BUSY roots required by the business.
        return @{ enforce = $true; groupCodes = @(); allEligibleRoots = $true }
    }

    # IMPORTANT:
    # Always read the normal user's CURRENT MobileUserPreference row from the
    # selected BUSY company. Do not trust only the permission snapshot embedded
    # in an existing auth token: the administrator may have changed M2 after the
    # token was issued. Using the live row also makes permission changes effective
    # immediately without requiring a new token.
    $groupCodes = @()
    $m2 = "{}"
    $liveProfileFound = $false

    if (
        $null -ne $AuthResult -and
        $null -ne $AuthResult.user -and
        -not [string]::IsNullOrWhiteSpace([string]$AuthResult.user.name)
    ) {
        try {
            $permissionResult = Get-UserPermissions `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode

            if ($permissionResult.success) {
                $activeName = ([string]$AuthResult.user.name).Trim().ToLowerInvariant()
                $profile = @($permissionResult.data) |
                    Where-Object {
                        ([string]$_.name).Trim().ToLowerInvariant() -eq $activeName
                    } |
                    Select-Object -First 1

                if ($null -ne $profile) {
                    $liveProfileFound = $true

                    if ($null -ne $profile.M2) {
                        $m2 = [string]$profile.M2
                    }
                }
            }
        }
        catch {
            # Fail closed below. We deliberately do not turn a DB read problem
            # into unrestricted Party access.
            $liveProfileFound = $false
            $m2 = "{}"
        }
    }

    # Compatibility fallback only when the live profile genuinely cannot be
    # found. This supports older/native tokens while still preferring the DB.
    if (
        -not $liveProfileFound -and
        $null -ne $AuthResult -and
        $null -ne $AuthResult.user -and
        $null -ne $AuthResult.user.permissions -and
        $null -ne $AuthResult.user.permissions.M2
    ) {
        $m2 = [string]$AuthResult.user.permissions.M2
    }

    try {
        $fieldsMap = $m2 | ConvertFrom-Json
        $prop = $fieldsMap.PSObject.Properties["$VchType"]

        if ($null -ne $prop -and $null -ne $prop.Value) {
            $voucherConfig = $prop.Value

            if ($null -ne $voucherConfig.partyGroupCodes) {
                foreach ($rawCode in @($voucherConfig.partyGroupCodes)) {
                    $code = 0
                    if ([int]::TryParse([string]$rawCode, [ref]$code) -and $code -gt 0) {
                        if ($groupCodes -notcontains $code) {
                            $groupCodes += $code
                        }
                    }
                }
            }
        }
    }
    catch {
        $groupCodes = @()
    }

    return @{
        enforce          = $true
        groupCodes       = @($groupCodes)
        allEligibleRoots = $false
    }
}

function Test-VoucherPartyAccountAccess {
    param(
        $AuthResult,
        $Data,
        [string]$InstanceId,
        [string]$CompanyCode,
        [bool]$RequireAuth = $true
    )

    if (-not $RequireAuth) {
        return @{ success = $true; allowed = $true; deniedAccounts = @() }
    }

    if ($null -eq $Data -or $null -eq $Data.vchType) {
        return @{ success = $true; allowed = $true; deniedAccounts = @() }
    }

    $vchType = [int]$Data.vchType

    if (-not (Test-IsPartyAccessVoucherType -VchType $vchType)) {
        return @{ success = $true; allowed = $true; deniedAccounts = @() }
    }

    $access = Get-PartyGroupAccessForAuthUser `
        -AuthResult $AuthResult `
        -VchType $vchType `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode `
        -RequireAuth $RequireAuth

    $accountNames = @()

    if ($vchType -eq 14 -or $vchType -eq 19) {
        # Receipt: Credit rows are the counterparty side.
        # Payment: Debit rows are the counterparty side.
        foreach ($account in @($Data.accounts)) {
            if ($null -eq $account) { continue }

            $dc = ([string]$account.dc).Trim().ToUpperInvariant()
            $isPartySide =
                ($vchType -eq 14 -and $dc -eq "C") -or
                ($vchType -eq 19 -and $dc -eq "D")

            if ($isPartySide) {
                $name = ([string]$account.accountName).Trim()
                if (-not [string]::IsNullOrWhiteSpace($name)) {
                    $accountNames += $name
                }
            }
        }
    }
    else {
        $partyName = ([string]$Data.party).Trim()
        if (-not [string]::IsNullOrWhiteSpace($partyName)) {
            $accountNames += $partyName
        }
    }

    return Test-PartyAccountsAllowed `
        -AccountNames @($accountNames) `
        -AllowedGroupCodes @($access.groupCodes) `
        -AllowAllEligibleRoots ([bool]$access.allEligibleRoots) `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode
}


# ===============================================================
# JOURNAL / CONTRA DEBIT/CREDIT ACCOUNT ACCESS
# ===============================================================

function Test-IsDebitCreditAccessVoucherType {
    param([int]$VchType)
    return @(15, 16) -contains $VchType
}

function Get-AccountSideAccessForAuthUser {
    param(
        $AuthResult,
        [int]$VchType,
        [string]$Dc,
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        [bool]$RequireAuth = $true
    )

    if (-not $RequireAuth) {
        return @{ enforce = $false; groupCodes = @(); accountCodes = @(); allAccounts = $true }
    }

    if (-not (Test-IsDebitCreditAccessVoucherType -VchType $VchType)) {
        return @{ enforce = $false; groupCodes = @(); accountCodes = @(); allAccounts = $false }
    }

    $side = ([string]$Dc).Trim().ToUpperInvariant()
    if ($side -ne "D" -and $side -ne "C") {
        return @{ enforce = $true; groupCodes = @(); accountCodes = @(); allAccounts = $false; invalidSide = $true }
    }

    if (
        $null -ne $AuthResult -and
        $null -ne $AuthResult.user -and
        (Test-IsPermissionAdminUser -User $AuthResult.user)
    ) {
        return @{ enforce = $true; groupCodes = @(); accountCodes = @(); allAccounts = $true }
    }

    $m2 = "{}"
    $liveProfileFound = $false

    if (
        $null -ne $AuthResult -and
        $null -ne $AuthResult.user -and
        -not [string]::IsNullOrWhiteSpace([string]$AuthResult.user.name)
    ) {
        try {
            $permissionResult = Get-UserPermissions `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode

            if ($permissionResult.success) {
                $activeName = ([string]$AuthResult.user.name).Trim().ToLowerInvariant()
                $profile = @($permissionResult.data) |
                    Where-Object { ([string]$_.name).Trim().ToLowerInvariant() -eq $activeName } |
                    Select-Object -First 1

                if ($null -ne $profile) {
                    $liveProfileFound = $true
                    if ($null -ne $profile.M2) { $m2 = [string]$profile.M2 }
                }
            }
        }
        catch {
            $liveProfileFound = $false
            $m2 = "{}"
        }
    }

    if (
        -not $liveProfileFound -and
        $null -ne $AuthResult -and
        $null -ne $AuthResult.user -and
        $null -ne $AuthResult.user.permissions -and
        $null -ne $AuthResult.user.permissions.M2
    ) {
        $m2 = [string]$AuthResult.user.permissions.M2
    }

    $groupCodes = @()

    try {
        $fieldsMap = $m2 | ConvertFrom-Json
        $prop = $fieldsMap.PSObject.Properties["$VchType"]

        if ($null -ne $prop -and $null -ne $prop.Value) {
            $voucherConfig = $prop.Value
            $groupProperty = if ($side -eq "D") { "debitGroupCodes" } else { "creditGroupCodes" }

            $groupProp = $voucherConfig.PSObject.Properties[$groupProperty]
            if ($null -ne $groupProp -and $null -ne $groupProp.Value) {
                foreach ($rawCode in @($groupProp.Value)) {
                    $code = 0
                    if ([int]::TryParse([string]$rawCode, [ref]$code) -and $code -gt 0 -and $groupCodes -notcontains $code) {
                        $groupCodes += $code
                    }
                }
            }
        }
    }
    catch {
        $groupCodes = @()
    }

    # V4 Journal/Contra policy is GROUP-ONLY. Legacy debitAccountCodes /
    # creditAccountCodes values are intentionally ignored server-side so an
    # old V3 profile cannot continue granting an individual ledger.
    return @{
        enforce      = $true
        groupCodes   = @($groupCodes)
        accountCodes = @()
        allAccounts  = $false
    }
}

function Test-VoucherDebitCreditAccountAccess {
    param(
        $AuthResult,
        $Data,
        [string]$InstanceId,
        [string]$CompanyCode,
        [bool]$RequireAuth = $true
    )

    if (-not $RequireAuth -or $null -eq $Data -or $null -eq $Data.vchType) {
        return @{ success = $true; allowed = $true; deniedAccounts = @() }
    }

    $vchType = [int]$Data.vchType
    if (-not (Test-IsDebitCreditAccessVoucherType -VchType $vchType)) {
        return @{ success = $true; allowed = $true; deniedAccounts = @() }
    }

    $debitNames = @()
    $creditNames = @()
    $invalidSideAccounts = @()

    foreach ($account in @($Data.accounts)) {
        if ($null -eq $account) { continue }
        $name = ([string]$account.accountName).Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { continue }

        $dc = ([string]$account.dc).Trim().ToUpperInvariant()
        if ($dc -eq "D") { $debitNames += $name }
        elseif ($dc -eq "C") { $creditNames += $name }
        else { $invalidSideAccounts += $name }
    }

    if ($invalidSideAccounts.Count -gt 0) {
        return @{ success=$true; allowed=$false; deniedAccounts=@($invalidSideAccounts); deniedDebitAccounts=@(); deniedCreditAccounts=@() }
    }

    $debitAccess = Get-AccountSideAccessForAuthUser `
        -AuthResult $AuthResult -VchType $vchType -Dc "D" `
        -InstanceId $InstanceId -CompanyCode $CompanyCode -RequireAuth $RequireAuth

    $creditAccess = Get-AccountSideAccessForAuthUser `
        -AuthResult $AuthResult -VchType $vchType -Dc "C" `
        -InstanceId $InstanceId -CompanyCode $CompanyCode -RequireAuth $RequireAuth

    $debitCheck = Test-AccountsAllowedByCodes `
        -AccountNames @($debitNames) `
        -AllowedGroupCodes @($debitAccess.groupCodes) `
        -AllowedAccountCodes @($debitAccess.accountCodes) `
        -AllowAllAccounts ([bool]$debitAccess.allAccounts) `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $debitCheck.success) {
        return @{ success=$false; allowed=$false; error=$debitCheck.error; deniedAccounts=@($debitCheck.deniedAccounts) }
    }

    $creditCheck = Test-AccountsAllowedByCodes `
        -AccountNames @($creditNames) `
        -AllowedGroupCodes @($creditAccess.groupCodes) `
        -AllowedAccountCodes @($creditAccess.accountCodes) `
        -AllowAllAccounts ([bool]$creditAccess.allAccounts) `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $creditCheck.success) {
        return @{ success=$false; allowed=$false; error=$creditCheck.error; deniedAccounts=@($creditCheck.deniedAccounts) }
    }

    $denied = @($debitCheck.deniedAccounts) + @($creditCheck.deniedAccounts)
    return @{
        success              = $true
        allowed              = ($denied.Count -eq 0)
        deniedAccounts       = @($denied)
        deniedDebitAccounts  = @($debitCheck.deniedAccounts)
        deniedCreditAccounts = @($creditCheck.deniedAccounts)
    }
}


# ===============================================================
# VOUCHER ITEM GROUP ACCESS PERMISSIONS
# ===============================================================

function Test-IsItemGroupAccessVoucherType {
    param([int]$VchType)

    return @(
        9, 2, 12, 13, 26, 27, 3, 10, 11, 4
    ) -contains $VchType
}

function Get-LiveItemPermissionM2Fast {
    param(
        [string]$UserName = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $trimmedUserName = ([string]$UserName).Trim()
    if ([string]::IsNullOrWhiteSpace($trimmedUserName)) {
        return @{ success = $true; found = $false; m2 = "{}" }
    }

    $conn = $null

    try {
        $found = Get-InstanceForCompany `
            -CompanyCode $CompanyCode `
            -InstanceId $InstanceId

        if (-not $found) {
            return @{
                success = $false
                found   = $false
                m2      = "{}"
                error   = "Company not found in instances.json"
            }
        }

        $inst = $found.instance
        $dbType = if ($null -ne $inst.dbType) { [int]$inst.dbType } else { 0 }

        if ($dbType -eq 1) {
            $sqlDb = Get-SqlDatabaseName `
                -CompanyCode $CompanyCode `
                -InstanceId $inst.id

            $conn = Open-SqlConnection `
                -SqlServer $inst.sqlServer `
                -Database $sqlDb `
                -SqlUser $inst.sqlUser `
                -SqlPassword $inst.sqlPassword

            $cmd = $conn.CreateCommand()
            $cmd.CommandText = @"
SELECT TOP 1
    CASE
        WHEN M2 IS NULL THEN '{}'
        ELSE CAST(M2 AS NVARCHAR(MAX))
    END
FROM MobileUserPreference
WHERE [Name] = @userName
"@
            $cmd.Parameters.AddWithValue("@userName", $trimmedUserName) | Out-Null

            $rawM2 = $cmd.ExecuteScalar()

            if ($null -eq $rawM2 -or $rawM2 -eq [System.DBNull]::Value) {
                return @{ success = $true; found = $false; m2 = "{}" }
            }

            return @{
                success = $true
                found   = $true
                m2      = [string]$rawM2
            }
        }

        $permissionResult = Get-UserPermissions `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (-not $permissionResult.success) {
            return @{
                success = $false
                found   = $false
                m2      = "{}"
                error   = $permissionResult.error
            }
        }

        $activeName = $trimmedUserName.ToLowerInvariant()
        $profile = @($permissionResult.data) |
            Where-Object {
                ([string]$_.name).Trim().ToLowerInvariant() -eq $activeName
            } |
            Select-Object -First 1

        if ($null -eq $profile) {
            return @{ success = $true; found = $false; m2 = "{}" }
        }

        $m2 = "{}"
        if ($null -ne $profile.M2) {
            $m2 = [string]$profile.M2
        }

        return @{ success = $true; found = $true; m2 = $m2 }
    }
    catch {
        return @{
            success = $false
            found   = $false
            m2      = "{}"
            error   = $_.Exception.Message
        }
    }
    finally {
        if ($null -ne $conn) {
            try { $conn.Close() } catch {}
        }
    }
}


function Get-NewPermissionProfileDefaultAccess {
    param(
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    # NEW profiles start with full branch access for:
    # - Party Account Groups
    # - Item Groups
    # - Journal/Contra Debit Account Groups
    # - Journal/Contra Credit Account Groups
    #
    # Voucher Add/Modify permissions remain separate and are not enabled here.

    $partyResult = Get-PartyAccountGroups `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $partyResult.success) {
        return @{
            success = $false
            error = "Could not load Party Account Group roots: $($partyResult.error)"
        }
    }

    $itemResult = Get-AllItemGroupPermissionNodes `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $itemResult.success) {
        return @{
            success = $false
            error = "Could not load Item Group roots: $($itemResult.error)"
        }
    }

    $accountResult = Get-AllAccountPermissionNodes `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $accountResult.success) {
        return @{
            success = $false
            error = "Could not load Chart of Accounts roots: $($accountResult.error)"
        }
    }

    $partyRootCodes = @(
        @($partyResult.data) |
        ForEach-Object { [int]$_.rootCode } |
        Where-Object { $_ -gt 0 } |
        Sort-Object -Unique
    )

    $itemRootCodes = @(
        @($itemResult.data) |
        ForEach-Object { [int]$_.rootCode } |
        Where-Object { $_ -gt 0 } |
        Sort-Object -Unique
    )

    $accountRootCodes = @(
        @($accountResult.data) |
        ForEach-Object { [int]$_.rootCode } |
        Where-Object { $_ -gt 0 } |
        Sort-Object -Unique
    )

    if ($partyRootCodes.Count -eq 0) {
        return @{ success = $false; error = "No eligible Party Account Group roots were found." }
    }

    if ($itemRootCodes.Count -eq 0) {
        return @{ success = $false; error = "No Item Group roots were found." }
    }

    if ($accountRootCodes.Count -eq 0) {
        return @{ success = $false; error = "No Chart of Accounts roots were found." }
    }

    $permissionMap = [ordered]@{}

    $partyVoucherTypes = @(9, 26, 12, 11, 3, 14, 2, 27, 13, 4, 10, 19)
    $itemVoucherTypes = @(9, 2, 12, 13, 26, 27, 3, 10, 11, 4)
    $debitCreditVoucherTypes = @(15, 16)

    foreach ($vchType in $partyVoucherTypes) {
        $key = [string]$vchType

        if (-not $permissionMap.Contains($key)) {
            $permissionMap[$key] = [ordered]@{}
        }

        $permissionMap[$key]["partyGroupCodes"] = @($partyRootCodes)
    }

    foreach ($vchType in $itemVoucherTypes) {
        $key = [string]$vchType

        if (-not $permissionMap.Contains($key)) {
            $permissionMap[$key] = [ordered]@{}
        }

        $permissionMap[$key]["itemGroupCodes"] = @($itemRootCodes)
    }

    foreach ($vchType in $debitCreditVoucherTypes) {
        $key = [string]$vchType

        if (-not $permissionMap.Contains($key)) {
            $permissionMap[$key] = [ordered]@{}
        }

        $permissionMap[$key]["debitGroupCodes"] = @($accountRootCodes)
        $permissionMap[$key]["creditGroupCodes"] = @($accountRootCodes)

        # Current Journal/Contra model is GROUP-ONLY.
        $permissionMap[$key]["debitAccountCodes"] = @()
        $permissionMap[$key]["creditAccountCodes"] = @()
    }

    $m2Json = $permissionMap | ConvertTo-Json -Depth 20 -Compress

    Write-Host (
        "   [PERMISSION DEFAULT ACCESS] company={0} partyRoots={1} itemRoots={2} accountRoots={3}" -f `
        $CompanyCode,
        $partyRootCodes.Count,
        $itemRootCodes.Count,
        $accountRootCodes.Count
    ) -ForegroundColor Green

    return @{
        success = $true
        data = @{
            M2 = $m2Json
            partyRootCodes = @($partyRootCodes)
            itemRootCodes = @($itemRootCodes)
            accountRootCodes = @($accountRootCodes)
        }
    }
}

function Get-ItemGroupAccessForAuthUser {
    param(
        $AuthResult,
        [int]$VchType,
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        [bool]$RequireAuth = $true
    )

    # Trusted bridge-secret/internal calls keep the existing unrestricted path.
    if (-not $RequireAuth) {
        return @{ enforce = $false; groupCodes = @(); allItems = $true }
    }

    if (-not (Test-IsItemGroupAccessVoucherType -VchType $VchType)) {
        return @{ enforce = $false; groupCodes = @(); allItems = $false }
    }

    if (
        $null -ne $AuthResult -and
        $null -ne $AuthResult.user -and
        (Test-IsPermissionAdminUser -User $AuthResult.user)
    ) {
        return @{ enforce = $true; groupCodes = @(); allItems = $true }
    }

    # Read the CURRENT MobileUserPreference.M2 row so changes made by an
    # administrator take effect immediately without requiring a new login token.
    #
    # SQL Server hot path: read only this user's M2 instead of loading all
    # MobileUserPreference rows on every item-search request.
    $groupCodes = @()
    $m2 = "{}"
    $liveProfileFound = $false

    if (
        $null -ne $AuthResult -and
        $null -ne $AuthResult.user -and
        -not [string]::IsNullOrWhiteSpace([string]$AuthResult.user.name)
    ) {
        $liveM2 = Get-LiveItemPermissionM2Fast `
            -UserName ([string]$AuthResult.user.name) `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if ($liveM2.success -and $liveM2.found) {
            $liveProfileFound = $true
            $m2 = [string]$liveM2.m2
        }
    }

    # Compatibility fallback for older/native tokens only when the live row
    # genuinely cannot be found.
    if (
        -not $liveProfileFound -and
        $null -ne $AuthResult -and
        $null -ne $AuthResult.user -and
        $null -ne $AuthResult.user.permissions -and
        $null -ne $AuthResult.user.permissions.M2
    ) {
        $m2 = [string]$AuthResult.user.permissions.M2
    }

    try {
        $fieldsMap = $m2 | ConvertFrom-Json
        $prop = $fieldsMap.PSObject.Properties["$VchType"]

        if ($null -ne $prop -and $null -ne $prop.Value) {
            $voucherConfig = $prop.Value
            $groupProp = $voucherConfig.PSObject.Properties["itemGroupCodes"]

            if ($null -ne $groupProp -and $null -ne $groupProp.Value) {
                foreach ($rawCode in @($groupProp.Value)) {
                    $code = 0
                    if (
                        [int]::TryParse([string]$rawCode, [ref]$code) -and
                        $code -gt 0 -and
                        $groupCodes -notcontains $code
                    ) {
                        $groupCodes += $code
                    }
                }
            }
        }
    }
    catch {
        $groupCodes = @()
    }

    return @{
        enforce    = $true
        groupCodes = @($groupCodes)
        allItems   = $false
    }
}

function Get-AllItemGroupPermissionNodes {
    param(
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) {
        return @{ success = $false; error = "BUSY connection failed"; data = @() }
    }

    try {
        # BUSY Item Groups are MasterType=5. Return ALL Item Groups, including
        # currently empty groups, so permissions remain valid when new items are
        # created later under an already-granted branch.
        $rst = $fi.GetRecordset(@"
SELECT
    Code,
    Name,
    Alias,
    ParentGrp
FROM Master1
WHERE MasterType = 5
ORDER BY Name
"@)

        $rows = @()
        $byCode = @{}

        if ($rst -and -not $rst.EOF) {
            try { $rst.MoveFirst() } catch {}

            while (-not $rst.EOF) {
                $code = 0
                $parentCode = 0
                $name = ""
                $alias = ""

                try {
                    $raw = $rst.Fields.Item("Code").Value
                    if ($null -ne $raw -and $raw -ne [System.DBNull]::Value) {
                        $code = [int]$raw
                    }
                } catch {}

                try {
                    $raw = $rst.Fields.Item("ParentGrp").Value
                    if ($null -ne $raw -and $raw -ne [System.DBNull]::Value) {
                        $parentCode = [int]$raw
                    }
                } catch {}

                try {
                    $raw = $rst.Fields.Item("Name").Value
                    if ($null -ne $raw -and $raw -ne [System.DBNull]::Value) {
                        $name = $raw.ToString().Trim()
                    }
                } catch {}

                try {
                    $raw = $rst.Fields.Item("Alias").Value
                    if ($null -ne $raw -and $raw -ne [System.DBNull]::Value) {
                        $alias = $raw.ToString().Trim()
                    }
                } catch {}

                if ($code -gt 0) {
                    $row = [pscustomobject]@{
                        code       = $code
                        name       = $name
                        alias      = $alias
                        parentCode = $parentCode
                    }
                    $rows += $row
                    $byCode[$code] = $row
                }

                $rst.MoveNext()
            }

            try { $rst.Close() } catch {}
        }

        $nodes = @()

        foreach ($row in $rows) {
            $chain = @()
            $current = [int]$row.code
            $visited = @{}

            for ($guard = 0; $guard -lt 100; $guard++) {
                if ($current -le 0 -or -not $byCode.ContainsKey($current)) {
                    break
                }

                if ($visited.ContainsKey($current)) {
                    break
                }

                $visited[$current] = $true
                $chain += $byCode[$current]
                $current = [int]$byCode[$current].parentCode
            }

            $ordered = @()
            for ($i = $chain.Count - 1; $i -ge 0; $i--) {
                $ordered += $chain[$i]
            }

            $pathNames = @($ordered | ForEach-Object { [string]$_.name })
            $rootCode = if ($ordered.Count -gt 0) { [int]$ordered[0].code } else { [int]$row.code }
            $rootName = if ($ordered.Count -gt 0) { [string]$ordered[0].name } else { [string]$row.name }

            $nodes += [pscustomobject]@{
                rootCode  = $rootCode
                rootName  = $rootName
                nodeType  = "GROUP"
                code      = [int]$row.code
                name      = [string]$row.name
                alias     = [string]$row.alias
                parentCode = [int]$row.parentCode
                level     = [Math]::Max(0, $ordered.Count - 1)
                pathText  = ($pathNames -join " > ")
            }
        }

        $sorted = @($nodes | Sort-Object pathText, name)

        return @{
            success = $true
            count   = $sorted.Count
            data    = $sorted
        }
    }
    catch {
        return @{ success = $false; error = $_.Exception.Message; data = @() }
    }
    finally {
        Disconnect-BUSY $fi
    }
}

function Get-ItemPermissionSnapshot {
    param(
        [int[]]$AllowedGroupCodes,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) {
        return @{ success = $false; error = "BUSY connection failed" }
    }

    try {
        $rst = $fi.GetRecordset(@"
SELECT
    Code,
    MasterType,
    Name,
    Alias,
    ParentGrp
FROM Master1
WHERE MasterType = 5 OR MasterType = 6
"@)

        $groups = @()
        $items = @()
        $groupByCode = @{}
        $itemByCode = @{}
        $itemByName = @{}

        if ($rst -and -not $rst.EOF) {
            try { $rst.MoveFirst() } catch {}

            while (-not $rst.EOF) {
                $code = 0
                $masterType = 0
                $parentCode = 0
                $name = ""
                $alias = ""

                try {
                    $raw = $rst.Fields.Item("Code").Value
                    if ($null -ne $raw -and $raw -ne [System.DBNull]::Value) {
                        $code = [int]$raw
                    }
                } catch {}

                try {
                    $raw = $rst.Fields.Item("MasterType").Value
                    if ($null -ne $raw -and $raw -ne [System.DBNull]::Value) {
                        $masterType = [int]$raw
                    }
                } catch {}

                try {
                    $raw = $rst.Fields.Item("ParentGrp").Value
                    if ($null -ne $raw -and $raw -ne [System.DBNull]::Value) {
                        $parentCode = [int]$raw
                    }
                } catch {}

                try {
                    $raw = $rst.Fields.Item("Name").Value
                    if ($null -ne $raw -and $raw -ne [System.DBNull]::Value) {
                        $name = $raw.ToString().Trim()
                    }
                } catch {}

                try {
                    $raw = $rst.Fields.Item("Alias").Value
                    if ($null -ne $raw -and $raw -ne [System.DBNull]::Value) {
                        $alias = $raw.ToString().Trim()
                    }
                } catch {}

                if ($code -gt 0 -and $masterType -eq 5) {
                    $row = [pscustomobject]@{
                        code       = $code
                        parentCode = $parentCode
                        name       = $name
                    }
                    $groups += $row
                    $groupByCode[$code] = $row
                }
                elseif ($code -gt 0 -and $masterType -eq 6) {
                    $row = [pscustomobject]@{
                        code       = $code
                        parentCode = $parentCode
                        name       = $name
                        alias      = $alias
                    }
                    $items += $row
                    $itemByCode[$code] = $row

                    $nameKey = $name.Trim().ToLowerInvariant()
                    if (-not [string]::IsNullOrWhiteSpace($nameKey)) {
                        $itemByName[$nameKey] = $row
                    }
                }

                $rst.MoveNext()
            }

            try { $rst.Close() } catch {}
        }

        $allowedGroupMap = @{}

        foreach ($rawCode in @($AllowedGroupCodes)) {
            $code = 0
            if (
                [int]::TryParse([string]$rawCode, [ref]$code) -and
                $code -gt 0 -and
                $groupByCode.ContainsKey($code)
            ) {
                $allowedGroupMap[$code] = $true
            }
        }

        # Expand every selected Item Group recursively.
        $changed = $true
        while ($changed) {
            $changed = $false

            foreach ($group in $groups) {
                $code = [int]$group.code
                $parentCode = [int]$group.parentCode

                if (
                    -not $allowedGroupMap.ContainsKey($code) -and
                    $allowedGroupMap.ContainsKey($parentCode)
                ) {
                    $allowedGroupMap[$code] = $true
                    $changed = $true
                }
            }
        }

        $allowedItemCodeMap = @{}

        foreach ($item in $items) {
            if ($allowedGroupMap.ContainsKey([int]$item.parentCode)) {
                $allowedItemCodeMap[[int]$item.code] = $true
            }
        }

        return @{
            success            = $true
            groupByCode        = $groupByCode
            itemByCode         = $itemByCode
            itemByName         = $itemByName
            allowedGroupMap    = $allowedGroupMap
            allowedItemCodeMap = $allowedItemCodeMap
        }
    }
    catch {
        return @{ success = $false; error = $_.Exception.Message }
    }
    finally {
        Disconnect-BUSY $fi
    }
}

function Test-ItemsAllowedByGroupCodes {
    param(
        $Items,
        [int[]]$AllowedGroupCodes,
        [bool]$AllowAllItems = $false,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ($AllowAllItems) {
        return @{ success = $true; allowed = $true; deniedItems = @() }
    }

    $submittedItems = @($Items)
    if ($submittedItems.Count -eq 0) {
        return @{ success = $true; allowed = $true; deniedItems = @() }
    }

    if (@($AllowedGroupCodes).Count -eq 0) {
        $denied = @(
            $submittedItems |
            ForEach-Object {
                $name = ([string]$_.itemName).Trim()
                if ([string]::IsNullOrWhiteSpace($name)) { $name = ([string]$_.name).Trim() }
                if ([string]::IsNullOrWhiteSpace($name)) { $name = "Unknown Item" }
                $name
            }
        )

        return @{ success = $true; allowed = $false; deniedItems = @($denied) }
    }

    $snapshot = Get-ItemPermissionSnapshot `
        -AllowedGroupCodes @($AllowedGroupCodes) `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $snapshot.success) {
        return @{
            success     = $false
            allowed     = $false
            error       = $snapshot.error
            deniedItems = @()
        }
    }

    $deniedItems = @()

    foreach ($item in $submittedItems) {
        if ($null -eq $item) { continue }

        $itemCode = 0

        foreach ($propName in @("itemCode", "code", "masterCode")) {
            try {
                $raw = $item.PSObject.Properties[$propName]
                if ($null -ne $raw -and $null -ne $raw.Value) {
                    $candidate = 0
                    if (
                        [int]::TryParse([string]$raw.Value, [ref]$candidate) -and
                        $candidate -gt 0
                    ) {
                        $itemCode = $candidate
                        break
                    }
                }
            }
            catch {}
        }

        $itemName = ([string]$item.itemName).Trim()
        if ([string]::IsNullOrWhiteSpace($itemName)) {
            $itemName = ([string]$item.name).Trim()
        }

        $resolved = $null

        if ($itemCode -gt 0 -and $snapshot.itemByCode.ContainsKey($itemCode)) {
            $resolved = $snapshot.itemByCode[$itemCode]
        }
        elseif (-not [string]::IsNullOrWhiteSpace($itemName)) {
            $nameKey = $itemName.ToLowerInvariant()
            if ($snapshot.itemByName.ContainsKey($nameKey)) {
                $resolved = $snapshot.itemByName[$nameKey]
                $itemCode = [int]$resolved.code
            }
        }

        if (
            $null -eq $resolved -or
            -not $snapshot.allowedItemCodeMap.ContainsKey([int]$resolved.code)
        ) {
            if ([string]::IsNullOrWhiteSpace($itemName) -and $itemCode -gt 0) {
                $itemName = "Item #$itemCode"
            }
            if ([string]::IsNullOrWhiteSpace($itemName)) {
                $itemName = "Unknown Item"
            }

            if ($deniedItems -notcontains $itemName) {
                $deniedItems += $itemName
            }
        }
    }

    return @{
        success     = $true
        allowed     = ($deniedItems.Count -eq 0)
        deniedItems = @($deniedItems)
    }
}

function Test-VoucherItemGroupAccess {
    param(
        $AuthResult,
        $Data,
        [string]$InstanceId,
        [string]$CompanyCode,
        [bool]$RequireAuth = $true
    )

    if (-not $RequireAuth -or $null -eq $Data -or $null -eq $Data.vchType) {
        return @{ success = $true; allowed = $true; deniedItems = @() }
    }

    $vchType = [int]$Data.vchType

    if (-not (Test-IsItemGroupAccessVoucherType -VchType $vchType)) {
        return @{ success = $true; allowed = $true; deniedItems = @() }
    }

    $access = Get-ItemGroupAccessForAuthUser `
        -AuthResult $AuthResult `
        -VchType $vchType `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode `
        -RequireAuth $RequireAuth

    return Test-ItemsAllowedByGroupCodes `
        -Items @($Data.items) `
        -AllowedGroupCodes @($access.groupCodes) `
        -AllowAllItems ([bool]$access.allItems) `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode
}


function Filter-VoucherItemSearchResultByAccess {
    param(
        $SearchResult,
        [int[]]$AllowedGroupCodes,
        [bool]$AllowAllItems = $false,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ($null -eq $SearchResult) {
        return @{ success = $true; total = 0; data = @() }
    }

    if ($AllowAllItems) {
        return $SearchResult
    }

    if (@($AllowedGroupCodes).Count -eq 0) {
        return @{ success = $true; total = 0; data = @() }
    }

    if ($SearchResult.success -eq $false) {
        return $SearchResult
    }

    $snapshot = Get-ItemPermissionSnapshot `
        -AllowedGroupCodes @($AllowedGroupCodes) `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $snapshot.success) {
        return @{
            success = $false
            error   = $snapshot.error
            total   = 0
            data    = @()
        }
    }

    $sourceData = @()
    try {
        if ($null -ne $SearchResult.data) {
            $sourceData = @($SearchResult.data)
        }
    }
    catch {}

    $filtered = @()

    foreach ($item in $sourceData) {
        $code = 0
        try {
            $raw = $item.code
            if ($null -ne $raw) {
                [void][int]::TryParse([string]$raw, [ref]$code)
            }
        }
        catch {}

        if ($code -gt 0 -and $snapshot.allowedItemCodeMap.ContainsKey($code)) {
            $filtered += $item
        }
    }

    return @{
        success = $true
        total   = $filtered.Count
        data    = @($filtered)
    }
}


function Start-BUSYServer {
    param(
        [int]$Port = 8081
    )

    $listener = New-Object System.Net.HttpListener
    try {
        $listener.Prefixes.Add("http://*:$Port/")
        Write-Host "  Added wildcard binding: http://*:$Port" -ForegroundColor Green
    } catch {
        Write-Host "  Could not add wildcard - run as Administrator" -ForegroundColor Yellow
        $listener = New-Object System.Net.HttpListener
        $listener.Prefixes.Add("http://localhost:$Port/")
        $listener.Prefixes.Add("http://127.0.0.1:$Port/")
    }

    try {
        $listener.Start()
        Write-Host "  Server started successfully!" -ForegroundColor Green
    } catch {
        Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    $config = Get-Config
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  BUSY 21 Enterprise API Bridge v4.5"   -ForegroundColor Green
    Write-Host "  Targeted Login + Audit Trail Engine"   -ForegroundColor Green
    Write-Host "  Server ID : $(Get-ServerId)"           -ForegroundColor Yellow
    Write-Host "  Port      : $Port"                     -ForegroundColor Yellow
    Write-Host "  Press Ctrl+C to stop."                 -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Cyan

    while ($listener.IsListening) {
        $context  = $listener.GetContext()
        $request  = $context.Request
        $response = $context.Response

        $response.Headers.Add("Content-Type",                 "application/json")
        $response.Headers.Add("Access-Control-Allow-Origin",  "*")
        $response.Headers.Add("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
        $response.Headers.Add("Access-Control-Allow-Headers", "Content-Type, Authorization, X-Bridge-Secret, X-Instance-ID, X-Company-Code, Idempotency-Key")

        if ($request.HttpMethod -eq "OPTIONS") {
            $response.StatusCode = 200
            $response.OutputStream.Close()
            continue
        }

        $path   = $request.Url.LocalPath
        $method = $request.HttpMethod
        Write-Host "[$method] $path" -ForegroundColor Yellow

        try {
            $result = $null
            $authResult = $null

            # ════════════════════════════════════════════════
            #  PUBLIC ENDPOINTS (no auth required)
            # ════════════════════════════════════════════════

            if ($path -eq "/health" -and $method -eq "GET") {
                $result = @{ success = $true; status = "BUSY Bridge is running"; version = "4.5"; serverId = Get-ServerId }
                Send-Response $response $result
                Write-Host "  [OK] $path" -ForegroundColor Green
                continue
            }

            # ── LEGACY TARGETED LOGIN (company required) ──
            if ($path -eq "/auth/login" -and $method -eq "POST") {
                $body = Read-RequestBody $request | ConvertFrom-Json

                if (-not $body.company -or -not $body.username -or -not $body.password) {
                    Send-Response $response @{success=$false; error="company, username, and password required"} 400
                    continue
                }

                $loginResult = Invoke-BusyLogin -CompanyIdentifier $body.company -Username $body.username -Password $body.password

                if ($loginResult.success) {
                    Send-Response $response @{ success=$true; token=$loginResult.token; user=$loginResult.user }
                    Write-Host "  [LOGIN SUCCESS] $($body.username) connected to $($body.company)" -ForegroundColor Green
                } else {
                    Send-Response $response @{ success=$false; error=$loginResult.error } 401
                    Write-Host "  [LOGIN FAIL] $($body.username) - $($loginResult.error)" -ForegroundColor Red
                }
                continue
            }

            # ── SCAN LOGIN (auto-detect company from credentials) ──
            if ($path -eq "/auth/scan-login" -and $method -eq "POST") {
                $body = Read-RequestBody $request | ConvertFrom-Json

                if (-not $body.username) {
                    Send-Response $response @{ success=$false; error="username is required" } 400
                    continue
                }

                $scanResult = Invoke-ScanLogin `
                    -Username      $body.username `
                    -Password      ([string]$body.password)

                if ($scanResult.success) {
                    Send-Response $response @{ success=$true; token=$scanResult.token; user=$scanResult.user }
                    Write-Host "  [SCAN LOGIN SUCCESS] $($body.username)" -ForegroundColor Green
                } else {
                    $status = if ($scanResult.error -eq "AMBIGUOUS_CREDENTIALS") { 409 } else { 401 }
                    Send-Response $response @{ success=$false; error=$scanResult.error; detail=$scanResult.detail } $status
                    Write-Host "  [SCAN LOGIN FAIL] $($body.username) - $($scanResult.error)" -ForegroundColor Red
                }
                continue
            }

            if ($path -eq "/auth/me" -and $method -eq "GET") {
                $authResult = Invoke-AuthCheck -Request $request
                if (-not $authResult.allowed) {
                    Send-Response $response @{success=$false; error=$authResult.reason} 401
                    Write-Host "  [AUTH FAIL] $($authResult.reason)" -ForegroundColor Red
                    continue
                }
                $user = $authResult.user

                $instArray = @()
                if ($null -ne $user.instances) { $instArray = @($user.instances) }

                $result = @{
                    success  = $true;
                    user     = @{ uid = $user.uid; name = $user.name; role = $user.role; permissions = $user.permissions };
                    instances = $instArray;
                    serverId = Get-ServerId
                }
                Send-Response $response $result
                Write-Host "  [OK] /auth/me ($($user.name))" -ForegroundColor Green
                continue
            }

            # ════════════════════════════════════════════════
            #  PROTECTED ENDPOINTS (Auth & Headers required)
            # ════════════════════════════════════════════════

            $instanceId   = $request.Headers["X-Instance-ID"]
            $companyCode  = $request.Headers["X-Company-Code"]
            $bridgeSecret = $request.Headers["X-Bridge-Secret"]
            $cfg          = Get-Config
            $requireAuth  = $true

            if ($bridgeSecret -eq $cfg.BRIDGE_SECRET -and -not $instanceId) {
                $requireAuth = $false
                $instanceId  = $cfg.INSTANCE_ID
                $companyCode = $cfg.COMP_CODE
            }

            if ($requireAuth) {
                if (-not $instanceId -or -not $companyCode) {
                    Send-Response $response @{success=$false; error="X-Instance-ID and X-Company-Code headers are required"} 400
                    continue
                }
                $authResult = Invoke-AuthCheck -Request $request -InstanceId $instanceId -CompanyCode $companyCode
                if (-not $authResult.allowed) {
                    Send-Response $response @{success=$false; error=$authResult.reason} 403
                    Write-Host "  [DENIED] $($authResult.reason)" -ForegroundColor Red
                    continue
                }
            }

            # --- COMPANY SETTINGS ---
            if ($path -eq "/busy/company" -and $method -eq "GET") {
                $result = Get-CompanyDetails -InstanceId $instanceId -CompanyCode $companyCode

            # --- VOUCHER MANAGEMENT ---
            } elseif ($path -eq "/busy/voucher" -and $method -eq "POST") {
                $bodyObj = Read-RequestBody $request | ConvertFrom-Json

                $salesmanCheck = Apply-SalesmanAssignmentToVoucherData `
                    -AuthResult $authResult `
                    -Data $bodyObj `
                    -InstanceId $instanceId `
                    -CompanyCode $companyCode `
                    -RequireAuth $requireAuth

                if (-not $salesmanCheck.success) {
                    Send-Response $response @{
                        success   = $false
                        errorCode = "SALESMAN_VALIDATION_FAILED"
                        error     = if ($salesmanCheck.error) { $salesmanCheck.error } else { "Could not validate Salesman assignment." }
                    } 500
                    continue
                }

                if (-not $salesmanCheck.allowed) {
                    Send-Response $response @{
                        success   = $false
                        errorCode = if ($salesmanCheck.errorCode) { $salesmanCheck.errorCode } else { "SALESMAN_ACCESS_DENIED" }
                        error     = if ($salesmanCheck.error) { $salesmanCheck.error } else { "Salesman access denied." }
                    } 403
                    continue
                }

                $bodyObj = $salesmanCheck.data

                $partyAccessCheck = Test-VoucherPartyAccountAccess `
                    -AuthResult $authResult `
                    -Data $bodyObj `
                    -InstanceId $instanceId `
                    -CompanyCode $companyCode `
                    -RequireAuth $requireAuth

                if (-not $partyAccessCheck.success) {
                    Send-Response $response @{
                        success   = $false
                        errorCode = "PARTY_ACCOUNT_VALIDATION_FAILED"
                        error     = if ($partyAccessCheck.error) { $partyAccessCheck.error } else { "Could not validate Party Account Access." }
                    } 500
                    continue
                }

                if (-not $partyAccessCheck.allowed) {
                    Send-Response $response @{
                        success        = $false
                        errorCode      = "PARTY_ACCOUNT_ACCESS_DENIED"
                        error          = "The selected Party/account is not allowed for this user and voucher type."
                        deniedAccounts = @($partyAccessCheck.deniedAccounts)
                    } 403
                    continue
                }

                $dcAccessCheck = Test-VoucherDebitCreditAccountAccess `
                    -AuthResult $authResult `
                    -Data $bodyObj `
                    -InstanceId $instanceId `
                    -CompanyCode $companyCode `
                    -RequireAuth $requireAuth

                if (-not $dcAccessCheck.success) {
                    Send-Response $response @{
                        success   = $false
                        errorCode = "ACCOUNT_SIDE_ACCESS_VALIDATION_FAILED"
                        error     = if ($dcAccessCheck.error) { $dcAccessCheck.error } else { "Could not validate Debit/Credit Account Access." }
                    } 500
                    continue
                }

                if (-not $dcAccessCheck.allowed) {
                    Send-Response $response @{
                        success              = $false
                        errorCode            = "ACCOUNT_SIDE_ACCESS_DENIED"
                        error                = "One or more Journal/Contra accounts are not allowed for their Debit/Credit side."
                        deniedAccounts       = @($dcAccessCheck.deniedAccounts)
                        deniedDebitAccounts  = @($dcAccessCheck.deniedDebitAccounts)
                        deniedCreditAccounts = @($dcAccessCheck.deniedCreditAccounts)
                    } 403
                    continue
                }

                $itemAccessCheck = Test-VoucherItemGroupAccess `
                    -AuthResult $authResult `
                    -Data $bodyObj `
                    -InstanceId $instanceId `
                    -CompanyCode $companyCode `
                    -RequireAuth $requireAuth

                if (-not $itemAccessCheck.success) {
                    Send-Response $response @{
                        success   = $false
                        errorCode = "ITEM_GROUP_ACCESS_VALIDATION_FAILED"
                        error     = if ($itemAccessCheck.error) { $itemAccessCheck.error } else { "Could not validate Item Group Access." }
                    } 500
                    continue
                }

                if (-not $itemAccessCheck.allowed) {
                    Send-Response $response @{
                        success     = $false
                        errorCode   = "ITEM_GROUP_ACCESS_DENIED"
                        error       = "One or more Items are outside the Item Group branches allowed for this user and voucher type."
                        deniedItems = @($itemAccessCheck.deniedItems)
                    } 403
                    continue
                }

                if ($null -ne $authResult -and $null -ne $authResult.user -and $authResult.user.name) {
                    $bodyObj | Add-Member -MemberType NoteProperty -Name "bridgeUserName" -Value $authResult.user.name -Force
                }

                $result = Create-Voucher -Data $bodyObj -InstanceId $instanceId -CompanyCode $companyCode

            # --- OFFLINE VOUCHER SYNCHRONIZATION ---
            } elseif ($path -eq "/busy/offline-vouchers/sync" -and $method -eq "POST") {
                if (-not (Get-Command Sync-OfflineVoucher -ErrorAction SilentlyContinue)) {
                    $result = @{
                        success   = $false
                        errorCode = "OFFLINE_SYNC_NOT_LOADED"
                        error     = "Offline synchronization functions are not loaded."
                    }
                    $response.StatusCode = 503
                }
                elseif (-not $requireAuth -or $null -eq $authResult -or $null -eq $authResult.user) {
                    $result = @{
                        success   = $false
                        errorCode = "AUTH_REQUIRED"
                        error     = "An authenticated user is required to synchronize offline vouchers."
                    }
                    $response.StatusCode = 401
                }
                else {
                    $bodyText = Read-RequestBody $request

                    if ([string]::IsNullOrWhiteSpace($bodyText)) {
                        $result = @{
                            success   = $false
                            errorCode = "EMPTY_REQUEST"
                            error     = "Offline voucher payload is required."
                        }
                        $response.StatusCode = 400
                    }
                    else {
                        $bodyObj = $null

                        try {
                            $bodyObj = $bodyText | ConvertFrom-Json
                        }
                        catch {
                            $result = @{
                                success   = $false
                                errorCode = "INVALID_JSON"
                                error     = "The offline voucher request body is not valid JSON."
                            }
                            $response.StatusCode = 400
                        }

                        if ($null -ne $bodyObj) {
                            $bodyObj | Add-Member -MemberType NoteProperty -Name "instanceId" -Value ([string]$instanceId) -Force
                            $bodyObj | Add-Member -MemberType NoteProperty -Name "companyCode" -Value ([string]$companyCode) -Force
                            $bodyObj | Add-Member -MemberType NoteProperty -Name "userName" -Value ([string]$authResult.user.name) -Force

                            $idempotencyKey = [string]$request.Headers["Idempotency-Key"]
                            if ([string]::IsNullOrWhiteSpace($idempotencyKey)) {
                                $idempotencyKey = [string]$bodyObj.localId
                            }

                            if ([string]::IsNullOrWhiteSpace($idempotencyKey)) {
                                $result = @{
                                    success   = $false
                                    errorCode = "IDEMPOTENCY_KEY_REQUIRED"
                                    error     = "Idempotency-Key header or localId is required."
                                }
                                $response.StatusCode = 400
                            }
                            elseif (-not [string]::IsNullOrWhiteSpace([string]$bodyObj.localId) -and ([string]$bodyObj.localId).Trim() -ne $idempotencyKey.Trim()) {
                                $result = @{
                                    success   = $false
                                    errorCode = "IDEMPOTENCY_KEY_MISMATCH"
                                    error     = "Idempotency-Key must match localId."
                                }
                                $response.StatusCode = 400
                            }
                            else {
                                $bodyObj | Add-Member -MemberType NoteProperty -Name "localId" -Value $idempotencyKey.Trim() -Force

                                try {
                                    $result = Sync-OfflineVoucher -Data $bodyObj -CurrentUser $authResult.user

                                    if ($null -eq $result) {
                                        $result = @{
                                            success   = $false
                                            localId   = $idempotencyKey.Trim()
                                            errorCode = "EMPTY_SYNC_RESULT"
                                            error     = "Offline synchronization returned no result."
                                        }
                                        $response.StatusCode = 500
                                    }
                                    elseif ($result.success -eq $false) {
                                        if ($result.conflict -eq $true) { $response.StatusCode = 409 }
                                        else { $response.StatusCode = 400 }
                                    }
                                    else {
                                        $response.StatusCode = 200
                                    }
                                }
                                catch {
                                    $message = $_.Exception.Message
                                    $statusCode = 500
                                    $errorCode = "OFFLINE_SYNC_FAILED"
                                    $isConflict = $false

                                    if ($message -match "already being processed") {
                                        $statusCode = 409
                                        $errorCode = "OFFLINE_SYNC_IN_PROGRESS"
                                        $isConflict = $true
                                    }
                                    elseif ($message -match "required" -or $message -match "invalid" -or $message -match "not found") {
                                        $statusCode = 400
                                        $errorCode = "OFFLINE_SYNC_VALIDATION_FAILED"
                                    }

                                    $result = @{
                                        success   = $false
                                        localId   = $idempotencyKey.Trim()
                                        conflict  = $isConflict
                                        errorCode = $errorCode
                                        error     = $message
                                    }
                                    $response.StatusCode = $statusCode
                                }
                            }
                        }
                    }
                }

            } elseif ($path -eq "/busy/voucher/modify" -and $method -eq "POST") {
                $bodyObj = Read-RequestBody $request | ConvertFrom-Json

                $salesmanCheck = Apply-SalesmanAssignmentToVoucherData `
                    -AuthResult $authResult `
                    -Data $bodyObj `
                    -InstanceId $instanceId `
                    -CompanyCode $companyCode `
                    -RequireAuth $requireAuth

                if (-not $salesmanCheck.success) {
                    Send-Response $response @{
                        success   = $false
                        errorCode = "SALESMAN_VALIDATION_FAILED"
                        error     = if ($salesmanCheck.error) { $salesmanCheck.error } else { "Could not validate Salesman assignment." }
                    } 500
                    continue
                }

                if (-not $salesmanCheck.allowed) {
                    Send-Response $response @{
                        success   = $false
                        errorCode = if ($salesmanCheck.errorCode) { $salesmanCheck.errorCode } else { "SALESMAN_ACCESS_DENIED" }
                        error     = if ($salesmanCheck.error) { $salesmanCheck.error } else { "Salesman access denied." }
                    } 403
                    continue
                }

                $bodyObj = $salesmanCheck.data

                $partyAccessCheck = Test-VoucherPartyAccountAccess `
                    -AuthResult $authResult `
                    -Data $bodyObj `
                    -InstanceId $instanceId `
                    -CompanyCode $companyCode `
                    -RequireAuth $requireAuth

                if (-not $partyAccessCheck.success) {
                    Send-Response $response @{
                        success   = $false
                        errorCode = "PARTY_ACCOUNT_VALIDATION_FAILED"
                        error     = if ($partyAccessCheck.error) { $partyAccessCheck.error } else { "Could not validate Party Account Access." }
                    } 500
                    continue
                }

                if (-not $partyAccessCheck.allowed) {
                    Send-Response $response @{
                        success        = $false
                        errorCode      = "PARTY_ACCOUNT_ACCESS_DENIED"
                        error          = "The selected Party/account is not allowed for this user and voucher type."
                        deniedAccounts = @($partyAccessCheck.deniedAccounts)
                    } 403
                    continue
                }

                $dcAccessCheck = Test-VoucherDebitCreditAccountAccess `
                    -AuthResult $authResult `
                    -Data $bodyObj `
                    -InstanceId $instanceId `
                    -CompanyCode $companyCode `
                    -RequireAuth $requireAuth

                if (-not $dcAccessCheck.success) {
                    Send-Response $response @{
                        success   = $false
                        errorCode = "ACCOUNT_SIDE_ACCESS_VALIDATION_FAILED"
                        error     = if ($dcAccessCheck.error) { $dcAccessCheck.error } else { "Could not validate Debit/Credit Account Access." }
                    } 500
                    continue
                }

                if (-not $dcAccessCheck.allowed) {
                    Send-Response $response @{
                        success              = $false
                        errorCode            = "ACCOUNT_SIDE_ACCESS_DENIED"
                        error                = "One or more Journal/Contra accounts are not allowed for their Debit/Credit side."
                        deniedAccounts       = @($dcAccessCheck.deniedAccounts)
                        deniedDebitAccounts  = @($dcAccessCheck.deniedDebitAccounts)
                        deniedCreditAccounts = @($dcAccessCheck.deniedCreditAccounts)
                    } 403
                    continue
                }

                $itemAccessCheck = Test-VoucherItemGroupAccess `
                    -AuthResult $authResult `
                    -Data $bodyObj `
                    -InstanceId $instanceId `
                    -CompanyCode $companyCode `
                    -RequireAuth $requireAuth

                if (-not $itemAccessCheck.success) {
                    Send-Response $response @{
                        success   = $false
                        errorCode = "ITEM_GROUP_ACCESS_VALIDATION_FAILED"
                        error     = if ($itemAccessCheck.error) { $itemAccessCheck.error } else { "Could not validate Item Group Access." }
                    } 500
                    continue
                }

                if (-not $itemAccessCheck.allowed) {
                    Send-Response $response @{
                        success     = $false
                        errorCode   = "ITEM_GROUP_ACCESS_DENIED"
                        error       = "One or more Items are outside the Item Group branches allowed for this user and voucher type."
                        deniedItems = @($itemAccessCheck.deniedItems)
                    } 403
                    continue
                }

                if ($null -ne $authResult -and $null -ne $authResult.user -and $authResult.user.name) {
                    $bodyObj | Add-Member -MemberType NoteProperty -Name "bridgeUserName" -Value $authResult.user.name -Force
                }

                $result = Modify-Voucher -Data $bodyObj -InstanceId $instanceId -CompanyCode $companyCode

            } elseif ($path -eq "/busy/voucher" -and $method -eq "DELETE") {
                $body   = Read-RequestBody $request
                $result = Delete-Voucher -Data ($body | ConvertFrom-Json) -InstanceId $instanceId -CompanyCode $companyCode

            } elseif ($path -eq "/busy/vouchers" -and $method -eq "GET") {
                $vchTypeStr = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                if (-not $vchTypeStr) {
                    $result = @{ success=$true; count=0; data=@() }
                } else {
                    $partyVal = Get-QueryStringValue $request.QueryString "party" ""
                    if ($partyVal -eq "") { $partyVal = Get-QueryStringValue $request.QueryString "params[party]" "" }

                    $fromVal = Get-QueryStringValue $request.QueryString "from" ""
                    if ($fromVal -eq "") { $fromVal = Get-QueryStringValue $request.QueryString "params[from]" "" }

                    $toVal = Get-QueryStringValue $request.QueryString "to" ""
                    if ($toVal -eq "") { $toVal = Get-QueryStringValue $request.QueryString "params[to]" "" }

                    $searchVal = Get-QueryStringValue $request.QueryString "search" ""
                    if ($searchVal -eq "") { $searchVal = Get-QueryStringValue $request.QueryString "params[search]" "" }

                    $result = Get-Vouchers `
                        -VchType     ([int]$vchTypeStr) `
                        -From        $fromVal `
                        -To          $toVal `
                        -Party       $partyVal `
                        -Search      $searchVal `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode
                }

            } elseif ($path -eq "/busy/voucher/detail" -and $method -eq "GET") {
                $vchTypeStr = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                $vchNo = Get-QueryStringValue $request.QueryString "vchNo" ""
                if ($vchNo -eq "") { $vchNo = Get-QueryStringValue $request.QueryString "params[vchNo]" "" }

                $vchSeries = Get-QueryStringValue $request.QueryString "vchSeries" "Main"
                if ($vchSeries -eq "Main" -or $vchSeries -eq "") {
                    $pSeries = Get-QueryStringValue $request.QueryString "params[vchSeries]" ""
                    if ($pSeries -ne "") { $vchSeries = $pSeries }
                }

                $vchDate = Get-QueryStringValue $request.QueryString "vchDate" ""
                if ($vchDate -eq "") { $vchDate = Get-QueryStringValue $request.QueryString "params[vchDate]" "" }

                if (-not $vchTypeStr -or $vchNo -eq "" -or $vchDate -eq "") {
                    $result = @{success=$false; error="vchType, vchNo, and vchDate required"}; $response.StatusCode=400
                } else {
                    $result = Get-VoucherDetail `
                        -VchType     ([int]$vchTypeStr) `
                        -VchNo       $vchNo `
                        -VchSeries   $vchSeries `
                        -VchDate     $vchDate `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode
                }

            } elseif ($path -eq "/busy/account-voucher/detail" -and $method -eq "GET") {
                $vchTypeStr = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                $vchNo = Get-QueryStringValue $request.QueryString "vchNo" ""
                if ($vchNo -eq "") { $vchNo = Get-QueryStringValue $request.QueryString "params[vchNo]" "" }

                $vchSeries = Get-QueryStringValue $request.QueryString "vchSeries" "Main"
                if ($vchSeries -eq "Main" -or $vchSeries -eq "") {
                    $pSeries = Get-QueryStringValue $request.QueryString "params[vchSeries]" ""
                    if ($pSeries -ne "") { $vchSeries = $pSeries }
                }

                $vchDate = Get-QueryStringValue $request.QueryString "vchDate" ""
                if ($vchDate -eq "") { $vchDate = Get-QueryStringValue $request.QueryString "params[vchDate]" "" }

                if (-not $vchTypeStr -or $vchNo -eq "" -or $vchDate -eq "") {
                    $result = @{success=$false; error="vchType, vchNo, and vchDate required"}; $response.StatusCode=400
                } else {
                    $result = Get-AccountVoucherDetail `
                        -VchType     ([int]$vchTypeStr) `
                        -VchNo       $vchNo `
                        -VchSeries   $vchSeries `
                        -VchDate     $vchDate `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode
                }

            # --- VOUCHER WORKFLOW (ORDERS & RETURNS) ---
            } elseif ($path -eq "/busy/pending-orders" -and $method -eq "GET") {
                $vchTypeStr = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                $party      = Get-QueryStringValue $request.QueryString "party" ""
                if ($party -eq "") { $party = Get-QueryStringValue $request.QueryString "params[party]" "" }

                if (-not $vchTypeStr -or $party -eq "") {
                    $result = @{ success=$true; count=0; data=@() }
                } else {
                    $result = Get-PendingOrders `
                        -VchType     ([int]$vchTypeStr) `
                        -Party       $party `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode
                }

            } elseif ($path -eq "/busy/pending-challans" -and $method -eq "GET") {
                $vchTypeStr = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                $party      = Get-QueryStringValue $request.QueryString "party" ""
                if ($party -eq "") { $party = Get-QueryStringValue $request.QueryString "params[party]" "" }

                if (-not $vchTypeStr -or $party -eq "") {
                    $result = @{ success=$true; count=0; data=@() }
                } else {
                    $result = Get-PendingChallans `
                        -VchType     ([int]$vchTypeStr) `
                        -Party       $party `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode
                }

            } elseif ($path -eq "/busy/search-for-return" -and $method -eq "GET") {
                # Used by:
                #   - Sale/Purchase return original-voucher search
                #   - Receipt F11 pending bills
                #   - Payment F11 pending bills
                #
                # Supported parent voucher types:
                #   2  = Purchase
                #   3  = Sale Return
                #   9  = Sale
                #   10 = Purchase Return

                $vchTypeStr = Get-QueryStringValue $request.QueryString "vchType" ""
                if ($vchTypeStr -eq "") {
                    $vchTypeStr = Get-QueryStringValue $request.QueryString "params[vchType]" ""
                }

                $parsedVchType = 0
                $validVchType = $false

                if (-not [string]::IsNullOrWhiteSpace($vchTypeStr)) {
                    $validVchType = [int]::TryParse(
                        $vchTypeStr.ToString(),
                        [ref]$parsedVchType
                    )
                }

                if (-not $validVchType -or $parsedVchType -notin @(2, 3, 9, 10)) {
                    $result = @{
                        success = $false
                        error   = "A valid vchType is required. Supported values are 2, 3, 9, and 10."
                        count   = 0
                        data    = @()
                    }
                    $response.StatusCode = 400
                } else {
                    $vchNoVal = Get-QueryStringValue $request.QueryString "vchNo" ""
                    if ($vchNoVal -eq "") {
                        $vchNoVal = Get-QueryStringValue $request.QueryString "params[vchNo]" ""
                    }

                    $partyVal = Get-QueryStringValue $request.QueryString "party" ""
                    if ($partyVal -eq "") {
                        $partyVal = Get-QueryStringValue $request.QueryString "params[party]" ""
                    }

                    $fromVal = Get-QueryStringValue $request.QueryString "from" ""
                    if ($fromVal -eq "") {
                        $fromVal = Get-QueryStringValue $request.QueryString "params[from]" ""
                    }

                    $toVal = Get-QueryStringValue $request.QueryString "to" ""
                    if ($toVal -eq "") {
                        $toVal = Get-QueryStringValue $request.QueryString "params[to]" ""
                    }

                    # Return the complete enriched result from Search-OriginalVouchers.
                    # The response includes:
                    #   transactionType
                    #   parentAmount
                    #   saleReturnedAmount
                    #   purchaseReturnedAmount
                    #   alreadyReceivedAmount
                    #   alreadyPaidAmount
                    #   pendingAmount
                    #   totalAmt / returnedAmt / netAmt compatibility fields
                    $result = Search-OriginalVouchers `
                        -VchType     $parsedVchType `
                        -VchNo       $vchNoVal `
                        -Party       $partyVal `
                        -FromDate    $fromVal `
                        -ToDate      $toVal `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode

                    if ($null -eq $result) {
                        $result = @{
                            success = $false
                            error   = "Search-OriginalVouchers returned no response."
                            count   = 0
                            data    = @()
                        }
                        $response.StatusCode = 500
                    }
                    elseif ($result.success -eq $false) {
                        $response.StatusCode = 500
                    }
                    else {
                        # Ensure predictable response structure without removing
                        # any of the enriched fields returned by vouchers.ps1.
                        if ($null -eq $result.data) {
                            $result.data = @()
                        }

                        if ($null -eq $result.count) {
                            $result.count = @($result.data).Count
                        }
                    }
                }

            } elseif ($path -eq "/busy/return-history" -and $method -eq "GET") {
                $vchTypeStr  = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                $vchCodeStr  = $request.QueryString["vchCode"]
                if (-not $vchCodeStr) { $vchCodeStr = $request.QueryString["params[vchCode]"] }

                if (-not $vchTypeStr -or -not $vchCodeStr) {
                    $result = @{ success=$true; count=0; data=@() }
                } else {
                    $result = Get-ReturnHistory `
                        -OrigVchType  ([int]$vchTypeStr) `
                        -OrigVchCode  ([int]$vchCodeStr) `
                        -InstanceId   $instanceId `
                        -CompanyCode  $companyCode
                }

            # ── DATABASE PERSISTENCE ENDPOINTS ──
            } elseif ($path -eq "/busy/column-config" -and $method -eq "GET") {
                $vchTypeStr = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                $deviceTypeStr = $request.QueryString["deviceType"]
                if (-not $deviceTypeStr) { $deviceTypeStr = $request.QueryString["params[deviceType]"] }

                if (-not $vchTypeStr -or $null -eq $deviceTypeStr) {
                    $result = @{ success = $false; error = "vchType and deviceType required" }
                    $response.StatusCode = 400
                } else {
                    $result = Get-ColumnConfig -VchType ([int]$vchTypeStr) -DeviceType ([int]$deviceTypeStr) -InstanceId $instanceId -CompanyCode $companyCode
                }

            } elseif ($path -eq "/busy/column-config" -and $method -eq "POST") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                $result = Save-ColumnConfig -Data $data -InstanceId $instanceId -CompanyCode $companyCode

            # --- VOUCHER NUMBERING ADMIN CONFIGURATION ---
            } elseif ($path -eq "/busy/voucher-numbering-admin" -and $method -eq "GET") {
                $vchTypeStr = Get-QueryStringValue $request.QueryString "vchType" ""
                if ([string]::IsNullOrWhiteSpace($vchTypeStr)) {
                    $vchTypeStr = Get-QueryStringValue $request.QueryString "params[vchType]" ""
                }

                $seriesName = Get-QueryStringValue $request.QueryString "seriesName" ""
                if ([string]::IsNullOrWhiteSpace($seriesName)) {
                    $seriesName = Get-QueryStringValue $request.QueryString "params[seriesName]" ""
                }

                $voucherDate = Get-QueryStringValue $request.QueryString "voucherDate" ""
                if ([string]::IsNullOrWhiteSpace($voucherDate)) {
                    $voucherDate = Get-QueryStringValue $request.QueryString "params[voucherDate]" ""
                }

                if (
                    [string]::IsNullOrWhiteSpace($vchTypeStr) -or
                    [string]::IsNullOrWhiteSpace($seriesName)
                ) {
                    $result = @{
                        success = $false
                        error   = "vchType and seriesName are required"
                    }
                    $response.StatusCode = 400
                }
                elseif (
                    -not [string]::IsNullOrWhiteSpace($voucherDate) -and
                    -not (
                        $voucherDate -match '^\d{4}-\d{2}-\d{2}$' -or
                        $voucherDate -match '^\d{2}-\d{2}-\d{4}$'
                    )
                ) {
                    $result = @{
                        success = $false
                        error   = "voucherDate must use yyyy-MM-dd or dd-MM-yyyy format"
                    }
                    $response.StatusCode = 400
                }
                else {
                    $result = Get-WebNumberingConfig `
                        -VchType ([int]$vchTypeStr) `
                        -SeriesName $seriesName.Trim() `
                        -VoucherDate $voucherDate `
                        -InstanceId $instanceId `
                        -CompanyCode $companyCode

                    if ($result.success -eq $false) {
                        $response.StatusCode = 400
                    }
                }

            } elseif ($path -eq "/busy/voucher-numbering-admin" -and $method -eq "POST") {
                $normalRole = ""

                if (
                    $requireAuth -and
                    $null -ne $authResult -and
                    $null -ne $authResult.user
                ) {
                    $normalRole = ([string]$authResult.user.role).Trim().ToLower()
                    $normalRole = $normalRole.Replace(" ", "").Replace("_", "").Replace("-", "")
                }

                if ($normalRole -ne "superadmin") {
                    $result = @{
                        success = $false
                        error   = "Only Super Admin can change voucher numbering configuration."
                    }
                    $response.StatusCode = 403
                }
                else {
                    $rawBody = Read-RequestBody $request

                    if ([string]::IsNullOrWhiteSpace($rawBody)) {
                        $result = @{
                            success = $false
                            error   = "Voucher numbering configuration payload is required."
                        }
                        $response.StatusCode = 400
                    }
                    else {
                        $data = $null

                        try {
                            $data = $rawBody | ConvertFrom-Json
                        }
                        catch {
                            $result = @{
                                success = $false
                                error   = "The request body is not valid JSON."
                            }
                            $response.StatusCode = 400
                        }

                        if ($null -ne $data) {
                            $effectiveInstanceId = $instanceId
                            $effectiveCompanyCode = $companyCode

                            if (
                                [string]::IsNullOrWhiteSpace($effectiveInstanceId) -and
                                $data.instance_id
                            ) {
                                $effectiveInstanceId = [string]$data.instance_id
                            }

                            if (
                                [string]::IsNullOrWhiteSpace($effectiveCompanyCode) -and
                                $data.company_code
                            ) {
                                $effectiveCompanyCode = [string]$data.company_code
                            }

                            $dateBasis = "VOUCHER_DATE"
                            $dateBasisProperty = $data.PSObject.Properties["date_basis"]

                            if ($null -ne $dateBasisProperty) {
                                $candidateDateBasis = ([string]$dateBasisProperty.Value).Trim().ToUpperInvariant()

                                if (
                                    $candidateDateBasis -notin @(
                                        "VOUCHER_DATE",
                                        "REAL_TIME"
                                    )
                                ) {
                                    $result = @{
                                        success = $false
                                        error   = "date_basis must be VOUCHER_DATE or REAL_TIME."
                                    }
                                    $response.StatusCode = 400
                                }
                                else {
                                    $dateBasis = $candidateDateBasis
                                }
                            }

                            if ($null -eq $result) {
                                $data |
                                    Add-Member `
                                        -MemberType NoteProperty `
                                        -Name "date_basis" `
                                        -Value $dateBasis `
                                        -Force

                                $updatedBy = ""

                                if (
                                    $null -ne $authResult -and
                                    $null -ne $authResult.user
                                ) {
                                    $updatedBy = [string]$authResult.user.name
                                }

                                $result = Save-WebNumberingConfig `
                                    -Data $data `
                                    -UpdatedBy $updatedBy `
                                    -InstanceId $effectiveInstanceId `
                                    -CompanyCode $effectiveCompanyCode

                                if ($result.success -eq $false) {
                                    $response.StatusCode = 400
                                }
                            }
                        }
                    }
                }

            # --- USER PERMISSIONS (db.bds OLEDB Integration) ---
            } elseif ($path -eq "/busy/permissions" -and $method -eq "GET") {
                if ($requireAuth -and -not (Test-IsPermissionAdminUser -User $authResult.user)) {
                    Send-Response $response @{ success=$false; error="Administrator permission is required." } 403
                    continue
                }
                $result = Get-UserPermissions -InstanceId $instanceId -CompanyCode $companyCode

            } elseif ($path -eq "/busy/permissions/my" -and $method -eq "GET") {
                $userRes = Get-UserPermissions -InstanceId $instanceId -CompanyCode $companyCode
                if ($userRes.success) {
                    $activeName = $authResult.user.name
                    $myPerm = $userRes.data | Where-Object { $_.name.Trim().ToLower() -eq $activeName.Trim().ToLower() }
                    if ($myPerm) {
                        $result = @{ success = $true; data = $myPerm }
                    } else {
                        $defaultPermissionValue = if (Test-IsPermissionAdminUser -User $authResult.user) { 1 } else { 0 }
                        $result = @{ success = $true; data = @{
                            name = $activeName
                            C1=$defaultPermissionValue;C2=$defaultPermissionValue;C3=$defaultPermissionValue;C4=$defaultPermissionValue;C5=$defaultPermissionValue;C6=$defaultPermissionValue;C7=$defaultPermissionValue;C8=$defaultPermissionValue;C9=$defaultPermissionValue;C10=$defaultPermissionValue
                            I1=$defaultPermissionValue;I2=$defaultPermissionValue;I3=$defaultPermissionValue;I4=$defaultPermissionValue;I5=$defaultPermissionValue;I6=$defaultPermissionValue;I7=$defaultPermissionValue;I8=$defaultPermissionValue;I9=$defaultPermissionValue;I10=$defaultPermissionValue;I11=$defaultPermissionValue;I12=$defaultPermissionValue;I13=$defaultPermissionValue;I14=$defaultPermissionValue;I15=$defaultPermissionValue;I16=$defaultPermissionValue;I17=$defaultPermissionValue;I18=$defaultPermissionValue;I19=$defaultPermissionValue;I20=$defaultPermissionValue;I21=$defaultPermissionValue;I22=$defaultPermissionValue
                            B33=$defaultPermissionValue;B34=$defaultPermissionValue;M1="{}";M2="{}"
                        } }
                    }
                } else {
                    $result = $userRes
                }

            } elseif ($path -eq "/busy/permissions/default-access" -and $method -eq "GET") {
                if ($requireAuth -and -not (Test-IsPermissionAdminUser -User $authResult.user)) {
                    Send-Response $response @{ success=$false; error="Administrator permission is required." } 403
                    continue
                }

                $result = Get-NewPermissionProfileDefaultAccess `
                    -InstanceId $instanceId `
                    -CompanyCode $companyCode

                if ($result.success -eq $false) {
                    Send-Response $response $result 500
                    continue
                }

            } elseif ($path -eq "/busy/permissions/save" -and $method -eq "POST") {
                if ($requireAuth -and -not (Test-IsPermissionAdminUser -User $authResult.user)) {
                    Send-Response $response @{ success=$false; error="Administrator permission is required." } 403
                    continue
                }

                $data = Read-RequestBody $request | ConvertFrom-Json
                $result = Save-UserPermissions -Data $data -InstanceId $instanceId -CompanyCode $companyCode

                if ($result.success -eq $false) {
                    Send-Response $response $result 500
                    continue
                }

            } elseif ($path -eq "/busy/users" -and $method -eq "GET") {
                if ($requireAuth -and -not (Test-IsPermissionAdminUser -User $authResult.user)) {
                    Send-Response $response @{ success=$false; error="Administrator permission is required." } 403
                    continue
                }
                $result = Get-CompanyUsers -InstanceId $instanceId -CompanyCode $companyCode

            # --- MASTER DATA ---

            # --- PRODUCTION BOM MASTER DATA ---
            } elseif ($path -eq "/busy/boms" -and $method -eq "GET") {
                $searchVal = Get-QueryStringValue $request.QueryString "search" ""
                if ($searchVal -eq "") {
                    $searchVal = Get-QueryStringValue $request.QueryString "params[search]" ""
                }

                $result = Get-BomList `
                    -Search       $searchVal `
                    -InstanceId   $instanceId `
                    -CompanyCode  $companyCode

            } elseif ($path -eq "/busy/bom/detail" -and $method -eq "GET") {
                $bomCodeStr = Get-QueryStringValue $request.QueryString "code" ""
                if ($bomCodeStr -eq "") {
                    $bomCodeStr = Get-QueryStringValue $request.QueryString "params[code]" ""
                }

                $bomName = Get-QueryStringValue $request.QueryString "name" ""
                if ($bomName -eq "") {
                    $bomName = Get-QueryStringValue $request.QueryString "params[name]" ""
                }

                $bomCode = 0
                $validBomCode = $false

                if (-not [string]::IsNullOrWhiteSpace($bomCodeStr)) {
                    $validBomCode = [int]::TryParse(
                        $bomCodeStr.ToString(),
                        [ref]$bomCode
                    )
                }

                if (
                    (-not $validBomCode -or $bomCode -le 0) -and
                    [string]::IsNullOrWhiteSpace($bomName)
                ) {
                    $result = @{
                        success = $false
                        error   = "BOM code or BOM name is required"
                    }
                    $response.StatusCode = 400
                } else {
                    $result = Get-BomDetail `
                        -BomCode     $bomCode `
                        -BomName     $bomName `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode
                }

            } elseif ($path -eq "/busy/bom" -and $method -eq "POST") {
                $data = Read-RequestBody $request | ConvertFrom-Json

                if (
                    -not $data.name -or
                    -not ($data.item -or $data.mainItemName) -or
                    -not ($data.unit -or $data.mainUnit)
                ) {
                    $result = @{
                        success = $false
                        error   = "BOM name, Item to Produce and Unit are required"
                    }
                    $response.StatusCode = 400
                }
                else {
                    $result = Create-Bom `
                        -Data        $data `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode

                    if ($result.success -eq $false) {
                        $response.StatusCode = 400
                    }
                }

            } elseif ($path -eq "/busy/bom" -and $method -eq "PUT") {
                $data = Read-RequestBody $request | ConvertFrom-Json

                if (
                    -not $data.name -or
                    -not ($data.item -or $data.mainItemName) -or
                    -not ($data.unit -or $data.mainUnit)
                ) {
                    $result = @{
                        success = $false
                        error   = "BOM name, Item to Produce and Unit are required"
                    }
                    $response.StatusCode = 400
                }
                else {
                    $result = Update-Bom `
                        -Data        $data `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode

                    if ($result.success -eq $false) {
                        $response.StatusCode = 400
                    }
                }

            } elseif ($path -eq "/busy/account-groups" -and $method -eq "GET") {
                $result = Get-AccountGroups -InstanceId $instanceId -CompanyCode $companyCode
            } elseif ($path -eq "/busy/account-group" -and $method -eq "POST") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                if (-not $data.name) { $result = @{success=$false;error="name required"}; $response.StatusCode=400 }
                else {
                    $pGrp = if ($data.parentGroup) { $data.parentGroup } else { "Primary" }
                    $result = Create-AccountGroup -Name $data.name -ParentGroup $pGrp -InstanceId $instanceId -CompanyCode $companyCode
                }
            } elseif ($path -eq "/busy/account-group" -and $method -eq "PUT") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                if (-not $data.name) { $result = @{success=$false;error="name required"}; $response.StatusCode=400 }
                else {
                    $pGrp = if ($data.parentGroup) { $data.parentGroup } else { "Primary" }
                    $result = Update-AccountGroup -Name $data.name -ParentGroup $pGrp -InstanceId $instanceId -CompanyCode $companyCode
                }
            } elseif ($path -eq "/busy/accounts" -and $method -eq "GET") {
                $pageVal = Get-QueryStringValue $request.QueryString "page" "1"
                if ($pageVal -eq "") { $pageVal = Get-QueryStringValue $request.QueryString "params[page]" "1" }

                $pageSizeVal = Get-QueryStringValue $request.QueryString "pageSize" "30"
                if ($pageSizeVal -eq "") { $pageSizeVal = Get-QueryStringValue $request.QueryString "params[pageSize]" "30" }

                $groupVal = $request.QueryString["group"]
                if ($groupVal -eq "") { $groupVal = Get-QueryStringValue $request.QueryString "params[group]" "" }

                $searchVal = Get-QueryStringValue $request.QueryString "search" ""
                if ($searchVal -eq "") { $searchVal = Get-QueryStringValue $request.QueryString "params[search]" "" }

                $result = Get-Accounts `
                    -GroupName   $groupVal `
                    -Search      $searchVal `
                    -Page        ([int]$pageVal) `
                    -PageSize    ([int]$pageSizeVal) `
                    -InstanceId  $instanceId `
                    -CompanyCode $companyCode

            } elseif ($path -eq "/busy/account" -and $method -eq "GET") {
                $name = $request.QueryString["name"]
                if (-not $name) { $result = @{success=$false;error="name required"}; $response.StatusCode=400 } else { $result = Get-AccountDetail -Name $name -InstanceId $instanceId -CompanyCode $companyCode }
            } elseif ($path -eq "/busy/account" -and $method -eq "POST") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                if (-not $data.name -or -not $data.group) { $result = @{success=$false;error="name and group required"}; $response.StatusCode=400 } else { $result = Create-Account -Data $data -InstanceId $instanceId -CompanyCode $companyCode }
            } elseif ($path -eq "/busy/account" -and $method -eq "PUT") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                if (-not $data.name -or -not $data.group) { $result = @{success=$false;error="name and group required"}; $response.StatusCode=400 } else { $result = Update-Account -Data $data -InstanceId $instanceId -CompanyCode $companyCode }

} elseif ($path -eq "/busy/reports/outstanding" -and $method -eq "GET") {
                $fromVal = Get-QueryStringValue $request.QueryString "from" ""
                $toVal = Get-QueryStringValue $request.QueryString "to" ""
                $asOfVal = Get-QueryStringValue $request.QueryString "asOf" ""

                $typeVal = Get-QueryStringValue $request.QueryString "type" "all"
                $accountVal = Get-QueryStringValue $request.QueryString "account" ""
                $searchVal = Get-QueryStringValue $request.QueryString "search" ""
                $groupVal = Get-QueryStringValue $request.QueryString "group" ""
                $statusVal = Get-QueryStringValue $request.QueryString "status" "all"
                $agingVal = Get-QueryStringValue $request.QueryString "aging" "all"

                $voucherTypeVal = Get-QueryStringValue $request.QueryString "voucherType" "0"
                $minAmountVal = Get-QueryStringValue $request.QueryString "minAmount" "0"
                $maxAmountVal = Get-QueryStringValue $request.QueryString "maxAmount" "0"
                $includeZeroVal = Get-QueryStringValue $request.QueryString "includeZero" "false"

                $pageVal = Get-QueryStringValue $request.QueryString "page" "1"
                $pageSizeVal = Get-QueryStringValue $request.QueryString "pageSize" "50"
                $sortByVal = Get-QueryStringValue $request.QueryString "sortBy" "dueDate"
                $sortDirectionVal = Get-QueryStringValue $request.QueryString "sortDirection" "asc"

                $safePage = 1
                $safePageSize = 50
                $safeVoucherType = 0
                $safeMinAmount = 0.0
                $safeMaxAmount = 0.0

                try {
                    $safePage = [Math]::Max(1, [int]$pageVal)
                }
                catch {
                    $safePage = 1
                }

                try {
                    $requestedPageSize = [int]$pageSizeVal

                    if ($requestedPageSize -eq 0) {
                        # pageSize=0 means "All matching rows".
                        $safePageSize = 0
                    }
                    else {
                        $safePageSize = [Math]::Min(
                            2500,
                            [Math]::Max(1, $requestedPageSize)
                        )
                    }
                }
                catch {
                    $safePageSize = 50
                }

                try {
                    $safeVoucherType = [Math]::Max(
                        0,
                        [int]$voucherTypeVal
                    )
                }
                catch {
                    $safeVoucherType = 0
                }

                # The report supports only:
                # 0  = All Sale/Purchase transaction types
                # 2  = Purchase
                # 3  = Sale Return
                # 9  = Sale
                # 10 = Purchase Return
                if ($safeVoucherType -notin @(0, 2, 3, 9, 10)) {
                    $safeVoucherType = 0
                }

                try {
                    $safeMinAmount = [Math]::Max(
                        0,
                        [double]$minAmountVal
                    )
                }
                catch {
                    $safeMinAmount = 0.0
                }

                try {
                    $safeMaxAmount = [Math]::Max(
                        0,
                        [double]$maxAmountVal
                    )
                }
                catch {
                    $safeMaxAmount = 0.0
                }

                if (
                    $safeMaxAmount -gt 0 -and
                    $safeMinAmount -gt $safeMaxAmount
                ) {
                    $temporaryAmount = $safeMinAmount
                    $safeMinAmount = $safeMaxAmount
                    $safeMaxAmount = $temporaryAmount
                }

                $validTypes = @(
                    "all",
                    "receivable",
                    "payable"
                )

                if ($typeVal -notin $validTypes) {
                    $typeVal = "all"
                }

                $validStatuses = @(
                    "all",
                    "due",
                    "overdue",
                    "not-due",
                    "partially-adjusted",
                    "unadjusted"
                )

                if ($statusVal -notin $validStatuses) {
                    $statusVal = "all"
                }

                $validAgingValues = @(
                    "all",
                    "notDue",
                    "days0To30",
                    "days31To60",
                    "days61To90",
                    "days91To180",
                    "above180"
                )

                if ($agingVal -notin $validAgingValues) {
                    $agingVal = "all"
                }

                $validSortFields = @(
                    "dueDate",
                    "refDate",
                    "account",
                    "pending",
                    "daysOverdue"
                )

                if ($sortByVal -notin $validSortFields) {
                    $sortByVal = "dueDate"
                }

                if ($sortDirectionVal -notin @("asc", "desc")) {
                    $sortDirectionVal = "asc"
                }

                $includeZero = (
                    $includeZeroVal -eq "true" -or
                    $includeZeroVal -eq "1"
                )

                Write-Host (
                    "  [OUTSTANDING-ROUTE] " +
                    "from='$fromVal'; to='$toVal'; asOf='$asOfVal'; " +
                    "type='$typeVal'; status='$statusVal'; aging='$agingVal'; " +
                    "account='$accountVal'; search='$searchVal'; group='$groupVal'; " +
                    "voucherType=$safeVoucherType; min=$safeMinAmount; " +
                    "max=$safeMaxAmount; includeZero=$includeZero; " +
                    "page=$safePage; pageSize=$safePageSize; " +
                    "sortBy='$sortByVal'; direction='$sortDirectionVal'"
                ) -ForegroundColor DarkCyan

                $result = Get-OutstandingReport `
                    -From $fromVal `
                    -To $toVal `
                    -AsOf $asOfVal `
                    -Type $typeVal `
                    -Account $accountVal `
                    -Search $searchVal `
                    -Group $groupVal `
                    -Status $statusVal `
                    -Aging $agingVal `
                    -VoucherType $safeVoucherType `
                    -MinAmount $safeMinAmount `
                    -MaxAmount $safeMaxAmount `
                    -IncludeZero $includeZero `
                    -Page $safePage `
                    -PageSize $safePageSize `
                    -SortBy $sortByVal `
                    -SortDirection $sortDirectionVal `
                    -InstanceId $instanceId `
                    -CompanyCode $companyCode

            } elseif ($path -eq "/busy/reports/stock-status" -and $method -eq "GET") {
                # Single-page BUSY-like Stock Status report.  Legacy query
                # parameters remain supported; the extra parameters power the
                # Detailed / Columnar / Grouped / Hierarchical web views.
                $fromVal = Get-QueryStringValue $request.QueryString "from" ""
                $toVal = Get-QueryStringValue $request.QueryString "to" ""
                $asOfVal = Get-QueryStringValue $request.QueryString "asOf" ""
                $viewVal = Get-QueryStringValue $request.QueryString "view" "balances"

                $materialCentreVal = Get-QueryStringValue $request.QueryString "materialCentre" ""
                $materialCentresVal = Get-QueryStringValue $request.QueryString "materialCentres" ""
                $itemGroupVal = Get-QueryStringValue $request.QueryString "itemGroup" ""
                $searchVal = Get-QueryStringValue $request.QueryString "search" ""
                $statusVal = Get-QueryStringValue $request.QueryString "status" "all"

                $includeZeroVal = Get-QueryStringValue $request.QueryString "includeZero" "true"
                $lowStockLevelVal = Get-QueryStringValue $request.QueryString "lowStockLevel" "5"
                $valueByVal = Get-QueryStringValue $request.QueryString "valueBy" "busy"
                $unitModeVal = Get-QueryStringValue $request.QueryString "unitMode" "both"
                $showValueVal = Get-QueryStringValue $request.QueryString "showValue" "true"
                $includeTransfersVal = Get-QueryStringValue $request.QueryString "includeStockTransfers" "true"
                $salePurchaseSeparateVal = Get-QueryStringValue $request.QueryString "showSalePurchaseSeparately" "true"
                $mastersModeVal = Get-QueryStringValue $request.QueryString "mastersMode" "moved-closing"
                $showParentGroupVal = Get-QueryStringValue $request.QueryString "showParentGroup" "true"

                $pageVal = Get-QueryStringValue $request.QueryString "page" "1"
                $pageSizeVal = Get-QueryStringValue $request.QueryString "pageSize" "100"

                $safeLowStockLevel = 5.0
                [double]::TryParse(
                    [string]$lowStockLevelVal,
                    [System.Globalization.NumberStyles]::Any,
                    [System.Globalization.CultureInfo]::InvariantCulture,
                    [ref]$safeLowStockLevel
                ) | Out-Null

                $safePage = 1
                [int]::TryParse([string]$pageVal, [ref]$safePage) | Out-Null
                if ($safePage -lt 1) { $safePage = 1 }

                # pageSize=0 means All.  Do not apply a hard-coded upper
                # limit here; Get-StockStatusReport queries the real company
                # item count and normalizes against that value.
                $safePageSize = 100
                [int]::TryParse([string]$pageSizeVal, [ref]$safePageSize) | Out-Null
                if ($safePageSize -lt 0) { $safePageSize = 100 }

                # Keep this log while validating server-side pagination. It proves
                # that the browser actually sent the selected page and pageSize.
                Write-Host (
                    "  [STOCK-PAGING] requested page=$pageVal pageSize=$pageSizeVal -> page=$safePage pageSize=$safePageSize"
                ) -ForegroundColor DarkCyan

                $toBool = {
                    param([string]$Value, [bool]$Default)
                    if ([string]::IsNullOrWhiteSpace($Value)) { return $Default }
                    $v = $Value.Trim().ToLowerInvariant()
                    if ($v -in @("1", "true", "yes", "y", "on")) { return $true }
                    if ($v -in @("0", "false", "no", "n", "off")) { return $false }
                    return $Default
                }

                $result = Get-StockStatusReport `
                    -From                    $fromVal `
                    -To                      $toVal `
                    -AsOf                    $asOfVal `
                    -View                    $viewVal `
                    -MaterialCentre          $materialCentreVal `
                    -MaterialCentres         $materialCentresVal `
                    -ItemGroup               $itemGroupVal `
                    -Search                  $searchVal `
                    -Status                  $statusVal `
                    -IncludeZero             (& $toBool $includeZeroVal $true) `
                    -LowStockLevel           $safeLowStockLevel `
                    -ValueBy                 $valueByVal `
                    -UnitMode                $unitModeVal `
                    -ShowValue               (& $toBool $showValueVal $true) `
                    -IncludeStockTransfers   (& $toBool $includeTransfersVal $true) `
                    -ShowSalePurchaseSeparately (& $toBool $salePurchaseSeparateVal $true) `
                    -MastersMode             $mastersModeVal `
                    -ShowParentGroup         (& $toBool $showParentGroupVal $true) `
                    -Page                    $safePage `
                    -PageSize                $safePageSize `
                    -InstanceId              $instanceId `
                    -CompanyCode             $companyCode

            } elseif ($path -eq "/busy/item-groups" -and $method -eq "GET") {
                $result = Get-ItemGroups -InstanceId $instanceId -CompanyCode $companyCode
            } elseif ($path -eq "/busy/item-group" -and $method -eq "POST") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                if (-not $data.name) { $result = @{success=$false;error="name required"}; $response.StatusCode=400 } else { $result = Create-ItemGroup -Data $data -InstanceId $instanceId -CompanyCode $companyCode }
            } elseif ($path -eq "/busy/item-group" -and $method -eq "PUT") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                if (-not $data.name) { $result = @{success=$false;error="name required"}; $response.StatusCode=400 } else { $result = Update-ItemGroup -Data $data -InstanceId $instanceId -CompanyCode $companyCode }

            } elseif ($path -eq "/busy/items" -and $method -eq "GET") {
                $pageVal = Get-QueryStringValue $request.QueryString "page" "1"
                if ($pageVal -eq "") { $pageVal = Get-QueryStringValue $request.QueryString "params[page]" "1" }

                $pageSizeVal = Get-QueryStringValue $request.QueryString "pageSize" "30"
                if ($pageSizeVal -eq "") { $pageSizeVal = Get-QueryStringValue $request.QueryString "params[pageSize]" "30" }

                $catVal = Get-QueryStringValue $request.QueryString "category" ""
                if ($catVal -eq "") { $catVal = Get-QueryStringValue $request.QueryString "params[category]" "" }

                $searchVal = Get-QueryStringValue $request.QueryString "search" ""
                if ($searchVal -eq "") { $searchVal = Get-QueryStringValue $request.QueryString "params[search]" "" }

                $result = Get-Items `
                    -Category    $catVal `
                    -Search      $searchVal `
                    -Page        ([int]$pageVal) `
                    -PageSize    ([int]$pageSizeVal) `
                    -InstanceId  $instanceId `
                    -CompanyCode $companyCode

            } elseif ($path -eq "/busy/vouchers/items" -and $method -eq "GET") {
                $searchVal = Get-QueryStringValue $request.QueryString "search" ""
                if ($searchVal -eq "") { $searchVal = Get-QueryStringValue $request.QueryString "params[search]" "" }

                $pageVal = Get-QueryStringValue $request.QueryString "page" "1"
                if ($pageVal -eq "") { $pageVal = Get-QueryStringValue $request.QueryString "params[page]" "1" }

                $pageSizeVal = Get-QueryStringValue $request.QueryString "pageSize" "30"
                if ($pageSizeVal -eq "") { $pageSizeVal = Get-QueryStringValue $request.QueryString "params[pageSize]" "30" }

                $itemPage = 1
                $itemPageSize = 30
                [void][int]::TryParse([string]$pageVal, [ref]$itemPage)
                [void][int]::TryParse([string]$pageSizeVal, [ref]$itemPageSize)
                if ($itemPage -lt 1) { $itemPage = 1 }
                if ($itemPageSize -lt 1) { $itemPageSize = 30 }
                if ($itemPageSize -gt 30) { $itemPageSize = 30 }

                $vchTypeVal = Get-QueryStringValue $request.QueryString "vchType" "0"
                if ($vchTypeVal -eq "") { $vchTypeVal = Get-QueryStringValue $request.QueryString "params[vchType]" "0" }

                $itemVchType = 0
                [void][int]::TryParse([string]$vchTypeVal, [ref]$itemVchType)

                # /busy/vouchers/items is the voucher picker endpoint. A normal
                # authenticated user must send the active voucher type so the
                # server can apply the correct M2.itemGroupCodes permission.
                # Admins and trusted internal calls retain their existing path.
                if (
                    $requireAuth -and
                    $itemVchType -le 0 -and
                    -not (Test-IsPermissionAdminUser -User $authResult.user)
                ) {
                    Write-Host "  [ITEM ACCESS] user=$($authResult.user.name) vchType=0 -> fail closed" -ForegroundColor Yellow
                    $result = @{ success=$true; total=0; data=@() }
                }
                else {
                    $itemAccess = Get-ItemGroupAccessForAuthUser `
                        -AuthResult $authResult `
                        -VchType $itemVchType `
                        -InstanceId $instanceId `
                        -CompanyCode $companyCode `
                        -RequireAuth $requireAuth

                    $userNameForLog = ""
                    if ($null -ne $authResult -and $null -ne $authResult.user) {
                        $userNameForLog = [string]$authResult.user.name
                    }

                    Write-Host (
                        "  [ITEM ACCESS] user={0} vchType={1} enforce={2} groups=[{3}] adminAll={4}" -f `
                        $userNameForLog, `
                        $itemVchType, `
                        [bool]$itemAccess.enforce, `
                        (@($itemAccess.groupCodes) -join ","), `
                        [bool]$itemAccess.allItems
                    ) -ForegroundColor DarkCyan

                    if (
                        $itemAccess.enforce -and
                        -not [bool]$itemAccess.allItems -and
                        @($itemAccess.groupCodes).Count -eq 0
                    ) {
                        $result = @{ success=$true; total=0; data=@() }
                    }
                    else {
                        if (
                            $itemAccess.enforce -and
                            -not [bool]$itemAccess.allItems
                        ) {
                            # Item searching belongs in items.ps1. The route only
                            # supplies the already-resolved permission group codes.
                            #
                            # Get-ItemsForVoucher applies the Item Group branch
                            # restriction INSIDE its existing Master1 search before
                            # TOP 30 is applied. This keeps the normal-user path as
                            # close as possible to the already-fast admin path.
                            $itemSearchStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

                            $result = Get-ItemsForVoucher `
                                -Search $searchVal `
                                -AllowedGroupCodes @($itemAccess.groupCodes) `
                                -EnforceGroupAccess $true `
                                -MaxResults $itemPageSize `
                                -InstanceId $instanceId `
                                -CompanyCode $companyCode

                            $itemSearchStopwatch.Stop()

                            Write-Host (
                                "  [ITEM ACCESS RESULT] user={0} vchType={1} search='{2}' returned={3} elapsedMs={4} path=ITEMS_MODULE" -f `
                                $userNameForLog, `
                                $itemVchType, `
                                $searchVal, `
                                @($result.data).Count, `
                                [int]$itemSearchStopwatch.ElapsedMilliseconds
                            ) -ForegroundColor DarkGreen
                        }
                        else {
                            # Admin / trusted internal path keeps the existing
                            # BUSY search behavior, including stock enrichment.
                            $result = Get-ItemsForVoucher `
                                -Search      $searchVal `
                                -InstanceId  $instanceId `
                                -CompanyCode $companyCode
                        }
                    }
                }

            } elseif ($path -eq "/busy/vouchers/item-cost" -and $method -eq "GET") {
                # ---------------------------------------------------------
                # Material-centre-specific current average inventory cost.
                #
                # Used by consumed rows in Stock Journal and Production.
                # Cost = remaining stock value / remaining stock quantity.
                # ---------------------------------------------------------
                $codeStr = Get-QueryStringValue $request.QueryString "code" ""
                if ($codeStr -eq "") {
                    $codeStr = Get-QueryStringValue $request.QueryString "params[code]" ""
                }

                $materialCentreVal = Get-QueryStringValue $request.QueryString "materialCentre" ""
                if ($materialCentreVal -eq "") {
                    $materialCentreVal = Get-QueryStringValue $request.QueryString "params[materialCentre]" ""
                }

                $asOfVal = Get-QueryStringValue $request.QueryString "asOf" ""
                if ($asOfVal -eq "") {
                    $asOfVal = Get-QueryStringValue $request.QueryString "params[asOf]" ""
                }
                if ($asOfVal -eq "") {
                    $asOfVal = Get-QueryStringValue $request.QueryString "date" ""
                }

                # Optional edit-mode safeguard: exclude the voucher being edited
                # so its already-posted stock movement does not affect the cost
                # calculated immediately before that voucher.
                $excludeVchCodeStr = Get-QueryStringValue $request.QueryString "excludeVchCode" "0"
                if ($excludeVchCodeStr -eq "") {
                    $excludeVchCodeStr = Get-QueryStringValue $request.QueryString "params[excludeVchCode]" "0"
                }
                $excludeVchCode = 0
                $null = [int]::TryParse($excludeVchCodeStr.ToString(), [ref]$excludeVchCode)

                $itemCode = 0
                $isValidCode = $false

                if (-not [string]::IsNullOrWhiteSpace($codeStr)) {
                    $isValidCode = [int]::TryParse(
                        $codeStr.ToString(),
                        [ref]$itemCode
                    )
                }

                if (-not $isValidCode -or $itemCode -le 0) {
                    $result = @{
                        success = $false
                        error   = "Valid item code is required"
                    }
                    $response.StatusCode = 400
                }
                elseif ([string]::IsNullOrWhiteSpace($materialCentreVal)) {
                    $result = @{
                        success = $false
                        error   = "Material centre is required"
                    }
                    $response.StatusCode = 400
                }
                else {
                    $result = Get-ItemMaterialCentreAverageCost `
                        -ItemCode       $itemCode `
                        -MaterialCentre $materialCentreVal `
                        -AsOf           $asOfVal `
                        -ExcludeVchCode $excludeVchCode `
                        -InstanceId     $instanceId `
                        -CompanyCode    $companyCode

                    if ($null -ne $result -and $result.success -eq $false) {
                        $response.StatusCode = 400
                    }
                }

            } elseif ($path -eq "/busy/vouchers/item-detail" -and $method -eq "GET") {
                $codeStr = Get-QueryStringValue $request.QueryString "code" ""
                if ($codeStr -eq "") {
                    $codeStr = Get-QueryStringValue $request.QueryString "params[code]" ""
                }

                $itemCode = 0
                $isValidCode = $false

                if (-not [string]::IsNullOrWhiteSpace($codeStr)) {
                    $isValidCode = [int]::TryParse(
                        $codeStr.ToString(),
                        [ref]$itemCode
                    )
                }

                if (-not $isValidCode -or $itemCode -le 0) {
                    $result = @{
                        success = $false
                        error   = "Valid item code is required"
                    }
                    $response.StatusCode = 400
                } else {
                    $result = Get-VoucherItemDetail `
                        -Code        $itemCode `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode
                }

            } elseif ($path -eq "/busy/item" -and $method -eq "GET") {
                $codeStr = $request.QueryString["code"]
                if (-not $codeStr) { $result = @{success=$false;error="code required"}; $response.StatusCode=400 } else { $result = Get-ItemByCode -Code ([int]$codeStr) -InstanceId $instanceId -CompanyCode $companyCode }
            } elseif ($path -eq "/busy/item/detail" -and $method -eq "GET") {
                $codeStr = $request.QueryString["code"]
                if (-not $codeStr) { $result = @{success=$false;error="code required"}; $response.StatusCode=400 } else { $result = Get-ItemDetail -Code ([int]$codeStr) -InstanceId $instanceId -CompanyCode $companyCode }
            } elseif ($path -eq "/busy/item" -and $method -eq "POST") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                if (-not $data.name -or -not $data.group) { $result = @{success=$false;error="name and group required"}; $response.StatusCode=400 } else { $result = Create-Item -Data $data -InstanceId $instanceId -CompanyCode $companyCode }
            } elseif ($path -eq "/busy/item" -and $method -eq "PUT") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                if (-not $data.name -or -not $data.group) { $result = @{success=$false;error="name and group required"}; $response.StatusCode=400 } else { $result = Update-Item -Data $data -InstanceId $instanceId -CompanyCode $companyCode }

            } elseif ($path -eq "/busy/item-group-access-tree" -and $method -eq "GET") {
                if ($requireAuth -and -not (Test-IsPermissionAdminUser -User $authResult.user)) {
                    Send-Response $response @{ success=$false; error="Administrator permission is required." } 403
                    continue
                }

                $result = Get-AllItemGroupPermissionNodes `
                    -InstanceId $instanceId `
                    -CompanyCode $companyCode

            } elseif ($path -eq "/busy/account-access-tree" -and $method -eq "GET") {
                if ($requireAuth -and -not (Test-IsPermissionAdminUser -User $authResult.user)) {
                    Send-Response $response @{ success=$false; error="Administrator permission is required." } 403
                    continue
                }

                $result = Get-AllAccountPermissionNodes `
                    -InstanceId $instanceId `
                    -CompanyCode $companyCode

            } elseif ($path -eq "/busy/party-account-groups" -and $method -eq "GET") {
                if ($requireAuth -and -not (Test-IsPermissionAdminUser -User $authResult.user)) {
                    Send-Response $response @{ success=$false; error="Administrator permission is required." } 403
                    continue
                }

                $result = Get-PartyAccountGroups `
                    -InstanceId $instanceId `
                    -CompanyCode $companyCode

            } elseif ($path -eq "/busy/parties" -and $method -eq "GET") {
                $searchVal = Get-QueryStringValue $request.QueryString "search" ""
                if ($searchVal -eq "") { $searchVal = Get-QueryStringValue $request.QueryString "params[search]" "" }

                $pageVal = Get-QueryStringValue $request.QueryString "page" "1"
                if ($pageVal -eq "") { $pageVal = Get-QueryStringValue $request.QueryString "params[page]" "1" }

                $pageSizeVal = Get-QueryStringValue $request.QueryString "pageSize" "30"
                if ($pageSizeVal -eq "") { $pageSizeVal = Get-QueryStringValue $request.QueryString "params[pageSize]" "30" }

                $cashBankVal = Get-QueryStringValue $request.QueryString "cashBankOnly" "false"
                if ($cashBankVal -eq "") { $cashBankVal = Get-QueryStringValue $request.QueryString "params[cashBankOnly]" "false" }
                $isCashBankOnly = ($cashBankVal -eq "true" -or $cashBankVal -eq "1")

                $vchTypeVal = Get-QueryStringValue $request.QueryString "vchType" "0"
                if ($vchTypeVal -eq "") { $vchTypeVal = Get-QueryStringValue $request.QueryString "params[vchType]" "0" }
                $partyVchType = 0
                [void][int]::TryParse([string]$vchTypeVal, [ref]$partyVchType)

                # Normal authenticated users must never get the generic,
                # unrestricted account list because vchType was accidentally
                # omitted. Admin screens may still use the generic endpoint.
                if (
                    $requireAuth -and
                    $partyVchType -le 0 -and
                    $null -ne $authResult -and
                    $null -ne $authResult.user -and
                    -not (Test-IsPermissionAdminUser -User $authResult.user)
                ) {
                    $result = @{
                        success    = $true
                        total      = 0
                        page       = ([int]$pageVal)
                        pageSize   = ([int]$pageSizeVal)
                        totalPages = 1
                        data       = @()
                    }
                    Send-Response $response $result
                    Write-Host "  [PARTY ACCESS] Blocked generic /busy/parties call for normal user because vchType was missing." -ForegroundColor Yellow
                    continue
                }

                $activePartyUser = if ($null -ne $authResult -and $null -ne $authResult.user) { [string]$authResult.user.name } else { "internal" }

                if (Test-IsDebitCreditAccessVoucherType -VchType $partyVchType) {
                    $dcVal = Get-QueryStringValue $request.QueryString "dc" ""
                    if ($dcVal -eq "") { $dcVal = Get-QueryStringValue $request.QueryString "params[dc]" "" }
                    $dcVal = ([string]$dcVal).Trim().ToUpperInvariant()

                    if (
                        $requireAuth -and
                        $null -ne $authResult -and
                        $null -ne $authResult.user -and
                        -not (Test-IsPermissionAdminUser -User $authResult.user) -and
                        ($dcVal -ne "D" -and $dcVal -ne "C")
                    ) {
                        $result = @{ success=$true; total=0; page=([int]$pageVal); pageSize=([int]$pageSizeVal); totalPages=1; data=@() }
                        Send-Response $response $result
                        Write-Host "  [ACCOUNT SIDE ACCESS] Blocked Journal/Contra account request because dc=D/C was missing." -ForegroundColor Yellow
                        continue
                    }

                    # Admins can request either side. If no dc is supplied for an
                    # admin utility screen, default to Debit only because both
                    # sides are unrestricted for admins anyway.
                    if ($dcVal -ne "D" -and $dcVal -ne "C") { $dcVal = "D" }

                    $sideAccess = Get-AccountSideAccessForAuthUser `
                        -AuthResult $authResult `
                        -VchType $partyVchType `
                        -Dc $dcVal `
                        -InstanceId $instanceId `
                        -CompanyCode $companyCode `
                        -RequireAuth $requireAuth

                    Write-Host "  [ACCOUNT SIDE ACCESS] user=$activePartyUser vchType=$partyVchType dc=$dcVal groups=[$(@($sideAccess.groupCodes) -join ',')] accounts=[$(@($sideAccess.accountCodes) -join ',')] adminAll=$($sideAccess.allAccounts)" -ForegroundColor DarkCyan

                    $result = Get-Parties `
                        -Search       $searchVal `
                        -CashBankOnly $isCashBankOnly `
                        -Page         ([int]$pageVal) `
                        -PageSize     ([int]$pageSizeVal) `
                        -InstanceId   $instanceId `
                        -CompanyCode  $companyCode `
                        -AllowedGroupCodes @($sideAccess.groupCodes) `
                        -AllowedAccountCodes @($sideAccess.accountCodes) `
                        -EnforceAllAccountAccess ([bool]$sideAccess.enforce) `
                        -AllowAllAccounts ([bool]$sideAccess.allAccounts)
                }
                else {
                    $partyAccess = Get-PartyGroupAccessForAuthUser `
                        -AuthResult $authResult `
                        -VchType $partyVchType `
                        -InstanceId $instanceId `
                        -CompanyCode $companyCode `
                        -RequireAuth $requireAuth

                    $activePartyGroups = (@($partyAccess.groupCodes) -join ",")
                    Write-Host "  [PARTY ACCESS] user=$activePartyUser vchType=$partyVchType enforce=$($partyAccess.enforce) groups=[$activePartyGroups] adminAll=$($partyAccess.allEligibleRoots)" -ForegroundColor DarkCyan

                    $result = Get-Parties `
                        -Search       $searchVal `
                        -CashBankOnly $isCashBankOnly `
                        -Page         ([int]$pageVal) `
                        -PageSize     ([int]$pageSizeVal) `
                        -InstanceId   $instanceId `
                        -CompanyCode  $companyCode `
                        -AllowedGroupCodes @($partyAccess.groupCodes) `
                        -EnforceGroupAccess ([bool]$partyAccess.enforce) `
                        -AllowAllEligibleRoots ([bool]$partyAccess.allEligibleRoots)
                }

            } elseif ($path -eq "/busy/cash-bank-accounts" -and $method -eq "GET") {
                $searchVal = Get-QueryStringValue $request.QueryString "search" ""
                if ($searchVal -eq "") { $searchVal = Get-QueryStringValue $request.QueryString "params[search]" "" }

                $result = Get-CashBankAccounts `
                    -Search      $searchVal `
                    -InstanceId  $instanceId `
                    -CompanyCode $companyCode
             }elseif ($path -eq "/busy/units" -and $method -eq "GET") {
                $result = Get-Units -InstanceId $instanceId -CompanyCode $companyCode
            } elseif ($path -eq "/busy/tax-categories" -and $method -eq "GET") {
                $result = Get-TaxCategories -InstanceId $instanceId -CompanyCode $companyCode
            } elseif ($path -eq "/busy/bill-sundries" -and $method -eq "GET") {
                $result = Get-BillSundries -InstanceId $instanceId -CompanyCode $companyCode
            } elseif ($path -eq "/busy/salesmen" -and $method -eq "GET") {
                $result = Get-Salesmen -InstanceId $instanceId -CompanyCode $companyCode
            } elseif ($path -eq "/busy/material-centers" -and $method -eq "GET") {
                $result = Get-MaterialCenters -InstanceId $instanceId -CompanyCode $companyCode

            # --- MISC UTILITIES & CONFIGURATIONS ---
            } elseif ($path -eq "/busy/numbering-config" -and $method -eq "GET") {
                $vchTypeStr = Get-QueryStringValue $request.QueryString "vchType" ""
                if ([string]::IsNullOrWhiteSpace($vchTypeStr)) {
                    $vchTypeStr = Get-QueryStringValue $request.QueryString "params[vchType]" ""
                }

                $seriesName = Get-QueryStringValue $request.QueryString "seriesName" ""
                if ([string]::IsNullOrWhiteSpace($seriesName)) {
                    $seriesName = Get-QueryStringValue $request.QueryString "params[seriesName]" ""
                }

                $voucherDate = Get-QueryStringValue $request.QueryString "voucherDate" ""
                if ([string]::IsNullOrWhiteSpace($voucherDate)) {
                    $voucherDate = Get-QueryStringValue $request.QueryString "params[voucherDate]" ""
                }

                if (
                    [string]::IsNullOrWhiteSpace($vchTypeStr) -or
                    [string]::IsNullOrWhiteSpace($seriesName)
                ) {
                    $result = @{
                        success = $false
                        error   = "vchType and seriesName required"
                    }
                    $response.StatusCode = 400
                } else {
                    $result = Get-EffectiveNumberingConfig `
                        -VchType ([int]$vchTypeStr) `
                        -SeriesName $seriesName.Trim() `
                        -VoucherDate $voucherDate `
                        -InstanceId $instanceId `
                        -CompanyCode $companyCode

                    if ($result.success -eq $false) {
                        $response.StatusCode = 400
                    }
                }

            } elseif ($path -eq "/busy/voucher-series" -and $method -eq "GET") {
                $vchTypeStr = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                $vchType = 0
                if ($vchTypeStr) { try { $vchType = [int]$vchTypeStr } catch {} }
                $result = Get-VoucherSeries -VchType $vchType -InstanceId $instanceId -CompanyCode $companyCode
            } elseif ($path -eq "/busy/voucher-types" -and $method -eq "GET") {
                $typeParam = Get-QueryStringValue $request.QueryString "type" "All"
                $result = Get-VoucherTypes -Type $typeParam -InstanceId $instanceId -CompanyCode $companyCode
            } elseif ($path -eq "/busy/input-types" -and $method -eq "GET") {
                $result = Get-VoucherInputTypes
            } elseif ($path -eq "/busy/cache/clear" -and $method -eq "POST") {
                Clear-Cache
                $result = @{success=$true; message="All caches cleared"}

            # --- OPTIONAL FIELDS ROUTING ---
            } elseif ($path -eq "/busy/voucher/optional-fields-config" -and $method -eq "GET") {
                $vchTypeStr = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                $seriesName = Get-QueryStringValue $request.QueryString "seriesName" ""
                if ($seriesName -eq "") { $seriesName = Get-QueryStringValue $request.QueryString "params[seriesName]" "" }

                if (-not $vchTypeStr -or $seriesName -eq "") {
                    $result = @{ success = $false; error = "vchType and seriesName required" }
                    $response.StatusCode = 400
                } else {
                    $result = Get-VoucherOptionalFields `
                        -VchType     ([int]$vchTypeStr) `
                        -SeriesName  $seriesName `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode
                }

            } elseif ($path -eq "/busy/voucher/optional-fields-values" -and $method -eq "GET") {
                $vchTypeStr = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                $seriesName = Get-QueryStringValue $request.QueryString "seriesName" ""
                if ($seriesName -eq "") { $seriesName = Get-QueryStringValue $request.QueryString "params[seriesName]" "" }

                $fieldKeyStr = Get-QueryStringValue $request.QueryString "fieldNo" ""
                if ($fieldKeyStr -eq "") { $fieldKeyStr = Get-QueryStringValue $request.QueryString "params[fieldNo]" "" }
                
                $fieldNo = 1
                if ($fieldKeyStr -match "OptionField(\d+)") {
                    $fieldNo = [int]$Matches[1]
                } elseif ($fieldKeyStr -ne "") {
                    $fieldNo = [int]$fieldKeyStr
                }

                if (-not $vchTypeStr -or $seriesName -eq "" -or $fieldNo -eq 0) {
                    $result = @{ success = $false; error = "vchType, seriesName, and fieldNo required" }
                    $response.StatusCode = 400
                } else {
                    $result = Get-OptionalFieldMasterValues `
                        -VchType     ([int]$vchTypeStr) `
                        -SeriesName  $seriesName `
                        -FieldNo     $fieldNo `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode
                }
            } else {
                $result = @{success=$false; error="Endpoint not found: $method $path"}
                $response.StatusCode = 404
            }

            Send-Response $response $result

            if ($result -and $result.success -eq $true) { Write-Host "[OK] $path" -ForegroundColor Green }
            else { Write-Host "  [FAIL] $($result.error)" -ForegroundColor Red }

        } catch {
            Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
            Send-Response $response @{success=$false; error=$_.Exception.Message} 500
        }
    }
}