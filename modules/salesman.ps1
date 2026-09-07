# modules/salesman.ps1
# BUSY Salesman / Sales Ref backend logic
# PowerShell 5.1 safe.
#
# BUSY stores Salesman as Broker master rows in Master1.MasterType = 19.
# This module owns:
#   - Salesman master loading
#   - per-user Salesman assignment lookup from MobileUserPreference.M2
#   - live BUSY master validation
#   - create/modify voucher Salesman enforcement
#
# Voucher XML serialization itself stays in vouchers.ps1 because that module
# owns BUSY voucher XML construction.

. "$PSScriptRoot\connection.ps1"
. "$PSScriptRoot\utils.ps1"

# BUSY Salesman master is stored as Broker master (Master1.MasterType = 19).
# The company-specific master list is exposed to the permission screen and
# administrator voucher form through /busy/salesmen.
function Get-Salesmen {
    param(
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )

    $cacheKey = "$InstanceId|$CompanyCode|salesmen"
    $cached = Get-Cache $cacheKey
    if ($cached) { return $cached }

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) {
        return @{ success = $false; error = "BUSY connection failed" }
    }

    try {
        $qry = "SELECT Name, Code, Alias FROM Master1 WHERE MasterType = 19 ORDER BY Name"
        $rst = $fi.GetRecordset($qry)

        $salesmen = Read-Recordset $rst {
            param($r)

            $alias = ""
            try {
                $rawAlias = $r.Fields.Item("Alias").Value
                if ($null -ne $rawAlias -and $rawAlias -ne [System.DBNull]::Value) {
                    $alias = ([string]$rawAlias).Trim()
                }
            }
            catch {}

            @{
                code  = [int][string]$r.Fields.Item("Code").Value
                name  = ([string]$r.Fields.Item("Name").Value).Trim()
                alias = $alias
            }
        }

        $result = @{
            success = $true
            count   = @($salesmen).Count
            data    = @($salesmen)
        }

        if (@($salesmen).Count -gt 0) {
            Set-Cache $cacheKey $result
        }

        return $result
    }
    catch {
        return @{ success = $false; error = $_.Exception.Message }
    }
    finally {
        Disconnect-BUSY $fi
    }
}

function Test-IsSalesmanVoucherType {
    param([int]$VchType)

    # Original BUSY voucher types:
    # 26 = Sales Quotation, 12 = Sale Order, 9 = Sale, 3 = Sale Return.
    return @(26, 12, 9, 3) -contains $VchType
}

function Get-SalesmanAssignmentForAuthUser {
    param(
        $AuthResult,
        [int]$VchType = 0,
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        [bool]$RequireAuth = $true
    )

    if (-not $RequireAuth) {
        return @{ success = $true; enforce = $false; isAdmin = $true; assignment = $null }
    }

    if ($null -eq $AuthResult -or $null -eq $AuthResult.user) {
        return @{
            success    = $false
            enforce    = $true
            isAdmin    = $false
            assignment = $null
            error      = "Authenticated user information is unavailable."
        }
    }

    if (Test-IsPermissionAdminUser -User $AuthResult.user) {
        return @{ success = $true; enforce = $false; isAdmin = $true; assignment = $null }
    }

    $m2 = "{}"
    $liveProfileFound = $false

    try {
        $permissionResult = Get-UserPermissions `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (-not $permissionResult.success) {
            return @{
                success    = $false
                enforce    = $true
                isAdmin    = $false
                assignment = $null
                error      = if ($permissionResult.error) {
                    [string]$permissionResult.error
                } else {
                    "Could not read MobileUserPreference."
                }
            }
        }

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
    catch {
        return @{
            success    = $false
            enforce    = $true
            isAdmin    = $false
            assignment = $null
            error      = $_.Exception.Message
        }
    }

    # Compatibility fallback only when the live preference row does not exist.
    if (
        -not $liveProfileFound -and
        $null -ne $AuthResult.user.permissions -and
        $null -ne $AuthResult.user.permissions.M2
    ) {
        $m2 = [string]$AuthResult.user.permissions.M2
    }

    try {
        if (-not (Test-IsSalesmanVoucherType -VchType $VchType)) {
            return @{ success = $true; enforce = $true; isAdmin = $false; assignment = $null }
        }

        $permissionMap = $m2 | ConvertFrom-Json

        # Salesman is stored inside the selected voucher's M2 config:
        # "9":  { "salesman": { ... } }
        # "12": { "salesman": { ... } }
        # "26": { "salesman": { ... } }
        # "3":  { "salesman": { ... } }
        #
        # The legacy top-level "salesman" property is intentionally ignored.
        $voucherProp = $permissionMap.PSObject.Properties[[string]$VchType]

        if ($null -eq $voucherProp -or $null -eq $voucherProp.Value) {
            return @{ success = $true; enforce = $true; isAdmin = $false; assignment = $null }
        }

        $voucherConfig = $voucherProp.Value
        $salesmanProp = $voucherConfig.PSObject.Properties["salesman"]

        if ($null -eq $salesmanProp -or $null -eq $salesmanProp.Value) {
            return @{ success = $true; enforce = $true; isAdmin = $false; assignment = $null }
        }

        $raw = $salesmanProp.Value
        $code = 0
        try { $code = [int]$raw.code } catch { $code = 0 }
        $name = ([string]$raw.name).Trim()

        if ($code -le 0 -and [string]::IsNullOrWhiteSpace($name)) {
            return @{ success = $true; enforce = $true; isAdmin = $false; assignment = $null }
        }

        return @{
            success = $true
            enforce = $true
            isAdmin = $false
            assignment = @{ code = $code; name = $name }
        }
    }
    catch {
        return @{
            success    = $false
            enforce    = $true
            isAdmin    = $false
            assignment = $null
            error      = "Invalid salesman assignment in MobileUserPreference.M2."
        }
    }
}

function Resolve-LiveSalesmanMaster {
    param(
        [int]$Code = 0,
        [string]$Name = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $salesmanResult = Get-Salesmen `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $salesmanResult.success) {
        return @{
            success = $false
            found   = $false
            error   = if ($salesmanResult.error) {
                [string]$salesmanResult.error
            } else {
                "Could not load BUSY Salesman master."
            }
        }
    }

    $match = $null

    if ($Code -gt 0) {
        $match = @($salesmanResult.data) |
            Where-Object { [int]$_.code -eq $Code } |
            Select-Object -First 1
    }

    if ($null -eq $match -and -not [string]::IsNullOrWhiteSpace($Name)) {
        $normalizedName = $Name.Trim().ToLowerInvariant()
        $match = @($salesmanResult.data) |
            Where-Object {
                ([string]$_.name).Trim().ToLowerInvariant() -eq $normalizedName
            } |
            Select-Object -First 1
    }

    if ($null -eq $match) {
        return @{ success = $true; found = $false }
    }

    return @{
        success = $true
        found   = $true
        data    = @{
            code  = [int]$match.code
            name  = ([string]$match.name).Trim()
            alias = ([string]$match.alias).Trim()
        }
    }
}

function Apply-SalesmanAssignmentToVoucherData {
    param(
        $AuthResult,
        $Data,
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        [bool]$RequireAuth = $true
    )

    $vchType = 0
    try { $vchType = [int]$Data.vchType } catch { $vchType = 0 }

    if (-not (Test-IsSalesmanVoucherType -VchType $vchType)) {
        return @{ success = $true; allowed = $true; data = $Data }
    }

    if (-not $RequireAuth) {
        return @{ success = $true; allowed = $true; data = $Data }
    }

    $assignmentResult = Get-SalesmanAssignmentForAuthUser `
        -AuthResult $AuthResult `
        -VchType $vchType `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode `
        -RequireAuth $RequireAuth

    if (-not $assignmentResult.success) {
        return @{ success = $false; allowed = $false; error = $assignmentResult.error }
    }

    # Admins may choose a Salesman in the voucher form. Validate live BUSY master.
    if ($assignmentResult.isAdmin) {
        $clientCode = 0
        try {
            if ($null -ne $Data.salesmanCode) {
                $clientCode = [int]$Data.salesmanCode
            }
        }
        catch { $clientCode = 0 }

        $clientName = ""
        try { $clientName = ([string]$Data.salesmanName).Trim() } catch {}

        if ($clientCode -le 0 -and [string]::IsNullOrWhiteSpace($clientName)) {
            return @{ success = $true; allowed = $true; data = $Data }
        }

        $resolved = Resolve-LiveSalesmanMaster `
            -Code $clientCode `
            -Name $clientName `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (-not $resolved.success) {
            return @{ success = $false; allowed = $false; error = $resolved.error }
        }

        if (-not $resolved.found) {
            return @{
                success   = $true
                allowed   = $false
                errorCode = "SALESMAN_INVALID"
                error     = "The selected Salesman no longer exists in the BUSY Salesman master."
            }
        }

        $Data | Add-Member -MemberType NoteProperty -Name "salesmanCode" -Value ([int]$resolved.data.code) -Force
        $Data | Add-Member -MemberType NoteProperty -Name "salesmanName" -Value ([string]$resolved.data.name) -Force

        return @{ success = $true; allowed = $true; data = $Data }
    }

    # Normal users cannot choose/override Salesman from the browser.
    # Salesman itself is OPTIONAL:
    # - assignment exists -> force that BUSY Salesman
    # - no assignment      -> force NO Salesman and allow the voucher to save
    if ($null -eq $assignmentResult.assignment) {
        $Data | Add-Member -MemberType NoteProperty -Name "salesmanCode" -Value 0 -Force
        $Data | Add-Member -MemberType NoteProperty -Name "salesmanName" -Value "" -Force

        return @{ success = $true; allowed = $true; data = $Data }
    }

    $resolved = Resolve-LiveSalesmanMaster `
        -Code ([int]$assignmentResult.assignment.code) `
        -Name ([string]$assignmentResult.assignment.name) `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $resolved.success) {
        return @{ success = $false; allowed = $false; error = $resolved.error }
    }

    if (-not $resolved.found) {
        return @{
            success   = $true
            allowed   = $false
            errorCode = "SALESMAN_ASSIGNMENT_INVALID"
            error     = "The Salesman assigned to this user no longer exists in the BUSY Salesman master."
        }
    }

    $Data | Add-Member -MemberType NoteProperty -Name "salesmanCode" -Value ([int]$resolved.data.code) -Force
    $Data | Add-Member -MemberType NoteProperty -Name "salesmanName" -Value ([string]$resolved.data.name) -Force

    return @{ success = $true; allowed = $true; data = $Data }
}
