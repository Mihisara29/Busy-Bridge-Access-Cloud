# modules/pocketbase.ps1
# Local PocketBase Authentication & Instances Config Integration

. "$PSScriptRoot\utils.ps1"

$script:_authCache = @{}
$script:_AUTH_CACHE_TTL = 300
$script:_instancesConfig = $null

# ═══════════════════════════════════════════════════════
#  LOAD INSTANCES CONFIG
# ═══════════════════════════════════════════════════════
function Get-InstancesConfig {
    if ($null -ne $script:_instancesConfig) { return $script:_instancesConfig }
    $path = "$PSScriptRoot\..\instances.json"
    if (-not (Test-Path $path)) {
        Write-ErrLog "instances.json not found at $path"
        return $null
    }
    $script:_instancesConfig = Get-Content $path -Raw | ConvertFrom-Json
    return $script:_instancesConfig
}

function Get-ServerId {
    $cfg = Get-InstancesConfig
    if ($cfg) { return $cfg.serverId }
    return "unknown"
}

# ═══════════════════════════════════════════════════════
#  AUTH CACHE HELPERS
# ═══════════════════════════════════════════════════════
function Get-CachedAuth {
    param([string]$Token)
    if (-not $Token) { return $null }
    $key = $Token.Substring([Math]::Max(0, $Token.Length - 32))
    if ($script:_authCache.ContainsKey($key)) {
        $entry = $script:_authCache[$key]
        if ((Get-Date) -lt $entry.expiry) { return $entry.user }
        $script:_authCache.Remove($key)
    }
    return $null
}

function Set-CachedAuth {
    param([string]$Token, $User)
    if (-not $Token) { return }
    $key = $Token.Substring([Math]::Max(0, $Token.Length - 32))
    $script:_authCache[$key] = @{ user = $User; expiry = (Get-Date).AddSeconds($script:_AUTH_CACHE_TTL) }
}

# ═══════════════════════════════════════════════════════
#  PERMISSION CHECK
# ═══════════════════════════════════════════════════════
function Test-UserHasAccess {
    param($User,[string]$InstanceId, [string]$CompanyCode)
    if ($User.role -eq "superadmin") { return $true }
    if ($User.allowedCompanies -eq "all") { return $true }

    $serverId = Get-ServerId
    foreach ($entry in $User.allowedCompanies) {
        if ($entry.serverId -eq $serverId -and $entry.instanceId -eq $InstanceId -and $entry.companyCodes -contains $CompanyCode) {
            return $true
        }
    }
    return $false
}

# ═══════════════════════════════════════════════════════
#  MAIN AUTH CHECK (POCKETBASE)
# ═══════════════════════════════════════════════════════
function Invoke-AuthCheck {
    param($Request, [string]$InstanceId = "",[string]$CompanyCode = "")

    # Extract Token from Authorization header
    $authHeader = $Request.Headers["Authorization"]
    if (-not $authHeader -or -not $authHeader.StartsWith("Bearer ")) {
        return @{ allowed = $false; reason = "Missing or invalid Authorization header" }
    }
    $token = $authHeader.Substring(7)

    # 1. Check Cache
    $cachedUser = Get-CachedAuth -Token $token
    if ($cachedUser) {
        if ($InstanceId -and $CompanyCode -and -not (Test-UserHasAccess -User $cachedUser -InstanceId $InstanceId -CompanyCode $CompanyCode)) {
            return @{ allowed = $false; reason = "Access denied to $InstanceId / $CompanyCode" }
        }
        return @{ allowed = $true; user = $cachedUser }
    }

    # 2. Verify with PocketBase
    $pbUrl = (Get-Config).POCKETBASE_URL
    try {
        $headers = @{ "Authorization" = "Bearer $token" }
        $response = Invoke-RestMethod -Uri "$pbUrl/api/collections/users/auth-refresh" -Method POST -Headers $headers -ErrorAction Stop
        
        $u = $response.record
        
        $allowedCompanies = "all"
        if ($u.allowedCompanies -and $u.allowedCompanies -ne "all" -and $u.allowedCompanies -ne "") {
            try {
                $allowedCompanies = $u.allowedCompanies | ConvertFrom-Json
            } catch {
                Write-WarnLog "Could not parse allowedCompanies JSON for user $($u.email)"
            }
        }

        $userDoc = @{
            uid = $u.id
            email = $u.email
            name = $u.name
            role = $u.role
            status = "active"
            allowedCompanies = $allowedCompanies
        }

        Set-CachedAuth -Token $token -User $userDoc

        if ($InstanceId -and $CompanyCode -and -not (Test-UserHasAccess -User $userDoc -InstanceId $InstanceId -CompanyCode $CompanyCode)) {
            return @{ allowed = $false; reason = "Access denied to $InstanceId / $CompanyCode" }
        }

        return @{ allowed = $true; user = $userDoc }

    } catch {
        Write-ErrLog "PocketBase Token Verification Failed: $($_.Exception.Message)"
        return @{ allowed = $false; reason = "Invalid or expired token" }
    }
}

# ═══════════════════════════════════════════════════════
#  GET ACCESSIBLE INSTANCES
# ═══════════════════════════════════════════════════════
function Get-UserAccessibleInstances {
    param($User)

    $config = Get-InstancesConfig
    if (-not $config) { return @() }

    $result = @()
    foreach ($instance in $config.instances) {
        $accessibleCompanies = @()
        foreach ($company in $instance.companies) {
            if (Test-UserHasAccess -User $User -InstanceId $instance.id -CompanyCode $company.code) {
                $accessibleCompanies += @{ code = $company.code; name = $company.name }
            }
        }
        if ($accessibleCompanies.Count -gt 0) {
            $result += @{
                instanceId  = $instance.id
                displayName = $instance.displayName
                companies   = $accessibleCompanies
            }
        }
    }
    return $result
}