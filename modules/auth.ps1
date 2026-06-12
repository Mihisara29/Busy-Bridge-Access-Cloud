# modules/auth.ps1
# Local PocketBase Authentication Integration

. "$PSScriptRoot\utils.ps1"

$script:_authCache = @{}
$script:_AUTH_CACHE_TTL = 300

# ═══════════════════════════════════════════════════════
#  HELPER TO RESOLVE CONFIG AND SERVER-ID DYNAMICALLY
# ═══════════════════════════════════════════════════════
function Get-InstancesConfig {
    $instancesPath = "$PSScriptRoot\..\instances.json"
    if (-not (Test-Path $instancesPath)) { return $null }
    try {
        return Get-Content $instancesPath -Raw | ConvertFrom-Json
    } catch {
        return $null
    }
}

function Get-CachedAuth {
    param([string]$Token)
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
    $key = $Token.Substring([Math]::Max(0, $Token.Length - 32))
    $script:_authCache[$key] = @{ user = $User; expiry = (Get-Date).AddSeconds($script:_AUTH_CACHE_TTL) }
}

function Test-UserHasAccess {
    param($User, [string]$InstanceId, [string]$CompanyCode)
    if ($User.role -eq "superadmin") { return $true }
    if ($User.allowedCompanies -eq "all") { return $true }

    $instancesConfig = Get-InstancesConfig
    if (-not $instancesConfig) { return $false }
    
    $serverId = $instancesConfig.serverId
    foreach ($entry in $User.allowedCompanies) {
        if ($entry.serverId -eq $serverId -and $entry.instanceId -eq $InstanceId -and $entry.companyCodes -contains $CompanyCode) {
            return $true
        }
    }
    return $false
}

function Invoke-AuthCheck {
    param($Request, [string]$InstanceId = "", [string]$CompanyCode = "")

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
        # auth-refresh returns 200 OK + user data if token is valid, 401 if expired/invalid
        $response = Invoke-RestMethod -Uri "$pbUrl/api/collections/users/auth-refresh" -Method POST -Headers $headers -ErrorAction Stop
        
        $u = $response.record
        
        # Parse allowedCompanies (PocketBase stores this as a JSON field)
        $allowedCompanies = "all"
        if ($u.allowedCompanies -and $u.allowedCompanies -ne "all" -and $u.allowedCompanies -ne "") {
            $allowedCompanies = $u.allowedCompanies | ConvertFrom-Json
        }

        $userDoc = @{
            uid = $u.id
            email = $u.email
            name = $u.name
            role = $u.role
            status = if ($u.disabled) { "disabled" } else { "active" }
            allowedCompanies = $allowedCompanies
        }

        if ($userDoc.status -ne "active") { return @{ allowed = $false; reason = "Account disabled" } }

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