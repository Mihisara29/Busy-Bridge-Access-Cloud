# modules/utils.ps1
# Utility / Logging + In-Memory Cache

if ($null -eq $script:Config) {
    . "$PSScriptRoot\config.ps1"
}

# ── Logging ───────────────────────────────────────────────────────────────────

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$Color = "White"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry  = "$timestamp [$Level] $Message"
    Write-Host $logEntry -ForegroundColor $Color

    try {
        $logDir = Split-Path $script:Config.LOG_FILE -Parent
        if (-not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        Add-Content -Path $script:Config.LOG_FILE -Value $logEntry
    } catch {
        # Silently ignore log-file write failures
    }
}

function Write-ErrLog     { param($Message) Write-Log $Message "ERROR"   "Red"    }
function Write-SuccessLog { param($Message) Write-Log $Message "SUCCESS" "Green"  }
function Write-WarnLog    { param($Message) Write-Log $Message "WARNING" "Yellow" }
function Write-DebugLog   {
    param($Message)
    if ($script:Config.DEBUG) { Write-Log $Message "DEBUG" "Gray" }
}

# ── In-Memory Cache ───────────────────────────────────────────────────────────
# Structure: $script:_cache = @{ key = @{ data = ...; expires = [datetime] } }

$script:_cache = @{}

function Get-Cache {
    param([string]$Key)
    if ($script:Config.CACHE_TTL -le 0) { return $null }
    if (-not $script:_cache.ContainsKey($Key)) { return $null }
    $entry = $script:_cache[$Key]
    if ((Get-Date) -gt $entry.expires) {
        $script:_cache.Remove($Key)
        return $null
    }
    Write-DebugLog "Cache HIT: $Key"
    return $entry.data
}

function Set-Cache {
    param([string]$Key, $Value)
    if ($script:Config.CACHE_TTL -le 0) { return }
    $script:_cache[$Key] = @{
        data    = $Value
        expires = (Get-Date).AddSeconds($script:Config.CACHE_TTL)
    }
    Write-DebugLog "Cache SET: $Key (TTL=$($script:Config.CACHE_TTL)s)"
}

function Clear-Cache {
    param([string]$Key = "")
    if ($Key -eq "") {
        $script:_cache = @{}
        Write-DebugLog "Cache CLEARED (all)"
    } else {
        $script:_cache.Remove($Key)
        Write-DebugLog "Cache CLEARED: $Key"
    }
}