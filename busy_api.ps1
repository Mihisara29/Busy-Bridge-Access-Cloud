# busy_api.ps1
# BUSY 21 Enterprise REST API Bridge v4.5 (Access/SQL Server Hybrid Version)
# ─────────────────────────────────────────────────────────────
# IMPORTANT: Run in 32-bit execution context:
#   C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe
#   -ExecutionPolicy Bypass -File busy_api.ps1
# ─────────────────────────────────────────────────────────────

param([int]$Port = 8081)

$moduleDir = Join-Path $PSScriptRoot "modules"

function Load-Module([string]$name) {
    $p = Join-Path $moduleDir $name
    if (-not (Test-Path $p)) {
        Write-Host "  [ERROR] Module not found: $p" -ForegroundColor Red
        exit 1
    }
    return $p
}

. (Load-Module "config.ps1")       ; Write-Host "  Loaded: config.ps1"       -ForegroundColor Gray
. (Load-Module "utils.ps1")        ; Write-Host "  Loaded: utils.ps1"        -ForegroundColor Gray
. (Load-Module "connection.ps1")   ; Write-Host "  Loaded: connection.ps1"   -ForegroundColor Gray
. (Load-Module "auth_native.ps1")  ; Write-Host "  Loaded: auth_native.ps1"  -ForegroundColor Gray
. (Load-Module "items.ps1")        ; Write-Host "  Loaded: items.ps1"        -ForegroundColor Gray
. (Load-Module "accounts.ps1")     ; Write-Host "  Loaded: accounts.ps1"     -ForegroundColor Gray
. (Load-Module "vouchers.ps1")     ; Write-Host "  Loaded: vouchers.ps1"     -ForegroundColor Gray
. (Load-Module "vch_setting.ps1")  ; Write-Host "  Loaded: vch_setting.ps1"  -ForegroundColor Gray
. (Load-Module "scan_login.ps1")   ; Write-Host "  Loaded: scan_login.ps1"   -ForegroundColor Gray
. (Load-Module "routes.ps1")       ; Write-Host "  Loaded: routes.ps1"       -ForegroundColor Gray

# Start server listener directly (No delegates needed anymore)
Start-BUSYServer -Port $Port