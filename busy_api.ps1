# busy_api.ps1
# BUSY 21 Enterprise REST API Bridge v4.5
# Access / SQL Server Hybrid Version
# ─────────────────────────────────────────────────────────────
# IMPORTANT: Run in 32-bit execution context:
# C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe
# -ExecutionPolicy Bypass -File busy_api.ps1
# ─────────────────────────────────────────────────────────────

param(
    [int]$Port = 8081
)

$moduleDir = Join-Path $PSScriptRoot "modules"

function Load-Module {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $modulePath = Join-Path $moduleDir $Name

    if (-not (Test-Path $modulePath)) {
        Write-Host "  [ERROR] Module not found: $modulePath" `
            -ForegroundColor Red

        exit 1
    }

    return $modulePath
}

. (Load-Module "config.ps1")
Write-Host "  Loaded: config.ps1" -ForegroundColor Gray

. (Load-Module "utils.ps1")
Write-Host "  Loaded: utils.ps1" -ForegroundColor Gray

. (Load-Module "connection.ps1")
Write-Host "  Loaded: connection.ps1" -ForegroundColor Gray

. (Load-Module "auth_native.ps1")
Write-Host "  Loaded: auth_native.ps1" -ForegroundColor Gray

. (Load-Module "items.ps1")
Write-Host "  Loaded: items.ps1" -ForegroundColor Gray

. (Load-Module "accounts.ps1")
Write-Host "  Loaded: accounts.ps1" -ForegroundColor Gray

. (Load-Module "vouchers.ps1")
Write-Host "  Loaded: vouchers.ps1" -ForegroundColor Gray

. (Load-Module "vch_setting.ps1")
Write-Host "  Loaded: vch_setting.ps1" -ForegroundColor Gray

# Reports must load before routes.ps1.
. (Load-Module "reports.ps1")
Write-Host "  Loaded: reports.ps1" -ForegroundColor Gray

. (Load-Module "scan_login.ps1")
Write-Host "  Loaded: scan_login.ps1" -ForegroundColor Gray

# Offline synchronization modules.
# These files currently enable StrictMode internally.
# They are loaded here, then StrictMode is turned off again so older
# BUSY modules keep their previous behavior with optional/missing properties.
. (Load-Module "offline_sync_db.ps1")
Write-Host "  Loaded: offline_sync_db.ps1" -ForegroundColor Gray

. (Load-Module "offline_sync_store.ps1")
Write-Host "  Loaded: offline_sync_store.ps1" -ForegroundColor Gray

. (Load-Module "offline_sync_validation.ps1")
Write-Host "  Loaded: offline_sync_validation.ps1" -ForegroundColor Gray

. (Load-Module "offline_sync_idempotency.ps1")
Write-Host "  Loaded: offline_sync_idempotency.ps1" -ForegroundColor Gray

. (Load-Module "offline_sync_numbering.ps1")
Write-Host "  Loaded: offline_sync_numbering.ps1" -ForegroundColor Gray

. (Load-Module "offline_sync_logging.ps1")
Write-Host "  Loaded: offline_sync_logging.ps1" -ForegroundColor Gray

. (Load-Module "offline_sync.ps1")
Write-Host "  Loaded: offline_sync.ps1" -ForegroundColor Gray

# IMPORTANT:
# Restore the original non-strict behavior used by existing vouchers.ps1.
# Without this, optional fields such as AltUnitReq throw:
# "The property 'AltUnitReq' cannot be found on this object."
Set-StrictMode -Off

# Load routes last because routes depend on all other functions.
. (Load-Module "routes.ps1")
Write-Host "  Loaded: routes.ps1" -ForegroundColor Gray

# Startup validation
$requiredFunctions = @(
    "Get-OutstandingReport",
    "Get-StockStatusReport",
    "Get-VoucherItemDetail",
    "Create-Voucher",
    "Sync-OfflineVoucher",
    "Start-BUSYServer"
)

foreach ($functionName in $requiredFunctions) {
    if (-not (Get-Command $functionName -ErrorAction SilentlyContinue)) {
        Write-Host `
            "  [ERROR] Required function was not loaded: $functionName" `
            -ForegroundColor Red

        exit 1
    }
}

Write-Host "  Required functions verified successfully." `
    -ForegroundColor Green

# Start server listener
Start-BUSYServer -Port $Port