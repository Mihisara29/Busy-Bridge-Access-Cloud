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
        Write-Host `
            "  [ERROR] Module not found: $modulePath" `
            -ForegroundColor Red
        exit 1
    }

    return $modulePath
}

# ============================================================
# CORE CONFIG / UTILITIES
# ============================================================

. (Load-Module "config.ps1")
Write-Host "  Loaded: config.ps1" -ForegroundColor Gray

. (Load-Module "utils.ps1")
Write-Host "  Loaded: utils.ps1" -ForegroundColor Gray

. (Load-Module "connection.ps1")
Write-Host "  Loaded: connection.ps1" -ForegroundColor Gray

. (Load-Module "auth_native.ps1")
Write-Host "  Loaded: auth_native.ps1" -ForegroundColor Gray


# ============================================================
# MASTER DATA MODULES
# ============================================================

. (Load-Module "items.ps1")
Write-Host "  Loaded: items.ps1" -ForegroundColor Gray

. (Load-Module "accounts.ps1")
Write-Host "  Loaded: accounts.ps1" -ForegroundColor Gray

# Salesman / Sales Ref master and per-user assignment logic.
# BUSY stores Salesman records as Master1.MasterType = 19.
. (Load-Module "salesman.ps1")
Write-Host "  Loaded: salesman.ps1" -ForegroundColor Gray


# ============================================================
# VOUCHERS
# ============================================================

. (Load-Module "vouchers.ps1")
Write-Host "  Loaded: vouchers.ps1" -ForegroundColor Gray


# ============================================================
# BILL OF MATERIALS MASTER
#
# IMPORTANT:
# Load after vouchers.ps1.
#
# The dedicated BOM module contains:
# - Get-BomList
# - Get-BomDetail
# - Create-Bom
# - Update-Bom
# - Build-BomXml
#
# If vouchers.ps1 still contains older Get-BomList /
# Get-BomDetail functions, these dedicated BOM functions will
# become the active versions.
# ============================================================

. (Load-Module "boms.ps1")
Write-Host "  Loaded: boms.ps1" -ForegroundColor Gray


# ============================================================
# VOUCHER SETTINGS
# ============================================================

. (Load-Module "vch_setting.ps1")
Write-Host "  Loaded: vch_setting.ps1" -ForegroundColor Gray


# ============================================================
# REPORTS
# Reports must load before routes.ps1.
# ============================================================

. (Load-Module "reports.ps1")
Write-Host "  Loaded: reports.ps1" -ForegroundColor Gray


# ============================================================
# SCAN LOGIN
# ============================================================

. (Load-Module "scan_login.ps1")
Write-Host "  Loaded: scan_login.ps1" -ForegroundColor Gray


# ============================================================
# OFFLINE / LOCAL VOUCHER SYNCHRONIZATION
#
# This replaces the old split offline_sync_*.ps1 modules.
#
# Required file:
#   modules\offline_voucher_sync.ps1
#
# It owns:
# - Local voucher synchronization
# - Idempotency / duplicate protection
# - Durable sync tracking
# - Retry / recovery state handling
# - Final numbering at synchronization time
# - Posting through the existing Create-Voucher pipeline
# ============================================================

. (Load-Module "offline_voucher_sync.ps1")
Write-Host "  Loaded: offline_voucher_sync.ps1" -ForegroundColor Gray


# ============================================================
# IMPORTANT
#
# Restore the original non-strict behavior used by existing
# vouchers.ps1 and other BUSY modules.
#
# Without this, optional fields such as AltUnitReq can throw:
# "The property 'AltUnitReq' cannot be found on this object."
# ============================================================

Set-StrictMode -Off


# ============================================================
# ROUTES
#
# Load routes last because routes depend on all other functions.
# ============================================================

. (Load-Module "routes.ps1")
Write-Host "  Loaded: routes.ps1" -ForegroundColor Gray


# ============================================================
# STARTUP VALIDATION
# ============================================================

$requiredFunctions = @(

    # Reports
    "Get-OutstandingReport",
    "Get-StockStatusReport",

    # Item / Voucher
    "Get-VoucherItemDetail",
    "Create-Voucher",

    # BusyCloud Voucher Approval
    "Get-VoucherApprovalConfig",
    "Save-VoucherApprovalConfig",
    "Ensure-BusyCloudVoucherApprovalAuditTable",
    "Initialize-BusyCloudApprovalStorage",
    "Approve-Voucher",
    "Unapprove-Voucher",

    # Salesman / Sales Ref
    "Get-Salesmen",
    "Get-SalesmanAssignmentForAuthUser",
    "Resolve-LiveSalesmanMaster",
    "Apply-SalesmanAssignmentToVoucherData",

    # BOM Master
    "Get-BomList",
    "Get-BomDetail",
    "Create-Bom",
    "Update-Bom",

    # Offline / Local Sync
    "Invoke-OfflineVoucherSync",

    # Server
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

Write-Host `
    "  Required functions verified successfully." `
    -ForegroundColor Green


# ============================================================
# BUSYCLOUD APPROVAL STORAGE INITIALIZATION
#
# Idempotent: creates BusyCloudVoucherApprovalAudit only when it
# does not already exist in each configured SQL / Access company.
#
# A failure in one offline company is logged but does not prevent
# the API from starting; each approval write also performs a lazy
# ensure check for safety.
# ============================================================

Write-Host `
    "  Checking BusyCloud voucher approval storage..." `
    -ForegroundColor Cyan

$approvalStorage = Initialize-BusyCloudApprovalStorage -VerboseOutput $true

if ($approvalStorage.failed -gt 0) {
    Write-Host `
        "  [WARN] Approval storage initialization completed with $($approvalStorage.failed) company warning(s)." `
        -ForegroundColor DarkYellow
}
else {
    Write-Host `
        "  Approval storage verified for $($approvalStorage.initialized) company database(s)." `
        -ForegroundColor Green
}


# ============================================================
# START SERVER
# ============================================================

Start-BUSYServer -Port $Port
