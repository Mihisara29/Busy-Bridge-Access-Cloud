# busy_api.ps1
# BUSY 21 Enterprise REST API Bridge v4.5
# Access / SQL Server Hybrid Version
# ─────────────────────────────────────────────────────────────
# IMPORTANT: Run in 32-bit execution context:
# C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe
# -ExecutionPolicy Bypass -File busy_api.ps1
# ─────────────────────────────────────────────────────────────

param(
    [int]$Port = 8081,
    [ValidateSet("Main", "Reports")]
    [string]$ServiceMode = "Main"
)
# BUSYWEB_REPORT_PORT_SPLIT_V1
$script:BusyBridgeServiceMode = $ServiceMode

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
# BUSYCLOUD WEB APPROVAL
#
# Loads after vch_setting.ps1 because it reuses the established
# active financial-year database resolver.
# ============================================================

. (Load-Module "web_approval.ps1")
Write-Host "  Loaded: web_approval.ps1" -ForegroundColor Gray

if (-not $script:BusyCloudWebApprovalModuleVersion) {
    Write-Host "  [ERROR] web_approval.ps1 is an older/incompatible build. Expected Web Approval manager/sync module." -ForegroundColor Red
    exit 1
}

Write-Host "  [WEB-APPROVAL] Active module version: $script:BusyCloudWebApprovalModuleVersion" -ForegroundColor DarkCyan


# ============================================================
# REPORTS
# Reports must load before routes.ps1.
# ============================================================

. (Load-Module "reports.ps1")
Write-Host "  Loaded: reports.ps1" -ForegroundColor Gray

. (Load-Module "account_ledger.ps1")
Write-Host "  Loaded: account_ledger.ps1" -ForegroundColor Gray


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

# Executive reports are isolated and must be present before routes are loaded.
. (Load-Module "dashboard.ps1")
Write-Host "  Loaded: dashboard.ps1" -ForegroundColor Gray
. (Load-Module "executive_insights.ps1")
Write-Host "  Loaded: executive_insights.ps1" -ForegroundColor Gray
. (Load-Module "executive_jobs.ps1")
Write-Host "  Loaded: executive_jobs.ps1" -ForegroundColor Gray
. (Load-Module "executive_automation.ps1")
Write-Host "  Loaded: executive_automation.ps1" -ForegroundColor Gray
if (-not (Get-Command Get-ExecutiveDashboardAsync -CommandType Function -ErrorAction SilentlyContinue)) {
    throw 'Executive job manager did not define Get-ExecutiveDashboardAsync.'
}
if (-not (Get-Command Get-ExecutiveSalesInsights -CommandType Function -ErrorAction SilentlyContinue)) {
    throw 'Executive insights module did not define Get-ExecutiveSalesInsights.'
}
. (Load-Module "dashboard_permissions.ps1")
Write-Host "  Loaded: dashboard_permissions.ps1" -ForegroundColor Gray
. (Load-Module "whatsapp_jobs.ps1")
Write-Host "  Loaded: whatsapp_jobs.ps1" -ForegroundColor Gray

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

    # BusyCloud Web Approval
    "Ensure-BusyCloudWebApprovalStorage",
    "Initialize-BusyCloudWebApprovalStorage",
    "Get-WebApprovalRoleDirectory",
    "Set-WebApprovalUserRole",
    "Get-WebApprovalManagersForSalesman",
    "Set-WebApprovalManagersForSalesman",
    "Invoke-WebApprovalVoucherCreateRouting",
    "Submit-WebApprovalVoucher",
    "Get-WebApprovalMySubmissions",
    "Get-WebApprovalNotifications",
    "Get-WebApprovalNumberChangeReason",
    "Set-WebApprovalNumberAudit",
    "Get-WebApprovalAssignedQueue",
    "Get-WebApprovalDetailForManager",
    "Approve-WebApprovalVoucher",
    "Reject-WebApprovalVoucher",
    "Sync-WebApprovalVoucher",

    # Persistent / Web Push Notifications
    "Get-WebPushPublicConfig",
    "Get-WebPushSubscriptionStatus",
    "Save-WebPushSubscription",
    "Disable-WebPushSubscription",
    "Set-WebApprovalNotificationRead",
    "Test-WebPushWorkerAuthorized",
    "Get-WebPushPendingJobsAllCompanies",
    "Complete-WebPushDelivery",
    "Initialize-BusyCloudPushRuntime",
    "Start-BusyCloudPushWorker",

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


if ($ServiceMode -eq 'Main') {
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
# BUSYCLOUD WEB APPROVAL STORAGE INITIALIZATION
#
# Idempotent automatic migration:
# - Permanent company DB:
#     MobileUserPreference.B35/B36
#     BusyCloudSalesmanManager
#     BusyCloudPushSubscription
# - Active financial-year DB:
#     BusyCloudWebApproval
#     BusyCloudWebApprovalAction
#     BusyCloudWebSyncLedger
#     BusyCloudNotification
#     BusyCloudNotificationDelivery
#
# This means a fresh VPS deployment does not require manual table
# creation. Existing tables are detected and left intact.
# ============================================================

Write-Host `
    "  Checking BusyCloud Web Approval storage..." `
    -ForegroundColor Cyan

$webApprovalStorage = Initialize-BusyCloudWebApprovalStorage -VerboseOutput $true

if ($webApprovalStorage.failed -gt 0) {
    Write-Host `
        "  [WARN] Web Approval storage initialization completed with $($webApprovalStorage.failed) company warning(s)." `
        -ForegroundColor DarkYellow
}
else {
    Write-Host `
        "  Web Approval storage verified for $($webApprovalStorage.initialized) company database(s)." `
        -ForegroundColor Green
}


# ============================================================
# WEB PUSH WORKER
#
# The backend automatically bootstraps Web Push on first start:
# portable Node.js (if needed), npm dependency, VAPID keys and
# worker secret. If bootstrap cannot complete (for example no
# Internet access), the API still starts normally and in-app
# notifications continue to work.
# ============================================================

$pushWorker = Start-BusyCloudPushWorker -ApiPort $Port

if ($pushWorker.started) {
    Write-Host `
        "  Web Push worker started (PID $($pushWorker.pid))." `
        -ForegroundColor Green
}
elseif ($pushWorker.alreadyRunning) {
    Write-Host `
        "  Web Push worker already running (PID $($pushWorker.pid))." `
        -ForegroundColor DarkGreen
}
elseif ($pushWorker.configured -eq $false) {
    Write-Host `
        "  [INFO] Web Push worker not started: $($pushWorker.reason)" `
        -ForegroundColor DarkYellow
}
elseif ($pushWorker.success -eq $false) {
    Write-Host `
        "  [WARN] Web Push worker failed to start: $($pushWorker.error)" `
        -ForegroundColor DarkYellow
}
else {
    Write-Host `
        "  [INFO] Web Push worker not started: $($pushWorker.reason)" `
        -ForegroundColor DarkYellow
}



} else {
    Write-Host "  [REPORTS 8085] Skipping approval migrations and Web Push startup." -ForegroundColor DarkCyan
}

# ============================================================
# START SERVER
# ============================================================

Start-BUSYServer -Port $Port
