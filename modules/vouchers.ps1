# modules/vouchers.ps1
# Voucher Management - Modular Loader
#
# This file intentionally stays small. Existing callers continue loading
# modules/vouchers.ps1 exactly as before; the implementation is split into
# feature-focused scripts under modules/vouchers/.

if ($null -eq $script:Config) {
    . "$PSScriptRoot\connection.ps1"
    . "$PSScriptRoot\utils.ps1"
    . "$PSScriptRoot\config.ps1"
}

$voucherComponentRoot = Join-Path $PSScriptRoot "vouchers"
$voucherComponents = @(
    "00-config.ps1",
    "01-core.ps1",
    "02-optional-fields.ps1",
    "03-numbering.ps1",
    "04-xml-items.ps1",
    "05-xml-accounting.ps1",
    "06-xml-builders.ps1",
    "07-lookups.ps1",
    "08-voucher-list.ps1",
    "09-item-cost.ps1",
    "10-voucher-details.ps1",
    "11-returns.ps1",
    "12-bom.ps1",
    "13-approval.ps1",
    "14-commands.ps1"
)

foreach ($component in $voucherComponents) {
    $componentPath = Join-Path $voucherComponentRoot $component
    if (-not (Test-Path -LiteralPath $componentPath)) {
        throw "Voucher component not found: $componentPath"
    }
    try {
        . $componentPath
    }
    catch {
        throw "Failed to load voucher component '$component': $($_.Exception.Message)"
    }
}

# Fail fast at startup if the public voucher surface was not loaded.
$requiredVoucherFunctions = @(
    "Create-Voucher",
    "Modify-Voucher",
    "Delete-Voucher",
    "Get-Vouchers",
    "Get-VoucherDetail",
    "Get-AccountVoucherDetail",
    "Get-VoucherApprovalQueue",
    "Approve-Voucher",
    "Unapprove-Voucher"
)

foreach ($functionName in $requiredVoucherFunctions) {
    if (-not (Get-Command $functionName -CommandType Function -ErrorAction SilentlyContinue)) {
        throw "Required voucher function was not loaded: $functionName"
    }
}
