# modules/config.ps1
# BUSY Configuration Module - Production Ready (Native & PocketBase Auth Version)
# ==============================================

$script:Config = @{
    BRIDGE_SECRET     = "YOUR-BRIDGE-SECRET-KEY-HERE"

    # Default fallbacks (overridden dynamically by instances.json)
    BUSY_PATH         = "C:\Enterprise02\"
    DATA_PATH         = "C:\Enterprise02\DATA\"
    BUSY_COM          = "Busy2L21.CFixedInterface"
    
    # 0 = MS Access Mode, 1 = SQL Server Mode (Dynamically overridden by instances.json)
    DB_TYPE           = 0  

    INSTANCE_ID       = "customer_induwara"
    COMP_CODE         = "COMP0001"
    PORT              = 8081
    DEBUG             = $true
    LOG_FILE          = "C:\busy-bridge\logs\busy_api.log"
    CACHE_TTL         = 300

    # PocketBase URL config used by modules/auth.ps1
    POCKETBASE_URL    = "http://127.0.0.1:8090"

    VOUCHER_TYPES = @{
        SALE=9; SALE_RETURN=3; PURCHASE=2; PURCHASE_RETURN=10; SALE_ORDER=12; PURCHASE_ORDER=13;
        SALE_QUOTATION=26; PURCHASE_QUOTATION=27; MATERIAL_RECEIPT=4; MATERIAL_ISSUE=11;
        STOCK_TRANSFER=5; STOCK_JOURNAL=8; PHYSICAL_STOCK=61; PRODUCTION=6; UNASSEMBLE=7;
        RECEIPT=14; PAYMENT=19; CONTRA=15; JOURNAL=16; DEBIT_NOTE=17; CREDIT_NOTE=18
    }

    MASTER_TYPES = @{
        ACCOUNT_GROUP=1; ACCOUNT=2; ITEM_GROUP=5; ITEM=6; UNIT=8; BILL_SUNDRY=9;
        MATERIAL_CENTER=11; SALE_TYPE=13; PURCHASE_TYPE=14; VOUCHER_SERIES=21; TAX_CATEGORY=25
    }

    VOUCHER_NAMES = @{
        2="Purchase"; 3="Sale Return"; 4="Material Receipt"; 5="Stock Transfer";
        6="Production"; 7="Unassemble"; 8="Stock Journal"; 9="Sale";
        10="Purchase Return"; 11="Material Issue"; 12="Sale Order";
        13="Purchase Order"; 14="Receipt"; 15="Contra"; 16="Journal";
        17="Debit Note"; 18="Credit Note"; 19="Payment"; 61="Physical Stock"
    }
}

function Get-Config { return $script:Config }
function Get-VoucherType { param([string]$Type); if ($script:Config.VOUCHER_TYPES.ContainsKey($Type)) { return $script:Config.VOUCHER_TYPES[$Type] } return $null }
function Get-MasterType { param([string]$Type); if ($script:Config.MASTER_TYPES.ContainsKey($Type)) { return $script:Config.MASTER_TYPES[$Type] } return $null }
function Get-VoucherName { param([int]$VchType); if ($script:Config.VOUCHER_NAMES.ContainsKey($VchType)) { return $script:Config.VOUCHER_NAMES[$VchType] } return "Unknown ($VchType)" }
function Get-VoucherXmlRoot {
    param([int]$VchType)
    $rootNames = @{ 2="Purchase"; 3="SaleReturn"; 4="MaterialReceipt"; 5="StockTransfer"; 6="Production"; 7="Unassemble"; 8="StockJournal"; 9="Sale"; 10="PurchaseReturn"; 11="MaterialIssue"; 12="SaleOrder"; 13="PurchaseOrder"; 14="Receipt"; 15="Contra"; 16="Journal"; 17="DebitNote"; 18="CreditNote"; 19="Payment"; 61="PhysicalStock" }
    if ($rootNames.ContainsKey($VchType)) { return $rootNames[$VchType] }
    return "Voucher"
}
function Test-BusyConnection { param($fi); return ($null -ne $fi) }