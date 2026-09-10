# 00-config.ps1
# Voucher type configuration shared by voucher components.

$script:VoucherConfig = @{
    9  = @{ xmlRoot = "Sale";           typeField = "STPTName"; typeDataKey = "saleType";     hasBillNo = $true;  isAccounting = $false; requiredKeys = @("vchSeries","date","party","saleType","items") }
    2  = @{ xmlRoot = "Purchase";       typeField = "STPTName"; typeDataKey = "purchaseType"; hasBillNo = $true;  isAccounting = $false; requiredKeys = @("vchSeries","date","party","purchaseType","items") }
    3  = @{ xmlRoot = "SaleReturn";     typeField = "STPTName"; typeDataKey = "saleType";     hasBillNo = $false; isAccounting = $false; requiredKeys = @("vchSeries","date","party","saleType","items") }
    10 = @{ xmlRoot = "PurchaseReturn"; typeField = "STPTName"; typeDataKey = "purchaseType"; hasBillNo = $false; isAccounting = $false; requiredKeys = @("vchSeries","date","party","purchaseType","items") }
    11 = @{ xmlRoot = "MaterialIssue";   typeField = "STPTName"; typeDataKey = "saleType";     hasBillNo = $false; isAccounting = $false; requiredKeys = @("vchSeries","date","party","saleType","items") }
    4  = @{ xmlRoot = "MaterialReceipt"; typeField = "STPTName"; typeDataKey = "purchaseType"; hasBillNo = $true;  isAccounting = $false; requiredKeys = @("vchSeries","date","party","purchaseType","items") }
    12 = @{ xmlRoot = "SaleOrder";      typeField = "STPTName"; typeDataKey = "saleType";     hasBillNo = $false; isAccounting = $false; requiredKeys = @("vchSeries","date","party","saleType","items") }
    13 = @{ xmlRoot = "PurchaseOrder";  typeField = "STPTName"; typeDataKey = "purchaseType"; hasBillNo = $true;  isAccounting = $false; requiredKeys = @("vchSeries","date","party","purchaseType","items") }
    14 = @{ xmlRoot = "Receipt";        hasBillNo = $false; isAccounting = $true;  requiredKeys = @("vchSeries","date","accounts") }
    15 = @{ xmlRoot = "Contra";         hasBillNo = $false; isAccounting = $true;  requiredKeys = @("vchSeries","date","accounts") }
    16 = @{ xmlRoot = "Journal";        hasBillNo = $false; isAccounting = $true;  requiredKeys = @("vchSeries","date","accounts") }
    19 = @{ xmlRoot = "Payment";        hasBillNo = $false; isAccounting = $true;  requiredKeys = @("vchSeries","date","accounts") }
    5  = @{ xmlRoot = "StockTransfer";  typeField = "VchSeriesName"; typeDataKey = "vchSeries"; hasBillNo = $false; isAccounting = $false; requiredKeys = @("vchSeries","date","party","items") }
    8  = @{ xmlRoot = "StockJournal";   typeField = "VchSeriesName"; typeDataKey = "vchSeries"; hasBillNo = $false; isAccounting = $false; requiredKeys = @("vchSeries","date","party","items") }
    6 = @{xmlRoot = "Production";  typeField = "VchSeriesName";  typeDataKey = "vchSeries";  hasBillNo = $false;  isAccounting = $false;
    requiredKeys = @("vchSeries","date","items")
    }
    26 = @{ xmlRoot = "SaleQuotation";     typeField = "STPTName"; typeDataKey = "saleType";     hasBillNo = $true; isAccounting = $false; requiredKeys = @("vchSeries","date","party","saleType","items") }
    27 = @{ xmlRoot = "PurchaseQuotation"; typeField = "STPTName"; typeDataKey = "purchaseType"; hasBillNo = $true; isAccounting = $false; requiredKeys = @("vchSeries","date","party","purchaseType","items") }

}
