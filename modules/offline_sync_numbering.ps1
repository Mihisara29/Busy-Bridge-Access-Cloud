function Get-OfflineVoucherFinalNumber {param($Data,$InstanceConfig)
 if(Get-Command Get-NextVoucherNumber -ErrorAction SilentlyContinue){return Get-NextVoucherNumber -VchType ([int]$Data.vchType) -SeriesName ([string]$Data.seriesName) -VoucherDate ([string]$Data.voucherDate) -InstanceConfig $InstanceConfig}
 throw 'Connect Get-OfflineVoucherFinalNumber to your existing numbering function.'}
