# =======================================================
# STANDALONE BUSY SQL UPDATE TEST SCRIPT V4
# =======================================================

if ([Environment]::Is64BitProcess) {
    Write-Host "Switching to 32-bit PowerShell..." -ForegroundColor Yellow
    Start-Sleep -Seconds 1
    Start-Process "$env:windir\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -NoProfile -File `"$PSCommandPath`"" -Wait
    exit
}

$busyPath = "C:\Enterprise02"
$dataPath = "C:\Enterprise02\DATA"
$companyCode = "COMP0002"
$voucherNo = "f111"

Write-Host "1. Initializing BUSY COM Object (32-bit)..." -ForegroundColor Cyan
$fi = New-Object -ComObject "Busy2L21.CFixedInterface"
$connected = $fi.OpenDB($busyPath, $dataPath, $companyCode, 0)

if ($connected -eq $true) {
    Write-Host "   -> Connection Successful!`n" -ForegroundColor Green

    $vchCode = 0
    $safeNo = $voucherNo.ToLower().Trim()
    
    # 1. Find the Voucher Code
    $rs = $fi.GetRecordset("SELECT VchCode FROM Tran1 WHERE VchType=9 AND VchNo LIKE '*$safeNo*'")
    
    if ($rs -and -not $rs.EOF) {
        $rs.MoveFirst()
        $vchCode = $rs.Fields.Item("VchCode").Value
        $rs.Close()
        
        Write-Host "Target VchCode: $vchCode`n" -ForegroundColor Cyan

        # ---------------------------------------------------------
        # EXECUTE QUERY: Update UserName
        # ---------------------------------------------------------
        Write-Host "--- Executing Update via ExecuteQuery ---" -ForegroundColor Yellow
        try {
            $fi.ExecuteQuery("UPDATE CheckList SET UserName='Test777' WHERE Code=$vchCode")
            Write-Host "ExecuteQuery command sent successfully." -ForegroundColor Green
        } catch {
            Write-Host "Error executing query: $_" -ForegroundColor Red
        }
        Write-Host ""
        
        # ---------------------------------------------------------
        # VERIFICATION: Read back the Username
        # ---------------------------------------------------------
        Write-Host "--- Verifying Update in CheckList Table ---" -ForegroundColor Yellow
        $check = $fi.GetRecordset("SELECT UserName FROM CheckList WHERE Code=$vchCode")
        
        if ($check -and -not $check.EOF) {
            $val = $check.Fields.Item("UserName").Value.ToString().Trim()
            Write-Host "Success! UserName for VchCode $vchCode is now: '$val'" -ForegroundColor Green
            $check.Close()
        } else {
            Write-Host "Failed! Could not find a CheckList record for this VchCode." -ForegroundColor Red
        }
        Write-Host ""

        # ---------------------------------------------------------
        # NEW: Read back the Voucher Number using the VchCode
        # ---------------------------------------------------------
        Write-Host "--- Fetching Voucher Details from Tran1 Table ---" -ForegroundColor Yellow
        # We use Tran1 because CheckList only stores the VchCode, not the actual Voucher Number
        $vchInfo = $fi.GetRecordset("SELECT VchNo, Date, CreatedBy FROM Tran1 WHERE VchCode=$vchCode")
        
        if ($vchInfo -and -not $vchInfo.EOF) {
            $dbVchNo = $vchInfo.Fields.Item("VchNo").Value.ToString().Trim()
            $dbDate = $vchInfo.Fields.Item("Date").Value
            $dbCreatedBy = $vchInfo.Fields.Item("CreatedBy").Value.ToString().Trim()
            
            Write-Host "Data retrieved for VchCode: $vchCode" -ForegroundColor Cyan
            Write-Host "  -> Voucher Number : $dbVchNo" -ForegroundColor White
            Write-Host "  -> Voucher Date   : $dbDate" -ForegroundColor White
            Write-Host "  -> Created By     : $dbCreatedBy" -ForegroundColor White
            
            $vchInfo.Close()
        } else {
            Write-Host "Error: Could not find Voucher details for VchCode $vchCode in Tran1." -ForegroundColor Red
        }
        Write-Host ""
        
    } else {
        Write-Host "Voucher not found! Make sure '$voucherNo' exists." -ForegroundColor Red
    }
    
    $fi.CloseDB()
} else {
    Write-Host "Failed to connect to database." -ForegroundColor Red
}

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($fi) | Out-Null
Write-Host "Press any key to exit..."
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null