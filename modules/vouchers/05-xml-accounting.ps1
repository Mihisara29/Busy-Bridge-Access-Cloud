# 05-xml-accounting.ps1
# Extracted from vouchers.ps1. Keep functions behavior-compatible with the monolith.

function Build-BillSundriesXml {
    param($billSundries)

    if (-not $billSundries -or $billSundries.Count -eq 0) {
        return ""
    }

    $xml = "<BillSundries>"
    foreach ($bs in $billSundries) {
        $xml += "<BSDetail>"
        $xml += "<SrNo>$($bs.srNo)</SrNo>"
        $xml += "<BSName>$([System.Security.SecurityElement]::Escape($bs.name))</BSName>"
        $xml += "<Amt>$($bs.amount)</Amt>"
        if ($bs.percentVal) {
            $xml += "<PercentVal>$($bs.percentVal)</PercentVal>"
        }
        $xml += "</BSDetail>"
    }
    $xml += "</BillSundries>"

    return $xml
}

function Build-AccEntriesXml {
    param($items, [string]$PartyName, [double]$TotalAmt,[int]$VchType, $Settlements)

    $salesAcc = "Sales"
    $purcAcc = "Purchase"
    
    foreach ($item in $items) {
        if ($item.salesAccount -and $item.salesAccount -ne "") {
            $salesAcc = $item.salesAccount
        }
        if ($item.purchaseAccount -and $item.purchaseAccount -ne "") {
            $purcAcc = $item.purchaseAccount
        }
    }

    $cashAmt = 0; $cardAmt = 0; $giftAmt = 0
    $cashAcc = "Cash"; $cardAcc = ""; $giftAcc = ""
    
    if ($null -ne $Settlements) {
        if ($Data.settlements.cash -and $Data.settlements.cash.amount) {
            $cashAmt = [double]$Data.settlements.cash.amount
            if ($Data.settlements.cash.account) { $cashAcc = $Data.settlements.cash.account }
        }
        if ($Data.settlements.card -and $Data.settlements.card.amount) {
            $cardAmt = [double]$Data.settlements.card.amount
            if ($Data.settlements.card.account) { $cardAcc = $Data.settlements.card.account }
        }
        if ($Data.settlements.gift -and $Data.settlements.gift.amount) {
            $giftAmt = [double]$Data.settlements.gift.amount
            if ($Data.settlements.gift.account) { $giftAcc = $Data.settlements.gift.account }
        }
    }

    $totalSettled = $cashAmt + $cardAmt + $giftAmt
    $partyBalance = $TotalAmt - $totalSettled

    $xml = "<AccEntries>"
    
    if ($VchType -eq 9 -or $VchType -eq 3) {
        
                        # 1. Dynamic assignments to handle Dr/Cr directions cleanly
            if ($VchType -eq 9) {
                # Sale Invoice: Party Dr (1, Negative), Sales Cr (2, Positive)
                $partyAmtType = 1; $partyAmt = -$partyBalance
                $saleAmtType  = 2; $saleAmt  = $TotalAmt
            } else {
                # Sale Return: Party Cr (2, Positive), Sales Dr (1, Negative)
                $partyAmtType = 2; $partyAmt = $partyBalance
                $saleAmtType  = 1; $saleAmt  = -$TotalAmt
            }
        
           # 2. Customer Party AccDetail (updated to use $partyAmtType)
            if ([Math]::Round($partyBalance, 2) -gt 0 -or $totalSettled -eq 0) {
                $xml += "<AccDetail>"
                $xml += "<AccountName>$([System.Security.SecurityElement]::Escape($PartyName))</AccountName>"
                $xml += "<AmountType>$partyAmtType</AmountType>"
                $xml += "<AmtMainCur>$($partyAmt.ToString('0.##',[System.Globalization.CultureInfo]::InvariantCulture))</AmtMainCur>"
                $xml += "</AccDetail>"
            }
        
                $xml += "<AccDetail>"
                $xml += "<AccountName>$([System.Security.SecurityElement]::Escape($salesAcc))</AccountName>"
                $xml += "<AmountType>$saleAmtType</AmountType>"
                $xml += "<AmtMainCur>$($saleAmt.ToString('0.##',[System.Globalization.CultureInfo]::InvariantCulture))</AmtMainCur>"
                $xml += "</AccDetail>"

        if ($cashAmt -gt 0 -and $cashAcc) {
            $cAmt = -$cashAmt
            $xml += "<AccDetail><AccountName>$([System.Security.SecurityElement]::Escape($cashAcc))</AccountName><AmountType>1</AmountType><AmtMainCur>$($cAmt.ToString('0.##',[System.Globalization.CultureInfo]::InvariantCulture))</AmtMainCur></AccDetail>"
        }
        if ($cardAmt -gt 0 -and $cardAcc) {
            $cAmt = -$cardAmt
            $xml += "<AccDetail><AccountName>$([System.Security.SecurityElement]::Escape($cardAcc))</AccountName><AmountType>1</AmountType><AmtMainCur>$($cAmt.ToString('0.##',[System.Globalization.CultureInfo]::InvariantCulture))</AmtMainCur></AccDetail>"
        }
        if ($giftAmt -gt 0 -and $giftAcc) {
            $cAmt = -$giftAmt
            $xml += "<AccDetail><AccountName>$([System.Security.SecurityElement]::Escape($giftAcc))</AccountName><AmountType>1</AmountType><AmtMainCur>$($cAmt.ToString('0.##',[System.Globalization.CultureInfo]::InvariantCulture))</AmtMainCur></AccDetail>"
        }
    } elseif ($VchType -eq 2 -or $VchType -eq 10) {

            # 1. Dynamic assignments to handle Dr/Cr directions cleanly
    if ($VchType -eq 2) {
        # Purchase: Party Cr (2, Positive), Purchase Dr (1, Negative)
        $partyAmtType = 2; $partyAmtVal = $TotalAmt
        $purcAmtType  = 1; $purcAmtVal  = -$TotalAmt
    } else {
        # Purchase Return: Party Dr (1, Negative), Purchase Cr (2, Positive)
        $partyAmtType = 1; $partyAmtVal = -$TotalAmt
        $purcAmtType  = 2; $purcAmtVal  = $TotalAmt
    }
    
        $negAmt = -$TotalAmt
        
        # 2. Supplier Party AccDetail
        $xml += "<AccDetail>"
        $xml += "<AccountName>$([System.Security.SecurityElement]::Escape($PartyName))</AccountName>"
        $xml += "<AmountType>$partyAmtType</AmountType>"
        $xml += "<AmtMainCur>$($partyAmtVal.ToString('0.##',[System.Globalization.CultureInfo]::InvariantCulture))</AmtMainCur>"
        $xml += "</AccDetail>"
        
        # 3. Purchase Account AccDetail
        $xml += "<AccDetail>"
        $xml += "<AccountName>$([System.Security.SecurityElement]::Escape($purcAcc))</AccountName>"
        $xml += "<AmountType>$purcAmtType</AmountType>"
        $xml += "<AmtMainCur>$($purcAmtVal.ToString('0.##',[System.Globalization.CultureInfo]::InvariantCulture))</AmtMainCur>"
        $xml += "</AccDetail>"
    }
    
    $xml += "</AccEntries>"
    return $xml
}

function Build-PendingBillDetailsXml {
    param($refEntries, [string]$PartyName, [string]$VchDate, [int]$VchType,[string]$VchNo)

    if ($null -eq $refEntries) { return "" }
    $refs = @($refEntries)
    if ($refs.Count -eq 0) { return "" }

    $xml = "<BillDetail>"
    $xml += "<MasterName1>$([System.Security.SecurityElement]::Escape($PartyName))</MasterName1>"

    foreach ($ref in $refs) {
        $rNo = if ($ref.refNo -eq "AUTO" -or [string]::IsNullOrWhiteSpace($ref.refNo)) { $VchNo } else { $ref.refNo }
        $dueDate = if ($ref.dueDate) { $ref.dueDate } else { $VchDate }
        
        $val = [double]$ref.amount
        if ($VchType -eq 9 -or $VchType -eq 10) {
            $val = -($val)
        }
        $valStr = $val.ToString("0.00",[System.Globalization.CultureInfo]::InvariantCulture)

        $xml += "<BillRefs>"
        $xml += "<Method>$($ref.refType)</Method>"
        $xml += "<SrNo>$($ref.srNo)</SrNo>"
        $xml += "<RefNo>$([System.Security.SecurityElement]::Escape($rNo))</RefNo>"
        $xml += "<Date>$VchDate</Date>"
        $dueDate = if ($ref.dueDate) { $ref.dueDate } else { $VchDate }
        $xml += "<DueDate>$dueDate</DueDate>"
        $xml += "<Value1>$valStr</Value1>"
        $xml += "<VchType>$VchType</VchType>"
        $xml += "<MfgDate>$VchDate</MfgDate>"
        $xml += "</BillRefs>"
    }

    $xml += "</BillDetail>"
    return $xml
}

function Build-PendingOrdersXml {
    param($items, [string]$PartyName, $linkedOrder, [int]$VchType, [string]$VchDate = "")

    if (-not $linkedOrder -or -not $linkedOrder.vchNo) {
        return ""
    }

    $xml = "<PendingOrders>"
    foreach ($item in $items) {
        $qty = [double]$item.qty
        $altQty = if ($item.altQty) { [double]$item.altQty } else { $qty }
        $price = if ($item.price) { [double]$item.price } else { 0 }
        $altPrice = if ($item.altPrice) { [double]$item.altPrice } else { $price }
        
        $qtyStr = (-$qty).ToString("0.000",[System.Globalization.CultureInfo]::InvariantCulture)
        $altQtyStr = (-$altQty).ToString("0.000",[System.Globalization.CultureInfo]::InvariantCulture)

        $xml += "<OrderDetail>"
        $xml += "<MasterName1>$([System.Security.SecurityElement]::Escape($item.itemName))</MasterName1>"
        $xml += "<MasterName2>$([System.Security.SecurityElement]::Escape($PartyName))</MasterName2>"
        $xml += "<OrderRefs>"
        $xml += "<Method>2</Method>"
        $xml += "<SrNo>1</SrNo>"
        $xml += "<RefNo>$([System.Security.SecurityElement]::Escape($linkedOrder.vchNo))</RefNo>"
        $xml += "<Date>$($linkedOrder.date)</Date>"
        $dueDate = if ($linkedOrder.date) { $linkedOrder.date } else { $VchDate }
        $xml += "<DueDate>$dueDate</DueDate>"
        $xml += "<Value1>$qtyStr</Value1>"
        $xml += "<Value2>$altQtyStr</Value2>"
        $xml += "<MainTranPrice>$price</MainTranPrice>"
        $xml += "<VchType>$VchType</VchType>"
        $xml += "<ItemSrNo>$($item.srNo)</ItemSrNo>"
        $xml += "<AltTranPrice>$altPrice</AltTranPrice>"
        $xml += "<MfgDate>$($linkedOrder.date)</MfgDate>"
        $xml += "</OrderRefs>"
        $xml += "</OrderDetail>"
    }
    
    $xml += "</PendingOrders>"
    return $xml
}

function Build-PendingChallansXml {
    param($items, [string]$PartyName, $linkedChallans, [int]$VchType, [string]$VchNo, [string]$VchDate, $fi)

    if ($VchType -eq 11 -or $VchType -eq 4) {
        $xml = "<PendingChallans>"
        foreach ($item in $items) {
            $qty        = [double]$item.qty
            $conFactor  = if ($item.altQtyConFactor -and [double]$item.altQtyConFactor -ne 0) { [double]$item.altQtyConFactor } else { 1 }
            $altQty     = if ($item.altQty -and [double]$item.altQty -ne 0) { [double]$item.altQty } else { [Math]::Round($qty * $conFactor, 3) }
            $price      = [double]$item.price
            $altPrice   = if ($altQty -ne 0) { [Math]::Round(($price * $qty) / $altQty, 4) } else { $price }

            $qtyStr     = (-$qty).ToString("0.000", [System.Globalization.CultureInfo]::InvariantCulture)
            $altQtyStr  = (-$altQty).ToString("0.000", [System.Globalization.CultureInfo]::InvariantCulture)
            $paddedVchNo = Pad-VchNo $VchNo

            $xml += "<ChallanDetail>"
            $xml += "<MasterName1>$([System.Security.SecurityElement]::Escape($item.itemName))</MasterName1>"
            $xml += "<MasterName2>$([System.Security.SecurityElement]::Escape($PartyName))</MasterName2>"
            $xml += "<ChallanRefs>"
            $xml += "<Method>1</Method>"
            $xml += "<SrNo>1</SrNo>"
            $xml += "<RefNo>$([System.Security.SecurityElement]::Escape($paddedVchNo))</RefNo>"
            $xml += "<Date>$VchDate</Date>"
            $xml += "<DueDate>$VchDate</DueDate>"
            $xml += "<Value1>$qtyStr</Value1>"
            $xml += "<Value2>$altQtyStr</Value2>"
            $xml += "<MainTranPrice>$price</MainTranPrice>"
            $xml += "<VchType>$VchType</VchType>"
            $xml += "<ItemSrNo>$($item.srNo)</ItemSrNo>"
            $xml += "<AltTranPrice>$altPrice</AltTranPrice>"
            $xml += "<MfgDate>$VchDate</MfgDate>"
            $xml += "</ChallanRefs>"
            $xml += "</ChallanDetail>"
        }
        $xml += "</PendingChallans>"
        return $xml
    }

    if (($VchType -eq 9 -or $VchType -eq 2) -and $linkedChallans -and @($linkedChallans).Count -gt 0 -and $null -ne $fi) {
        $xml = "<PendingChallans>"
        
        $partyCode = ""
        $safeParty = $PartyName -replace "'", "''"
        $mRst = $fi.GetRecordset("SELECT Code FROM Master1 WHERE Name = '$safeParty'")
        if ($mRst -and -not $mRst.EOF) {
            $partyCode = $mRst.Fields.Item("Code").Value.ToString()
        }
        if ($mRst) { try { $mRst.Close() } catch {} }

        $masterMap = @{}
        $itemQryRst = $fi.GetRecordset("SELECT Code, Name FROM Master1 WHERE MasterType = 6")
        if ($itemQryRst -and -not $itemQryRst.EOF) {
            while (-not $itemQryRst.EOF) {
                $masterMap[$itemQryRst.Fields.Item("Code").Value.ToString()] = $itemQryRst.Fields.Item("Name").Value.ToString()
                $itemQryRst.MoveNext()
            }
            $itemQryRst.Close()
        }

        $dbType = 0
        if ($script:ActiveConnection -ne $null) {
            $instance = Get-InstanceConfig -InstanceId $script:ActiveInstanceId
            if ($null -ne $instance -and $null -ne $instance.dbType) {
                $dbType = [int]$instance.dbType
            }
        }
        $wildcard = if ($dbType -eq 1) { "%" } else { "*" }

        foreach ($item in $items) {
            $qty      = [double]$item.qty
            $altQty   = if ($item.altQty  -and [double]$item.altQty  -ne 0) { [double]$item.altQty  } else { $qty }
            $altPrice = if ($item.altPrice -and [double]$item.altPrice -ne 0) { [double]$item.altPrice } else { [double]$item.price }
            $itemSrNo = [int]$item.srNo
            
            foreach ($lc in $linkedChallans) {
                $paddedRefNo = Pad-VchNo $lc.vchNo
                
                $resolvedRefCode  = ""
                $resolvedItemCode = ""
                if ($partyCode -ne "") {
                    $safeNo = $lc.vchNo.Trim() -replace "'", "''"
                    $targetRecType = if ($VchType -eq 9) { 3 } else { 4 }
                    
                    $refQry = "SELECT RefCode, MasterCode1 FROM Tran3 WHERE Method=1 AND RecType=$targetRecType AND MasterCode2=$partyCode AND [No] LIKE '$wildcard$safeNo$wildcard' AND ItemSrNo=$itemSrNo"
                    $refRst = $fi.GetRecordset($refQry)
                    if ($refRst -and -not $refRst.EOF) {
                        $resolvedRefCode  = $refRst.Fields.Item("RefCode").Value.ToString()
                        $resolvedItemCode = $refRst.Fields.Item("MasterCode1").Value.ToString()
                    }
                    if ($refRst) { try { $refRst.Close() } catch {} }
                }

                if ($resolvedItemCode -eq "") {
                    $safeItem = $item.itemName -replace "'", "''"
                    $iRst = $fi.GetRecordset("SELECT Code FROM Master1 WHERE Name = '$safeItem'")
                    if ($iRst -and -not $iRst.EOF) {
                        $resolvedItemCode = $iRst.Fields.Item("Code").Value.ToString()
                    }
                    if ($iRst) { try { $iRst.Close() } catch {} }
                }

                $resolvedItemName = if ($null -ne $masterMap -and $masterMap.ContainsKey($resolvedItemCode)) { $masterMap[$resolvedItemCode] } else { $item.itemName }

                $xml += "<ChallanDetail>"
                $xml += "<MasterName1>$([System.Security.SecurityElement]::Escape($resolvedItemName))</MasterName1>"
                $xml += "<MasterName2>$([System.Security.SecurityElement]::Escape($PartyName))</MasterName2>"
                if ($resolvedItemCode -ne "") { $xml += "<tmpMasterCode1>$resolvedItemCode</tmpMasterCode1>" }
                if ($partyCode -ne "")        { $xml += "<tmpMasterCode2>$partyCode</tmpMasterCode2>" }
                
                $xml += "<ChallanRefs>"
                $xml += "<Method>2</Method>"
                $xml += "<SrNo>1</SrNo>"
                $xml += "<RefNo>$([System.Security.SecurityElement]::Escape($paddedRefNo))</RefNo>"
                $xml += "<Date>$VchDate</Date>"
                $dueDate = if ($lc.date) { $lc.date } else { $VchDate }
                $xml += "<DueDate>$dueDate</DueDate>"
                $xml += "<Value1>$qty</Value1>"
                $xml += "<Value2>$altQty</Value2>"
                $xml += "<MainTranPrice>$([double]$item.price)</MainTranPrice>"
                $xml += "<VchType>$VchType</VchType>"
                $xml += "<ItemSrNo>$itemSrNo</ItemSrNo>"
                $xml += "<AltTranPrice>$altPrice</AltTranPrice>"
                $xml += "<MfgDate>$VchDate</MfgDate>"
                
                if ($resolvedRefCode -ne "") {
                    $xml += "<tmpRefCode>$resolvedRefCode</tmpRefCode>"
                    $xml += "<tmpRecType>$targetRecType</tmpRecType>"
                    if ($resolvedItemCode -ne "") { $xml += "<tmpMasterCode1>$resolvedItemCode</tmpMasterCode1>" }
                    if ($partyCode -ne "")        { $xml += "<tmpMasterCode2>$partyCode</tmpMasterCode2>" }
                    $xml += "<TranType>6</TranType>"
                }
                
                $xml += "</ChallanRefs>"
                $xml += "</ChallanDetail>"
            }
        }
        $xml += "</PendingChallans>"
        return $xml
    }
    return ""
}

