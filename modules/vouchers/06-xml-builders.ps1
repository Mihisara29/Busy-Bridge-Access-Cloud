# 06-xml-builders.ps1
# Extracted from vouchers.ps1. Keep functions behavior-compatible with the monolith.

function Build-VoucherXml {
    param($Data, $Cfg, [int]$VchType, [string]$VchNo, [bool]$SkipBBA = $false, $fi = $null)

    # Resolve sequential count to populate AutoVchNo natively in database
    $autoVchNo = Get-BusyAutoVchNoFromVoucherNo `
    -fi $fi `
    -VchType $VchType `
    -SeriesName ([string]$Data.vchSeries) `
    -VchNo $VchNo

    $root = $Cfg.xmlRoot
    $typeField = $Cfg.typeField
    $typeValue = [string]$Data.($Cfg.typeDataKey)
    $matCentre = if ($Data.matCentre) { $Data.matCentre } else { "Main Store" }
    $narration = if ($Data.narration) { $Data.narration } else { "" }
    $totalAmt = ($Data.items | ForEach-Object { [double]$_.amount } | Measure-Object -Sum).Sum
    $inputType = if ($Data.inputType) { [int]$Data.inputType } else { 1 }

    $usesSalesman = @(26, 12, 9, 3) -contains $VchType
    $salesmanName = ""
    $salesmanCode = 0

    if ($usesSalesman) {
        try { $salesmanName = ([string]$Data.salesmanName).Trim() } catch {}
        try {
            if ($null -ne $Data.salesmanCode) {
                $salesmanCode = [int]$Data.salesmanCode
            }
        }
        catch { $salesmanCode = 0 }
    }

    # ------------------------------------------------------------
    # Challan types require BUSY TranType = 3.
    #
    # SkipBBA itself is decided by Create-Voucher / Modify-Voucher
    # and passed into this function explicitly.
    # ------------------------------------------------------------
    $isChallanType = ($vchType -eq 11 -or $vchType -eq 4)

    $tranType = 0
    if ($isChallanType) {
        $tranType = 3
    } elseif ($inputType -eq 2) {
        $tranType = 6
    }

    $xml = "<$root>"
    $seriesNameToSend = $Data.vchSeries
    $xml += "<VchSeriesName>$([System.Security.SecurityElement]::Escape($seriesNameToSend))</VchSeriesName>"
    $xml += "<Date>$($Data.date)</Date>"
    $xml += "<VchType>$VchType</VchType>"
    $xml += "<TranType>$tranType</TranType>"
    $xml += "<VchNo>$([System.Security.SecurityElement]::Escape($VchNo))</VchNo>"

    # Inject sequence number tag
    if ($autoVchNo -gt 0) {
        $xml += "<AutoVchNo>$autoVchNo</AutoVchNo>"
    }

$xml += "<$typeField>$([System.Security.SecurityElement]::Escape($typeValue))</$typeField>"

# ============================================================
# PRODUCTION VOUCHER - VchType 6
# ============================================================
if ($VchType -eq 6) {

    # Production XML mapping discovered from actual BUSY voucher:
    #
    # MasterName1       = Generated Material Centre
    # MasterName2       = BOM Name
    # ConMCName         = Consumed Material Centre
    # ExtraExpenseInBOM = Extra Expense / Unit

    $generatedMC = if ($Data.matCentre) {
        ([string]$Data.matCentre).Trim()
    } else {
        ""
    }

    $consumedMC = if ($Data.party) {
        ([string]$Data.party).Trim()
    } else {
        ""
    }

    $bomName = if ($Data.bomName) {
        ([string]$Data.bomName).Trim()
    } else {
        ""
    }

    $extraExpenseInBOM = 0.0

    try {
        if ($null -ne $Data.extraExpenseInBOM) {
            $extraExpenseInBOM =
                [double]$Data.extraExpenseInBOM
        }
    }
    catch {
        $extraExpenseInBOM = 0.0
    }

    if ([string]::IsNullOrWhiteSpace($bomName)) {
        throw "Production BOM name is required."
    }

    if ([string]::IsNullOrWhiteSpace($generatedMC)) {
        throw "Generated material centre is required."
    }

    if ([string]::IsNullOrWhiteSpace($consumedMC)) {
        throw "Consumed material centre is required."
    }

    $xml += "<MasterName1>$([System.Security.SecurityElement]::Escape($generatedMC))</MasterName1>"

    $xml += "<MasterName2>$([System.Security.SecurityElement]::Escape($bomName))</MasterName2>"

    $xml += "<ConMCName>$([System.Security.SecurityElement]::Escape($consumedMC))</ConMCName>"

    $extraExpenseText =
        $extraExpenseInBOM.ToString(
            "0.####",
            [System.Globalization.CultureInfo]::InvariantCulture
        )

    $xml += "<ExtraExpenseInBOM>$extraExpenseText</ExtraExpenseInBOM>"

    # Frontend sends the actual selected BOM code.
    # This matches BUSY's native Production XML tmpMasterCode2.
    $bomCode = 0

    try {
        if ($null -ne $Data.bomCode) {
            $bomCode = [int]$Data.bomCode
        }
    }
    catch {
        $bomCode = 0
    }

    if ($bomCode -gt 0) {
        $xml += "<tmpMasterCode2>$bomCode</tmpMasterCode2>"
    }
}
else {

    # ========================================================
    # EXISTING NON-PRODUCTION VOUCHER BEHAVIOUR
    # ========================================================

    $xml += "<MasterName1>$([System.Security.SecurityElement]::Escape([string]$Data.party))</MasterName1>"

    $xml += "<MasterName2>$([System.Security.SecurityElement]::Escape($matCentre))</MasterName2>"

    # Stock Journal consumed material centre
    if ($VchType -eq 8) {
        $xml += "<ConMCName>$([System.Security.SecurityElement]::Escape([string]$Data.party))</ConMCName>"
    }
}

$xml += "<TranCurName>Rs.</TranCurName>"

    if ($usesSalesman -and -not [string]::IsNullOrWhiteSpace($salesmanName)) {
        $xml += "<BrokerInvolved>True</BrokerInvolved>"
        $xml += "<BrokerName>$([System.Security.SecurityElement]::Escape($salesmanName))</BrokerName>"
    }

    $xml += "<InputType>$inputType</InputType>"
    $xml += "<BillingDetails><PartyName>$([System.Security.SecurityElement]::Escape($Data.party))</PartyName></BillingDetails>"
    
    $xml += "<VchOtherInfoDetails>"

    if ($Data.optionalFields) {
        $xml += "<OFInfo>"
        # Populate the XML elements OF1 through OF20 sequentially
        for ($i = 1; $i -le 20; $i++) {
            $key = "OptionField$i"
            $val = ""

            # DUAL-COMPATIBLE PROPERTY DETECTION: Handles both [PSCustomObject] and [Hashtable] safely
            if ($Data.optionalFields.psobject.Properties[$key]) {
                $val = [string]$Data.optionalFields.$key
            } elseif ($Data.optionalFields.GetType().Name -eq "Hashtable" -and $Data.optionalFields.ContainsKey($key)) {
                $val = [string]$Data.optionalFields[$key]
            }

            if ($val -ne "") {
                # AUTOMATIC DATE FORMATTER: Detect and reformat standard HTML5 yyyy-MM-dd values to dd-MM-yyyy
                if ($val -match "^(\d{4})-(\d{2})-(\d{2})$") {
                    $val = "$($Matches[3])-$($Matches[2])-$($Matches[1])"
                }
                $xml += "<OF$i>$([System.Security.SecurityElement]::Escape($val))</OF$i>"
            }
        }
        $xml += "</OFInfo>"
    }

    if ($Cfg.hasBillNo) {
        $billNo = if ($Data.supplierBillNo) { $Data.supplierBillNo } elseif ($Data.purchaseBillNo) { $Data.purchaseBillNo } else { "" }
        $xml += "<PurchaseBillNo>$([System.Security.SecurityElement]::Escape($billNo))</PurchaseBillNo>"
    }
    
    $xml += "<Narration1>$([System.Security.SecurityElement]::Escape($narration))</Narration1>"
    $xml += "</VchOtherInfoDetails>"

if ($VchType -eq 8 -or $VchType -eq 6) {

    # ========================================================
    # STOCK JOURNAL + PRODUCTION
    #
    # ItemEntries  = Items Generated
    # ItemEntries1 = Items Consumed
    # ========================================================

    $genItems = @(
        $Data.items | Where-Object {
            $null -eq $_.itemType -or
            [int]$_.itemType -eq 1
        }
    )

    $conItems = @(
        $Data.items | Where-Object {
            $null -ne $_.itemType -and
            [int]$_.itemType -eq 2
        }
    )

    $debugLabel =
        if ($VchType -eq 6) {
            "Production"
        }
        else {
            "StockJournal"
        }

    Write-Host "[$debugLabel Save DEBUG] Generated items=$($genItems.Count), Consumed items=$($conItems.Count)" -ForegroundColor Cyan
    Write-Host "[$debugLabel Save DEBUG] Generated MC=$matCentre, Consumed MC=$($Data.party)" -ForegroundColor Cyan

    # Generated section
    $xml += Build-ItemsXml `
        -items $genItems `
        -defaultMC $matCentre

    # Consumed section
    $xml += Build-ConsumedItemsXml `
        -items $conItems `
        -defaultMC ([string]$Data.party)
}
else {

    # Existing behaviour for Sales / Purchase / Challan etc.
    $xml += Build-ItemsXml `
        -items $Data.items `
        -defaultMC $matCentre `
        -salesmanName $salesmanName
}
    $xml += Build-BillSundriesXml -billSundries $Data.billSundries

    if (-not $isChallanType) {
        $xml += Build-AccEntriesXml -items $Data.items -PartyName $Data.party -TotalAmt $totalAmt -VchType $VchType -Settlements $Data.settlements
    }

    $cAmt = 0; $cc1Amt = 0; $cc2Amt = 0
    $cAcc = "Cash"; $cc1Acc = ""; $cc2Acc = ""
    $tendered = 0

    if ($Data.settlements) {
        if ($Data.settlements.cash.amount) { $cAmt = [double]$Data.settlements.cash.amount }
        if ($Data.settlements.cash.account) { $cAcc = $Data.settlements.cash.account }

        if ($Data.settlements.card.amount) { $cc1Amt = [double]$Data.settlements.card.amount }
        if ($Data.settlements.card.account) { $cc1Acc = $Data.settlements.card.account }

        if ($Data.settlements.gift.amount) { $cc2Amt = [double]$Data.settlements.gift.amount }
        if ($Data.settlements.gift.account) { $cc2Acc = $Data.settlements.gift.account }

        $tendered = $cAmt
        if ($Data.settlements.cashTendered) { $tendered = [double]$Data.settlements.cashTendered }
        if ($tendered -lt $cAmt) { $tendered = $cAmt }
    }

    $bbaBlocks = ""

    if (-not $SkipBBA -and -not $isChallanType) {
        if (-not $Data.refEntries -or @($Data.refEntries).Count -eq 0) {
            $partyBalance = $totalAmt
            if ($Data.settlements) {
                $partyBalance = $totalAmt - ($cAmt + $cc1Amt + $cc2Amt)
            }
            if ([Math]::Round($partyBalance, 2) -gt 0) {
                $autoRef = @( @{ srNo = 1; refType = 1; refNo = $VchNo; amount = $partyBalance; dueDate = $Data.date } )
                $bbaBlocks += Build-PendingBillDetailsXml -refEntries $autoRef -PartyName $Data.party -VchDate $Data.date -VchType $VchType -VchNo $VchNo
            }
        } else {
            $bbaBlocks += Build-PendingBillDetailsXml -refEntries $Data.refEntries -PartyName $Data.party -VchDate $Data.date -VchType $VchType -VchNo $VchNo
        }

        if ($Data.settlements) {
            if ($cc1Amt -gt 0 -and $cc1Acc -ne "") {
                $autoRefCard = @( @{ srNo = 1; refType = 1; refNo = $VchNo; amount = $cc1Amt; dueDate = $Data.date } )
                $bbaBlocks += Build-PendingBillDetailsXml -refEntries $autoRefCard -PartyName $cc1Acc -VchDate $Data.date -VchType $VchType -VchNo $VchNo
            }
            if ($cc2Amt -gt 0 -and $cc2Acc -ne "") {
                $autoRefGift = @( @{ srNo = 1; refType = 1; refNo = $VchNo; amount = $cc2Amt; dueDate = $Data.date } )
                $bbaBlocks += Build-PendingBillDetailsXml -refEntries $autoRefGift -PartyName $cc2Acc -VchDate $Data.date -VchType $VchType -VchNo $VchNo
            }
        }
    }

    if ($bbaBlocks -ne "") {
        $xml += "<PendingBillDetails>" + $bbaBlocks + "</PendingBillDetails>"
    }

    if ($Data.linkedOrder -and $Data.linkedOrder.vchNo) {
        $xml += Build-PendingOrdersXml -items $Data.items -PartyName $Data.party -linkedOrder $Data.linkedOrder -VchType $VchType
    }

    if (-not [string]::IsNullOrWhiteSpace($VchNo)) {
        $xml += Build-PendingChallansXml -items $Data.items -PartyName $Data.party -linkedChallans $Data.linkedChallans -VchType $VchType -VchNo $VchNo -VchDate $Data.date -fi $fi
    }

    if ($Data.settlements) {
        $xml += "<PosEnabled>True</PosEnabled>"
        $xml += "<POSVchData>"
        $xml += "<CashAmt>$($cAmt.ToString('0.00',[System.Globalization.CultureInfo]::InvariantCulture))</CashAmt>"
        $xml += "<CashRecvdAmt>$($tendered.ToString('0.00',[System.Globalization.CultureInfo]::InvariantCulture))</CashRecvdAmt>"
        $xml += "<CreditCardAmt1>$($cc1Amt.ToString('0.00',[System.Globalization.CultureInfo]::InvariantCulture))</CreditCardAmt1>"
        $xml += "<CreditCardAmt2>$($cc2Amt.ToString('0.00',[System.Globalization.CultureInfo]::InvariantCulture))</CreditCardAmt2>"
        $xml += "<CarrySettlement>True</CarrySettlement>"
        $xml += "<CashAccName>$([System.Security.SecurityElement]::Escape($cAcc))</CashAccName>"
        $xml += "<CCAccName1>$([System.Security.SecurityElement]::Escape($cc1Acc))</CCAccName1>"
        $xml += "<CCAccName2>$([System.Security.SecurityElement]::Escape($cc2Acc))</CCAccName2>"
        $xml += "</POSVchData>"
    }

    if ($usesSalesman -and $salesmanCode -gt 0) {
        $xml += "<tmpBrokerCode>$salesmanCode</tmpBrokerCode>"
    }

    $xml += "</$root>"
    return $xml
}

# FIXED: Re-balanced XML mapping of accounting vouchers to support standard ledger balance checks

function Build-AccountingVoucherXml {
    param($Data, $Cfg,[int]$VchType,[string]$VchNo, [bool]$SkipBBA = $false, $fi = $null)

    # Resolve sequential count to populate AutoVchNo natively in database
    $autoVchNo = Get-BusyAutoVchNoFromVoucherNo `
    -fi $fi `
    -VchType $VchType `
    -SeriesName ([string]$Data.vchSeries) `
    -VchNo $VchNo

    $root = $Cfg.xmlRoot
    $narration = if ($Data.narration) { $Data.narration } else { "" }

    # Query if "Double Entry System for Payment & Receipt" is enabled (RecType=1, L1=16, L2=1 means Double Entry)
    $isDoubleEntry = $false
    if ($fi -and ($VchType -eq 14 -or $VchType -eq 19)) {
        try {
            $cfgRst = $fi.GetRecordset("SELECT L2 FROM Config WHERE RecType=1 AND L1=16")
            if ($cfgRst -and -not $cfgRst.EOF) {
                if ($cfgRst.Fields.Item("L2").Value -ne [System.DBNull]::Value) {
                    $isDoubleEntry = ([int]$cfgRst.Fields.Item("L2").Value -eq 1)
                }
                $cfgRst.Close()
            }
        } catch {
            $isDoubleEntry = $false
        }
    }

    $xml = "<$root>"
    $xml += "<VchSeriesName>$([System.Security.SecurityElement]::Escape($Data.vchSeries))</VchSeriesName>"
    $xml += "<Date>$($Data.date)</Date>"
    $xml += "<VchType>$VchType</VchType>"
    $xml += "<VchNo>$([System.Security.SecurityElement]::Escape($VchNo))</VchNo>"

    # Inject sequence number tag
    if ($autoVchNo -gt 0) {
        $xml += "<AutoVchNo>$autoVchNo</AutoVchNo>"
    }

    $xml += "<TranCurName>Rs.</TranCurName>"

    # Single Entry logic: Specifying MasterName1 Mode at the root
    if (-not $isDoubleEntry -and ($VchType -eq 14 -or $VchType -eq 19) -and $Data.accounts.Count -gt 0) {
        $modeAccount = $Data.accounts[0]
        $xml += "<MasterName1>$([System.Security.SecurityElement]::Escape($modeAccount.accountName))</MasterName1>"
    }

    # AccEntries must contain ALL standard Debit and Credit account lines to satisfy double-entry balancing
    $xml += "<AccEntries>"
    foreach ($acc in $Data.accounts) {
        $amtType = if ($acc.dc -eq 'D') { 1 } else { 2 }
        $amt = [double]$acc.amount 
        
        $xml += "<AccDetail>"
        $xml += "<SrNo>$($acc.srNo)</SrNo>"
        $xml += "<AccountName>$([System.Security.SecurityElement]::Escape($acc.accountName))</AccountName>"
        $xml += "<AmountType>$amtType</AmountType>"
        $xml += "<AmtMainCur>$amt</AmtMainCur>"
        if ($acc.shortNarration) {
            $xml += "<ShortNar>$([System.Security.SecurityElement]::Escape($acc.shortNarration))</ShortNar>"
        }
        
        if (-not $SkipBBA -and $acc.bbaEntries -and @($acc.bbaEntries).Count -gt 0) {
            $xml += "<BillRefs>"
            $bbaSrNo = 1
            foreach ($ref in $acc.bbaEntries) {
                $xml += "<BillDetails>"
                $xml += "<Method>$($ref.refType)</Method>"
                $xml += "<SrNo>$bbaSrNo</SrNo>"
                $rNo = if ($ref.refNo -eq "AUTO" -or [string]::IsNullOrWhiteSpace($ref.refNo)) { $VchNo } else { $ref.refNo }
                $xml += "<RefNo>$([System.Security.SecurityElement]::Escape($rNo))</RefNo>"
                $xml += "<Date>$($Data.date)</Date>"
                $dueDate = if ($ref.dueDate) { $ref.dueDate } else { $Data.date }
                $xml += "<DueDate>$dueDate</DueDate>"
                
                $val1 = if ($amtType -eq 1) { -[Math]::Abs([double]$ref.amount) } else { [Math]::Abs([double]$ref.amount) }
                $xml += "<Value1>$val1</Value1>"
                $xml += "<VchType>$VchType</VchType>"
                $xml += "</BillDetails>"
                $bbaSrNo++
            }
            $xml += "</BillRefs>"
        }
        $xml += "</AccDetail>"
    }
    $xml += "</AccEntries>"

    if (-not $SkipBBA) {
        $hasBba = $false
        foreach ($acc in $Data.accounts) {
            if ($acc.bbaEntries -and @($acc.bbaEntries).Count -gt 0) { $hasBba = $true; break }
        }

        if ($hasBba) {
            $xml += "<PendingBillDetails>"
            foreach ($acc in $Data.accounts) {
                if ($acc.bbaEntries -and @($acc.bbaEntries).Count -gt 0) {
                    $amtType = if ($acc.dc -eq 'D') { 1 } else { 2 }
                    $xml += "<BillDetail>"
                    $xml += "<MasterName1>$([System.Security.SecurityElement]::Escape($acc.accountName))</MasterName1>"
                    $bbaSrNo = 1
                    foreach ($ref in $acc.bbaEntries) {
                        $xml += "<BillRefs>"
                        $xml += "<Method>$($ref.refType)</Method>"
                        $xml += "<SrNo>$bbaSrNo</SrNo>"
                        $rNo = if ($ref.refNo -eq "AUTO" -or [string]::IsNullOrWhiteSpace($ref.refNo)) { $VchNo } else { $ref.refNo }
                        $xml += "<RefNo>$([System.Security.SecurityElement]::Escape($rNo))</RefNo>"
                        $xml += "<Date>$($Data.date)</Date>"
                        $dueDate = if ($ref.dueDate) { $ref.dueDate } else { $Data.date }
                        $xml += "<DueDate>$dueDate</DueDate>"
                        $val1 = if ($amtType -eq 1) { -[Math]::Abs([double]$ref.amount) } else { [Math]::Abs([double]$ref.amount) }
                        $xml += "<Value1>$val1</Value1>"
                        $xml += "<VchType>$VchType</VchType>"
                        $xml += "</BillRefs>"
                        $bbaSrNo++
                    }
                    $xml += "</BillDetail>"
                }
            }
            $xml += "</PendingBillDetails>"
        }
    }

    if ($narration) { $xml += "<VchOtherInfoDetails><Narration1>$([System.Security.SecurityElement]::Escape($narration))</Narration1></VchOtherInfoDetails>" }
    $xml += "</$root>"
    
    return $xml
}

# Helper to natively resolve VchCode directly via high-performance direct DB query

