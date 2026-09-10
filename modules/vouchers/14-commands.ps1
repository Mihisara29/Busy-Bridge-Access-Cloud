# 14-commands.ps1
# Extracted from vouchers.ps1. Keep functions behavior-compatible with the monolith.

function Create-Voucher {
    param($Data, [string]$InstanceId = "", [string]$CompanyCode = "")

    $vchType = [int]$Data.vchType
    $cfg = $script:VoucherConfig[$vchType]
    if (-not $cfg) { return @{ success=$false; error="Unsupported vchType: $vchType" } }

    $validationError = Validate-VoucherData $Data $cfg
    if ($validationError) { return @{ success=$false; error=$validationError } }

    # ------------------------------------------------------------
    # VOUCHERS THAT MUST NOT GENERATE BBA
    #
    # 11 = Delivery Order / Material Issue
    # 4  = Goods Received Note / Material Receipt
    # 26 = Sales Quotation
    # 27 = Purchase Quotation
    #
    # Quotations are non-accounting documents, so they must not
    # create Pending Bill / Bill-by-Bill Adjustment references.
    # ------------------------------------------------------------
    $isChallanType = ($vchType -eq 11 -or $vchType -eq 4)
    $isQuotationType = ($vchType -eq 26 -or $vchType -eq 27)

    $skipBBA = [bool]($isChallanType -or $isQuotationType)

    $maxAttempts = 2
    $attempt = 1
    $lastExceptionMsg = ""
    $attempt = 1
    $lastExceptionMsg = ""

    while ($attempt -le $maxAttempts) {
        $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
        if (-not $fi) { return @{ success=$false; error="BUSY connection failed" } }

        try {
            $seriesName = ([string]$Data.vchSeries).Trim()
            $busyDateText = [string]$Data.date
            $voucherDate = [datetime]::Now.Date
            if ($busyDateText -match '^\d{2}-\d{2}-\d{4}$') {
                $voucherDate = [datetime]::ParseExact($busyDateText, 'dd-MM-yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
            } elseif ($busyDateText -match '^\d{4}-\d{2}-\d{2}$') {
                $voucherDate = [datetime]::ParseExact($busyDateText, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
            }

            # Read the administrator-selected source for this type + series.
            $adminConfigResult = Get-WebNumberingConfig `
                -VchType $vchType `
                -SeriesName $seriesName `
                -VoucherDate $voucherDate.ToString('yyyy-MM-dd') `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode

            if (-not $adminConfigResult.success) { throw $adminConfigResult.error }
            $adminConfig = $adminConfigResult.data

            # WEB source is authoritative. BUSY source keeps the number submitted by
            # the existing BUSY numbering flow.
            $vchNo = if ($adminConfig.source -eq 'WEB' -and $adminConfig.is_active) {
                ([string]$adminConfig.next_vch_no).Trim()
            } else {
                ([string]$Data.vchNo).Trim()
            }

            if ([string]::IsNullOrWhiteSpace($vchNo)) { return @{ success=$false; error='Voucher number is required.' } }

            $exists = Test-VoucherNumberExists -fi $fi -VchType $vchType -SeriesName $seriesName -VchNo $vchNo
            if ($exists) {
                return @{ success=$false; error="Voucher number '$vchNo' already exists in series '$seriesName'. Refresh and try again." }
            }

            # Ensure XML builders receive the final backend-selected number.
            $Data.vchNo = $vchNo
            $xml = if ($cfg.isAccounting) {
                Build-AccountingVoucherXml -Data $Data -Cfg $cfg -VchType $vchType -VchNo $vchNo -SkipBBA $false -fi $fi
            } else {
                Build-VoucherXml -Data $Data -Cfg $cfg -VchType $vchType -VchNo $vchNo -SkipBBA $skipBBA -fi $fi
            }

            $errMsg = ""
            $saved = $fi.SaveVchFromXML($vchType, $xml, [ref]$errMsg)
            if ($saved -ne $true) {
                return @{ success=$false; error=if ($errMsg) { $errMsg } else { 'Unknown BUSY error' } }
            }

            # Apply the administrator's BusyCloud approval policy immediately
            # after BUSY creates the voucher:
            #   policy OFF -> ApprovalStatus 0 (Approval Not Required)
            #   policy ON  -> ApprovalStatus 2 (To be Approved)
            $approvalStateResult = Set-WebCreatedVoucherApprovalState `
                -fi $fi `
                -VchType $vchType `
                -VchNo $vchNo `
                -VchSeries ([string]$Data.vchSeries) `
                -VchDate ([string]$Data.date) `
                -UserName ([string]$Data.bridgeUserName) `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode

            if (-not $approvalStateResult.success) {
                # The voucher already exists at this point. Explicitly return a
                # partial-success response so the client never retries creation.
                return @{
                    success = $false
                    created = $true
                    error = "Voucher was created in BUSY, but BusyCloud could not apply its configured approval state. Do not create it again. Details: $($approvalStateResult.error)"
                    data = @{
                        vchType = $vchType
                        vchSeries = $seriesName
                        vchNo = $vchNo
                        date = $Data.date
                        vchCode = $approvalStateResult.vchCode
                    }
                }
            }

            Clear-StockCaches -InstanceId $InstanceId -CompanyCode $CompanyCode
            return @{
                success=$true
                message="$($cfg.xmlRoot) created successfully"
                data=@{
                    vchType=$vchType
                    vchSeries=$seriesName
                    vchNo=$vchNo
                    date=$Data.date
                    party=$Data.party
                    numberingSource=$adminConfig.source
                    vchCode=$approvalStateResult.vchCode
                    approvalRequired=$approvalStateResult.approvalRequired
                    approvalStatus=$approvalStateResult.approvalStatus
                    approvalStatusText=$approvalStateResult.approvalStatusText
                }
            }
        } catch {
            $lastExceptionMsg = $_.Exception.Message
            $script:ActiveConnection = $null
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($fi) | Out-Null } catch {}
            $attempt++
        } finally {
            Disconnect-BUSY $fi
        }
    }

    return @{ success=$false; error="Database error. Connection reset. Details: $lastExceptionMsg" }
}

function Modify-Voucher {
    param($Data, [string]$InstanceId = "",[string]$CompanyCode = "")

    $vchType = [int]$Data.vchType
    $cfg = $script:VoucherConfig[$vchType]
    
    if (-not $cfg) { return @{ success = $false; error = "Unsupported vchType" } }
    if (-not $Data.vchNo -or [string]$Data.vchNo -eq "") { return @{ success = $false; error = "vchNo is required for modify" } }

        # ------------------------------------------------------------
    # Keep BBA behaviour consistent between Create and Modify.
    # ------------------------------------------------------------
    $isChallanType = ($vchType -eq 11 -or $vchType -eq 4)
    $isQuotationType = ($vchType -eq 26 -or $vchType -eq 27)

    $skipBBA = [bool]($isChallanType -or $isQuotationType)
    
    $maxAttempts = 2
    $attempt = 1
    $lastExceptionMsg = ""

    while ($attempt -le $maxAttempts) {
        $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
        if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
        
        try {
            # routes.ps1 resolves the existing voucher before modify and places
            # its current approval state here. Preserve 0/1/2 exactly across
            # SaveVchFromXML; otherwise BUSY may auto-change approval state.
            $approvalStatusToPreserve = -1
            try {
                if ($null -ne $Data.preserveApprovalStatus) {
                    $candidateStatus = [int]$Data.preserveApprovalStatus
                    if ($candidateStatus -in @(0,1,2)) { $approvalStatusToPreserve = $candidateStatus }
                }
            } catch {}

            $approvalVchCode = 0
            $approvalMarkerSnapshot = $null

            if ($approvalStatusToPreserve -in @(0,1,2)) {
                $beforeState = Get-VoucherApprovalStateByIdentity `
                    -VchType $vchType `
                    -VchNo ([string]$Data.vchNo) `
                    -VchDate ([string]$Data.date) `
                    -VchSeries ([string]$Data.vchSeries) `
                    -InstanceId $InstanceId `
                    -CompanyCode $CompanyCode `
                    -ExistingFi $fi

                if ($beforeState.success -and $beforeState.found) {
                    $approvalVchCode = [int]$beforeState.vchCode
                }

                # Approved vouchers must retain the original Action=3 marker so
                # BUSY continues to show the original Approved By / Approved At.
                if ($approvalStatusToPreserve -eq 1 -and $approvalVchCode -gt 0) {
                    $approvalRs = $fi.GetRecordset("SELECT TOP 1 Type,ActionTime,UserName,D1,D2,D3,D4,D5,Notes,ComputerName FROM CheckList WHERE Code=$approvalVchCode AND Action=3 ORDER BY ActionTime DESC")
                    if ($approvalRs -and -not $approvalRs.EOF) {
                        function Read-ModifyApprovalField { param($R,[string]$F,$D); try { $v=$R.Fields.Item($F).Value; if ($null -ne $v -and $v -ne [System.DBNull]::Value) { return $v } } catch {}; return $D }
                        $approvalMarkerSnapshot = @{
                            Type = [int](Read-ModifyApprovalField $approvalRs "Type" 2)
                            ActionTime = Read-ModifyApprovalField $approvalRs "ActionTime" (Get-Date)
                            UserName = [string](Read-ModifyApprovalField $approvalRs "UserName" "")
                            D1 = Read-ModifyApprovalField $approvalRs "D1" 0
                            D2 = Read-ModifyApprovalField $approvalRs "D2" 0
                            D3 = Read-ModifyApprovalField $approvalRs "D3" 0
                            D4 = Read-ModifyApprovalField $approvalRs "D4" 0
                            D5 = Read-ModifyApprovalField $approvalRs "D5" 0
                            Notes = [string](Read-ModifyApprovalField $approvalRs "Notes" "")
                            ComputerName = [string](Read-ModifyApprovalField $approvalRs "ComputerName" "BUSYCLOUD")
                        }
                    }
                    if ($approvalRs) { try { $approvalRs.Close() } catch {} }
                }
            }

           $xml = if ($cfg.isAccounting) {
                Build-AccountingVoucherXml `
                    -Data $Data `
                    -Cfg $cfg `
                    -VchType $vchType `
                    -VchNo ([string]$Data.vchNo) `
                    -SkipBBA $false `
                    -fi $fi
            } else {
                Build-VoucherXml `
                    -Data $Data `
                    -Cfg $cfg `
                    -VchType $vchType `
                    -VchNo ([string]$Data.vchNo) `
                    -SkipBBA $skipBBA `
                    -fi $fi
            }

            $errMsg = ""
            $saved = $fi.SaveVchFromXML($vchType, $xml, [ref]$errMsg, $true)
            
            if ($saved -eq $true) {
                # Restore the exact approval state that existed before modify.
                # 0 = Approval Not Required, 1 = Approved, 2 = To be Approved.
                if ($approvalStatusToPreserve -in @(0,1,2)) {
                    if ($approvalVchCode -le 0) {
                        $afterState = Get-VoucherApprovalStateByIdentity `
                            -VchType $vchType `
                            -VchNo ([string]$Data.vchNo) `
                            -VchDate ([string]$Data.date) `
                            -VchSeries ([string]$Data.vchSeries) `
                            -InstanceId $InstanceId `
                            -CompanyCode $CompanyCode `
                            -ExistingFi $fi
                        if ($afterState.success -and $afterState.found) { $approvalVchCode = [int]$afterState.vchCode }
                    }

                    if ($approvalVchCode -gt 0) {
                        $fi.ExecuteQuery("UPDATE Tran1 SET ApprovalStatus=$approvalStatusToPreserve WHERE VchCode=$approvalVchCode")

                        if ($approvalStatusToPreserve -ne 1) {
                            # Pending / Not Required must not carry an approval marker.
                            $fi.ExecuteQuery("DELETE FROM CheckList WHERE Code=$approvalVchCode AND Action=3")
                        }
                        else {
                            # Keep exactly the original approval marker. If BUSY
                            # generates/removes Action=3 during SaveVchFromXML,
                            # normalize it back to the pre-modify Approved By / At.
                            if ($null -ne $approvalMarkerSnapshot) {
                                function To-ModifyNumberLiteral { param($Value); try { return ([double]$Value).ToString("0.################", [System.Globalization.CultureInfo]::InvariantCulture) } catch { return "0" } }
                                $found = Get-InstanceForCompany -CompanyCode $CompanyCode -InstanceId $InstanceId
                                $dbType = if ($found -and $null -ne $found.instance.dbType) { [int]$found.instance.dbType } else { 0 }
                                $dt = [datetime]$approvalMarkerSnapshot.ActionTime
                                $dateExpr = if ($dbType -eq 1) { "'" + $dt.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'" } else { "#" + $dt.ToString("MM/dd/yyyy HH:mm:ss") + "#" }
                                $safeApprovedBy = ([string]$approvalMarkerSnapshot.UserName).Replace("'", "''")
                                $safeApprovalNotes = ([string]$approvalMarkerSnapshot.Notes).Replace("'", "''")
                                $safeApprovalComputer = ([string]$approvalMarkerSnapshot.ComputerName).Replace("'", "''")
                                $d1 = To-ModifyNumberLiteral $approvalMarkerSnapshot.D1
                                $d2 = To-ModifyNumberLiteral $approvalMarkerSnapshot.D2
                                $d3 = To-ModifyNumberLiteral $approvalMarkerSnapshot.D3
                                $d4 = To-ModifyNumberLiteral $approvalMarkerSnapshot.D4
                                $d5 = To-ModifyNumberLiteral $approvalMarkerSnapshot.D5
                                $markerType = [int]$approvalMarkerSnapshot.Type
                                $fi.ExecuteQuery("DELETE FROM CheckList WHERE Code=$approvalVchCode AND Action=3")
                                $fi.ExecuteQuery("INSERT INTO CheckList (Type,Code,Action,ActionTime,UserName,D1,D2,D3,D4,D5,Notes,ComputerName) VALUES ($markerType,$approvalVchCode,3,$dateExpr,'$safeApprovedBy',$d1,$d2,$d3,$d4,$d5,'$safeApprovalNotes','$safeApprovalComputer')")
                            }
                        }
                    }
                }

                if ($Data.bridgeUserName) {
                    Update-CheckListCreator -fi $fi -VchType $vchType -VchNo ([string]$Data.vchNo) -VchDate $Data.date -VchSeries ([string]$Data.vchSeries) -UserName $Data.bridgeUserName -InstanceId $InstanceId -CompanyCode $CompanyCode
                }

                Clear-StockCaches -InstanceId $InstanceId -CompanyCode $CompanyCode
                return @{
                    success = $true
                    message = "$($cfg.xmlRoot) modified"
                    data = @{
                        vchType = $vchType
                        vchSeries = $Data.vchSeries
                        vchNo = $Data.vchNo
                        date = $Data.date
                    }
                }
            } else {
                return @{ success = $false; error = if ($errMsg) { $errMsg } else { "Unknown BUSY error" } }
            }
        } catch {
            $lastExceptionMsg = $_.Exception.Message
            $script:ActiveConnection = $null
            try { $script:ActiveConnection.CloseDB() } catch {}
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($fi) | Out-Null } catch {}
            
            if ($attempt -eq 1) {
                Write-Host "2026-06-26 [WARNING] Idle connection timeout detected on Modify. Reconnecting for Attempt 2..." -ForegroundColor Yellow
            }
            $attempt++
        } finally {
            Disconnect-BUSY $fi
        }
    }

    return @{ success = $false; error = "Database error. Connection reset. Details: $lastExceptionMsg" }
}



# ═══════════════════════════════════════════════════════════════
#  DELETE VOUCHER
# ═══════════════════════════════════════════════════════════════

function Delete-Voucher {
    param($Data, [string]$InstanceId = "", [string]$CompanyCode = "")

    $cfg = $script:VoucherConfig[[int]$Data.vchType]
    if (-not $cfg) { return @{ success = $false; error = "Unsupported vchType" } }
    
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
    
    try {
        $errMsg = ""
        $deleted = $fi.DeleteVch($Data.vchType, $Data.vchSeries, $Data.vchDate, $Data.vchNo, [ref]$errMsg)
        
        if ($deleted -eq $true) {
            Clear-StockCaches -InstanceId $InstanceId -CompanyCode $CompanyCode
            return @{ success = $true; message = "Voucher deleted successfully" }
        } else {
            return @{ success = $false; error = if ($errMsg) { $errMsg } else { "Unknown BUSY error" } }
        }
    } catch {
        $script:ActiveConnection = $null
        try { $script:ActiveConnection.CloseDB() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($fi) | Out-Null } catch {}
        return @{ success = $false; error = "Database error. Connection reset. Please try deleting again." }
    } finally {
        Disconnect-BUSY $fi
    }
}

