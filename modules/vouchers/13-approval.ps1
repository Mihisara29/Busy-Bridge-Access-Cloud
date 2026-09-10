# 13-approval.ps1
# Extracted from vouchers.ps1. Keep functions behavior-compatible with the monolith.

function Get-VoucherApprovalStateByIdentity {
    param(
        [int]$VchType,
        [string]$VchNo,
        [string]$VchDate = "",
        [string]$VchSeries = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        $ExistingFi = $null
    )

    if ($VchType -le 0 -or [string]::IsNullOrWhiteSpace($VchNo)) {
        return @{ success=$false; found=$false; error="vchType and vchNo are required." }
    }

    $fi = $ExistingFi
    $ownsConnection = $false
    if (-not $fi) {
        $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
        $ownsConnection = $true
    }
    if (-not $fi) { return @{ success=$false; found=$false; error="BUSY database connection failed" } }

    try {
        $dbType = 0
        $foundCompany = Get-InstanceForCompany -CompanyCode $CompanyCode -InstanceId $InstanceId
        if ($foundCompany -and $null -ne $foundCompany.instance.dbType) { $dbType = [int]$foundCompany.instance.dbType }
        $wildcard = if ($dbType -eq 1) { "%" } else { "*" }

        $targetNo = $VchNo.Trim()
        $targetSeries = ([string]$VchSeries).Trim()
        $targetDate = $null
        if (-not [string]::IsNullOrWhiteSpace($VchDate)) {
            try { $targetDate = Safe-ParseDate ([string]$VchDate) } catch {}
        }

        $safeNo = $targetNo.Replace("'", "''")
        $rst = $fi.GetRecordset(@"
SELECT VchCode,VchType,VchNo,VchSeriesCode,[Date],ApprovalStatus
FROM Tran1
WHERE VchType=$VchType
  AND VchNo LIKE '$wildcard$safeNo$wildcard'
ORDER BY VchCode DESC
"@)

        if ($rst) {
            while (-not $rst.EOF) {
                $dbNo = ""
                try { $dbNo = ([string]$rst.Fields.Item("VchNo").Value).Trim() } catch {}

                if ($dbNo.Equals($targetNo, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $seriesCode = 0
                    $dateValue = $null
                    try { $seriesCode = [int]$rst.Fields.Item("VchSeriesCode").Value } catch {}
                    try { $dateValue = $rst.Fields.Item("Date").Value } catch {}

                    $dateMatches = $true
                    if ($null -ne $targetDate -and $null -ne $dateValue) {
                        try { $dateMatches = ([datetime]$dateValue).Date -eq ([datetime]$targetDate).Date } catch { $dateMatches = $false }
                    }

                    $seriesMatches = $true
                    $seriesName = ""
                    if (-not [string]::IsNullOrWhiteSpace($targetSeries)) {
                        $seriesMatches = $false
                        if ($seriesCode -gt 0) {
                            $sr = $fi.GetRecordset("SELECT Name FROM Master1 WHERE Code=$seriesCode")
                            if ($sr -and -not $sr.EOF) {
                                try { $seriesName = ([string]$sr.Fields.Item("Name").Value).Trim() } catch {}
                            }
                            if ($sr) { try { $sr.Close() } catch {} }
                        }

                        $prefix = "{0:D2}" -f $VchType
                        if ($seriesName.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                            $seriesName = $seriesName.Substring($prefix.Length)
                        }

                        $seriesMatches = $seriesName.Equals($targetSeries, [System.StringComparison]::OrdinalIgnoreCase)
                    }

                    if ($dateMatches -and $seriesMatches) {
                        $result = @{
                            success = $true
                            found = $true
                            vchCode = [int]$rst.Fields.Item("VchCode").Value
                            vchType = [int]$rst.Fields.Item("VchType").Value
                            vchNo = $dbNo
                            vchSeriesCode = $seriesCode
                            vchSeries = $seriesName
                            date = $dateValue
                            approvalStatus = [int]$rst.Fields.Item("ApprovalStatus").Value
                        }
                        try { $rst.Close() } catch {}
                        return $result
                    }
                }
                $rst.MoveNext()
            }
            try { $rst.Close() } catch {}
        }

        return @{ success=$true; found=$false }
    }
    catch { return @{ success=$false; found=$false; error=$_.Exception.Message } }
    finally { if ($ownsConnection) { Disconnect-BUSY $fi } }
}

function Set-WebCreatedVoucherApprovalState {
    param(
        $fi,
        [int]$VchType,
        [string]$VchNo,
        [string]$VchSeries = "",
        [string]$VchDate = "",
        [string]$UserName,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    # Verified BUSY mapping:
    # 0 Approval Not Required | 1 Approved | 2 To be Approved
    # CheckList: 1 Created | 2 Modified | 3 Approved
    if ([string]::IsNullOrWhiteSpace($VchNo)) {
        return @{ success=$false; error="Voucher number is required to set approval state." }
    }

    try {
        $config = Get-VoucherApprovalConfig -VchType $VchType -InstanceId $InstanceId -CompanyCode $CompanyCode -ExistingFi $fi
        if (-not $config.success) { return @{ success=$false; error="Could not read voucher approval configuration. $($config.error)" } }

        $approvalRequired = [bool]$config.data.approval_required
        $targetStatus = if ($approvalRequired) { 2 } else { 0 }

        $state = Get-VoucherApprovalStateByIdentity -VchType $VchType -VchNo $VchNo -VchSeries $VchSeries -VchDate $VchDate -InstanceId $InstanceId -CompanyCode $CompanyCode -ExistingFi $fi
        if (-not $state.success) { return $state }
        if (-not $state.found) { return @{ success=$false; error="Created voucher was saved, but its VchCode could not be resolved for approval processing." } }

        $vchCode = [int]$state.vchCode
        $fi.ExecuteQuery("UPDATE Tran1 SET ApprovalStatus=$targetStatus WHERE VchCode=$vchCode")

        # SaveVchFromXML can auto-approve at creation. Normalize a newly-created
        # voucher to the BusyCloud policy before any human can act on it.
        $fi.ExecuteQuery("DELETE FROM CheckList WHERE Code=$vchCode AND Action=3")

        if (-not [string]::IsNullOrWhiteSpace($UserName)) {
            $safeUser = $UserName.Trim().Replace("'", "''")
            $fi.ExecuteQuery("UPDATE CheckList SET UserName='$safeUser' WHERE Code=$vchCode AND Action=1")
        }

        $verify = $fi.GetRecordset("SELECT ApprovalStatus FROM Tran1 WHERE VchCode=$vchCode")
        $actual = -1
        if ($verify -and -not $verify.EOF) { $actual = [int]$verify.Fields.Item("ApprovalStatus").Value }
        if ($verify) { try { $verify.Close() } catch {} }
        if ($actual -ne $targetStatus) {
            return @{ success=$false; error="Voucher was created, but BUSY approval status could not be normalized."; vchCode=$vchCode; approvalStatus=$actual; expectedApprovalStatus=$targetStatus }
        }

        $text = if ($targetStatus -eq 2) { "To be Approved" } else { "Approval Not Required" }
        return @{ success=$true; vchCode=$vchCode; approvalRequired=$approvalRequired; approvalStatus=$targetStatus; approvalStatusText=$text }
    }
    catch { return @{ success=$false; error=$_.Exception.Message } }
}

function Set-WebCreatedVoucherPendingApproval {
    param($fi,[int]$VchType,[string]$VchNo,[string]$VchSeries="",[string]$VchDate="",[string]$UserName,[string]$InstanceId="",[string]$CompanyCode="")
    return Set-WebCreatedVoucherApprovalState -fi $fi -VchType $VchType -VchNo $VchNo -VchSeries $VchSeries -VchDate $VchDate -UserName $UserName -InstanceId $InstanceId -CompanyCode $CompanyCode
}

# -----------------------------------------------------------------------------
# BusyCloud approval read/write helpers.
# IMPORTANT: Approval pages/actions must never open BUSY COM. SQL companies can
# take 1-4 minutes to initialize OpenCSDB on some installations. The approval
# feature only needs Tran1 / CheckList / Master1, so use the direct fiscal DB
# connection resolved by vch_setting.ps1 for both SQL Server and Access/BDS.
# -----------------------------------------------------------------------------

function Read-BusyCloudApprovalReaderValue {
    param($Reader, [string]$Field, $Default = $null)
    try {
        $idx = $Reader.GetOrdinal($Field)
        if ($idx -ge 0 -and -not $Reader.IsDBNull($idx)) { return $Reader.GetValue($idx) }
    } catch {}
    return $Default
}

function Convert-BusyCloudApprovalNumberLiteral {
    param($Value)
    try { return ([double]$Value).ToString('0.################', [System.Globalization.CultureInfo]::InvariantCulture) }
    catch { return '0' }
}

function Get-BusyCloudApprovalDateLiteral {
    param([datetime]$Value, [int]$DbType)
    if ($DbType -eq 1) { return "'" + $Value.ToString('yyyy-MM-dd HH:mm:ss.fff') + "'" }
    return '#' + $Value.ToString('MM/dd/yyyy HH:mm:ss') + '#'
}

function Invoke-BusyCloudApprovalNonQuery {
    param($Connection, $Transaction, [string]$Sql)
    $cmd = $Connection.CreateCommand()
    if ($null -ne $Transaction) { $cmd.Transaction = $Transaction }
    try { $cmd.CommandTimeout = 15 } catch {}
    $cmd.CommandText = $Sql
    return $cmd.ExecuteNonQuery()
}

function Get-VoucherApprovalQueue-Direct {
    param(
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$Status = 'pending',
        [int]$VchType = 0,
        [int]$Limit = 200,
        [string]$InstanceId = '',
        [string]$CompanyCode = ''
    )

    if ([string]::IsNullOrWhiteSpace($UserName)) {
        return @{ success=$false; httpStatus=401; error='Authenticated BUSY user is required.' }
    }

    $allowed = Get-VoucherApprovalTypesForUser -UserName $UserName -IsAdmin:$IsAdmin -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $allowed.success) { return $allowed }
    $allowedTypes = @($allowed.data | ForEach-Object { [int]$_ } | Where-Object { $_ -gt 0 } | Sort-Object -Unique)

    if ($VchType -gt 0) {
        if ($allowedTypes -notcontains $VchType) {
            return @{ success=$false; httpStatus=403; error="You are not allowed to approve voucher type $VchType." }
        }
        $effectiveTypes = @($VchType)
    } else {
        $effectiveTypes = @($allowedTypes)
    }

    $normalized = ([string]$Status).Trim().ToLowerInvariant()
    if ($normalized -notin @('pending','approved','history')) {
        return @{ success=$false; httpStatus=400; error='status must be pending, approved, or history.' }
    }

    if ($effectiveTypes.Count -eq 0) {
        return @{ success=$true; data=@{ allowed_vch_types=@(); status=$normalized; items=@() } }
    }

    if ($normalized -eq 'history') {
        $historyUser = if ($IsAdmin) { '' } else { $UserName }
        $history = Get-BusyCloudVoucherApprovalHistory `
            -VchTypes $effectiveTypes `
            -VchType $VchType `
            -ActionBy $historyUser `
            -Limit $Limit `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
        if (-not $history.success) { return $history }
        return @{ success=$true; data=@{ allowed_vch_types=@($allowedTypes); status='history'; items=@($history.data) } }
    }

    $approvalStatus = if ($normalized -eq 'approved') { 1 } else { 2 }
    if ($Limit -lt 1) { $Limit = 1 }
    if ($Limit -gt 500) { $Limit = 500 }

    $ctx = $null
    try {
        $ctx = Get-BusyCloudApprovalDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $conn = $ctx.connection
        $typesSql = ($effectiveTypes -join ',')

        # 1) Read voucher headers in one query.
        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 15 } catch {}
        $cmd.CommandText = @"
SELECT TOP $Limit
    VchCode,VchType,VchNo,[Date],VchSeriesCode,MasterCode1,
    VchAmtBaseCur,VchSalePurcAmt,ApprovalStatus
FROM Tran1
WHERE VchType IN ($typesSql)
  AND ApprovalStatus=$approvalStatus
ORDER BY VchCode DESC
"@
        $rdr = $cmd.ExecuteReader()
        $rows = @()
        $voucherCodes = @()
        $masterCodes = @()
        while ($rdr.Read()) {
            $code = [int](Read-BusyCloudApprovalReaderValue $rdr 'VchCode' 0)
            $type = [int](Read-BusyCloudApprovalReaderValue $rdr 'VchType' 0)
            $seriesCode = [int](Read-BusyCloudApprovalReaderValue $rdr 'VchSeriesCode' 0)
            $partyCode = [int](Read-BusyCloudApprovalReaderValue $rdr 'MasterCode1' 0)
            $rawDate = Read-BusyCloudApprovalReaderValue $rdr 'Date' $null
            $dateText = ''
            if ($null -ne $rawDate) {
                try { $dateText = ([datetime]$rawDate).ToString('yyyy-MM-dd') }
                catch { $dateText = [string]$rawDate }
            }
            $amount = [double](Read-BusyCloudApprovalReaderValue $rdr 'VchAmtBaseCur' 0)
            if ([math]::Abs($amount) -lt 0.000001) {
                $amount = [double](Read-BusyCloudApprovalReaderValue $rdr 'VchSalePurcAmt' 0)
            }

            $rows += @{
                vchCode=$code
                vchType=$type
                vchNo=([string](Read-BusyCloudApprovalReaderValue $rdr 'VchNo' '')).Trim()
                date=$dateText
                seriesCode=$seriesCode
                partyCode=$partyCode
                amount=[math]::Abs($amount)
            }
            if ($code -gt 0 -and $voucherCodes -notcontains $code) { $voucherCodes += $code }
            foreach ($mc in @($seriesCode,$partyCode)) {
                if ($mc -gt 0 -and $masterCodes -notcontains $mc) { $masterCodes += $mc }
            }
        }
        $rdr.Close()

        if ($rows.Count -eq 0) {
            return @{ success=$true; data=@{ allowed_vch_types=@($allowedTypes); status=$normalized; items=@() } }
        }

        # 2) Read all relevant creator/approval audit markers in one query.
        $createdMap = @{}
        $approvedMap = @{}
        $codesSql = ($voucherCodes -join ',')
        if ($codesSql) {
            $cmd = $conn.CreateCommand()
            try { $cmd.CommandTimeout = 15 } catch {}
            $cmd.CommandText = "SELECT Code,Action,ActionTime,UserName,Notes FROM CheckList WHERE Code IN ($codesSql) AND Action IN (1,3) ORDER BY Code,ActionTime,Action"
            $cr = $cmd.ExecuteReader()
            while ($cr.Read()) {
                $code = [int](Read-BusyCloudApprovalReaderValue $cr 'Code' 0)
                $action = [int](Read-BusyCloudApprovalReaderValue $cr 'Action' 0)
                $user = ([string](Read-BusyCloudApprovalReaderValue $cr 'UserName' '')).Trim()
                $notes = [string](Read-BusyCloudApprovalReaderValue $cr 'Notes' '')
                $rawTime = Read-BusyCloudApprovalReaderValue $cr 'ActionTime' $null
                $timeText = ''
                if ($null -ne $rawTime) { try { $timeText = ([datetime]$rawTime).ToString('yyyy-MM-dd HH:mm:ss') } catch {} }

                if ($action -eq 1 -and -not $createdMap.ContainsKey($code)) {
                    $createdMap[$code] = @{ user=$user; time=$timeText }
                } elseif ($action -eq 3) {
                    # Rows are ordered by time ascending; last Action=3 wins.
                    $approvedMap[$code] = @{ user=$user; time=$timeText; notes=$notes }
                }
            }
            $cr.Close()
        }

        # 3) Resolve only the Master1 names needed by these rows.
        $masterMap = @{}
        if ($masterCodes.Count -gt 0) {
            $mastersSql = ($masterCodes -join ',')
            $cmd = $conn.CreateCommand()
            try { $cmd.CommandTimeout = 15 } catch {}
            $cmd.CommandText = "SELECT Code,Name FROM Master1 WHERE Code IN ($mastersSql)"
            $mr = $cmd.ExecuteReader()
            while ($mr.Read()) {
                $mc = [int](Read-BusyCloudApprovalReaderValue $mr 'Code' 0)
                if ($mc -gt 0) { $masterMap[$mc] = [string](Read-BusyCloudApprovalReaderValue $mr 'Name' '') }
            }
            $mr.Close()
        }

        $items = @()
        foreach ($row in $rows) {
            $code = [int]$row.vchCode
            $type = [int]$row.vchType
            $seriesCode = [int]$row.seriesCode
            $partyCode = [int]$row.partyCode

            $seriesName = if ($masterMap.ContainsKey($seriesCode)) { ([string]$masterMap[$seriesCode]).Trim() } else { 'Main' }
            $prefix = '{0:D2}' -f $type
            if ($seriesName.StartsWith($prefix,[System.StringComparison]::OrdinalIgnoreCase)) {
                $seriesName = $seriesName.Substring($prefix.Length)
            }

            $created = if ($createdMap.ContainsKey($code)) { $createdMap[$code] } else { @{user='';time=''} }
            $approved = if ($approvedMap.ContainsKey($code)) { $approvedMap[$code] } else { @{user='';time='';notes=''} }

            $items += @{
                vchCode=$code
                vchType=$type
                vchNo=[string]$row.vchNo
                date=[string]$row.date
                vchSeriesCode=$seriesCode
                vchSeries=$seriesName
                party=if($masterMap.ContainsKey($partyCode)){[string]$masterMap[$partyCode]}else{''}
                amount=[double]$row.amount
                approvalStatus=$approvalStatus
                approvalStatusText=if($approvalStatus -eq 1){'Approved'}else{'To be Approved'}
                createdBy=[string]$created.user
                createdAt=[string]$created.time
                approvedBy=[string]$approved.user
                approvedAt=[string]$approved.time
                approvalRemarks=[string]$approved.notes
            }
        }

        return @{ success=$true; data=@{ allowed_vch_types=@($allowedTypes); status=$normalized; items=@($items) } }
    }
    catch { return @{ success=$false; error=$_.Exception.Message } }
    finally {
        if ($ctx -and $ctx.connection) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}

function Approve-Voucher-Direct {
    param(
        [int]$VchCode,
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$Remarks = '',
        [string]$InstanceId = '',
        [string]$CompanyCode = ''
    )

    if ($VchCode -le 0) { return @{ success=$false; httpStatus=400; error='vchCode is required.' } }
    if ([string]::IsNullOrWhiteSpace($UserName)) { return @{ success=$false; httpStatus=401; error='Authenticated BUSY user is required.' } }

    $ctx = $null
    $txn = $null
    try {
        $ctx = Get-BusyCloudApprovalDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $conn = $ctx.connection
        $dbType = [int]$ctx.dbType

        $cmd = $conn.CreateCommand(); try{$cmd.CommandTimeout=15}catch{}
        $cmd.CommandText = "SELECT TOP 1 VchType,VchNo,VchSeriesCode,[Date],ApprovalStatus FROM Tran1 WHERE VchCode=$VchCode"
        $r = $cmd.ExecuteReader()
        if (-not $r.Read()) { $r.Close(); return @{success=$false;httpStatus=404;error='Voucher not found.'} }
        $vchType=[int](Read-BusyCloudApprovalReaderValue $r 'VchType' 0)
        $vchNo=([string](Read-BusyCloudApprovalReaderValue $r 'VchNo' '')).Trim()
        $seriesCode=[int](Read-BusyCloudApprovalReaderValue $r 'VchSeriesCode' 0)
        $vchDate=Read-BusyCloudApprovalReaderValue $r 'Date' (Get-Date)
        $current=[int](Read-BusyCloudApprovalReaderValue $r 'ApprovalStatus' -1)
        $r.Close()

        $access=Test-VoucherApprover -UserName $UserName -VchType $vchType -IsAdmin:$IsAdmin -InstanceId $InstanceId -CompanyCode $CompanyCode
        if(-not $access.success){return $access}
        if(-not $access.allowed){return @{success=$false;httpStatus=403;error='You are not allowed to approve this voucher type.'}}
        if($current -eq 0){return @{success=$false;httpStatus=409;error='Approval is not required for this voucher.'}}
        if($current -eq 1){return @{success=$true;message='Voucher is already approved.';data=@{vchCode=$VchCode;vchType=$vchType;vchNo=$vchNo;approvalStatus=1}}}
        if($current -ne 2){return @{success=$false;httpStatus=409;error='Voucher is not in a pending approval state.'}}

        $cmd=$conn.CreateCommand();try{$cmd.CommandTimeout=15}catch{}
        $cmd.CommandText="SELECT TOP 1 Type,D1,D2,D3,D4,D5 FROM CheckList WHERE Code=$VchCode AND Action=1 ORDER BY ActionTime"
        $cr=$cmd.ExecuteReader()
        $checkType=2;$d1='0';$d2='0';$d3='0';$d4='0';$d5='0'
        if($cr.Read()){
            $checkType=[int](Read-BusyCloudApprovalReaderValue $cr 'Type' 2)
            $d1=Convert-BusyCloudApprovalNumberLiteral (Read-BusyCloudApprovalReaderValue $cr 'D1' 0)
            $d2=Convert-BusyCloudApprovalNumberLiteral (Read-BusyCloudApprovalReaderValue $cr 'D2' 0)
            $d3=Convert-BusyCloudApprovalNumberLiteral (Read-BusyCloudApprovalReaderValue $cr 'D3' 0)
            $d4=Convert-BusyCloudApprovalNumberLiteral (Read-BusyCloudApprovalReaderValue $cr 'D4' 0)
            $d5=Convert-BusyCloudApprovalNumberLiteral (Read-BusyCloudApprovalReaderValue $cr 'D5' 0)
        }
        $cr.Close()

        $seriesName='Main'
        if($seriesCode -gt 0){
            $cmd=$conn.CreateCommand();try{$cmd.CommandTimeout=15}catch{}
            $cmd.CommandText="SELECT TOP 1 Name FROM Master1 WHERE Code=$seriesCode"
            $raw=$cmd.ExecuteScalar();if($null-ne$raw-and$raw-ne[System.DBNull]::Value){$seriesName=([string]$raw).Trim()}
        }
        $prefix='{0:D2}' -f $vchType
        if($seriesName.StartsWith($prefix,[System.StringComparison]::OrdinalIgnoreCase)){$seriesName=$seriesName.Substring($prefix.Length)}

        $safeUser=$UserName.Trim().Replace("'","''")
        $safeRemarks=([string]$Remarks).Trim().Replace("'","''")
        $computer=if($env:COMPUTERNAME){$env:COMPUTERNAME}else{'BUSYCLOUD'}
        $safeComputer=$computer.Replace("'","''")
        $now=Get-Date
        $dateExpr=Get-BusyCloudApprovalDateLiteral -Value $now -DbType $dbType

        $txn=$conn.BeginTransaction()
        $updated=Invoke-BusyCloudApprovalNonQuery -Connection $conn -Transaction $txn -Sql "UPDATE Tran1 SET ApprovalStatus=1 WHERE VchCode=$VchCode AND ApprovalStatus=2"
        if([int]$updated -le 0){throw 'Voucher approval state changed before this request completed. Refresh and try again.'}
        [void](Invoke-BusyCloudApprovalNonQuery -Connection $conn -Transaction $txn -Sql "DELETE FROM CheckList WHERE Code=$VchCode AND Action=3")
        [void](Invoke-BusyCloudApprovalNonQuery -Connection $conn -Transaction $txn -Sql "INSERT INTO CheckList (Type,Code,Action,ActionTime,UserName,D1,D2,D3,D4,D5,Notes,ComputerName) VALUES ($checkType,$VchCode,3,$dateExpr,'$safeUser',$d1,$d2,$d3,$d4,$d5,'$safeRemarks','$safeComputer')")
        $txn.Commit();$txn=$null

        $audit=Write-BusyCloudVoucherApprovalAudit -VchCode $VchCode -VchType $vchType -VchNo $vchNo -VchSeries $seriesName -VchDate $vchDate -EventType 'APPROVE' -PreviousStatus 2 -NewStatus 1 -ActionBy $UserName -Remarks ([string]$Remarks).Trim() -InstanceId $InstanceId -CompanyCode $CompanyCode
        $result=@{success=$true;message='Voucher approved successfully.';data=@{vchCode=$VchCode;vchType=$vchType;vchNo=$vchNo;approvalStatus=1;approvalStatusText='Approved';approvedBy=$UserName;approvedAt=$now.ToString('yyyy-MM-dd HH:mm:ss');remarks=([string]$Remarks).Trim()}}
        if(-not$audit.success){$result.warning="Voucher was approved, but BusyCloud history could not be written: $($audit.error)"}
        return $result
    }
    catch {
        if($txn){try{$txn.Rollback()}catch{}}
        return @{success=$false;error=$_.Exception.Message}
    }
    finally {
        if($ctx -and $ctx.connection){try{$ctx.connection.Close()}catch{};try{$ctx.connection.Dispose()}catch{}}
    }
}

function Unapprove-Voucher-Direct {
    param(
        [int]$VchCode,
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$Reason = '',
        [string]$InstanceId = '',
        [string]$CompanyCode = ''
    )

    if($VchCode -le 0){return @{success=$false;httpStatus=400;error='vchCode is required.'}}
    $reasonText=([string]$Reason).Trim()
    if([string]::IsNullOrWhiteSpace($reasonText)){return @{success=$false;httpStatus=400;error='An unapprove reason is required.'}}

    $ctx=$null;$txn=$null
    try {
        $ctx=Get-BusyCloudApprovalDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $conn=$ctx.connection

        $cmd=$conn.CreateCommand();try{$cmd.CommandTimeout=15}catch{}
        $cmd.CommandText="SELECT TOP 1 VchType,VchNo,VchSeriesCode,[Date],ApprovalStatus FROM Tran1 WHERE VchCode=$VchCode"
        $r=$cmd.ExecuteReader()
        if(-not$r.Read()){$r.Close();return @{success=$false;httpStatus=404;error='Voucher not found.'}}
        $vchType=[int](Read-BusyCloudApprovalReaderValue $r 'VchType' 0)
        $vchNo=([string](Read-BusyCloudApprovalReaderValue $r 'VchNo' '')).Trim()
        $seriesCode=[int](Read-BusyCloudApprovalReaderValue $r 'VchSeriesCode' 0)
        $vchDate=Read-BusyCloudApprovalReaderValue $r 'Date' (Get-Date)
        $current=[int](Read-BusyCloudApprovalReaderValue $r 'ApprovalStatus' -1)
        $r.Close()

        $access=Test-VoucherApprover -UserName $UserName -VchType $vchType -IsAdmin:$IsAdmin -InstanceId $InstanceId -CompanyCode $CompanyCode
        if(-not$access.success){return $access}
        if(-not$access.allowed){return @{success=$false;httpStatus=403;error='You are not allowed to unapprove this voucher type.'}}
        if($current -eq 0){return @{success=$false;httpStatus=409;error='Approval is not required for this voucher.'}}
        if($current -eq 2){return @{success=$true;message='Voucher is already pending approval.';data=@{vchCode=$VchCode;approvalStatus=2}}}
        if($current -ne 1){return @{success=$false;httpStatus=409;error='Voucher is not currently approved.'}}

        $cmd=$conn.CreateCommand();try{$cmd.CommandTimeout=15}catch{}
        $cmd.CommandText="SELECT TOP 1 Action FROM CheckList WHERE Code=$VchCode AND Action=3 ORDER BY ActionTime DESC"
        $marker=$cmd.ExecuteScalar()
        if($null-eq$marker-or$marker-eq[System.DBNull]::Value){return @{success=$false;httpStatus=409;error='The approved voucher has no CheckList Action=3 marker; refusing to create an inconsistent unapprove state.'}}

        $seriesName='Main'
        if($seriesCode -gt 0){
            $cmd=$conn.CreateCommand();try{$cmd.CommandTimeout=15}catch{}
            $cmd.CommandText="SELECT TOP 1 Name FROM Master1 WHERE Code=$seriesCode"
            $raw=$cmd.ExecuteScalar();if($null-ne$raw-and$raw-ne[System.DBNull]::Value){$seriesName=([string]$raw).Trim()}
        }
        $prefix='{0:D2}' -f $vchType
        if($seriesName.StartsWith($prefix,[System.StringComparison]::OrdinalIgnoreCase)){$seriesName=$seriesName.Substring($prefix.Length)}

        $txn=$conn.BeginTransaction()
        $updated=Invoke-BusyCloudApprovalNonQuery -Connection $conn -Transaction $txn -Sql "UPDATE Tran1 SET ApprovalStatus=2 WHERE VchCode=$VchCode AND ApprovalStatus=1"
        if([int]$updated -le 0){throw 'Voucher approval state changed before this request completed. Refresh and try again.'}
        [void](Invoke-BusyCloudApprovalNonQuery -Connection $conn -Transaction $txn -Sql "DELETE FROM CheckList WHERE Code=$VchCode AND Action=3")
        $txn.Commit();$txn=$null

        $now=Get-Date
        $audit=Write-BusyCloudVoucherApprovalAudit -VchCode $VchCode -VchType $vchType -VchNo $vchNo -VchSeries $seriesName -VchDate $vchDate -EventType 'UNAPPROVE' -PreviousStatus 1 -NewStatus 2 -ActionBy $UserName -Remarks $reasonText -InstanceId $InstanceId -CompanyCode $CompanyCode
        $result=@{success=$true;message='Voucher moved back to To be Approved.';data=@{vchCode=$VchCode;vchType=$vchType;vchNo=$vchNo;approvalStatus=2;approvalStatusText='To be Approved';unapprovedBy=$UserName;unapprovedAt=$now.ToString('yyyy-MM-dd HH:mm:ss');reason=$reasonText}}
        if(-not$audit.success){$result.warning="Voucher was unapproved, but BusyCloud history could not be written: $($audit.error)"}
        return $result
    }
    catch {
        if($txn){try{$txn.Rollback()}catch{}}
        return @{success=$false;error=$_.Exception.Message}
    }
    finally {
        if($ctx -and $ctx.connection){try{$ctx.connection.Close()}catch{};try{$ctx.connection.Dispose()}catch{}}
    }
}

function Get-VoucherApprovalQueue-AccessCom {
    param(
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$Status = "pending",
        [int]$VchType = 0,
        [int]$Limit = 200,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ([string]::IsNullOrWhiteSpace($UserName)) { return @{ success=$false; httpStatus=401; error="Authenticated BUSY user is required." } }

    $allowed = Get-VoucherApprovalTypesForUser -UserName $UserName -IsAdmin:$IsAdmin -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $allowed.success) { return $allowed }
    $allowedTypes = @($allowed.data | ForEach-Object { [int]$_ } | Where-Object { $_ -gt 0 } | Sort-Object -Unique)

    if ($VchType -gt 0) {
        if ($allowedTypes -notcontains $VchType) { return @{ success=$false; httpStatus=403; error="You are not allowed to approve voucher type $VchType." } }
        $effectiveTypes = @($VchType)
    } else { $effectiveTypes = @($allowedTypes) }

    $normalized = ([string]$Status).Trim().ToLowerInvariant()
    if ($normalized -notin @("pending","approved","history")) { return @{ success=$false; httpStatus=400; error="status must be pending, approved, or history." } }

    if ($effectiveTypes.Count -eq 0) {
        return @{ success=$true; data=@{ allowed_vch_types=@(); status=$normalized; items=@() } }
    }

    if ($normalized -eq "history") {
        $historyUser = if ($IsAdmin) { "" } else { $UserName }
        $history = Get-BusyCloudVoucherApprovalHistory `
            -VchTypes $effectiveTypes `
            -VchType $VchType `
            -ActionBy $historyUser `
            -Limit $Limit `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
        if (-not $history.success) { return $history }
        return @{ success=$true; data=@{ allowed_vch_types=@($allowedTypes); status="history"; items=@($history.data) } }
    }

    $approvalStatus = if ($normalized -eq "approved") { 1 } else { 2 }
    if ($Limit -lt 1) { $Limit=1 }; if ($Limit -gt 500) { $Limit=500 }

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success=$false; error="BUSY database connection failed" } }

    try {
        function Read-ApprovalField { param($R,[string]$F,$D); try { $v=$R.Fields.Item($F).Value; if ($null -ne $v -and $v -ne [System.DBNull]::Value) { return $v } } catch {}; return $D }

        $masterMap = @{}
        $m = $fi.GetRecordset("SELECT Code,Name FROM Master1")
        if ($m) {
            while (-not $m.EOF) { try { $c=[int]$m.Fields.Item("Code").Value; if ($c -gt 0) { $masterMap[$c]=[string]$m.Fields.Item("Name").Value } } catch {}; $m.MoveNext() }
            try { $m.Close() } catch {}
        }

        $typesSql = ($effectiveTypes -join ",")
        $rst = $fi.GetRecordset(@"
SELECT TOP $Limit VchCode,VchType,VchNo,[Date],VchSeriesCode,MasterCode1,VchAmtBaseCur,VchSalePurcAmt,ApprovalStatus
FROM Tran1
WHERE VchType IN ($typesSql) AND ApprovalStatus=$approvalStatus
ORDER BY VchCode DESC
"@)

        $items = @()
        if ($rst) {
            while (-not $rst.EOF) {
                $code=[int](Read-ApprovalField $rst "VchCode" 0); $type=[int](Read-ApprovalField $rst "VchType" 0)
                $createdBy=""; $createdAt=""; $approvedBy=""; $approvedAt=""; $approvalRemarks=""

                $createdRs = $fi.GetRecordset("SELECT TOP 1 UserName,ActionTime FROM CheckList WHERE Code=$code AND Action=1 ORDER BY ActionTime")
                if ($createdRs -and -not $createdRs.EOF) {
                    $createdBy=([string](Read-ApprovalField $createdRs "UserName" "")).Trim(); $raw=Read-ApprovalField $createdRs "ActionTime" $null
                    if ($null -ne $raw) { try { $createdAt=([datetime]$raw).ToString("yyyy-MM-dd HH:mm:ss") } catch {} }
                }
                if ($createdRs) { try { $createdRs.Close() } catch {} }

                if ($approvalStatus -eq 1) {
                    $a = $fi.GetRecordset("SELECT TOP 1 UserName,ActionTime,Notes FROM CheckList WHERE Code=$code AND Action=3 ORDER BY ActionTime DESC")
                    if ($a -and -not $a.EOF) {
                        $approvedBy=([string](Read-ApprovalField $a "UserName" "")).Trim(); $approvalRemarks=[string](Read-ApprovalField $a "Notes" ""); $raw=Read-ApprovalField $a "ActionTime" $null
                        if ($null -ne $raw) { try { $approvedAt=([datetime]$raw).ToString("yyyy-MM-dd HH:mm:ss") } catch {} }
                    }
                    if ($a) { try { $a.Close() } catch {} }
                }

                # Any currently assigned approver can Unapprove an approved
                # voucher in an allowed voucher type. Therefore the Approved
                # list is type-scoped, not "approved by me" scoped.
                $include = $true

                if ($include) {
                    $seriesCode=[int](Read-ApprovalField $rst "VchSeriesCode" 0); $masterCode1=[int](Read-ApprovalField $rst "MasterCode1" 0)
                    $seriesName=if($masterMap.ContainsKey($seriesCode)){[string]$masterMap[$seriesCode]}else{"Main"}; $prefix="{0:D2}" -f $type
                    if($seriesName.StartsWith($prefix,[System.StringComparison]::OrdinalIgnoreCase)){$seriesName=$seriesName.Substring($prefix.Length)}
                    $rawDate=Read-ApprovalField $rst "Date" $null; $dateText=""; if($null-ne $rawDate){try{$dateText=([datetime]$rawDate).ToString("yyyy-MM-dd")}catch{$dateText=[string]$rawDate}}
                    $amount=[double](Read-ApprovalField $rst "VchAmtBaseCur" 0); if([math]::Abs($amount)-lt 0.000001){$amount=[double](Read-ApprovalField $rst "VchSalePurcAmt" 0)}
                    $items += @{
                        vchCode=$code; vchType=$type; vchNo=[string](Read-ApprovalField $rst "VchNo" ""); date=$dateText; vchSeriesCode=$seriesCode; vchSeries=$seriesName;
                        party=if($masterMap.ContainsKey($masterCode1)){[string]$masterMap[$masterCode1]}else{""}; amount=[math]::Abs($amount);
                        approvalStatus=$approvalStatus; approvalStatusText=if($approvalStatus-eq1){"Approved"}else{"To be Approved"};
                        createdBy=$createdBy; createdAt=$createdAt; approvedBy=$approvedBy; approvedAt=$approvedAt; approvalRemarks=$approvalRemarks
                    }
                }
                $rst.MoveNext()
            }
            try { $rst.Close() } catch {}
        }

        return @{ success=$true; data=@{ allowed_vch_types=@($allowedTypes); status=$normalized; items=@($items) } }
    }
    catch { return @{ success=$false; error=$_.Exception.Message } }
    finally { Disconnect-BUSY $fi }
}

function Approve-Voucher-AccessCom {
    param(
        [int]$VchCode,
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$Remarks = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ($VchCode -le 0) { return @{ success=$false; httpStatus=400; error="vchCode is required." } }
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success=$false; error="BUSY database connection failed" } }

    try {
        $rst=$fi.GetRecordset("SELECT VchType,VchNo,VchSeriesCode,[Date],ApprovalStatus FROM Tran1 WHERE VchCode=$VchCode")
        if(-not $rst -or $rst.EOF){return @{success=$false;httpStatus=404;error="Voucher not found."}}
        $vchType=[int]$rst.Fields.Item("VchType").Value; $vchNo=([string]$rst.Fields.Item("VchNo").Value).Trim(); $seriesCode=[int]$rst.Fields.Item("VchSeriesCode").Value; $vchDate=$rst.Fields.Item("Date").Value; $current=[int]$rst.Fields.Item("ApprovalStatus").Value
        try{$rst.Close()}catch{}

        $access=Test-VoucherApprover -UserName $UserName -VchType $vchType -IsAdmin:$IsAdmin -InstanceId $InstanceId -CompanyCode $CompanyCode -ExistingFi $fi
        if(-not $access.success){return $access}; if(-not $access.allowed){return @{success=$false;httpStatus=403;error="You are not allowed to approve this voucher type."}}
        if($current-eq0){return @{success=$false;httpStatus=409;error="Approval is not required for this voucher."}}
        if($current-eq1){return @{success=$true;message="Voucher is already approved.";data=@{vchCode=$VchCode;vchType=$vchType;vchNo=$vchNo;approvalStatus=1}}}
        if($current-ne2){return @{success=$false;httpStatus=409;error="Voucher is not in a pending approval state."}}

        $storage=Ensure-BusyCloudVoucherApprovalAuditTable -InstanceId $InstanceId -CompanyCode $CompanyCode
        if(-not $storage.success){return @{success=$false;error="Approval audit storage is unavailable. $($storage.error)"}}

        function RAF { param($R,[string]$F,$D);try{$v=$R.Fields.Item($F).Value;if($null-ne$v-and$v-ne[System.DBNull]::Value){return$v}}catch{};return$D }
        function NumLit { param($V);try{return([double]$V).ToString("0.################",[System.Globalization.CultureInfo]::InvariantCulture)}catch{return"0"} }

        $src=$fi.GetRecordset("SELECT TOP 1 Type,D1,D2,D3,D4,D5 FROM CheckList WHERE Code=$VchCode AND Action=1 ORDER BY ActionTime")
        $checkType=2;$d1="0";$d2="0";$d3="0";$d4="0";$d5="0"
        if($src-and-not$src.EOF){$checkType=[int](RAF $src "Type" 2);$d1=NumLit(RAF $src "D1" 0);$d2=NumLit(RAF $src "D2" 0);$d3=NumLit(RAF $src "D3" 0);$d4=NumLit(RAF $src "D4" 0);$d5=NumLit(RAF $src "D5" 0)}
        if($src){try{$src.Close()}catch{}}

        $seriesName="Main";$sr=$fi.GetRecordset("SELECT Name FROM Master1 WHERE Code=$seriesCode");if($sr-and-not$sr.EOF){$seriesName=([string]$sr.Fields.Item("Name").Value).Trim()};if($sr){try{$sr.Close()}catch{}};$prefix="{0:D2}"-f$vchType;if($seriesName.StartsWith($prefix,[System.StringComparison]::OrdinalIgnoreCase)){$seriesName=$seriesName.Substring($prefix.Length)}

        $safeUser=$UserName.Trim().Replace("'","''");$safeRemarks=([string]$Remarks).Trim().Replace("'","''");$computer=if($env:COMPUTERNAME){$env:COMPUTERNAME}else{"BUSYCLOUD"};$safeComputer=$computer.Replace("'","''")
        $found=Get-InstanceForCompany -CompanyCode $CompanyCode -InstanceId $InstanceId;$dbType=if($found-and$null-ne$found.instance.dbType){[int]$found.instance.dbType}else{0};$nowExpr=if($dbType-eq1){"GETDATE()"}else{"Now()"}

        $fi.ExecuteQuery("UPDATE Tran1 SET ApprovalStatus=1 WHERE VchCode=$VchCode AND ApprovalStatus=2")
        try {
            $fi.ExecuteQuery("DELETE FROM CheckList WHERE Code=$VchCode AND Action=3")
            $fi.ExecuteQuery("INSERT INTO CheckList (Type,Code,Action,ActionTime,UserName,D1,D2,D3,D4,D5,Notes,ComputerName) VALUES ($checkType,$VchCode,3,$nowExpr,'$safeUser',$d1,$d2,$d3,$d4,$d5,'$safeRemarks','$safeComputer')")
            $audit=Write-BusyCloudVoucherApprovalAudit -VchCode $VchCode -VchType $vchType -VchNo $vchNo -VchSeries $seriesName -VchDate $vchDate -EventType "APPROVE" -PreviousStatus 2 -NewStatus 1 -ActionBy $UserName -Remarks ([string]$Remarks).Trim() -InstanceId $InstanceId -CompanyCode $CompanyCode
            if(-not$audit.success){throw "Could not write BusyCloud approval history. $($audit.error)"}
        } catch {
            try{$fi.ExecuteQuery("DELETE FROM CheckList WHERE Code=$VchCode AND Action=3")}catch{};try{$fi.ExecuteQuery("UPDATE Tran1 SET ApprovalStatus=2 WHERE VchCode=$VchCode")}catch{};throw
        }

        return @{success=$true;message="Voucher approved successfully.";data=@{vchCode=$VchCode;vchType=$vchType;vchNo=$vchNo;approvalStatus=1;approvalStatusText="Approved";approvedBy=$UserName;approvedAt=$audit.actionTime;remarks=([string]$Remarks).Trim()}}
    }
    catch{return @{success=$false;error=$_.Exception.Message}}
    finally{Disconnect-BUSY $fi}
}

function Unapprove-Voucher-AccessCom {
    param(
        [int]$VchCode,
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$Reason = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if($VchCode-le0){return @{success=$false;httpStatus=400;error="vchCode is required."}}
    $reasonText=([string]$Reason).Trim();if([string]::IsNullOrWhiteSpace($reasonText)){return @{success=$false;httpStatus=400;error="An unapprove reason is required."}}
    $fi=Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode;if(-not$fi){return @{success=$false;error="BUSY database connection failed"}}

    try {
        function URF { param($R,[string]$F,$D);try{$v=$R.Fields.Item($F).Value;if($null-ne$v-and$v-ne[System.DBNull]::Value){return$v}}catch{};return$D }
        function UNum { param($V);try{return([double]$V).ToString("0.################",[System.Globalization.CultureInfo]::InvariantCulture)}catch{return"0"} }

        $rst=$fi.GetRecordset("SELECT VchType,VchNo,VchSeriesCode,[Date],ApprovalStatus FROM Tran1 WHERE VchCode=$VchCode");if(-not$rst-or$rst.EOF){return @{success=$false;httpStatus=404;error="Voucher not found."}}
        $vchType=[int]$rst.Fields.Item("VchType").Value;$vchNo=([string]$rst.Fields.Item("VchNo").Value).Trim();$seriesCode=[int]$rst.Fields.Item("VchSeriesCode").Value;$vchDate=$rst.Fields.Item("Date").Value;$current=[int]$rst.Fields.Item("ApprovalStatus").Value;try{$rst.Close()}catch{}

        $access=Test-VoucherApprover -UserName $UserName -VchType $vchType -IsAdmin:$IsAdmin -InstanceId $InstanceId -CompanyCode $CompanyCode -ExistingFi $fi;if(-not$access.success){return$access};if(-not$access.allowed){return @{success=$false;httpStatus=403;error="You are not allowed to unapprove this voucher type."}}
        if($current-eq0){return @{success=$false;httpStatus=409;error="Approval is not required for this voucher."}};if($current-eq2){return @{success=$true;message="Voucher is already pending approval.";data=@{vchCode=$VchCode;approvalStatus=2}}};if($current-ne1){return @{success=$false;httpStatus=409;error="Voucher is not currently approved."}}

        $approval=$fi.GetRecordset("SELECT TOP 1 Type,ActionTime,UserName,D1,D2,D3,D4,D5,Notes,ComputerName FROM CheckList WHERE Code=$VchCode AND Action=3 ORDER BY ActionTime DESC")
        if(-not$approval-or$approval.EOF){return @{success=$false;httpStatus=409;error="The approved voucher has no CheckList Action=3 marker; refusing to create an inconsistent unapprove state."}}
        $aType=[int](URF $approval "Type" 2);$aTime=URF $approval "ActionTime" (Get-Date);$aUser=[string](URF $approval "UserName" "");$ad1=UNum(URF $approval "D1" 0);$ad2=UNum(URF $approval "D2" 0);$ad3=UNum(URF $approval "D3" 0);$ad4=UNum(URF $approval "D4" 0);$ad5=UNum(URF $approval "D5" 0);$aNotes=[string](URF $approval "Notes" "");$aComputer=[string](URF $approval "ComputerName" "BUSYCLOUD");try{$approval.Close()}catch{}

        $seriesName="Main";$sr=$fi.GetRecordset("SELECT Name FROM Master1 WHERE Code=$seriesCode");if($sr-and-not$sr.EOF){$seriesName=([string]$sr.Fields.Item("Name").Value).Trim()};if($sr){try{$sr.Close()}catch{}};$prefix="{0:D2}"-f$vchType;if($seriesName.StartsWith($prefix,[System.StringComparison]::OrdinalIgnoreCase)){$seriesName=$seriesName.Substring($prefix.Length)}
        $storage=Ensure-BusyCloudVoucherApprovalAuditTable -InstanceId $InstanceId -CompanyCode $CompanyCode;if(-not$storage.success){return @{success=$false;error="Approval audit storage is unavailable. $($storage.error)"}}

        $fi.ExecuteQuery("UPDATE Tran1 SET ApprovalStatus=2 WHERE VchCode=$VchCode AND ApprovalStatus=1")
        try {
            $fi.ExecuteQuery("DELETE FROM CheckList WHERE Code=$VchCode AND Action=3")
            $audit=Write-BusyCloudVoucherApprovalAudit -VchCode $VchCode -VchType $vchType -VchNo $vchNo -VchSeries $seriesName -VchDate $vchDate -EventType "UNAPPROVE" -PreviousStatus 1 -NewStatus 2 -ActionBy $UserName -Remarks $reasonText -InstanceId $InstanceId -CompanyCode $CompanyCode
            if(-not$audit.success){throw "Could not write BusyCloud unapprove history. $($audit.error)"}
        } catch {
            try{$fi.ExecuteQuery("UPDATE Tran1 SET ApprovalStatus=1 WHERE VchCode=$VchCode")}catch{}
            try {
                $found=Get-InstanceForCompany -CompanyCode $CompanyCode -InstanceId $InstanceId;$dbType=if($found-and$null-ne$found.instance.dbType){[int]$found.instance.dbType}else{0};$dt=[datetime]$aTime;$dateExpr=if($dbType-eq1){"'"+$dt.ToString("yyyy-MM-dd HH:mm:ss.fff")+"'"}else{"#"+$dt.ToString("MM/dd/yyyy HH:mm:ss")+"#"}
                $su=$aUser.Replace("'","''");$sn=$aNotes.Replace("'","''");$sc=$aComputer.Replace("'","''")
                $fi.ExecuteQuery("INSERT INTO CheckList (Type,Code,Action,ActionTime,UserName,D1,D2,D3,D4,D5,Notes,ComputerName) VALUES ($aType,$VchCode,3,$dateExpr,'$su',$ad1,$ad2,$ad3,$ad4,$ad5,'$sn','$sc')")
            } catch {}
            throw
        }

        return @{success=$true;message="Voucher moved back to To be Approved.";data=@{vchCode=$VchCode;vchType=$vchType;vchNo=$vchNo;approvalStatus=2;approvalStatusText="To be Approved";unapprovedBy=$UserName;unapprovedAt=$audit.actionTime;reason=$reasonText}}
    }
    catch{return @{success=$false;error=$_.Exception.Message}}
    finally{Disconnect-BUSY $fi}
}




# -----------------------------------------------------------------------------
# HOTFIX V5: SQL approval operations stay on the fast direct DB path.
# Access/BDS approval operations use BUSY COM because the direct db.bds used for
# authentication does not expose the fiscal Tran1 / CheckList / Master1 / Config
# tables consistently. Access COM initialization is local and fast in the
# observed installation, while SQL COM remains deliberately avoided.
# -----------------------------------------------------------------------------

function Test-BusyCloudApprovalAccessRuntime {
    param(
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    try {
        $found = Get-InstanceForCompany -CompanyCode $CompanyCode -InstanceId $InstanceId
        if (-not $found) { return $false }
        $dbType = if ($null -ne $found.instance.dbType) { [int]$found.instance.dbType } else { 0 }
        return ($dbType -eq 0)
    }
    catch {
        return $false
    }
}

function Get-VoucherApprovalQueue {
    param(
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$Status = 'pending',
        [int]$VchType = 0,
        [int]$Limit = 200,
        [string]$InstanceId = '',
        [string]$CompanyCode = ''
    )

    if (Test-BusyCloudApprovalAccessRuntime -InstanceId $InstanceId -CompanyCode $CompanyCode) {
        return Get-VoucherApprovalQueue-AccessCom `
            -UserName $UserName `
            -IsAdmin:$IsAdmin `
            -Status $Status `
            -VchType $VchType `
            -Limit $Limit `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }

    return Get-VoucherApprovalQueue-Direct `
        -UserName $UserName `
        -IsAdmin:$IsAdmin `
        -Status $Status `
        -VchType $VchType `
        -Limit $Limit `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode
}

function Approve-Voucher {
    param(
        [int]$VchCode,
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$Remarks = '',
        [string]$InstanceId = '',
        [string]$CompanyCode = ''
    )

    if (Test-BusyCloudApprovalAccessRuntime -InstanceId $InstanceId -CompanyCode $CompanyCode) {
        return Approve-Voucher-AccessCom `
            -VchCode $VchCode `
            -UserName $UserName `
            -IsAdmin:$IsAdmin `
            -Remarks $Remarks `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }

    return Approve-Voucher-Direct `
        -VchCode $VchCode `
        -UserName $UserName `
        -IsAdmin:$IsAdmin `
        -Remarks $Remarks `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode
}

function Unapprove-Voucher {
    param(
        [int]$VchCode,
        [string]$UserName,
        [bool]$IsAdmin = $false,
        [string]$Reason = '',
        [string]$InstanceId = '',
        [string]$CompanyCode = ''
    )

    if (Test-BusyCloudApprovalAccessRuntime -InstanceId $InstanceId -CompanyCode $CompanyCode) {
        return Unapprove-Voucher-AccessCom `
            -VchCode $VchCode `
            -UserName $UserName `
            -IsAdmin:$IsAdmin `
            -Reason $Reason `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }

    return Unapprove-Voucher-Direct `
        -VchCode $VchCode `
        -UserName $UserName `
        -IsAdmin:$IsAdmin `
        -Reason $Reason `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode
}


# =============================================================================
# BusyCloud Voucher Approval v6 list/search/pagination
# =============================================================================

function Convert-BusyCloudApprovalFilterDate {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $formats = @('yyyy-MM-dd','dd-MM-yyyy','MM/dd/yyyy','yyyy/MM/dd')
    foreach ($fmt in $formats) {
        $dt = [datetime]::MinValue
        if ([datetime]::TryParseExact($Value.Trim(), $fmt, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
            return $dt.Date
        }
    }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($Value, [ref]$parsed)) { return $parsed.Date }
    return $null
}

function New-BusyCloudApprovalPagination {
    param([int]$Page, [int]$PageSize, [int]$Total)
    if ($Page -lt 1) { $Page = 1 }
    if ($PageSize -le 0) {
        return @{
            page=1; page_size='all'; total=$Total; total_pages=1;
            has_previous=$false; has_next=$false
        }
    }
    $totalPages = if ($Total -le 0) { 1 } else { [int][math]::Ceiling($Total / [double]$PageSize) }
    if ($Page -gt $totalPages) { $Page = $totalPages }
    return @{
        page=$Page; page_size=$PageSize; total=$Total; total_pages=$totalPages;
        has_previous=($Page -gt 1); has_next=($Page -lt $totalPages)
    }
}

function Get-BusyCloudVoucherApprovalHistoryV6 {
    param(
        [int[]]$VchTypes=@(),
        [int]$VchType=0,
        [string]$ActionBy='',
        [string]$Search='',
        [string]$FromDate='',
        [string]$ToDate='',
        [int]$Page=1,
        [int]$PageSize=50,
        [string]$InstanceId='',
        [string]$CompanyCode=''
    )
    $ensure = Ensure-BusyCloudVoucherApprovalAuditTable -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $ensure.success) { return $ensure }

    $effectiveTypes = @($VchTypes | ForEach-Object { [int]$_ } | Where-Object { Test-IsBusyCloudApprovalVoucherType -VchType $_ } | Sort-Object -Unique)
    if ($VchType -gt 0) { $effectiveTypes = @($VchType) }
    if ($effectiveTypes.Count -eq 0) {
        return @{ success=$true; data=@(); pagination=(New-BusyCloudApprovalPagination -Page 1 -PageSize $PageSize -Total 0) }
    }

    if ($Page -lt 1) { $Page=1 }
    if ($PageSize -notin @(0,50,100,200,300,500)) { $PageSize=50 }
    $from = Convert-BusyCloudApprovalFilterDate $FromDate
    $to = Convert-BusyCloudApprovalFilterDate $ToDate
    $needle = ([string]$Search).Trim().ToLowerInvariant()

    try {
        $items = @()
        foreach ($line in @(Get-Content -Path $ensure.path -ErrorAction SilentlyContinue)) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            try { $r = $line | ConvertFrom-Json } catch { continue }
            $t=0; try { $t=[int]$r.vchType } catch {}
            if ($effectiveTypes -notcontains $t) { continue }
            if ($ActionBy -and -not ([string]$r.actionBy).Equals($ActionBy,[System.StringComparison]::OrdinalIgnoreCase)) { continue }

            $actionDt=$null
            try { $actionDt=[datetime]$r.actionTime } catch {}
            if ($from -and $actionDt -and $actionDt.Date -lt $from.Date) { continue }
            if ($to -and $actionDt -and $actionDt.Date -gt $to.Date) { continue }

            if ($needle) {
                $hay = @(
                    [string]$r.vchNo, [string]$r.vchSeries, [string]$r.eventType,
                    [string]$r.actionBy, [string]$r.remarks, [string]$r.vchCode
                ) -join ' '
                if (-not $hay.ToLowerInvariant().Contains($needle)) { continue }
            }
            $items += $r
        }

        $items = @($items | Sort-Object @{Expression={ try{[datetime]$_.actionTime}catch{[datetime]::MinValue} };Descending=$true}, @{Expression={ try{[long]$_.id}catch{0} };Descending=$true})
        $total = $items.Count
        if ($PageSize -le 0) {
            $paged = @($items)
            $page = 1
        } else {
            $totalPages = if ($total -le 0) { 1 } else { [int][math]::Ceiling($total/[double]$PageSize) }
            if ($Page -gt $totalPages) { $Page=$totalPages }
            $skip = ($Page-1)*$PageSize
            $paged = @($items | Select-Object -Skip $skip -First $PageSize)
        }
        return @{ success=$true; data=@($paged); pagination=(New-BusyCloudApprovalPagination -Page $Page -PageSize $PageSize -Total $total) }
    }
    catch { return @{ success=$false; error=$_.Exception.Message } }
}

function Get-VoucherApprovalQueue-DirectV6 {
    param(
        [string]$UserName,
        [bool]$IsAdmin=$false,
        [string]$Status='pending',
        [int]$VchType=0,
        [string]$Search='',
        [string]$FromDate='',
        [string]$ToDate='',
        [int]$Page=1,
        [int]$PageSize=50,
        [string]$InstanceId='',
        [string]$CompanyCode=''
    )
    if ([string]::IsNullOrWhiteSpace($UserName)) { return @{success=$false;httpStatus=401;error='Authenticated BUSY user is required.'} }
    $allowed = Get-VoucherApprovalTypesForUser -UserName $UserName -IsAdmin:$IsAdmin -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $allowed.success) { return $allowed }
    $allowedTypes=@($allowed.data | ForEach-Object{[int]$_} | Where-Object{$_ -gt 0} | Sort-Object -Unique)
    if ($VchType -le 0) {
        return @{success=$true;data=@{allowed_vch_types=@($allowedTypes);status=([string]$Status).Trim().ToLowerInvariant();items=@();pagination=(New-BusyCloudApprovalPagination -Page 1 -PageSize $PageSize -Total 0)}}
    }
    if ($allowedTypes -notcontains $VchType) { return @{success=$false;httpStatus=403;error="Approval processing is not enabled for voucher type $VchType, or you are not an approver for it."} }
    $normalized=([string]$Status).Trim().ToLowerInvariant()
    if($normalized -notin @('pending','approved','history')){return @{success=$false;httpStatus=400;error='status must be pending, approved, or history.'}}
    if($Page -lt 1){$Page=1}; if($PageSize -notin @(0,50,100,200,300,500)){$PageSize=50}

    if($normalized -eq 'history'){
        $historyUser=if($IsAdmin){''}else{$UserName}
        $h=Get-BusyCloudVoucherApprovalHistoryV6 -VchTypes @($VchType) -VchType $VchType -ActionBy $historyUser -Search $Search -FromDate $FromDate -ToDate $ToDate -Page $Page -PageSize $PageSize -InstanceId $InstanceId -CompanyCode $CompanyCode
        if(-not $h.success){return$h}
        return @{success=$true;data=@{allowed_vch_types=@($allowedTypes);status='history';items=@($h.data);pagination=$h.pagination}}
    }

    $approvalStatus=if($normalized -eq 'approved'){1}else{2}
    $ctx=$null
    try{
        $ctx=Get-BusyCloudApprovalDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $conn=$ctx.connection
        $where=@("T.VchType=$VchType","T.ApprovalStatus=$approvalStatus")
        $from=Convert-BusyCloudApprovalFilterDate $FromDate
        $to=Convert-BusyCloudApprovalFilterDate $ToDate
        if($from){$where += "T.[Date] >= '"+$from.ToString('yyyy-MM-dd')+"'"}
        if($to){$where += "T.[Date] < '"+$to.AddDays(1).ToString('yyyy-MM-dd')+"'"}
        $needle=([string]$Search).Trim()
        if($needle){
            $safe=$needle.Replace("'","''")
            $like="%$safe%"
            $where += "(T.VchNo LIKE '$like' OR CAST(T.VchCode AS VARCHAR(30)) LIKE '$like' OR ISNULL(P.Name,'') LIKE '$like' OR ISNULL(S.Name,'') LIKE '$like' OR EXISTS (SELECT 1 FROM CheckList C WHERE C.Code=T.VchCode AND C.Action IN (1,3) AND (ISNULL(C.UserName,'') LIKE '$like' OR ISNULL(C.Notes,'') LIKE '$like')))"
        }
        $whereSql=$where -join ' AND '

        $countCmd=$conn.CreateCommand();try{$countCmd.CommandTimeout=20}catch{}
        $countCmd.CommandText="SELECT COUNT(*) FROM Tran1 T LEFT JOIN Master1 P ON P.Code=T.MasterCode1 LEFT JOIN Master1 S ON S.Code=T.VchSeriesCode WHERE $whereSql"
        $total=[int]$countCmd.ExecuteScalar()
        $pagination=New-BusyCloudApprovalPagination -Page $Page -PageSize $PageSize -Total $total
        $Page=[int]$pagination.page

        $sql=@"
SELECT T.VchCode,T.VchType,T.VchNo,T.[Date],T.VchSeriesCode,T.MasterCode1,
       T.VchAmtBaseCur,T.VchSalePurcAmt,T.ApprovalStatus,
       ISNULL(P.Name,'') AS PartyName, ISNULL(S.Name,'') AS SeriesName
FROM Tran1 T
LEFT JOIN Master1 P ON P.Code=T.MasterCode1
LEFT JOIN Master1 S ON S.Code=T.VchSeriesCode
WHERE $whereSql
ORDER BY T.VchCode DESC
"@
        if($PageSize -gt 0){
            $offset=($Page-1)*$PageSize
            $sql += " OFFSET $offset ROWS FETCH NEXT $PageSize ROWS ONLY"
        }
        $cmd=$conn.CreateCommand();try{$cmd.CommandTimeout=20}catch{};$cmd.CommandText=$sql
        $rdr=$cmd.ExecuteReader();$rows=@();$codes=@()
        while($rdr.Read()){
            $code=[int](Read-BusyCloudApprovalReaderValue $rdr 'VchCode' 0)
            $rawDate=Read-BusyCloudApprovalReaderValue $rdr 'Date' $null;$dateText='';if($rawDate){try{$dateText=([datetime]$rawDate).ToString('yyyy-MM-dd')}catch{$dateText=[string]$rawDate}}
            $amount=[double](Read-BusyCloudApprovalReaderValue $rdr 'VchAmtBaseCur' 0);if([math]::Abs($amount) -lt .000001){$amount=[double](Read-BusyCloudApprovalReaderValue $rdr 'VchSalePurcAmt' 0)}
            $seriesName=([string](Read-BusyCloudApprovalReaderValue $rdr 'SeriesName' 'Main')).Trim();if(-not $seriesName){$seriesName='Main'};$prefix='{0:D2}' -f $VchType;if($seriesName.StartsWith($prefix,[System.StringComparison]::OrdinalIgnoreCase)){$seriesName=$seriesName.Substring($prefix.Length)}
            $rows += @{vchCode=$code;vchType=$VchType;vchNo=([string](Read-BusyCloudApprovalReaderValue $rdr 'VchNo' '')).Trim();date=$dateText;vchSeriesCode=[int](Read-BusyCloudApprovalReaderValue $rdr 'VchSeriesCode' 0);vchSeries=$seriesName;party=([string](Read-BusyCloudApprovalReaderValue $rdr 'PartyName' '')).Trim();amount=[math]::Abs($amount)}
            if($code -gt 0){$codes += $code}
        };$rdr.Close()

        $createdMap=@{};$approvedMap=@{}
        if($codes.Count -gt 0){
            $codesSql=($codes -join ',');$ac=$conn.CreateCommand();try{$ac.CommandTimeout=20}catch{};$ac.CommandText="SELECT Code,Action,ActionTime,UserName,Notes FROM CheckList WHERE Code IN ($codesSql) AND Action IN (1,3) ORDER BY Code,ActionTime,Action"
            $ar=$ac.ExecuteReader();while($ar.Read()){
                $code=[int](Read-BusyCloudApprovalReaderValue $ar 'Code' 0);$action=[int](Read-BusyCloudApprovalReaderValue $ar 'Action' 0);$user=([string](Read-BusyCloudApprovalReaderValue $ar 'UserName' '')).Trim();$notes=[string](Read-BusyCloudApprovalReaderValue $ar 'Notes' '');$raw=Read-BusyCloudApprovalReaderValue $ar 'ActionTime' $null;$time='';if($raw){try{$time=([datetime]$raw).ToString('yyyy-MM-dd HH:mm:ss')}catch{}}
                if($action -eq 1 -and -not $createdMap.ContainsKey($code)){$createdMap[$code]=@{user=$user;time=$time}}elseif($action -eq 3){$approvedMap[$code]=@{user=$user;time=$time;notes=$notes}}
            };$ar.Close()
        }
        $items=@();foreach($row in $rows){$c=[int]$row.vchCode;$cr=if($createdMap.ContainsKey($c)){$createdMap[$c]}else{@{user='';time=''}};$ap=if($approvedMap.ContainsKey($c)){$approvedMap[$c]}else{@{user='';time='';notes=''}};$items+=@{vchCode=$c;vchType=$row.vchType;vchNo=$row.vchNo;date=$row.date;vchSeriesCode=$row.vchSeriesCode;vchSeries=$row.vchSeries;party=$row.party;amount=$row.amount;approvalStatus=$approvalStatus;approvalStatusText=if($approvalStatus -eq 1){'Approved'}else{'To be Approved'};createdBy=$cr.user;createdAt=$cr.time;approvedBy=$ap.user;approvedAt=$ap.time;approvalRemarks=$ap.notes}}
        return @{success=$true;data=@{allowed_vch_types=@($allowedTypes);status=$normalized;items=@($items);pagination=$pagination}}
    }catch{return @{success=$false;error=$_.Exception.Message}}finally{if($ctx -and $ctx.connection){try{$ctx.connection.Close()}catch{};try{$ctx.connection.Dispose()}catch{}}}
}

function Get-VoucherApprovalQueue-AccessComV6 {
    param(
        [string]$UserName,[bool]$IsAdmin=$false,[string]$Status='pending',[int]$VchType=0,
        [string]$Search='',[string]$FromDate='',[string]$ToDate='',[int]$Page=1,[int]$PageSize=50,
        [string]$InstanceId='',[string]$CompanyCode=''
    )
    if([string]::IsNullOrWhiteSpace($UserName)){return @{success=$false;httpStatus=401;error='Authenticated BUSY user is required.'}}
    $allowed=Get-VoucherApprovalTypesForUser -UserName $UserName -IsAdmin:$IsAdmin -InstanceId $InstanceId -CompanyCode $CompanyCode
    if(-not $allowed.success){return$allowed};$allowedTypes=@($allowed.data|ForEach-Object{[int]$_}|Where-Object{$_ -gt 0}|Sort-Object -Unique)
    if($VchType -le 0){return @{success=$true;data=@{allowed_vch_types=@($allowedTypes);status=([string]$Status).Trim().ToLowerInvariant();items=@();pagination=(New-BusyCloudApprovalPagination -Page 1 -PageSize $PageSize -Total 0)}}}
    if($allowedTypes -notcontains $VchType){return @{success=$false;httpStatus=403;error="Approval processing is not enabled for voucher type $VchType, or you are not an approver for it."}}
    $normalized=([string]$Status).Trim().ToLowerInvariant();if($normalized -notin @('pending','approved','history')){return @{success=$false;httpStatus=400;error='status must be pending, approved, or history.'}}
    if($Page -lt 1){$Page=1};if($PageSize -notin @(0,50,100,200,300,500)){$PageSize=50}
    if($normalized -eq 'history'){$historyUser=if($IsAdmin){''}else{$UserName};$h=Get-BusyCloudVoucherApprovalHistoryV6 -VchTypes @($VchType) -VchType $VchType -ActionBy $historyUser -Search $Search -FromDate $FromDate -ToDate $ToDate -Page $Page -PageSize $PageSize -InstanceId $InstanceId -CompanyCode $CompanyCode;if(-not $h.success){return$h};return @{success=$true;data=@{allowed_vch_types=@($allowedTypes);status='history';items=@($h.data);pagination=$h.pagination}}}

    $approvalStatus=if($normalized -eq 'approved'){1}else{2};$fi=Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode;if(-not $fi){return @{success=$false;error='BUSY database connection failed'}}
    try{
        function Read-ApprovalFieldV6{param($R,[string]$F,$D);try{$v=$R.Fields.Item($F).Value;if($null -ne $v -and $v -ne [System.DBNull]::Value){return$v}}catch{};return$D}
        $masterMap=@{};$m=$fi.GetRecordset('SELECT Code,Name FROM Master1');if($m){while(-not $m.EOF){try{$c=[int]$m.Fields.Item('Code').Value;if($c -gt 0){$masterMap[$c]=[string]$m.Fields.Item('Name').Value}}catch{};$m.MoveNext()};try{$m.Close()}catch{}}
        $from=Convert-BusyCloudApprovalFilterDate $FromDate;$to=Convert-BusyCloudApprovalFilterDate $ToDate;$dateSql='';if($from){$dateSql += " AND [Date]>=#"+$from.ToString('MM/dd/yyyy')+'#'};if($to){$dateSql += " AND [Date]<#"+$to.AddDays(1).ToString('MM/dd/yyyy')+'#'}
        $rst=$fi.GetRecordset("SELECT VchCode,VchType,VchNo,[Date],VchSeriesCode,MasterCode1,VchAmtBaseCur,VchSalePurcAmt,ApprovalStatus FROM Tran1 WHERE VchType=$VchType AND ApprovalStatus=$approvalStatus$dateSql ORDER BY VchCode DESC")
        $headers=@();$codes=@();if($rst){while(-not $rst.EOF){$code=[int](Read-ApprovalFieldV6 $rst 'VchCode' 0);$seriesCode=[int](Read-ApprovalFieldV6 $rst 'VchSeriesCode' 0);$partyCode=[int](Read-ApprovalFieldV6 $rst 'MasterCode1' 0);$rawDate=Read-ApprovalFieldV6 $rst 'Date' $null;$dateText='';if($rawDate){try{$dateText=([datetime]$rawDate).ToString('yyyy-MM-dd')}catch{$dateText=[string]$rawDate}};$amt=[double](Read-ApprovalFieldV6 $rst 'VchAmtBaseCur' 0);if([math]::Abs($amt) -lt .000001){$amt=[double](Read-ApprovalFieldV6 $rst 'VchSalePurcAmt' 0)};$headers+=@{vchCode=$code;vchNo=([string](Read-ApprovalFieldV6 $rst 'VchNo' '')).Trim();date=$dateText;seriesCode=$seriesCode;partyCode=$partyCode;amount=[math]::Abs($amt)};if($code -gt 0){$codes+=$code};$rst.MoveNext()};try{$rst.Close()}catch{}}
        $createdMap=@{};$approvedMap=@{};for($i=0;$i -lt $codes.Count;$i+=150){$end=[math]::Min($i+149,$codes.Count-1);$batch=@($codes[$i..$end]);if($batch.Count -eq 0){continue};$cs=$batch-join',';$a=$fi.GetRecordset("SELECT Code,Action,ActionTime,UserName,Notes FROM CheckList WHERE Code IN ($cs) AND Action IN (1,3) ORDER BY Code,ActionTime,Action");if($a){while(-not $a.EOF){$c=[int](Read-ApprovalFieldV6 $a 'Code' 0);$act=[int](Read-ApprovalFieldV6 $a 'Action' 0);$u=([string](Read-ApprovalFieldV6 $a 'UserName' '')).Trim();$n=[string](Read-ApprovalFieldV6 $a 'Notes' '');$raw=Read-ApprovalFieldV6 $a 'ActionTime' $null;$tm='';if($raw){try{$tm=([datetime]$raw).ToString('yyyy-MM-dd HH:mm:ss')}catch{}};if($act -eq 1 -and -not $createdMap.ContainsKey($c)){$createdMap[$c]=@{user=$u;time=$tm}}elseif($act -eq 3){$approvedMap[$c]=@{user=$u;time=$tm;notes=$n}};$a.MoveNext()};try{$a.Close()}catch{}}}
        $needle=([string]$Search).Trim().ToLowerInvariant();$all=@();foreach($h in $headers){$c=[int]$h.vchCode;$series=if($masterMap.ContainsKey([int]$h.seriesCode)){([string]$masterMap[([int]$h.seriesCode)]).Trim()}else{'Main'};$prefix='{0:D2}'-f$VchType;if($series.StartsWith($prefix,[System.StringComparison]::OrdinalIgnoreCase)){$series=$series.Substring($prefix.Length)};$party=if($masterMap.ContainsKey([int]$h.partyCode)){([string]$masterMap[([int]$h.partyCode)]).Trim()}else{''};$cr=if($createdMap.ContainsKey($c)){$createdMap[$c]}else{@{user='';time=''}};$ap=if($approvedMap.ContainsKey($c)){$approvedMap[$c]}else{@{user='';time='';notes=''}};if($needle){$hay=(@($h.vchNo,$series,$party,$cr.user,$ap.user,$ap.notes,$c)-join' ').ToLowerInvariant();if(-not $hay.Contains($needle)){continue}};$all+=@{vchCode=$c;vchType=$VchType;vchNo=$h.vchNo;date=$h.date;vchSeriesCode=$h.seriesCode;vchSeries=$series;party=$party;amount=$h.amount;approvalStatus=$approvalStatus;approvalStatusText=if($approvalStatus -eq 1){'Approved'}else{'To be Approved'};createdBy=$cr.user;createdAt=$cr.time;approvedBy=$ap.user;approvedAt=$ap.time;approvalRemarks=$ap.notes}}
        $total=$all.Count;if($PageSize -le 0){$paged=@($all);$Page=1}else{$tp=if($total -le 0){1}else{[int][math]::Ceiling($total/[double]$PageSize)};if($Page -gt $tp){$Page=$tp};$skip=($Page-1)*$PageSize;$paged=@($all|Select-Object -Skip $skip -First $PageSize)}
        return @{success=$true;data=@{allowed_vch_types=@($allowedTypes);status=$normalized;items=@($paged);pagination=(New-BusyCloudApprovalPagination -Page $Page -PageSize $PageSize -Total $total)}}
    }catch{return @{success=$false;error=$_.Exception.Message}}finally{Disconnect-BUSY $fi}
}

function Get-VoucherApprovalQueue {
    param(
        [string]$UserName,[bool]$IsAdmin=$false,[string]$Status='pending',[int]$VchType=0,
        [string]$Search='',[string]$FromDate='',[string]$ToDate='',[int]$Page=1,[int]$PageSize=50,
        [int]$Limit=0,[string]$InstanceId='',[string]$CompanyCode=''
    )
    # Backward compatibility: old callers used Limit only.
    if($PageSize -eq 50 -and $Limit -gt 0){if($Limit -in @(50,100,200,300,500)){$PageSize=$Limit}}
    if(Test-BusyCloudApprovalAccessRuntime -InstanceId $InstanceId -CompanyCode $CompanyCode){
        return Get-VoucherApprovalQueue-AccessComV6 -UserName $UserName -IsAdmin:$IsAdmin -Status $Status -VchType $VchType -Search $Search -FromDate $FromDate -ToDate $ToDate -Page $Page -PageSize $PageSize -InstanceId $InstanceId -CompanyCode $CompanyCode
    }
    return Get-VoucherApprovalQueue-DirectV6 -UserName $UserName -IsAdmin:$IsAdmin -Status $Status -VchType $VchType -Search $Search -FromDate $FromDate -ToDate $ToDate -Page $Page -PageSize $PageSize -InstanceId $InstanceId -CompanyCode $CompanyCode
}

# -----------------------------------------------------------------------------
# HOTFIX V7: SQL approval-list query hardening
# - keeps default server-side paging at 50
# - removes unnecessary joins from COUNT queries
# - pages Tran1 before joining Master1
# - batches CheckList lookups so "All" does not create one huge IN clause
# - uses read-only NOLOCK reads for list/search screens to avoid blocking BUSY
# - preserves Access/BDS COM path from V6
# -----------------------------------------------------------------------------

function Get-VoucherApprovalQueue-DirectV7 {
    param(
        [string]$UserName,
        [bool]$IsAdmin=$false,
        [string]$Status='pending',
        [int]$VchType=0,
        [string]$Search='',
        [string]$FromDate='',
        [string]$ToDate='',
        [int]$Page=1,
        [int]$PageSize=50,
        [string]$InstanceId='',
        [string]$CompanyCode=''
    )

    if ([string]::IsNullOrWhiteSpace($UserName)) {
        return @{ success=$false; httpStatus=401; error='Authenticated BUSY user is required.' }
    }

    $allowed = Get-VoucherApprovalTypesForUser `
        -UserName $UserName `
        -IsAdmin:$IsAdmin `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $allowed.success) { return $allowed }

    $allowedTypes = @(
        $allowed.data |
            ForEach-Object { [int]$_ } |
            Where-Object { $_ -gt 0 } |
            Sort-Object -Unique
    )

    $normalized = ([string]$Status).Trim().ToLowerInvariant()
    if ($normalized -notin @('pending','approved','history')) {
        return @{ success=$false; httpStatus=400; error='status must be pending, approved, or history.' }
    }

    if ($Page -lt 1) { $Page = 1 }
    if ($PageSize -notin @(0,50,100,200,300,500)) { $PageSize = 50 }

    # Metadata-only request used by the frontend to discover enabled/allowed types.
    if ($VchType -le 0) {
        return @{
            success=$true
            data=@{
                allowed_vch_types=@($allowedTypes)
                status=$normalized
                items=@()
                pagination=(New-BusyCloudApprovalPagination -Page 1 -PageSize $PageSize -Total 0)
            }
        }
    }

    if ($allowedTypes -notcontains $VchType) {
        return @{
            success=$false
            httpStatus=403
            error="Approval processing is not enabled for voucher type $VchType, or you are not an approver for it."
        }
    }

    if ($normalized -eq 'history') {
        $historyUser = if ($IsAdmin) { '' } else { $UserName }
        $history = Get-BusyCloudVoucherApprovalHistoryV6 `
            -VchTypes @($VchType) `
            -VchType $VchType `
            -ActionBy $historyUser `
            -Search $Search `
            -FromDate $FromDate `
            -ToDate $ToDate `
            -Page $Page `
            -PageSize $PageSize `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (-not $history.success) { return $history }

        return @{
            success=$true
            data=@{
                allowed_vch_types=@($allowedTypes)
                status='history'
                items=@($history.data)
                pagination=$history.pagination
            }
        }
    }

    $approvalStatus = if ($normalized -eq 'approved') { 1 } else { 2 }
    $ctx = $null

    try {
        $ctx = Get-BusyCloudApprovalDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
        $conn = $ctx.connection

        $whereParts = New-Object System.Collections.Generic.List[string]
        [void]$whereParts.Add('T.VchType = @VchType')
        [void]$whereParts.Add('T.ApprovalStatus = @ApprovalStatus')

        $from = Convert-BusyCloudApprovalFilterDate $FromDate
        $to = Convert-BusyCloudApprovalFilterDate $ToDate
        if ($from) { [void]$whereParts.Add('T.[Date] >= @FromDate') }
        if ($to) { [void]$whereParts.Add('T.[Date] < @ToDateExclusive') }

        $needle = ([string]$Search).Trim()
        if (-not [string]::IsNullOrWhiteSpace($needle)) {
            [void]$whereParts.Add(@"
(
    T.VchNo LIKE @SearchLike
    OR CONVERT(VARCHAR(30), T.VchCode) LIKE @SearchLike
    OR EXISTS (
        SELECT 1 FROM Master1 PM WITH (NOLOCK)
        WHERE PM.Code = T.MasterCode1 AND PM.Name LIKE @SearchLike
    )
    OR EXISTS (
        SELECT 1 FROM Master1 SM WITH (NOLOCK)
        WHERE SM.Code = T.VchSeriesCode AND SM.Name LIKE @SearchLike
    )
    OR EXISTS (
        SELECT 1 FROM CheckList CL WITH (NOLOCK)
        WHERE CL.Code = T.VchCode
          AND CL.Action IN (1,3)
          AND (ISNULL(CL.UserName,'') LIKE @SearchLike OR ISNULL(CL.Notes,'') LIKE @SearchLike)
    )
)
"@)
        }

        $whereSql = $whereParts -join ' AND '
        $timeoutSeconds = if ($PageSize -le 0) { 120 } else { 30 }

        # COUNT only Tran1. Do not join Master1 unless search explicitly needs it.
        $countCmd = $conn.CreateCommand()
        $countCmd.CommandTimeout = $timeoutSeconds
        $countCmd.CommandText = "SELECT COUNT_BIG(1) FROM Tran1 T WITH (NOLOCK) WHERE $whereSql"
        [void]$countCmd.Parameters.AddWithValue('@VchType', $VchType)
        [void]$countCmd.Parameters.AddWithValue('@ApprovalStatus', $approvalStatus)
        if ($from) { [void]$countCmd.Parameters.AddWithValue('@FromDate', $from) }
        if ($to) { [void]$countCmd.Parameters.AddWithValue('@ToDateExclusive', $to.AddDays(1)) }
        if (-not [string]::IsNullOrWhiteSpace($needle)) {
            [void]$countCmd.Parameters.AddWithValue('@SearchLike', ('%' + $needle + '%'))
        }

        $total64 = [long]$countCmd.ExecuteScalar()
        $total = if ($total64 -gt [int]::MaxValue) { [int]::MaxValue } else { [int]$total64 }
        $pagination = New-BusyCloudApprovalPagination -Page $Page -PageSize $PageSize -Total $total
        $Page = [int]$pagination.page

        # First page Tran1, then join Master1 only for those rows.
        $pagingSql = 'OFFSET 0 ROWS'
        if ($PageSize -gt 0) {
            $offset = ($Page - 1) * $PageSize
            $pagingSql = "OFFSET $offset ROWS FETCH NEXT $PageSize ROWS ONLY"
        }

        $dataSql = @"
;WITH PageRows AS (
    SELECT
        T.VchCode,
        T.VchType,
        T.VchNo,
        T.[Date],
        T.VchSeriesCode,
        T.MasterCode1,
        T.VchAmtBaseCur,
        T.VchSalePurcAmt,
        T.ApprovalStatus
    FROM Tran1 T WITH (NOLOCK)
    WHERE $whereSql
    ORDER BY T.VchCode DESC
    $pagingSql
)
SELECT
    R.VchCode,
    R.VchType,
    R.VchNo,
    R.[Date],
    R.VchSeriesCode,
    R.MasterCode1,
    R.VchAmtBaseCur,
    R.VchSalePurcAmt,
    R.ApprovalStatus,
    ISNULL(P.Name,'') AS PartyName,
    ISNULL(S.Name,'') AS SeriesName
FROM PageRows R
LEFT JOIN Master1 P WITH (NOLOCK) ON P.Code = R.MasterCode1
LEFT JOIN Master1 S WITH (NOLOCK) ON S.Code = R.VchSeriesCode
ORDER BY R.VchCode DESC
"@

        $cmd = $conn.CreateCommand()
        $cmd.CommandTimeout = $timeoutSeconds
        $cmd.CommandText = $dataSql
        [void]$cmd.Parameters.AddWithValue('@VchType', $VchType)
        [void]$cmd.Parameters.AddWithValue('@ApprovalStatus', $approvalStatus)
        if ($from) { [void]$cmd.Parameters.AddWithValue('@FromDate', $from) }
        if ($to) { [void]$cmd.Parameters.AddWithValue('@ToDateExclusive', $to.AddDays(1)) }
        if (-not [string]::IsNullOrWhiteSpace($needle)) {
            [void]$cmd.Parameters.AddWithValue('@SearchLike', ('%' + $needle + '%'))
        }

        $rows = @()
        $codes = @()
        $rdr = $cmd.ExecuteReader()
        try {
            while ($rdr.Read()) {
                $code = [int](Read-BusyCloudApprovalReaderValue $rdr 'VchCode' 0)
                $rawDate = Read-BusyCloudApprovalReaderValue $rdr 'Date' $null
                $dateText = ''
                if ($rawDate) {
                    try { $dateText = ([datetime]$rawDate).ToString('yyyy-MM-dd') }
                    catch { $dateText = [string]$rawDate }
                }

                $amount = [double](Read-BusyCloudApprovalReaderValue $rdr 'VchAmtBaseCur' 0)
                if ([math]::Abs($amount) -lt 0.000001) {
                    $amount = [double](Read-BusyCloudApprovalReaderValue $rdr 'VchSalePurcAmt' 0)
                }

                $seriesName = ([string](Read-BusyCloudApprovalReaderValue $rdr 'SeriesName' 'Main')).Trim()
                if (-not $seriesName) { $seriesName = 'Main' }
                $prefix = '{0:D2}' -f $VchType
                if ($seriesName.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $seriesName = $seriesName.Substring($prefix.Length)
                }

                $rows += @{
                    vchCode=$code
                    vchType=$VchType
                    vchNo=([string](Read-BusyCloudApprovalReaderValue $rdr 'VchNo' '')).Trim()
                    date=$dateText
                    vchSeriesCode=[int](Read-BusyCloudApprovalReaderValue $rdr 'VchSeriesCode' 0)
                    vchSeries=$seriesName
                    party=([string](Read-BusyCloudApprovalReaderValue $rdr 'PartyName' '')).Trim()
                    amount=[math]::Abs($amount)
                }
                if ($code -gt 0) { $codes += $code }
            }
        }
        finally {
            try { $rdr.Close() } catch {}
        }

        # Read checklist metadata in bounded batches. This is especially important
        # when the user explicitly selects Rows=All.
        $createdMap = @{}
        $approvedMap = @{}
        $batchSize = 250

        for ($start = 0; $start -lt $codes.Count; $start += $batchSize) {
            $end = [math]::Min($start + $batchSize - 1, $codes.Count - 1)
            $batch = @($codes[$start..$end])
            if ($batch.Count -eq 0) { continue }

            $codesSql = $batch -join ','
            $auditCmd = $conn.CreateCommand()
            $auditCmd.CommandTimeout = $timeoutSeconds
            $auditCmd.CommandText = "SELECT Code,Action,ActionTime,UserName,Notes FROM CheckList WITH (NOLOCK) WHERE Code IN ($codesSql) AND Action IN (1,3) ORDER BY Code,ActionTime,Action"

            $auditReader = $auditCmd.ExecuteReader()
            try {
                while ($auditReader.Read()) {
                    $code = [int](Read-BusyCloudApprovalReaderValue $auditReader 'Code' 0)
                    $action = [int](Read-BusyCloudApprovalReaderValue $auditReader 'Action' 0)
                    $user = ([string](Read-BusyCloudApprovalReaderValue $auditReader 'UserName' '')).Trim()
                    $notes = [string](Read-BusyCloudApprovalReaderValue $auditReader 'Notes' '')
                    $rawTime = Read-BusyCloudApprovalReaderValue $auditReader 'ActionTime' $null
                    $time = ''
                    if ($rawTime) {
                        try { $time = ([datetime]$rawTime).ToString('yyyy-MM-dd HH:mm:ss') } catch {}
                    }

                    if ($action -eq 1 -and -not $createdMap.ContainsKey($code)) {
                        $createdMap[$code] = @{ user=$user; time=$time }
                    }
                    elseif ($action -eq 3) {
                        $approvedMap[$code] = @{ user=$user; time=$time; notes=$notes }
                    }
                }
            }
            finally {
                try { $auditReader.Close() } catch {}
            }
        }

        $items = @()
        foreach ($row in $rows) {
            $code = [int]$row.vchCode
            $created = if ($createdMap.ContainsKey($code)) { $createdMap[$code] } else { @{ user=''; time='' } }
            $approved = if ($approvedMap.ContainsKey($code)) { $approvedMap[$code] } else { @{ user=''; time=''; notes='' } }

            $items += @{
                vchCode=$code
                vchType=$row.vchType
                vchNo=$row.vchNo
                date=$row.date
                vchSeriesCode=$row.vchSeriesCode
                vchSeries=$row.vchSeries
                party=$row.party
                amount=$row.amount
                approvalStatus=$approvalStatus
                approvalStatusText=if ($approvalStatus -eq 1) { 'Approved' } else { 'To be Approved' }
                createdBy=$created.user
                createdAt=$created.time
                approvedBy=$approved.user
                approvedAt=$approved.time
                approvalRemarks=$approved.notes
            }
        }

        return @{
            success=$true
            data=@{
                allowed_vch_types=@($allowedTypes)
                status=$normalized
                items=@($items)
                pagination=$pagination
            }
        }
    }
    catch {
        return @{ success=$false; error=$_.Exception.Message }
    }
    finally {
        if ($ctx -and $ctx.connection) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}

# Final runtime dispatcher override. Access/BDS keeps the V6 COM implementation;
# SQL Server uses the optimized V7 direct implementation above.

function Get-VoucherApprovalQueue {
    param(
        [string]$UserName,
        [bool]$IsAdmin=$false,
        [string]$Status='pending',
        [int]$VchType=0,
        [string]$Search='',
        [string]$FromDate='',
        [string]$ToDate='',
        [int]$Page=1,
        [int]$PageSize=50,
        [int]$Limit=0,
        [string]$InstanceId='',
        [string]$CompanyCode=''
    )

    if ($PageSize -eq 50 -and $Limit -gt 0 -and $Limit -in @(50,100,200,300,500)) {
        $PageSize = $Limit
    }

    if (Test-BusyCloudApprovalAccessRuntime -InstanceId $InstanceId -CompanyCode $CompanyCode) {
        return Get-VoucherApprovalQueue-AccessComV6 `
            -UserName $UserName `
            -IsAdmin:$IsAdmin `
            -Status $Status `
            -VchType $VchType `
            -Search $Search `
            -FromDate $FromDate `
            -ToDate $ToDate `
            -Page $Page `
            -PageSize $PageSize `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }

    return Get-VoucherApprovalQueue-DirectV7 `
        -UserName $UserName `
        -IsAdmin:$IsAdmin `
        -Status $Status `
        -VchType $VchType `
        -Search $Search `
        -FromDate $FromDate `
        -ToDate $ToDate `
        -Page $Page `
        -PageSize $PageSize `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode
}

