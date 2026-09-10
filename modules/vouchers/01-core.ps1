# 01-core.ps1
# Extracted from vouchers.ps1. Keep functions behavior-compatible with the monolith.

function Safe-ParseDate {
    param([string]$dateStr)
    if ([string]::IsNullOrWhiteSpace($dateStr)) { return $null }
    $formats = @("dd-MM-yyyy", "yyyy-MM-dd", "MM/dd/yyyy", "d-M-yyyy", "yyyy/MM/dd")
    foreach ($fmt in $formats) {
        try {
            return [datetime]::ParseExact($dateStr, $fmt, [System.Globalization.CultureInfo]::InvariantCulture)
        } catch {}
    }
    try {
        return [datetime]::Parse($dateStr)
    } catch {}
    return $null
}

function Get-VchCode-Direct {
    param($conn, [int]$VchType, [string]$VchNo, [string]$VchDate, [bool]$isSql)
    
    $dateFilter = ""
    if ($VchDate -match "\d{2}-\d{2}-\d{4}") {
        $d = [datetime]::ParseExact($VchDate, "dd-MM-yyyy", $null)
        $dateFilter = if ($isSql) { " AND [Date]='" + $d.ToString("yyyy-MM-dd") + "'" } else { " AND [Date]=#" + $d.ToString("MM/dd/yyyy") + "#" }
    } elseif ($VchDate -match "\d{4}-\d{2}-\d{2}") {
        $d = [datetime]::ParseExact($VchDate, "yyyy-MM-dd", $null)
        $dateFilter = if ($isSql) { " AND [Date]='" + $d.ToString("yyyy-MM-dd") + "'" } else { " AND [Date]=#" + $d.ToString("MM/dd/yyyy") + "#" }
    }

    $targetNo = $VchNo.Trim().ToLower()
    $vchCode = 0

    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT [VchCode], [VchNo], [AutoVchNo] FROM [Tran1] WHERE [VchType] = $VchType" + $dateFilter
        $rdr = $cmd.ExecuteReader()
        while ($rdr.Read()) {
            $dbVchNo = ""
            $v = $rdr["VchNo"]
            if ($null -ne $v -and "$v" -ne "") { $dbVchNo = $v.ToString().Trim().ToLower() }
            if ($dbVchNo -eq "") {
                $av = $rdr["AutoVchNo"]
                if ($null -ne $av -and "$v" -ne "") { $dbVchNo = $av.ToString().Trim().ToLower() }
            }
            if ($dbVchNo -eq $targetNo) {
                $vchCode = [int]$rdr["VchCode"]
                break
            }
        }
        $rdr.Close()
    } catch {}

    if ($vchCode -eq 0) {
        try {
            $cmd = $conn.CreateCommand()
            $cmd.CommandText = "SELECT [VchCode], [VchNo], [AutoVchNo] FROM [Tran1] WHERE [VchType] = $VchType"
            $rdr = $cmd.ExecuteReader()
            while ($rdr.Read()) {
                $dbVchNo = ""
                $v = $rdr["VchNo"]
                if ($null -ne $v -and "$v" -ne "") { $dbVchNo = $v.ToString().Trim().ToLower() }
                if ($dbVchNo -eq "") {
                    $av = $rdr["AutoVchNo"]
                    if ($null -ne $av -and "$av" -ne "") { $dbVchNo = $av.ToString().Trim().ToLower() }
                }
                if ($dbVchNo -eq $targetNo) {
                    $vchCode = [int]$rdr["VchCode"]
                    break
                }
            }
            $rdr.Close()
        } catch {}
    }
    return $vchCode
}


# -----------------------------------------------------------------------------
# Fast SQL fiscal-database resolver used by voucher list/detail reads.
#
# The legacy Get-DirectConnection resolves BUSY's active fiscal database by
# opening the BUSY COM/OpenCSDB layer first. On some SQL installations that
# initialization takes 20-60+ seconds. vch_setting.ps1 already contains a
# direct SQL fiscal database resolver for the approval feature, so voucher
# read-only SQL paths reuse it when available.
#
# Access/BDS behaviour is unchanged. Write/save operations are unchanged.
# -----------------------------------------------------------------------------
function Get-FastVoucherDirectConnection {
    param(
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    try {
        $inst = Get-InstanceConfig -InstanceId $InstanceId
        $dbType = 0
        if ($null -ne $inst -and $null -ne $inst.dbType) {
            $dbType = [int]$inst.dbType
        }

        if ($dbType -eq 1) {
            $resolver = Get-Command Get-BusyCloudApprovalDbContext -ErrorAction SilentlyContinue
            if ($null -ne $resolver) {
                $ctx = Get-BusyCloudApprovalDbContext -InstanceId $InstanceId -CompanyCode $CompanyCode
                if ($null -ne $ctx -and $null -ne $ctx.connection) {
                    $conn = $ctx.connection
                    # Existing voucher functions expect a closed connection and
                    # call .Open() themselves. Normalize to that contract.
                    try {
                        if ($conn.State -eq [System.Data.ConnectionState]::Open) {
                            $conn.Close()
                        }
                    } catch {}

                    Write-Host "  [FAST-VOUCHER-DB] $InstanceId/$CompanyCode -> $($ctx.database)" -ForegroundColor DarkCyan
                    return @{
                        connection = $conn
                        database   = $ctx.database
                        dbType     = 1
                        fast       = $true
                    }
                }
            }
        }
    }
    catch {
        Write-Host "  [WARN] Fast voucher DB resolver failed; using legacy resolver: $($_.Exception.Message)" -ForegroundColor DarkYellow
    }

    # Safety fallback keeps old behaviour if the direct fiscal resolver is not
    # available for any reason.
    return Get-DirectConnection -InstanceId $InstanceId -CompanyCode $CompanyCode
}

# ===============================================================
#  CORE UTILITY HELPERS
# ===============================================================

function Validate-VoucherData {
    param($Data, $Cfg)

    foreach ($key in $Cfg.requiredKeys) {
        $val = $Data.$key
        if ($key -eq "items") {
            if (-not $val -or $val.Count -eq 0) {
                return "At least one item is required"
            }
        } elseif ($key -eq "accounts") {
            if (-not $val -or $val.Count -eq 0) {
                return "At least one account entry is required"
            }
        } else {
            if (-not $val -or [string]$val -eq "") {
                return "$key is required"
            }
        }
    }
    return $null
}

function Pad-VchNo {
    param([string]$VchNo)
    if ($null -eq $VchNo) { return "" }
    if ($VchNo.Length -ge 25) { return $VchNo }
    return $VchNo.PadLeft(25)
}

function Clear-StockCaches {
    param([string]$InstanceId = "", [string]$CompanyCode = "")
    
    $prefix = "$InstanceId|$CompanyCode|"
    $keysToRemove = @($script:_cache.Keys | Where-Object {
        $_ -like "${prefix}items|*" -or
        $_ -like "${prefix}item|*" -or
        $_ -like "${prefix}parties|*" -or
        $_ -like "${prefix}vouchers|*" -or
        $_ -like "${prefix}voucherseries|*"
    })
    
    foreach ($k in $keysToRemove) {
        $script:_cache.Remove($k)
    }
}

function Update-CheckListCreator {
    param(
        $fi,
        [int]$VchType,
        [string]$VchNo,
        [string]$VchDate,
        [string]$VchSeries = "",
        [string]$UserName,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    # Historical name retained for compatibility. On MODIFY this function must
    # update only the newest CheckList Action=2 row. Updating every row would
    # overwrite Created By (Action=1) and Approved By (Action=3).
    if ([string]::IsNullOrWhiteSpace($UserName) -or [string]::IsNullOrWhiteSpace($VchNo)) { return }

    try {
        $state = Get-VoucherApprovalStateByIdentity `
            -VchType $VchType `
            -VchNo $VchNo `
            -VchDate $VchDate `
            -VchSeries $VchSeries `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode `
            -ExistingFi $fi

        if (-not $state.success -or -not $state.found -or [int]$state.vchCode -le 0) { return }

        $vchCode = [int]$state.vchCode
        $safeUser = $UserName.Trim().Replace("'", "''")
        $fi.ExecuteQuery(@"
UPDATE CheckList
SET UserName='$safeUser'
WHERE Code=$vchCode
  AND Action=2
  AND ActionTime=(SELECT MAX(ActionTime) FROM CheckList WHERE Code=$vchCode AND Action=2)
"@)
    } catch {
        Write-Host "  [WARN] Could not update latest modifier audit row: $($_.Exception.Message)" -ForegroundColor DarkYellow
    }
}
