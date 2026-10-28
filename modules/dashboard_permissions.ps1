# BUSYWEB_DASHBOARD_PERMISSIONS_V1
# Independent, company-scoped BusyWeb-only grants. Windows PowerShell 5.1.
# This module is loaded by both 8081 and 8085 from ONE backend directory.
$script:BusyDashboardPermissionKeys = @(
 'netSales','netPurchases','receipts','payments','performance','distribution',
 'inventoryLandscape','workingCapital','customerIntelligence','productIntelligence'
)
$script:BusyDashboardPermissionRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'dashboard-permissions'

function Get-BusyDashboardValue {
    param($Data, [string]$Name, $Default = $null)
    if ($null -eq $Data) { return $Default }
    if ($Data -is [System.Collections.IDictionary]) {
        if ($Data.Contains($Name)) { return $Data[$Name] }
        return $Default
    }
    $prop = $Data.PSObject.Properties[$Name]
    if ($null -ne $prop) { return $prop.Value }
    return $Default
}

function Test-BusyDashboardAdmin {
    param($User)
    $role = [string](Get-BusyDashboardValue $User 'role' '')
    return $role -in @('superadmin','companyadmin')
}

function New-BusyDashboardGrants {
    param([bool]$All = $false)
    $grants = @{}
    foreach ($key in $script:BusyDashboardPermissionKeys) { $grants[$key] = $All }
    return $grants
}

function Get-BusyDashboardGrantPath {
    param([string]$InstanceId, [string]$CompanyCode, [string]$UserName)
    if ($InstanceId -notmatch '^[a-zA-Z0-9_-]{1,70}$' -or
        $CompanyCode -notmatch '^[a-zA-Z0-9_-]{1,70}$' -or
        [string]::IsNullOrWhiteSpace($UserName) -or $UserName.Length -gt 120) {
        throw 'Invalid dashboard permission scope.'
    }
    $identity = ($InstanceId.Trim().ToLowerInvariant() + [char]0 +
        $CompanyCode.Trim().ToLowerInvariant() + [char]0 + $UserName.Trim().ToLowerInvariant())
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($identity))).Replace('-','').ToLowerInvariant()
    } finally { $sha.Dispose() }
    return (Join-Path $script:BusyDashboardPermissionRoot ($hash + '.json'))
}

function Read-BusyDashboardGrantRecord {
    param([string]$InstanceId, [string]$CompanyCode, [string]$UserName)
    $file = Get-BusyDashboardGrantPath -InstanceId $InstanceId -CompanyCode $CompanyCode -UserName $UserName
    if (-not [IO.File]::Exists($file)) { return $null }
    $record = [IO.File]::ReadAllText($file, [Text.Encoding]::UTF8) | ConvertFrom-Json
    if ($null -eq $record -or [string]$record.instanceId -cne $InstanceId -or
        [string]$record.companyCode -cne $CompanyCode -or
        -not ([string]$record.userName).Equals($UserName, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Dashboard grants failed integrity/scope validation.'
    }
    return $record
}

function Get-BusyDashboardGrants {
    param([string]$InstanceId, [string]$CompanyCode, $User)
    if ($null -eq $User) { throw 'Authenticated user required.' }
    $name = [string](Get-BusyDashboardValue $User 'name' '')
    if (Test-BusyDashboardAdmin $User) { return (New-BusyDashboardGrants -All $true) }
    $record = Read-BusyDashboardGrantRecord -InstanceId $InstanceId -CompanyCode $CompanyCode -UserName $name
    $grants = New-BusyDashboardGrants
    if ($null -eq $record) { return $grants } # Explicit deny until admin grants.
    foreach ($key in $script:BusyDashboardPermissionKeys) {
        $raw = Get-BusyDashboardValue $record.grants $key $false
        $grants[$key] = ($raw -is [bool] -and $raw -eq $true)
    }
    return $grants
}

function Set-BusyDashboardGrants {
    param([string]$InstanceId, [string]$CompanyCode, [string]$UserName,
          $Submitted, [string]$ExpectedVersion, [string]$UpdatedBy)
    $path = Get-BusyDashboardGrantPath -InstanceId $InstanceId -CompanyCode $CompanyCode -UserName $UserName
    $previous = Read-BusyDashboardGrantRecord -InstanceId $InstanceId -CompanyCode $CompanyCode -UserName $UserName
    $previousVersion = if ($null -eq $previous) { '' } else { [string]$previous.version }
    if ($ExpectedVersion -cne $previousVersion) {
        return @{success=$false;httpStatus=409;error='Permissions were changed by another administrator. Reload before saving.'}
    }
    $grants = New-BusyDashboardGrants
    foreach ($key in $script:BusyDashboardPermissionKeys) {
        $raw = Get-BusyDashboardValue $Submitted $key $null
        if ($raw -isnot [bool]) {
            return @{success=$false;httpStatus=400;error="Permission '$key' must be a true/false value."}
        }
        $grants[$key] = [bool]$raw
    }
    if (-not [IO.Directory]::Exists($script:BusyDashboardPermissionRoot)) {
        [void][IO.Directory]::CreateDirectory($script:BusyDashboardPermissionRoot)
    }
    $record = @{
        schema=1;instanceId=$InstanceId;companyCode=$CompanyCode;userName=$UserName
        grants=$grants;version=[Guid]::NewGuid().ToString('N')
        updatedAt=[DateTime]::UtcNow.ToString('o');updatedBy=$UpdatedBy
    }
    $temporary = $path + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
    try {
        $json = ConvertTo-Json $record -Depth 8
        [IO.File]::WriteAllText($temporary, $json, ([Text.UTF8Encoding]::new($false)))
        if ([IO.File]::Exists($path)) { [IO.File]::Replace($temporary, $path, $null) }
        else { [IO.File]::Move($temporary, $path) }
    } finally { if ([IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) } }
    Write-Host "[DASH-PERMS] Grants saved for selected company/user." -ForegroundColor Green
    return @{success=$true;data=@{name=$UserName;grants=$grants;version=$record.version;updatedAt=$record.updatedAt}}
}

function Test-BusyDashboardAny {
    param($Grants, [string[]]$Keys)
    foreach ($key in $Keys) { if ($Grants[$key] -eq $true) { return $true } }
    return $false
}

# Create fresh responses; NEVER return an unfiltered copy of cached company-wide data.
function Protect-BusyDashboardResult {
    param([string]$Section, $Payload, $Grants)
    if ($null -eq $Payload) { return $Payload }
    $safe = @{}
    foreach ($key in @('success','section','asOf','financialYearStart','generatedAt','currency',
                      'instanceId','companyCode','snapshotStale','refreshState','warnings')) {
        $v = Get-BusyDashboardValue $Payload $key $null
        if ($null -ne $v) { $safe[$key] = $v }
    }
    $safe['success'] = $true
    if ($Section -eq 'core') {
        $periods = @{}
        foreach ($period in @('today','mtd','fytd')) {
            $values = @{}
            $p = Get-BusyDashboardValue (Get-BusyDashboardValue $Payload 'periods') $period
            $sales = $Grants['netSales'] -or $Grants['distribution']
            $purchases = $Grants['netPurchases'] -or $Grants['distribution']
            $receipts = $Grants['receipts'] -or $Grants['distribution']
            $payments = $Grants['payments'] -or $Grants['distribution']
            $fields = @()
            if ($sales) { $fields += @('netSales') }
            if ($purchases) { $fields += @('netPurchases') }
            if ($receipts) { $fields += @('receiptVouchers') }
            if ($payments) { $fields += @('paymentVouchers') }
            foreach ($field in $fields) {
                $v = Get-BusyDashboardValue $p $field $null
                if ($null -ne $v) { $values[$field] = $v }
            }
            $periods[$period] = $values
        }
        $safe['periods'] = $periods
        $rawTrend = @(Get-BusyDashboardValue $Payload 'trend' @())
        if ($Grants['performance']) {
            $safe['trend'] = $rawTrend
        } else {
            # KPI micro-charts can see their own granted time-series only.
            $trend = @()
            foreach ($day in $rawTrend) {
                $item = @{date=(Get-BusyDashboardValue $day 'date' '')}
                foreach ($tuple in @(@('netSales','sales'),@('netPurchases','purchases'),@('receipts','receipts'),@('payments','payments'))) {
                    if ($Grants[$tuple[0]]) {
                        $item[$tuple[1]] = Get-BusyDashboardValue $day $tuple[1] 0
                    }
                }
                $trend += $item
            }
            $safe['trend'] = $trend
        }
        $safe['recent'] = @() # No current client panel authorizes a raw recent-voucher list.
        $safe['databaseMode'] = Get-BusyDashboardValue $Payload 'databaseMode' ''
        $safe['source'] = Get-BusyDashboardValue $Payload 'source' ''
        $safe['definitions'] = @{}
    } elseif ($Section -eq 'exposure') {
        $stock = Get-BusyDashboardValue $Payload 'stock'
        $outstanding = Get-BusyDashboardValue $Payload 'outstanding'
        $safe['stock'] = if ($Grants['inventoryLandscape'] -and $null -ne $stock) {
            @{stockValue=(Get-BusyDashboardValue $stock 'stockValue' 0);
              topValuedItems=@(Get-BusyDashboardValue $stock 'topValuedItems' @())}
        } else { $null }
        $safe['outstanding'] = if ($Grants['workingCapital'] -and $null -ne $outstanding) {
            @{receivable=(Get-BusyDashboardValue $outstanding 'receivable' 0);
              payable=(Get-BusyDashboardValue $outstanding 'payable' 0);
              netPosition=(Get-BusyDashboardValue $outstanding 'netPosition' 0)}
        } else { $null }
    } elseif ($Section -eq 'insights') {
        $safe['topCustomers'] = if ($Grants['customerIntelligence']) { @(Get-BusyDashboardValue $Payload 'topCustomers' @()) } else { @() }
        $safe['bestSellingItems'] = if ($Grants['productIntelligence']) { @(Get-BusyDashboardValue $Payload 'bestSellingItems' @()) } else { @() }
        if ($Grants['customerIntelligence']) {
            $safe['outstandingAvailable'] = Get-BusyDashboardValue $Payload 'outstandingAvailable' $false
            $safe['outstandingWarning'] = Get-BusyDashboardValue $Payload 'outstandingWarning'
        }
    }
    return $safe
}
