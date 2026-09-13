# modules/accounts.ps1
# Account & Master Data Management (PowerShell 5.1 Safe) - Multi-Instance Version

. "$PSScriptRoot\connection.ps1"
. "$PSScriptRoot\utils.ps1"

# ═══════════════════════════════════════════════════════
#  INTERNAL HELPER - Escape XML special characters
# ═══════════════════════════════════════════════════════
Write-Host "  [PARTIES-FISCAL-DB-V2] Party/account reads use the active fiscal database." -ForegroundColor DarkCyan

Write-Host "  [PARTIES-QUERY-V3] Page-first party/account query loaded." -ForegroundColor DarkCyan

Write-Host "  [PARTY-GROUPS-DIRECT-V1] Party Account Group permission tree is COM-free." -ForegroundColor DarkCyan

function ConvertTo-XmlSafe {
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return "" }
    $Value = $Value -replace '&',  '&amp;'
    $Value = $Value -replace '<',  '&lt;'
    $Value = $Value -replace '>',  '&gt;'
    $Value = $Value -replace '"',  '&quot;'
    $Value = $Value -replace "'",  '&apos;'
    return $Value
}

# ═══════════════════════════════════════════════════════
#  INTERNAL HELPER - Shared Field Mapping
# ═══════════════════════════════════════════════════════
function Get-AccountCommonFieldsXml {
    param($Data)
    $alias     = ConvertTo-XmlSafe ([string]$Data.alias)
    $printName = ConvertTo-XmlSafe ([string]$Data.printName)
    $itPan     = ConvertTo-XmlSafe ([string]$Data.itPan)
    $vat       = ConvertTo-XmlSafe ([string]$Data.vat)
    $ward      = ConvertTo-XmlSafe ([string]$Data.ward)

    $opBalRaw  = 0.0
    if ($Data.opBal) { $opBalRaw = [double]$Data.opBal }
    $opBalType = if ($Data.opBalType) { ([string]$Data.opBalType).ToUpper().Trim() } else { "D" }
    $opBal     = if ($opBalType -eq "D") { -[Math]::Abs($opBalRaw) } else { [Math]::Abs($opBalRaw) }

    $pyBalRaw  = 0.0
    if ($Data.prevYearBal) { $pyBalRaw = [double]$Data.prevYearBal }
    $pyBalType = if ($Data.prevYearBalType) { ([string]$Data.prevYearBalType).ToUpper().Trim() } else { "D" }
    $pyBal     = if ($pyBalType -eq "D") { -[Math]::Abs($pyBalRaw) } else { [Math]::Abs($pyBalRaw) }

    $xml  = ""
    if ($alias)     { $xml += "<Alias>$alias</Alias>" }
    if ($printName) { $xml += "<PrintName>$printName</PrintName><ChequePrintName>$printName</ChequePrintName>" }
    $xml += "<OPBal>$opBal</OPBal>"
    $xml += "<PYBal>$pyBal</PYBal>"
    $xml += "<Address>"
    if ($Data.address) {
        $lines = ([string]$Data.address) -split "`n"
        if ($lines.Count -gt 0) { $xml += "<Address1>$(ConvertTo-XmlSafe $lines[0].Trim())</Address1>" }
        if ($lines.Count -gt 1) { $xml += "<Address2>$(ConvertTo-XmlSafe $lines[1].Trim())</Address2>" }
        if ($lines.Count -gt 2) { $xml += "<Address3>$(ConvertTo-XmlSafe $lines[2].Trim())</Address3>" }
        if ($lines.Count -gt 3) { $xml += "<Address4>$(ConvertTo-XmlSafe $lines[3].Trim())</Address4>" }
    }
    $xml += "<TelNo>$(ConvertTo-XmlSafe $Data.telNo)</TelNo>"
    $xml += "<Fax>$(ConvertTo-XmlSafe $Data.fax)</Fax>"
    $xml += "<Email>$(ConvertTo-XmlSafe $Data.email)</Email>"
    $xml += "<Mobile>$(ConvertTo-XmlSafe $Data.mobileNo)</Mobile>"
    $xml += "<WhatsAppNo>$(ConvertTo-XmlSafe $Data.whatsappNo)</WhatsAppNo>"
    $xml += "<ITPAN>$itPan</ITPAN>"
    $xml += "<ITWard>$ward</ITWard>"
    $xml += "<Contact>$(ConvertTo-XmlSafe $Data.contactPerson)</Contact>"
    $xml += "<TINNo>$vat</TINNo>"
    $svat = ConvertTo-XmlSafe ([string]$Data.svat)
    $xml += "<CST>$svat</CST>"
    $xml += "<CountryName>Sri Lanka</CountryName>"
    $xml += "<AreaName>---Others---</AreaName>"
    $xml += "<Transport>$(ConvertTo-XmlSafe $Data.transport)</Transport>"
    $xml += "<Station>$(ConvertTo-XmlSafe $Data.station)</Station>"
    $xml += "</Address>"
    $taxType = if ($Data.taxType) { [string]$Data.taxType } else { "Others" }
    $xml += "<SupplierType>1</SupplierType><PriceLevel>@</PriceLevel><TaxType>$taxType</TaxType>"
    return $xml
}

# ═══════════════════════════════════════════════════════
#  UPDATE ACCOUNT
# ═══════════════════════════════════════════════════════
function Update-Account {
    param(
        $Data,
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
    try {
        $originalName = if ($Data._originalName) { [string]$Data._originalName } else { [string]$Data.name }
        $rawBBB       = [string]$Data.maintainBillByBill
        $billByBill   = if ($rawBBB -eq "True" -or $rawBBB -eq "true" -or $rawBBB -eq "1") { "True" } else { "False" }

        $xml  = "<Account>"
        $xml += "<Name>$(ConvertTo-XmlSafe $originalName)</Name>"
        $xml += "<ParentGroup>$(ConvertTo-XmlSafe ([string]$Data.group))</ParentGroup>"
        if ($billByBill -eq "True") { $xml += "<BillByBillBalancing>True</BillByBillBalancing>" }
        $xml += Get-AccountCommonFieldsXml -Data $Data
        $xml += "</Account>"

        $err   = ""
        $saved = $fi.SaveMasterFromXML(2, $xml, [ref]$err, $true)

        if ($saved -eq $true) {
            Clear-AccountCaches -InstanceId $InstanceId -CompanyCode $CompanyCode
            return @{ success = $true; message = "Account updated successfully" }
        } else {
            $errOut = if ($err) { $err } else { "Account not found or locked" }
            return @{ success = $false; error = $errOut }
        }
    } finally { Disconnect-BUSY $fi }
}

# ═══════════════════════════════════════════════════════
#  CREATE ACCOUNT
# ═══════════════════════════════════════════════════════
function Create-Account {
    param(
        $Data,
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
    try {
        $rawBBB     = [string]$Data.maintainBillByBill
        $billByBill = if ($rawBBB -eq "True" -or $rawBBB -eq "true" -or $rawBBB -eq "1") { "True" } else { "False" }

        $xml  = "<Account>"
        $xml += "<Name>$(ConvertTo-XmlSafe $Data.name)</Name>"
        $xml += "<ParentGroup>$(ConvertTo-XmlSafe $Data.group)</ParentGroup>"
        if ($billByBill -eq "True") { $xml += "<BillByBillBalancing>True</BillByBillBalancing>" }
        $xml += Get-AccountCommonFieldsXml -Data $Data
        $xml += "</Account>"

        $err   = ""
        $saved = $fi.SaveMasterFromXML(2, $xml, [ref]$err, $false)
        if ($saved -eq $true) { Clear-AccountCaches -InstanceId $InstanceId -CompanyCode $CompanyCode; return @{ success = $true; message = "Account created" } }
        return @{ success = $false; error = $err }
    } finally { Disconnect-BUSY $fi }
}

# ═══════════════════════════════════════════════════════
#  GET ACCOUNT DETAIL
# ═══════════════════════════════════════════════════════
function Get-AccountDetail {
    param(
        [string]$Name,
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
    try {
        $mastCode = $fi.MasterName2Code($Name, 2)
        if ($mastCode -le 0) { return @{ success = $false; error = "Account not found" } }
        $xmlStr = $fi.GetMasterXML($mastCode)
        $xml    = [xml]$xmlStr

        $opRaw     = if ($xml.Account.OPBal) { [double]$xml.Account.OPBal } else { 0.0 }
        $opBal     = [Math]::Abs($opRaw)
        $opBalType = if ($opRaw -lt 0) { "D" } else { "C" }
        $pyRaw     = if ($xml.Account.PYBal) { [double]$xml.Account.PYBal } else { 0.0 }
        $pyBal     = [Math]::Abs($pyRaw)
        $pyBalType = if ($pyRaw -lt 0) { "D" } else { "C" }

        $lines = @()
        if ($xml.Account.Address.Address1) { $lines += [string]$xml.Account.Address.Address1 }
        if ($xml.Account.Address.Address2) { $lines += [string]$xml.Account.Address.Address2 }
        if ($xml.Account.Address.Address3) { $lines += [string]$xml.Account.Address.Address3 }
        if ($xml.Account.Address.Address4) { $lines += [string]$xml.Account.Address.Address4 }
        $address = $lines -join "`n"

        return @{ success = $true; data = @{
            name               = [string]$xml.Account.Name
            alias              = [string]$xml.Account.Alias
            printName          = [string]$xml.Account.PrintName
            group              = [string]$xml.Account.ParentGroup
            opBal              = $opBal
            opBalType          = $opBalType
            prevYearBal        = $pyBal
            prevYearBalType    = $pyBalType
            address            = $address
            mobileNo           = [string]$xml.Account.Address.Mobile
            whatsappNo         = [string]$xml.Account.Address.WhatsAppNo
            telNo              = [string]$xml.Account.Address.TelNo
            fax                = [string]$xml.Account.Address.Fax
            email              = [string]$xml.Account.Address.Email
            contactPerson      = [string]$xml.Account.Address.Contact
            transport          = [string]$xml.Account.Address.Transport
            station            = [string]$xml.Account.Address.Station
            itPan              = [string]$xml.Account.Address.ITPAN
            ward               = [string]$xml.Account.Address.ITWard
            vat                = [string]$xml.Account.Address.TINNo
            svat               = [string]$xml.Account.Address.CST
            taxType            = [string]$xml.Account.TaxType
            maintainBillByBill = ($null -ne $xml.Account.BillByBillBalancing -and $xml.Account.BillByBillBalancing -ne "")
        }}
    } finally { Disconnect-BUSY $fi }
}

# ═══════════════════════════════════════════════════════
#  CACHE HELPERS
# ═══════════════════════════════════════════════════════
function Clear-AccountCaches {
    param(
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )
    $prefix = "$InstanceId|$CompanyCode|"
    Clear-Cache "${prefix}acct-groups"
    Clear-Cache "${prefix}acct-all"
    Clear-Cache "${prefix}parties"
    if ($null -ne $script:_cache) {
        $keysToRemove = @()
        foreach ($k in $script:_cache.Keys) {
            if ($k -like "${prefix}acct-grp|*") { $keysToRemove += $k }
        }
        foreach ($k in $keysToRemove) { $script:_cache.Remove($k) }
    }

    if ($null -ne $script:BusyCloudAccountAccessTreeCache) {
        $treeKey = (
            "{0}|{1}" -f
            $InstanceId,
            $CompanyCode
        ).ToLowerInvariant()

        if ($script:BusyCloudAccountAccessTreeCache.ContainsKey($treeKey)) {
            try {
                $script:BusyCloudAccountAccessTreeCache.Remove($treeKey)
            }
            catch {
            }
        }
    }
}

function Read-Recordset {
    param($rst, [scriptblock]$RowMapper)
    $results = @()
    if ($null -eq $rst) { return $results }
    try {
        if ($rst.RecordCount -lt 0) { $rst.MoveLast() | Out-Null }
        if ($rst.RecordCount -gt 0) {
            $rst.MoveFirst()
            while (-not $rst.EOF) {
                $row = & $RowMapper $rst
                if ($null -ne $row) { $results += $row }
                $rst.MoveNext()
            }
        }
    } catch { }
    return $results
}

# ═══════════════════════════════════════════════════════
#  GET ACCOUNT GROUPS
# ═══════════════════════════════════════════════════════
function Get-AccountGroups {
    param(
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )
    $cacheKey = "$InstanceId|$CompanyCode|acct-groups"
    $cached = Get-Cache $cacheKey
    if ($cached) { return $cached }
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
    try {
        $qry = "SELECT Master1.Code, Master1.Name,
                    (SELECT M1.Name FROM Master1 M1 WHERE M1.Code = Master1.ParentGrp) AS ParentName
                FROM Master1
                WHERE Master1.MasterType = 1
                ORDER BY Master1.Name"
        $rst    = $fi.GetRecordset($qry)
        $groups = Read-Recordset $rst {
            param($r)
            $parentVal = $r.Fields.Item("ParentName").Value
            $parent    = if ($parentVal -and $parentVal -ne [System.DBNull]::Value) { [string]$parentVal } else { "" }
            $nameVal   = $r.Fields.Item("Name").Value
            $name      = if ($nameVal -and $nameVal -ne [System.DBNull]::Value) { [string]$nameVal } else { "" }
            $codeVal   = $r.Fields.Item("Code").Value
            $code      = if ($codeVal -and $codeVal -ne [System.DBNull]::Value) { [int][string]$codeVal } else { 0 }
            @{ code = $code; name = $name; parent = $parent }
        }
        $dataArray = @($groups)
        $result    = @{ success = $true; count = $dataArray.Count; data = $dataArray }
        Set-Cache $cacheKey $result
        return $result
    } finally { Disconnect-BUSY $fi }
}

# ═══════════════════════════════════════════════════════
#  GET PAGINATED ACCOUNTS
# ═══════════════════════════════════════════════════════
function Get-Accounts {
    param(
        [string]$GroupName   = "",
        [string]$Search      = "",
        [string]$InstanceId  = "",
        [string]$CompanyCode = "",
        [int]$Page           = 1,
        [int]$PageSize       = 30
    )

    if ($Page -lt 1) { $Page = 1 }
    if ($PageSize -lt 1) { $PageSize = 30 }

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }

    try {
        # Dynamically set the wildcard string based on dbType
        $targetInst = Get-InstanceConfig -InstanceId $InstanceId
        $dbType = 0
        if ($null -ne $targetInst -and $null -ne $targetInst.dbType) { $dbType = [int]$targetInst.dbType }
        $wildcard = if ($dbType -eq 1) { "%" } else { "*" }

        # Build dynamic MS Access WHERE filter block
        $where = "Master1.MasterType = 2"
        if ($GroupName -and $GroupName -ne "") {
            $safeGroup = $GroupName -replace "'", "''"
            $where += " AND (
                Master1.ParentGrp = (SELECT Code FROM Master1 WHERE Name = '$safeGroup' AND MasterType = 1)
                OR
                Master1.ParentGrp IN (
                    SELECT Code FROM Master1
                    WHERE ParentGrp = (SELECT Code FROM Master1 WHERE Name = '$safeGroup' AND MasterType = 1)
                      AND MasterType = 1
                )
            )"
        }

        # FIX: Apply dynamic database wildcards
        if ($Search -and $Search -ne "") {
            $safeSearch = $Search -replace "'", "''"
            $where += " AND (Master1.Name LIKE '$wildcard$safeSearch$wildcard' OR Master1.Alias LIKE '$wildcard$safeSearch$wildcard')"
        }

        # 1. Fetch exact total records count
        $countQry = "SELECT COUNT(*) AS TotalCount FROM Master1 WHERE $where"
        $countRst = $fi.GetRecordset($countQry)
        $totalRecords = 0
        if ($countRst -and -not $countRst.EOF) {
            $totalRecords = [int]$countRst.Fields.Item("TotalCount").Value
            $countRst.Close()
        }

        # 2. Query target page dataset
        $qry = "SELECT Master1.Code, Master1.Name, Master1.Alias,
                    Master1.CM3 AS PrintName,
                    (SELECT M1.Name FROM Master1 M1 WHERE M1.Code = Master1.ParentGrp) AS GroupName
                FROM Master1
                WHERE $where
                ORDER BY Master1.Name"

        $rst      = $fi.GetRecordset($qry)
        $accounts = @()

        # Pagination indexing parameters
        $startIndex   = ($Page - 1) * $PageSize
        $endIndex     = $startIndex + $PageSize - 1
        $currentIndex = 0

        # FIX: Added EOF safety check to prevent COM exception 3021
        if ($rst -and -not $rst.EOF) {
            $rst.MoveFirst()
            while (-not $rst.EOF) {
                # Skip items before target page
                if ($currentIndex -lt $startIndex) {
                    $currentIndex++
                    $rst.MoveNext()
                    continue
                }
                
                # Break once page is full
                if ($currentIndex -gt $endIndex) {
                    break
                }

                $code = [int][string]$rst.Fields.Item("Code").Value

                # Load detailed fields only for this page's 30 items
                $xmlStr = $fi.GetMasterXML($code)
                $xml    = [xml]$xmlStr

                $opRaw     = if ($xml.Account.OPBal) { [double]$xml.Account.OPBal } else { 0.0 }
                $opBal     = [Math]::Abs($opRaw)
                $opBalType = if ($opRaw -lt 0) { "D" } elseif ($opRaw -gt 0) { "C" } else { "D" }

                $mobileNo = [string]$xml.Account.Address.Mobile
                $email    = [string]$xml.Account.Address.Email

                $accounts += @{
                    code      = $code
                    name      = [string]$rst.Fields.Item("Name").Value
                    alias     = [string]$rst.Fields.Item("Alias").Value
                    printName = [string]$rst.Fields.Item("PrintName").Value
                    mobileNo  = $mobileNo
                    email     = $email
                    group     = [string]$rst.Fields.Item("GroupName").Value
                    opBal     = $opBal
                    opBalType = $opBalType
                }

                $currentIndex++
                $rst.MoveNext()
            }
            try { $rst.Close() } catch {}
        }

        # Calculate Total Pages
        $totalPages = [Math]::Ceiling($totalRecords / $PageSize)
        if ($totalPages -lt 1) { $totalPages = 1 }

        return @{ 
            success    = $true 
            total      = $totalRecords
            page       = $Page
            pageSize   = $PageSize
            totalPages = $totalPages
            data       = @($accounts) 
        }
    } finally { 
        Disconnect-BUSY $fi 
    }
}

# ═══════════════════════════════════════════════════════
#  GET ACCOUNTS BY GROUP (Backwards-compatibility route forwarding)
# ═══════════════════════════════════════════════════════
function Get-AccountsByGroup {
    param(
        [string]$GroupName,
        [string]$InstanceId  = "",
        [string]$CompanyCode = "",
        [int]$Page           = 1,
        [int]$PageSize       = 30
    )
    return Get-Accounts `
        -GroupName   $GroupName `
        -Page        $Page `
        -PageSize    $PageSize `
        -InstanceId  $InstanceId `
        -CompanyCode $CompanyCode
}

# ═══════════════════════════════════════════════════════
#  CREATE ACCOUNT GROUP
# ═══════════════════════════════════════════════════════
function Create-AccountGroup {
    param(
        $Name,
        $ParentGroup = "Primary",
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
    try {
        $xml = "<AccountGroup><Name>$(ConvertTo-XmlSafe $Name)</Name><ParentGroupName>$(ConvertTo-XmlSafe $ParentGroup)</ParentGroupName></AccountGroup>"
        $err = ""
        $saved = $fi.SaveMasterFromXML(1, $xml, [ref]$err, $false)
        if ($saved -eq $true) {
            Clear-AccountCaches -InstanceId $InstanceId -CompanyCode $CompanyCode
            return @{ success = $true; message = "Group created" }
        } else {
            $errOut = if ($err) { $err } else { "Group may already exist" }
            return @{ success = $false; error = $errOut }
        }
    } finally { Disconnect-BUSY $fi }
}

# ═══════════════════════════════════════════════════════
#  UPDATE ACCOUNT GROUP
# ═══════════════════════════════════════════════════════
function Update-AccountGroup {
    param(
        $Name,
        $ParentGroup,
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
    try {
        $xml = "<AccountGroup><Name>$(ConvertTo-XmlSafe $Name)</Name><ParentGroupName>$(ConvertTo-XmlSafe $ParentGroup)</ParentGroupName></AccountGroup>"
        $err = ""
        $saved = $fi.SaveMasterFromXML(1, $xml, [ref]$err, $true)
        if ($saved -eq $true) {
            Clear-AccountCaches -InstanceId $InstanceId -CompanyCode $CompanyCode
            return @{ success = $true; message = "Group updated" }
        } else {
            $errOut = if ($err) { $err } else { "Group not found" }
            return @{ success = $false; error = $errOut }
        }
    } finally { Disconnect-BUSY $fi }
}

function Get-CashBankAccounts {
    param(
        [string]$Search      = "",
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )

    $startedAt = [System.Diagnostics.Stopwatch]::StartNew()
    $ctx = $null
    $reader = $null
    $cmd = $null

    try {
        $resolver = Get-Command `
            Get-BusyCloudFastConfigDbContext `
            -ErrorAction SilentlyContinue

        if ($null -eq $resolver) {
            throw "Fast fiscal database resolver is unavailable."
        }

        $ctx = Get-BusyCloudFastConfigDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (
            $null -eq $ctx -or
            $null -eq $ctx.connection
        ) {
            throw "Direct fiscal database connection is unavailable."
        }

        $conn = $ctx.connection
        $dbType = [int]$ctx.dbType
        $wildcard = if ($dbType -eq 1) { "%" } else { "*" }

        $where = @"
Master1.MasterType = 2
AND Master1.ParentGrp IN (
    SELECT Code
    FROM Master1
    WHERE Name IN ('Cash-in-hand', 'Bank Accounts')
)
"@

        if (-not [string]::IsNullOrWhiteSpace($Search)) {
            $safeSearch = $Search.Trim().Replace("'", "''")

            $where += @"
 AND (
    Master1.Name LIKE '$wildcard$safeSearch$wildcard'
    OR Master1.Alias LIKE '$wildcard$safeSearch$wildcard'
 )
"@
        }

        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}

        $cmd.CommandText = @"
SELECT
    Master1.Code,
    Master1.Name,
    Master1.Alias,
    (
        SELECT M1.Name
        FROM Master1 M1
        WHERE M1.Code = Master1.ParentGrp
    ) AS ParentGrpName
FROM Master1
WHERE $where
ORDER BY Master1.Name
"@

        $reader = $cmd.ExecuteReader()

        $accounts = @()

        while ($reader.Read()) {
            $code = 0
            $name = ""
            $alias = ""
            $parentGrp = ""

            try {
                if (-not $reader.IsDBNull(0)) {
                    $code = [int][string]$reader.GetValue(0)
                }
            }
            catch {
                $code = 0
            }

            try {
                if (-not $reader.IsDBNull(1)) {
                    $name = [string]$reader.GetValue(1)
                }
            }
            catch {
                $name = ""
            }

            try {
                if (-not $reader.IsDBNull(2)) {
                    $alias = [string]$reader.GetValue(2)
                }
            }
            catch {
                $alias = ""
            }

            try {
                if (-not $reader.IsDBNull(3)) {
                    $parentGrp = [string]$reader.GetValue(3)
                }
            }
            catch {
                $parentGrp = ""
            }

            if ($code -le 0) {
                continue
            }

            $kind = if ($parentGrp -eq "Bank Accounts") {
                "Bank"
            }
            else {
                "Cash"
            }

            $accounts += @{
                code  = $code
                name  = $name
                alias = $alias
                group = $parentGrp
                type  = $kind
            }
        }

        $startedAt.Stop()

        Write-Host (
            "  [CASH-BANK-FAST] {0}/{1} db={2} rows={3} elapsedMs={4}" -f
            $InstanceId,
            $CompanyCode,
            [string]$ctx.database,
            @($accounts).Count,
            [int]$startedAt.ElapsedMilliseconds
        ) -ForegroundColor DarkCyan

        return @{
            success = $true
            count   = @($accounts).Count
            data    = @($accounts)
        }
    }
    catch {
        if ($startedAt.IsRunning) {
            $startedAt.Stop()
        }

        Write-Host (
            "  [CASH-BANK-FAST FAIL] {0}/{1} elapsedMs={2} error={3}" -f
            $InstanceId,
            $CompanyCode,
            [int]$startedAt.ElapsedMilliseconds,
            $_.Exception.Message
        ) -ForegroundColor Red

        return @{
            success = $false
            error   = $_.Exception.Message
            data    = @()
        }
    }
    finally {
        if ($reader) {
            try { $reader.Close() } catch {}
            try { $reader.Dispose() } catch {}
        }

        if ($cmd) {
            try { $cmd.Dispose() } catch {}
        }

        if (
            $ctx -and
            $ctx.connection
        ) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}


# ═══════════════════════════════════════════════════════
#  GET PARTIES (Paginated, Searchable & Cash/Bank Capable)
# ═══════════════════════════════════════════════════════
# ===============================================================
# PARTY ACCOUNT GROUP PERMISSIONS
# ===============================================================
# The permission tree is built from BUSY Master1:
#   MasterType = 1 -> Account Group
#   MasterType = 2 -> Account
#   ParentGrp       -> parent Account Group Code
#
# Only the three built-in BUSY party roots are eligible:
#   Code 111 -> Cash-in-hand
#   Code 116 -> original BUSY group "Sundry Debtors"
#   Code 117 -> original BUSY group "Sundry Creditors"
#
# IMPORTANT:
# BUSY allows these built-in group captions to be renamed (for example
# "Sundry Debtors" -> "Account Receivables" and "Sundry Creditors" ->
# "Accounts Payables"). Therefore Party permissions MUST identify these
# roots by their stable BUSY system Code, not by the current Name/caption.
#
# Similar custom sibling groups are intentionally excluded.

function Get-PartyGroupNormalizedName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return ""
    }

    $value = $Name.Trim().ToLowerInvariant()
    $value = $value -replace "[-_]+", " "
    $value = $value -replace "\s+", " "
    return $value.Trim()
}

function Get-PartyPermissionGroupRows {
    param($fi)

    $qry = @"
SELECT
    Code,
    Name,
    ParentGrp
FROM Master1
WHERE MasterType = 1
ORDER BY Name
"@

    $rst = $fi.GetRecordset($qry)
    $groups = @()

    if ($rst -and -not $rst.EOF) {
        try { $rst.MoveFirst() } catch {}

        while (-not $rst.EOF) {
            $codeValue = $rst.Fields.Item("Code").Value
            $nameValue = $rst.Fields.Item("Name").Value
            $parentValue = $rst.Fields.Item("ParentGrp").Value

            $code = if ($codeValue -ne [System.DBNull]::Value) {
                [int][string]$codeValue
            } else { 0 }

            $name = if ($nameValue -ne [System.DBNull]::Value) {
                ([string]$nameValue).Trim()
            } else { "" }

            $parentGrp = if ($parentValue -ne [System.DBNull]::Value) {
                [int][string]$parentValue
            } else { 0 }

            if ($code -gt 0) {
                $groups += @{
                    code      = $code
                    name      = $name
                    parentGrp = $parentGrp
                }
            }

            $rst.MoveNext()
        }

        try { $rst.Close() } catch {}
    }

    return @($groups)
}

function Get-PartyPermissionTreeInfo {
    param([array]$Groups)

    # Stable BUSY system root identities. Do not match these by Name:
    # users can change the visible captions in BUSY while the system Codes
    # remain the identity of the original built-in groups.
    $partyRootCodes = @(111, 116, 117)

    $groupByCode = @{}
    $childrenByParent = @{}

    foreach ($group in @($Groups)) {
        $code = [int]$group.code
        $parentCode = [int]$group.parentGrp
        $groupByCode["$code"] = $group

        if (-not $childrenByParent.ContainsKey("$parentCode")) {
            $childrenByParent["$parentCode"] = @()
        }
        $childrenByParent["$parentCode"] += $group
    }

    $roots = @(
        @($Groups) |
        Where-Object {
            $partyRootCodes -contains [int]$_.code
        }
    )

    $targetCodes = @{}
    $rootForCode = @{}
    $levelForCode = @{}
    $pathForCode = @{}

    $queue = New-Object System.Collections.Queue

    foreach ($root in $roots) {
        $queue.Enqueue(@{
            group    = $root
            rootCode = [int]$root.code
            level    = 0
            path     = @([string]$root.name)
        })
    }

    while ($queue.Count -gt 0) {
        $node = $queue.Dequeue()
        $group = $node.group
        $code = [int]$group.code
        $key = "$code"

        if ($targetCodes.ContainsKey($key)) {
            continue
        }

        $targetCodes[$key] = $true
        $rootForCode[$key] = [int]$node.rootCode
        $levelForCode[$key] = [int]$node.level
        $pathForCode[$key] = @($node.path)

        if ($childrenByParent.ContainsKey($key)) {
            foreach ($child in @($childrenByParent[$key])) {
                $queue.Enqueue(@{
                    group    = $child
                    rootCode = [int]$node.rootCode
                    level    = ([int]$node.level + 1)
                    path     = @($node.path) + @([string]$child.name)
                })
            }
        }
    }

    return @{
        groups           = @($Groups)
        roots            = @($roots)
        groupByCode      = $groupByCode
        childrenByParent = $childrenByParent
        targetCodes      = $targetCodes
        rootForCode      = $rootForCode
        levelForCode     = $levelForCode
        pathForCode      = $pathForCode
    }
}

function Resolve-AllowedPartyGroupCodes {
    param(
        [array]$Groups,
        [int[]]$SelectedGroupCodes = @()
    )

    if ($null -eq $SelectedGroupCodes -or @($SelectedGroupCodes).Count -eq 0) {
        return @()
    }

    $tree = Get-PartyPermissionTreeInfo -Groups $Groups
    $granted = @{}
    $queue = New-Object System.Collections.Queue

    foreach ($rawCode in @($SelectedGroupCodes)) {
        $code = [int]$rawCode
        $key = "$code"

        # Ignore invalid groups and any group outside the exact three target
        # root trees. This keeps the server authoritative even if M2 is edited.
        if ($code -gt 0 -and $tree.targetCodes.ContainsKey($key)) {
            $queue.Enqueue($code)
        }
    }

    while ($queue.Count -gt 0) {
        $code = [int]$queue.Dequeue()
        $key = "$code"

        if ($granted.ContainsKey($key)) {
            continue
        }

        $granted[$key] = $true

        if ($tree.childrenByParent.ContainsKey($key)) {
            foreach ($child in @($tree.childrenByParent[$key])) {
                $queue.Enqueue([int]$child.code)
            }
        }
    }

    return @(
        $granted.Keys |
        ForEach-Object { [int]$_ } |
        Sort-Object
    )
}

function Get-PartyAccountGroups {
    param(
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )

    if ($null -eq $script:BusyCloudPartyAccountGroupsCache) {
        $script:BusyCloudPartyAccountGroupsCache = @{}
    }

    $cacheKey = (
        "{0}|{1}" -f
        $InstanceId,
        $CompanyCode
    ).ToLowerInvariant()

    if ($script:BusyCloudPartyAccountGroupsCache.ContainsKey($cacheKey)) {
        $cached = $script:BusyCloudPartyAccountGroupsCache[$cacheKey]

        if (
            $cached -and
            $cached.expires -and
            (Get-Date) -lt $cached.expires
        ) {
            Write-Host (
                "  [PARTY-GROUPS-FAST] cache HIT {0}/{1} rows={2}" -f
                $InstanceId,
                $CompanyCode,
                @($cached.data).Count
            ) -ForegroundColor DarkCyan

            return @{
                success = $true
                count   = @($cached.data).Count
                data    = @($cached.data)
            }
        }

        try {
            $script:BusyCloudPartyAccountGroupsCache.Remove($cacheKey)
        }
        catch {
        }
    }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $ctx = $null

    try {
        $resolver = Get-Command `
            Get-BusyCloudFastConfigDbContext `
            -ErrorAction SilentlyContinue

        if ($null -eq $resolver) {
            throw "Fast fiscal database resolver is unavailable."
        }

        $ctx = Get-BusyCloudFastConfigDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (
            $null -eq $ctx -or
            $null -eq $ctx.connection
        ) {
            throw "Direct fiscal database connection is unavailable."
        }

        # Reuse the existing direct party-group loader and exact tree builder.
        # No BUSY COM is required for this read-only permission master.
        $groups = @(
            Get-PartyPermissionGroupRowsDirect `
                -Connection $ctx.connection
        )

        $tree = Get-PartyPermissionTreeInfo `
            -Groups $groups

        $output = @()

        foreach ($group in @($groups)) {
            $code = [int]$group.code
            $key = "$code"

            if (-not $tree.targetCodes.ContainsKey($key)) {
                continue
            }

            $rootCode = [int]$tree.rootForCode[$key]
            $root = $tree.groupByCode["$rootCode"]
            $pathNames = @($tree.pathForCode[$key])

            $output += @{
                rootCode   = $rootCode
                rootName   = [string]$root.name
                code       = $code
                name       = [string]$group.name
                parentCode = [int]$group.parentGrp
                level      = [int]$tree.levelForCode[$key]
                pathText   = ($pathNames -join " > ")
            }
        }

        $output = @(
            $output |
            Sort-Object rootName, pathText
        )

        $script:BusyCloudPartyAccountGroupsCache[$cacheKey] = @{
            expires = (Get-Date).AddMinutes(30)
            data    = @($output)
        }

        $sw.Stop()

        Write-Host (
            "  [PARTY-GROUPS-FAST] {0}/{1} db={2} rows={3} elapsedMs={4}" -f
            $InstanceId,
            $CompanyCode,
            [string]$ctx.database,
            @($output).Count,
            [int]$sw.ElapsedMilliseconds
        ) -ForegroundColor DarkCyan

        return @{
            success = $true
            count   = @($output).Count
            data    = @($output)
        }
    }
    catch {
        if ($sw.IsRunning) {
            $sw.Stop()
        }

        Write-Host (
            "  [PARTY-GROUPS-FAST FAIL] {0}/{1} elapsedMs={2} error={3}" -f
            $InstanceId,
            $CompanyCode,
            [int]$sw.ElapsedMilliseconds,
            $_.Exception.Message
        ) -ForegroundColor Red

        return @{
            success = $false
            error   = $_.Exception.Message
            data    = @()
        }
    }
    finally {
        if (
            $ctx -and
            $ctx.connection
        ) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}

function Test-PartyAccountsAllowed {
    param(
        [string[]]$AccountNames = @(),
        [int[]]$AllowedGroupCodes = @(),
        [bool]$AllowAllEligibleRoots = $false,
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )

    $requestedNames = @(
        @($AccountNames) |
        ForEach-Object { ([string]$_).Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique
    )

    if ($requestedNames.Count -eq 0) {
        return @{ success = $true; allowed = $true; deniedAccounts = @() }
    }

    if (
        -not $AllowAllEligibleRoots -and
        ($null -eq $AllowedGroupCodes -or @($AllowedGroupCodes).Count -eq 0)
    ) {
        return @{
            success        = $true
            allowed        = $false
            deniedAccounts = @($requestedNames)
        }
    }

    $fi = Connect-BUSY `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $fi) {
        return @{ success = $false; allowed = $false; error = "BUSY connection failed"; deniedAccounts = $requestedNames }
    }

    try {
        $groups = Get-PartyPermissionGroupRows -fi $fi
        $selectedCodes = @($AllowedGroupCodes)

        if ($AllowAllEligibleRoots) {
            $tree = Get-PartyPermissionTreeInfo -Groups $groups
            $selectedCodes = @($tree.roots | ForEach-Object { [int]$_.code })
        }

        $expandedCodes = @(
            Resolve-AllowedPartyGroupCodes `
                -Groups $groups `
                -SelectedGroupCodes $selectedCodes
        )

        # Normal target-voucher users fail closed when no valid group branch
        # has been assigned.
        if ($expandedCodes.Count -eq 0) {
            return @{
                success        = $true
                allowed        = $false
                deniedAccounts = @($requestedNames)
            }
        }

        $allowedSet = @{}
        foreach ($code in $expandedCodes) {
            $allowedSet["$([int]$code)"] = $true
        }

        $safeNames = @(
            $requestedNames |
            ForEach-Object { "'$(($_ -replace "'", "''"))'" }
        )

        $qry = @"
SELECT
    Name,
    ParentGrp
FROM Master1
WHERE MasterType = 2
  AND Name IN ($($safeNames -join ','))
"@

        $rst = $fi.GetRecordset($qry)
        $allowedNames = @{}

        if ($rst -and -not $rst.EOF) {
            try { $rst.MoveFirst() } catch {}

            while (-not $rst.EOF) {
                $nameValue = $rst.Fields.Item("Name").Value
                $parentValue = $rst.Fields.Item("ParentGrp").Value

                $name = if ($nameValue -ne [System.DBNull]::Value) {
                    ([string]$nameValue).Trim()
                } else { "" }

                $parentCode = if ($parentValue -ne [System.DBNull]::Value) {
                    [int][string]$parentValue
                } else { 0 }

                if (
                    -not [string]::IsNullOrWhiteSpace($name) -and
                    $allowedSet.ContainsKey("$parentCode")
                ) {
                    $allowedNames[$name.ToLowerInvariant()] = $true
                }

                $rst.MoveNext()
            }

            try { $rst.Close() } catch {}
        }

        $denied = @()
        foreach ($name in $requestedNames) {
            if (-not $allowedNames.ContainsKey($name.ToLowerInvariant())) {
                $denied += $name
            }
        }

        return @{
            success        = $true
            allowed        = ($denied.Count -eq 0)
            deniedAccounts = @($denied)
        }
    }
    catch {
        return @{
            success        = $false
            allowed        = $false
            error          = $_.Exception.Message
            deniedAccounts = @($requestedNames)
        }
    }
    finally {
        Disconnect-BUSY $fi
    }
}


# ===============================================================
# JOURNAL / CONTRA FULL ACCOUNT ACCESS PERMISSIONS
# ===============================================================
# These permissions are intentionally different from Party Account Access.
# Journal and Contra may use ANY BUSY Account Group / Account, and Debit and
# Credit selections are stored independently in MobileUserPreference.M2.

function Resolve-AllowedAccountGroupCodes {
    param(
        [array]$Groups,
        [int[]]$SelectedGroupCodes = @()
    )

    if ($null -eq $SelectedGroupCodes -or @($SelectedGroupCodes).Count -eq 0) {
        return @()
    }

    $groupByCode = @{}
    $childrenByParent = @{}

    foreach ($group in @($Groups)) {
        $code = [int]$group.code
        $parentCode = [int]$group.parentGrp
        $groupByCode["$code"] = $group

        if (-not $childrenByParent.ContainsKey("$parentCode")) {
            $childrenByParent["$parentCode"] = @()
        }
        $childrenByParent["$parentCode"] += $group
    }

    $granted = @{}
    $queue = New-Object System.Collections.Queue

    foreach ($rawCode in @($SelectedGroupCodes)) {
        $code = 0
        if ([int]::TryParse([string]$rawCode, [ref]$code) -and $code -gt 0) {
            if ($groupByCode.ContainsKey("$code")) {
                $queue.Enqueue($code)
            }
        }
    }

    while ($queue.Count -gt 0) {
        $code = [int]$queue.Dequeue()
        $key = "$code"

        if ($granted.ContainsKey($key)) {
            continue
        }

        $granted[$key] = $true

        if ($childrenByParent.ContainsKey($key)) {
            foreach ($child in @($childrenByParent[$key])) {
                $queue.Enqueue([int]$child.code)
            }
        }
    }

    return @(
        $granted.Keys |
        ForEach-Object { [int]$_ } |
        Sort-Object
    )
}

function Get-AllAccountPermissionNodes {
    param(
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )

    # Journal / Contra permissions require the FULL Chart of Accounts:
    #   MasterType = 1 -> account groups
    #   MasterType = 2 -> ledger accounts
    #
    # The earlier performance patch accidentally returned GROUP nodes only.
    # This implementation restores GROUP + ACCOUNT while remaining COM-free.

    if ($null -eq $script:BusyCloudAccountAccessTreeCache) {
        $script:BusyCloudAccountAccessTreeCache = @{}
    }

    $cacheKey = (
        "{0}|{1}" -f
        $InstanceId,
        $CompanyCode
    ).ToLowerInvariant()

    if ($script:BusyCloudAccountAccessTreeCache.ContainsKey($cacheKey)) {
        $cached = $script:BusyCloudAccountAccessTreeCache[$cacheKey]

        if (
            $cached -and
            $cached.expires -and
            (Get-Date) -lt $cached.expires
        ) {
            Write-Host (
                "  [ACCOUNT-TREE-FAST-V2] cache HIT {0}/{1} rows={2}" -f
                $InstanceId,
                $CompanyCode,
                @($cached.data).Count
            ) -ForegroundColor DarkCyan

            return @{
                success = $true
                count   = @($cached.data).Count
                data    = @($cached.data)
                cached  = $true
            }
        }

        try {
            $script:BusyCloudAccountAccessTreeCache.Remove($cacheKey)
        }
        catch {
        }
    }

    $startedAt = [System.Diagnostics.Stopwatch]::StartNew()
    $ctx = $null
    $reader = $null
    $cmd = $null

    try {
        $resolver = Get-Command `
            Get-BusyCloudFastConfigDbContext `
            -ErrorAction SilentlyContinue

        if ($null -eq $resolver) {
            throw "Fast fiscal database resolver is unavailable."
        }

        $ctx = Get-BusyCloudFastConfigDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (
            $null -eq $ctx -or
            $null -eq $ctx.connection
        ) {
            throw "Direct fiscal database connection is unavailable."
        }

        $conn = $ctx.connection

        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 8 } catch {}

        # No ORDER BY here: large Master1 tables can waste time sorting.
        # We sort the final small in-memory tree instead.
        $cmd.CommandText = @"
SELECT
    MasterType,
    Code,
    Name,
    Alias,
    ParentGrp
FROM Master1
WHERE MasterType IN (1, 2)
"@

        $reader = $cmd.ExecuteReader()

        $groups = @()
        $accounts = @()

        while ($reader.Read()) {
            $masterType = 0
            $code = 0
            $name = ""
            $alias = ""
            $parentGrp = 0

            try {
                if (-not $reader.IsDBNull(0)) {
                    $masterType = [int][string]$reader.GetValue(0)
                }
            }
            catch {}

            try {
                if (-not $reader.IsDBNull(1)) {
                    $code = [int][string]$reader.GetValue(1)
                }
            }
            catch {}

            try {
                if (-not $reader.IsDBNull(2)) {
                    $name = ([string]$reader.GetValue(2)).Trim()
                }
            }
            catch {}

            try {
                if (-not $reader.IsDBNull(3)) {
                    $alias = ([string]$reader.GetValue(3)).Trim()
                }
            }
            catch {}

            try {
                if (-not $reader.IsDBNull(4)) {
                    $parentGrp = [int][string]$reader.GetValue(4)
                }
            }
            catch {}

            if ($code -le 0) {
                continue
            }

            $row = @{
                masterType = $masterType
                code       = $code
                name       = $name
                alias      = $alias
                parentGrp  = $parentGrp
            }

            if ($masterType -eq 1) {
                $groups += $row
            }
            elseif ($masterType -eq 2) {
                $accounts += $row
            }
        }

        try { $reader.Close() } catch {}
        try { $reader.Dispose() } catch {}
        $reader = $null

        try { $cmd.Dispose() } catch {}
        $cmd = $null

        $groupByCode = @{}
        foreach ($group in @($groups)) {
            $groupByCode[[int]$group.code] = $group
        }

        $pathCache = @{}

        function Resolve-AccountPermissionGroupPathV2 {
            param(
                [int]$StartCode,
                [hashtable]$GroupByCode,
                [hashtable]$PathCache
            )

            if ($PathCache.ContainsKey($StartCode)) {
                return @($PathCache[$StartCode])
            }

            $reverse = @()
            $visited = @{}
            $current = $StartCode

            for ($guard = 0; $guard -lt 100; $guard++) {
                if (
                    $current -le 0 -or
                    -not $GroupByCode.ContainsKey($current) -or
                    $visited.ContainsKey($current)
                ) {
                    break
                }

                $visited[$current] = $true

                $group = $GroupByCode[$current]
                $reverse += $group
                $current = [int]$group.parentGrp
            }

            $ordered = @()
            for ($i = $reverse.Count - 1; $i -ge 0; $i--) {
                $ordered += $reverse[$i]
            }

            $PathCache[$StartCode] = @($ordered)
            return @($ordered)
        }

        $nodes = @()

        # GROUP nodes.
        foreach ($group in @($groups)) {
            $path = @(
                Resolve-AccountPermissionGroupPathV2 `
                    -StartCode ([int]$group.code) `
                    -GroupByCode $groupByCode `
                    -PathCache $pathCache
            )

            $root = if ($path.Count -gt 0) {
                $path[0]
            }
            else {
                $group
            }

            $pathNames = @(
                $path |
                ForEach-Object { [string]$_.name }
            )

            $nodes += @{
                rootCode   = [int]$root.code
                rootName   = [string]$root.name
                nodeType   = "GROUP"
                code       = [int]$group.code
                name       = [string]$group.name
                alias      = [string]$group.alias
                parentCode = [int]$group.parentGrp
                level      = [Math]::Max(0, $path.Count - 1)
                pathText   = ($pathNames -join " > ")
            }
        }

        # ACCOUNT nodes.
        foreach ($account in @($accounts)) {
            $groupPath = @(
                Resolve-AccountPermissionGroupPathV2 `
                    -StartCode ([int]$account.parentGrp) `
                    -GroupByCode $groupByCode `
                    -PathCache $pathCache
            )

            $root = if ($groupPath.Count -gt 0) {
                $groupPath[0]
            }
            elseif ($groupByCode.ContainsKey([int]$account.parentGrp)) {
                $groupByCode[[int]$account.parentGrp]
            }
            else {
                @{
                    code = [int]$account.parentGrp
                    name = ""
                }
            }

            $pathNames = @(
                $groupPath |
                ForEach-Object { [string]$_.name }
            )

            $pathNames += [string]$account.name

            $nodes += @{
                rootCode   = [int]$root.code
                rootName   = [string]$root.name
                nodeType   = "ACCOUNT"
                code       = [int]$account.code
                name       = [string]$account.name
                alias      = [string]$account.alias
                parentCode = [int]$account.parentGrp
                level      = [Math]::Max(0, $groupPath.Count)
                pathText   = ($pathNames -join " > ")
            }
        }

        $nodes = @(
            $nodes |
            Sort-Object pathText, nodeType, name
        )

        $script:BusyCloudAccountAccessTreeCache[$cacheKey] = @{
            expires = (Get-Date).AddMinutes(30)
            data    = @($nodes)
        }

        $startedAt.Stop()

        $groupCount = @(
            $nodes |
            Where-Object { $_.nodeType -eq "GROUP" }
        ).Count

        $accountCount = @(
            $nodes |
            Where-Object { $_.nodeType -eq "ACCOUNT" }
        ).Count

        Write-Host (
            "  [ACCOUNT-TREE-FAST-V2] {0}/{1} db={2} groups={3} accounts={4} total={5} elapsedMs={6}" -f
            $InstanceId,
            $CompanyCode,
            [string]$ctx.database,
            $groupCount,
            $accountCount,
            $nodes.Count,
            [int]$startedAt.ElapsedMilliseconds
        ) -ForegroundColor DarkCyan

        return @{
            success = $true
            count   = $nodes.Count
            data    = @($nodes)
            cached  = $false
        }
    }
    catch {
        if ($startedAt.IsRunning) {
            $startedAt.Stop()
        }

        Write-Host (
            "  [ACCOUNT-TREE-FAST-V2 FAIL] {0}/{1} elapsedMs={2} error={3}" -f
            $InstanceId,
            $CompanyCode,
            [int]$startedAt.ElapsedMilliseconds,
            $_.Exception.Message
        ) -ForegroundColor Red

        return @{
            success = $false
            error = $_.Exception.Message
            data = @()
        }
    }
    finally {
        if ($reader) {
            try { $reader.Close() } catch {}
            try { $reader.Dispose() } catch {}
        }

        if ($cmd) {
            try { $cmd.Dispose() } catch {}
        }

        if (
            $ctx -and
            $ctx.connection
        ) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}

function Test-AccountsAllowedByCodes {
    param(
        [string[]]$AccountNames = @(),
        [int[]]$AllowedGroupCodes = @(),
        [int[]]$AllowedAccountCodes = @(),
        [bool]$AllowAllAccounts = $false,
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )

    $requestedNames = @(
        @($AccountNames) |
        ForEach-Object { ([string]$_).Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique
    )

    if ($requestedNames.Count -eq 0 -or $AllowAllAccounts) {
        return @{ success = $true; allowed = $true; deniedAccounts = @() }
    }

    if (
        ($null -eq $AllowedGroupCodes -or @($AllowedGroupCodes).Count -eq 0) -and
        ($null -eq $AllowedAccountCodes -or @($AllowedAccountCodes).Count -eq 0)
    ) {
        return @{ success = $true; allowed = $false; deniedAccounts = @($requestedNames) }
    }

    $fi = Connect-BUSY `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $fi) {
        return @{ success = $false; allowed = $false; error = "BUSY connection failed"; deniedAccounts = @($requestedNames) }
    }

    try {
        $groups = Get-PartyPermissionGroupRows -fi $fi
        $expandedGroups = @(Resolve-AllowedAccountGroupCodes -Groups $groups -SelectedGroupCodes @($AllowedGroupCodes))

        $allowedGroupSet = @{}
        foreach ($code in $expandedGroups) { $allowedGroupSet["$([int]$code)"] = $true }

        $allowedAccountSet = @{}
        foreach ($rawCode in @($AllowedAccountCodes)) {
            $code = 0
            if ([int]::TryParse([string]$rawCode, [ref]$code) -and $code -gt 0) {
                $allowedAccountSet["$code"] = $true
            }
        }

        $safeNames = @($requestedNames | ForEach-Object { "'$(($_ -replace "'", "''"))'" })
        $qry = @"
SELECT
    Code,
    Name,
    ParentGrp
FROM Master1
WHERE MasterType = 2
  AND Name IN ($($safeNames -join ','))
"@

        $rst = $fi.GetRecordset($qry)
        $allowedNames = @{}

        if ($rst -and -not $rst.EOF) {
            try { $rst.MoveFirst() } catch {}

            while (-not $rst.EOF) {
                $codeValue = $rst.Fields.Item("Code").Value
                $nameValue = $rst.Fields.Item("Name").Value
                $parentValue = $rst.Fields.Item("ParentGrp").Value

                $code = if ($codeValue -ne [System.DBNull]::Value) { [int][string]$codeValue } else { 0 }
                $name = if ($nameValue -ne [System.DBNull]::Value) { ([string]$nameValue).Trim() } else { "" }
                $parentCode = if ($parentValue -ne [System.DBNull]::Value) { [int][string]$parentValue } else { 0 }

                if (
                    -not [string]::IsNullOrWhiteSpace($name) -and
                    ($allowedAccountSet.ContainsKey("$code") -or $allowedGroupSet.ContainsKey("$parentCode"))
                ) {
                    $allowedNames[$name.ToLowerInvariant()] = $true
                }

                $rst.MoveNext()
            }

            try { $rst.Close() } catch {}
        }

        $denied = @()
        foreach ($name in $requestedNames) {
            if (-not $allowedNames.ContainsKey($name.ToLowerInvariant())) {
                $denied += $name
            }
        }

        return @{
            success        = $true
            allowed        = ($denied.Count -eq 0)
            deniedAccounts = @($denied)
        }
    }
    catch {
        return @{ success = $false; allowed = $false; error = $_.Exception.Message; deniedAccounts = @($requestedNames) }
    }
    finally {
        Disconnect-BUSY $fi
    }
}


# ===============================================================
# FAST DIRECT-DATABASE PARTY READER
# ===============================================================
# Read-only party/account lookups must not initialize BUSY COM/OpenCSDB.
# On SQL companies that COM startup can take 20-60+ seconds.  The server
# already has a fiscal-database resolver used by voucher/settings reads, so
# this helper reuses that direct ADO.NET connection at request time.
#
# The legacy COM implementation below is intentionally retained as a safety
# fallback.  If the direct resolver is unavailable or a company has an
# unexpected schema, Get-Parties falls back to the existing behaviour.
function Get-PartyPermissionGroupRowsDirect {
    param($Connection)

    $groups = @()
    $cmd = $null
    $rdr = $null

    try {
        $cmd = $Connection.CreateCommand()
        try { $cmd.CommandTimeout = 8 } catch {}
        $cmd.CommandText = @"
SELECT
    Code,
    Name,
    ParentGrp
FROM Master1
WHERE MasterType = 1
ORDER BY Name
"@

        $rdr = $cmd.ExecuteReader()
        while ($rdr.Read()) {
            $code = 0
            $name = ""
            $parentGrp = 0

            try {
                if ($rdr["Code"] -ne [System.DBNull]::Value) {
                    $code = [int][string]$rdr["Code"]
                }
            } catch {}

            try {
                if ($rdr["Name"] -ne [System.DBNull]::Value) {
                    $name = ([string]$rdr["Name"]).Trim()
                }
            } catch {}

            try {
                if ($rdr["ParentGrp"] -ne [System.DBNull]::Value) {
                    $parentGrp = [int][string]$rdr["ParentGrp"]
                }
            } catch {}

            if ($code -gt 0) {
                $groups += @{
                    code      = $code
                    name      = $name
                    parentGrp = $parentGrp
                }
            }
        }
    }
    finally {
        if ($rdr) { try { $rdr.Close() } catch {}; try { $rdr.Dispose() } catch {} }
        if ($cmd) { try { $cmd.Dispose() } catch {} }
    }

    return @($groups)
}

function Get-PartiesDirectDatabase {
    param(
        [string]$Search      = "",
        [bool]$CashBankOnly  = $false,
        [int]$Page           = 1,
        [int]$PageSize       = 30,
        [string]$InstanceId  = "",
        [string]$CompanyCode = "",
        [int[]]$AllowedGroupCodes = @(),
        [bool]$EnforceGroupAccess = $false,
        [bool]$AllowAllEligibleRoots = $false,
        [int[]]$AllowedAccountCodes = @(),
        [bool]$EnforceAllAccountAccess = $false,
        [bool]$AllowAllAccounts = $false
    )

    # Fast read-only party/account list.
    #
    # V3 removes the expensive full-table LEFT JOIN to MasterAddressInfo.
    # The old query joined address data for ALL ~10k accounts before returning
    # only 30 rows. Under SQL locking/load that query could hit the 8s timeout,
    # after which Get-Parties fell back to BUSY COM and froze the API for
    # another ~20-30 seconds.
    #
    # New flow:
    #   1) count/filter Master1 only
    #   2) page Master1 to 30 rows
    #   3) fetch address data only for those 30 account codes
    #
    # SQL reads use NOLOCK because this endpoint is an autocomplete/read-only
    # picker; it must not wait behind BUSY write locks.

    $resolver = Get-Command `
        Get-BusyCloudFastConfigDbContext `
        -ErrorAction SilentlyContinue

    if ($null -eq $resolver) {
        return $null
    }

    $ctx = $null
    $conn = $null
    $reader = $null
    $cmd = $null
    $countCmd = $null
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $dbType = 0

    try {
        if ($Page -lt 1) {
            $Page = 1
        }

        if ($PageSize -lt 1) {
            $PageSize = 30
        }

        $ctx = Get-BusyCloudFastConfigDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (
            $null -eq $ctx -or
            $null -eq $ctx.connection
        ) {
            throw "Direct BUSY fiscal database context was not available."
        }

        $conn = $ctx.connection
        $dbType = if ($null -ne $ctx.dbType) {
            [int]$ctx.dbType
        }
        else {
            0
        }

        $wildcard = if ($dbType -eq 1) {
            "%"
        }
        else {
            "*"
        }

        $masterRef = if ($dbType -eq 1) {
            "Master1 AS M WITH (NOLOCK)"
        }
        else {
            "Master1 AS M"
        }

        $groupRef = if ($dbType -eq 1) {
            "Master1 AS G WITH (NOLOCK)"
        }
        else {
            "Master1 AS G"
        }

        $addressRef = if ($dbType -eq 1) {
            "MasterAddressInfo AS AX WITH (NOLOCK)"
        }
        else {
            "MasterAddressInfo AS AX"
        }

        $where = "M.MasterType = 2"

        # -------------------------------------------------------------
        # Preserve existing account-side permission logic.
        # -------------------------------------------------------------
        if ($EnforceAllAccountAccess) {
            if (-not $AllowAllAccounts) {
                $groupRows = @(
                    Get-PartyPermissionGroupRowsDirect `
                        -Connection $conn
                )

                $expandedAllowedGroups = @(
                    Resolve-AllowedAccountGroupCodes `
                        -Groups $groupRows `
                        -SelectedGroupCodes @($AllowedGroupCodes)
                )

                $validAccountCodes = @(
                    @($AllowedAccountCodes) |
                    ForEach-Object {
                        $parsed = 0
                        if (
                            [int]::TryParse(
                                [string]$_,
                                [ref]$parsed
                            ) -and
                            $parsed -gt 0
                        ) {
                            $parsed
                        }
                    } |
                    Select-Object -Unique
                )

                $accessConditions = @()

                if ($expandedAllowedGroups.Count -gt 0) {
                    $allowedGroupSql = (
                        $expandedAllowedGroups |
                        ForEach-Object { [string][int]$_ }
                    ) -join ","

                    $accessConditions +=
                        "M.ParentGrp IN ($allowedGroupSql)"
                }

                if ($validAccountCodes.Count -gt 0) {
                    $allowedAccountSql = (
                        $validAccountCodes |
                        ForEach-Object { [string][int]$_ }
                    ) -join ","

                    $accessConditions +=
                        "M.Code IN ($allowedAccountSql)"
                }

                if ($accessConditions.Count -eq 0) {
                    return @{
                        success    = $true
                        total      = 0
                        page       = $Page
                        pageSize   = $PageSize
                        totalPages = 1
                        data       = @()
                    }
                }

                $where +=
                    " AND (" +
                    ($accessConditions -join " OR ") +
                    ")"
            }
        }
        elseif ($EnforceGroupAccess) {
            $groupRows = @(
                Get-PartyPermissionGroupRowsDirect `
                    -Connection $conn
            )

            $selectedCodes = @($AllowedGroupCodes)

            if ($AllowAllEligibleRoots) {
                $tree =
                    Get-PartyPermissionTreeInfo `
                        -Groups $groupRows

                $selectedCodes = @(
                    $tree.roots |
                    ForEach-Object { [int]$_.code }
                )
            }

            $expandedAllowedGroups = @(
                Resolve-AllowedPartyGroupCodes `
                    -Groups $groupRows `
                    -SelectedGroupCodes $selectedCodes
            )

            if ($expandedAllowedGroups.Count -eq 0) {
                return @{
                    success    = $true
                    total      = 0
                    page       = $Page
                    pageSize   = $PageSize
                    totalPages = 1
                    data       = @()
                }
            }

            $allowedGroupSql = (
                $expandedAllowedGroups |
                ForEach-Object { [string][int]$_ }
            ) -join ","

            $where +=
                " AND M.ParentGrp IN ($allowedGroupSql)"
        }

        if ($CashBankOnly) {
            if ($dbType -eq 1) {
                $where += @"
 AND M.ParentGrp IN (
    SELECT Code
    FROM Master1 WITH (NOLOCK)
    WHERE Name IN ('Cash-in-hand', 'Bank Accounts')
)
"@
            }
            else {
                $where += @"
 AND M.ParentGrp IN (
    SELECT Code
    FROM Master1
    WHERE Name IN ('Cash-in-hand', 'Bank Accounts')
)
"@
            }
        }

        # -------------------------------------------------------------
        # Preserve Name/Alias/phone/mobile/TIN search without joining the
        # whole address table. Address matching is done with EXISTS.
        # -------------------------------------------------------------
        if (-not [string]::IsNullOrWhiteSpace($Search)) {
            $safeSearch =
                $Search.Trim().Replace("'", "''")

            $where += @"
 AND (
      M.Name LIKE '$wildcard$safeSearch$wildcard'
      OR M.Alias LIKE '$wildcard$safeSearch$wildcard'
      OR EXISTS (
          SELECT 1
          FROM $addressRef
          WHERE AX.MasterCode = M.Code
            AND (
                 AX.TelNo LIKE '$wildcard$safeSearch$wildcard'
                 OR AX.Mobile LIKE '$wildcard$safeSearch$wildcard'
                 OR AX.TINNo LIKE '$wildcard$safeSearch$wildcard'
            )
      )
 )
"@
        }

        # -------------------------------------------------------------
        # 1) Count Master1 rows only. No full address LEFT JOIN.
        # -------------------------------------------------------------
        $countCmd = $conn.CreateCommand()
        try {
            $countCmd.CommandTimeout = 5
        }
        catch {
        }

        $countCmd.CommandText = @"
SELECT COUNT(*)
FROM $masterRef
WHERE $where
"@

        $totalRecords = 0
        $countValue = $countCmd.ExecuteScalar()

        if (
            $null -ne $countValue -and
            $countValue -ne [System.DBNull]::Value
        ) {
            $totalRecords = [int][string]$countValue
        }

        try { $countCmd.Dispose() } catch {}
        $countCmd = $null

        if ($totalRecords -le 0) {
            $sw.Stop()

            Write-Host (
                "  [PARTIES-FAST-V3] {0}/{1} db={2} rows=0 total=0 elapsedMs={3}" -f
                $InstanceId,
                $CompanyCode,
                [string]$ctx.database,
                [int]$sw.ElapsedMilliseconds
            ) -ForegroundColor DarkCyan

            return @{
                success    = $true
                total      = 0
                page       = $Page
                pageSize   = $PageSize
                totalPages = 1
                data       = @()
            }
        }

        # -------------------------------------------------------------
        # 2) Read ONLY the requested page from Master1.
        # -------------------------------------------------------------
        $startIndex = ($Page - 1) * $PageSize

        $cmd = $conn.CreateCommand()
        try {
            $cmd.CommandTimeout = 5
        }
        catch {
        }

        if ($dbType -eq 1) {
            $cmd.CommandText = @"
SELECT
    M.Code,
    M.Name,
    M.Alias,
    M.ParentGrp AS ParentGroupCode,
    G.Name AS ParentGrpName
FROM $masterRef
LEFT JOIN $groupRef
    ON G.Code = M.ParentGrp
WHERE $where
ORDER BY M.Name, M.Code
OFFSET $startIndex ROWS
FETCH NEXT $PageSize ROWS ONLY
"@
        }
        else {
            # Jet/ACE has no OFFSET/FETCH. Access companies in this bridge are
            # small, so read the filtered account master and slice in PS.
            $cmd.CommandText = @"
SELECT
    M.Code,
    M.Name,
    M.Alias,
    M.ParentGrp AS ParentGroupCode,
    G.Name AS ParentGrpName
FROM $masterRef
LEFT JOIN $groupRef
    ON G.Code = M.ParentGrp
WHERE $where
ORDER BY M.Name, M.Code
"@
        }

        $reader = $cmd.ExecuteReader()

        $rawPageRows = @()
        $currentIndex = 0
        $endIndex = $startIndex + $PageSize - 1

        while ($reader.Read()) {
            if ($dbType -ne 1) {
                if ($currentIndex -lt $startIndex) {
                    $currentIndex++
                    continue
                }

                if ($currentIndex -gt $endIndex) {
                    break
                }
            }

            $readText = {
                param([int]$Index)

                try {
                    if (-not $reader.IsDBNull($Index)) {
                        return (
                            [string]$reader.GetValue($Index)
                        ).Trim()
                    }
                }
                catch {
                }

                return ""
            }

            $readInt = {
                param([int]$Index)

                try {
                    if (-not $reader.IsDBNull($Index)) {
                        return [int][string]$reader.GetValue($Index)
                    }
                }
                catch {
                }

                return 0
            }

            $code = & $readInt 0

            if ($code -gt 0) {
                $rawPageRows += @{
                    code            = $code
                    name            = & $readText 1
                    alias           = & $readText 2
                    parentGroupCode = & $readInt 3
                    group           = & $readText 4
                }
            }

            $currentIndex++

            if (
                $dbType -eq 1 -and
                $rawPageRows.Count -ge $PageSize
            ) {
                break
            }
        }

        try { $reader.Close() } catch {}
        try { $reader.Dispose() } catch {}
        $reader = $null

        try { $cmd.Dispose() } catch {}
        $cmd = $null

        # -------------------------------------------------------------
        # 3) Address information ONLY for page codes (max PageSize rows).
        # -------------------------------------------------------------
        $addressByCode = @{}

        $pageCodes = @(
            $rawPageRows |
            ForEach-Object { [int]$_.code } |
            Where-Object { $_ -gt 0 } |
            Select-Object -Unique
        )

        if ($pageCodes.Count -gt 0) {
            $codeSql = (
                $pageCodes |
                ForEach-Object { [string][int]$_ }
            ) -join ","

            $cmd = $conn.CreateCommand()
            try {
                $cmd.CommandTimeout = 5
            }
            catch {
            }

            if ($dbType -eq 1) {
                $cmd.CommandText = @"
SELECT
    MasterCode,
    Address1,
    Address2,
    Address3,
    Address4,
    TelNo,
    Mobile,
    Email,
    TINNo
FROM MasterAddressInfo WITH (NOLOCK)
WHERE MasterCode IN ($codeSql)
"@
            }
            else {
                $cmd.CommandText = @"
SELECT
    MasterCode,
    Address1,
    Address2,
    Address3,
    Address4,
    TelNo,
    Mobile,
    Email,
    TINNo
FROM MasterAddressInfo
WHERE MasterCode IN ($codeSql)
"@
            }

            try {
                $reader = $cmd.ExecuteReader()

                while ($reader.Read()) {
                    $masterCode = 0

                    try {
                        if (-not $reader.IsDBNull(0)) {
                            $masterCode =
                                [int][string]$reader.GetValue(0)
                        }
                    }
                    catch {
                    }

                    if (
                        $masterCode -le 0 -or
                        $addressByCode.ContainsKey($masterCode)
                    ) {
                        continue
                    }

                    $addressByCode[$masterCode] = @{
                        address1 = if (-not $reader.IsDBNull(1)) {
                            ([string]$reader.GetValue(1)).Trim()
                        } else { "" }

                        address2 = if (-not $reader.IsDBNull(2)) {
                            ([string]$reader.GetValue(2)).Trim()
                        } else { "" }

                        address3 = if (-not $reader.IsDBNull(3)) {
                            ([string]$reader.GetValue(3)).Trim()
                        } else { "" }

                        address4 = if (-not $reader.IsDBNull(4)) {
                            ([string]$reader.GetValue(4)).Trim()
                        } else { "" }

                        telNo = if (-not $reader.IsDBNull(5)) {
                            ([string]$reader.GetValue(5)).Trim()
                        } else { "" }

                        mobileNo = if (-not $reader.IsDBNull(6)) {
                            ([string]$reader.GetValue(6)).Trim()
                        } else { "" }

                        email = if (-not $reader.IsDBNull(7)) {
                            ([string]$reader.GetValue(7)).Trim()
                        } else { "" }

                        taxNo = if (-not $reader.IsDBNull(8)) {
                            ([string]$reader.GetValue(8)).Trim()
                        } else { "" }
                    }
                }
            }
            catch {
                # Preserve the original behavior: address enrichment problems
                # must not prevent the account/party list from loading.
                Write-Host (
                    "  [PARTIES-FAST-V3] address enrichment skipped: {0}" -f
                    $_.Exception.Message
                ) -ForegroundColor DarkYellow
            }
            finally {
                if ($reader) {
                    try { $reader.Close() } catch {}
                    try { $reader.Dispose() } catch {}
                    $reader = $null
                }

                if ($cmd) {
                    try { $cmd.Dispose() } catch {}
                    $cmd = $null
                }
            }
        }

        # -------------------------------------------------------------
        # 4) Build the SAME API response fields.
        # -------------------------------------------------------------
        $parties = @()

        foreach ($row in @($rawPageRows)) {
            $code = [int]$row.code

            $addr = if ($addressByCode.ContainsKey($code)) {
                $addressByCode[$code]
            }
            else {
                @{
                    address1 = ""
                    address2 = ""
                    address3 = ""
                    address4 = ""
                    telNo = ""
                    mobileNo = ""
                    email = ""
                    taxNo = ""
                }
            }

            $addressLines = @(
                $addr.address1,
                $addr.address2,
                $addr.address3,
                $addr.address4
            ) |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace([string]$_)
            }

            $address = @($addressLines) -join ", "

            $phoneParts = @()

            if (
                -not [string]::IsNullOrWhiteSpace(
                    [string]$addr.telNo
                )
            ) {
                $phoneParts += [string]$addr.telNo
            }

            if (
                -not [string]::IsNullOrWhiteSpace(
                    [string]$addr.mobileNo
                ) -and
                [string]$addr.mobileNo -ne [string]$addr.telNo
            ) {
                $phoneParts += [string]$addr.mobileNo
            }

            $phone = $phoneParts -join ", "

            $parentGrp = [string]$row.group
            $name = [string]$row.name

            $partyType = "Other"

            if ($parentGrp -match "Debtor|Customer|Receivable") {
                $partyType = "Customer"
            }
            elseif ($parentGrp -match "Creditor|Supplier|Payable") {
                $partyType = "Supplier"
            }
            elseif (
                $name -eq "Cash" -or
                $parentGrp -match "Cash"
            ) {
                $partyType = "Cash"
            }

            $parties += @{
                code            = $code
                name            = $name
                alias           = [string]$row.alias
                group           = $parentGrp
                parentGroupCode = [int]$row.parentGroupCode
                type            = $partyType

                address  = $address
                address1 = [string]$addr.address1
                address2 = [string]$addr.address2
                address3 = [string]$addr.address3
                address4 = [string]$addr.address4

                telNo    = [string]$addr.telNo
                mobileNo = [string]$addr.mobileNo
                phone    = $phone
                email    = [string]$addr.email

                taxNo    = [string]$addr.taxNo
                vat      = [string]$addr.taxNo
                tinNo    = [string]$addr.taxNo
            }
        }

        $totalPages = [Math]::Ceiling(
            $totalRecords / [double]$PageSize
        )

        if ($totalPages -lt 1) {
            $totalPages = 1
        }

        $sw.Stop()

        Write-Host (
            "  [PARTIES-FAST-V3] {0}/{1} db={2} rows={3} total={4} elapsedMs={5}" -f
            $InstanceId,
            $CompanyCode,
            [string]$ctx.database,
            @($parties).Count,
            $totalRecords,
            [int]$sw.ElapsedMilliseconds
        ) -ForegroundColor DarkCyan

        return @{
            success    = $true
            total      = $totalRecords
            page       = $Page
            pageSize   = $PageSize
            totalPages = $totalPages
            data       = @($parties)
        }
    }
    catch {
        if ($sw.IsRunning) {
            $sw.Stop()
        }

        $message = $_.Exception.Message

        # IMPORTANT: Do not turn a transient SQL read timeout into a 20-30 sec
        # BUSY COM freeze. Return a fast error instead. Access keeps the legacy
        # compatibility fallback for schema/provider-specific problems.
        if (
            $dbType -eq 1 -and
            $message -match "Timeout|timed out|Execution Timeout"
        ) {
            Write-Host (
                "  [PARTIES-FAST-V3 FAIL-NO-COM] {0}/{1} elapsedMs={2} error={3}" -f
                $InstanceId,
                $CompanyCode,
                [int]$sw.ElapsedMilliseconds,
                $message
            ) -ForegroundColor Red

            return @{
                success    = $false
                error      = $message
                total      = 0
                page       = $Page
                pageSize   = $PageSize
                totalPages = 1
                data       = @()
            }
        }

        Write-Host (
            "  [WARN] Fiscal direct party lookup failed; compatibility COM fallback allowed: {0}" -f
            $message
        ) -ForegroundColor DarkYellow

        return $null
    }
    finally {
        if ($reader) {
            try { $reader.Close() } catch {}
            try { $reader.Dispose() } catch {}
        }

        if ($cmd) {
            try { $cmd.Dispose() } catch {}
        }

        if ($countCmd) {
            try { $countCmd.Dispose() } catch {}
        }

        if ($conn) {
            try { $conn.Close() } catch {}
            try { $conn.Dispose() } catch {}
        }
    }
}

function Get-Parties {
    param(
        [string]$Search      = "",
        [bool]$CashBankOnly  = $false,
        [int]$Page           = 1,
        [int]$PageSize       = 30,
        [string]$InstanceId  = "",
        [string]$CompanyCode = "",
        [int[]]$AllowedGroupCodes = @(),
        [bool]$EnforceGroupAccess = $false,
        [bool]$AllowAllEligibleRoots = $false,
        [int[]]$AllowedAccountCodes = @(),
        [bool]$EnforceAllAccountAccess = $false,
        [bool]$AllowAllAccounts = $false
    )

    if ($Page -lt 1) {
        $Page = 1
    }

    if ($PageSize -lt 1) {
        $PageSize = 30
    }

    if (
        $EnforceAllAccountAccess -and
        -not $AllowAllAccounts -and
        ($null -eq $AllowedGroupCodes -or @($AllowedGroupCodes).Count -eq 0) -and
        ($null -eq $AllowedAccountCodes -or @($AllowedAccountCodes).Count -eq 0)
    ) {
        return @{
            success    = $true
            total      = 0
            page       = $Page
            pageSize   = $PageSize
            totalPages = 1
            data       = @()
        }
    }

    if (
        $EnforceGroupAccess -and
        -not $AllowAllEligibleRoots -and
        ($null -eq $AllowedGroupCodes -or @($AllowedGroupCodes).Count -eq 0)
    ) {
        return @{
            success    = $true
            total      = 0
            page       = $Page
            pageSize   = $PageSize
            totalPages = 1
            data       = @()
        }
    }

    # Fast path: direct SQL/OLEDB read. This avoids starting BUSY COM merely
    # to populate an autocomplete/filter list. The original COM implementation
    # remains below as a compatibility fallback.
    $directResult = Get-PartiesDirectDatabase `
        -Search $Search `
        -CashBankOnly $CashBankOnly `
        -Page $Page `
        -PageSize $PageSize `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode `
        -AllowedGroupCodes @($AllowedGroupCodes) `
        -EnforceGroupAccess $EnforceGroupAccess `
        -AllowAllEligibleRoots $AllowAllEligibleRoots `
        -AllowedAccountCodes @($AllowedAccountCodes) `
        -EnforceAllAccountAccess $EnforceAllAccountAccess `
        -AllowAllAccounts $AllowAllAccounts

    if ($null -ne $directResult) {
        return $directResult
    }

    $fi = Connect-BUSY `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $fi) {
        return @{
            success = $false
            error   = "BUSY connection failed"
        }
    }

    try {
        $targetInst = Get-InstanceConfig `
            -InstanceId $InstanceId

        $dbType = 0

        if (
            $null -ne $targetInst -and
            $null -ne $targetInst.dbType
        ) {
            $dbType = [int]$targetInst.dbType
        }

        # SQL Server uses %, Access uses *.
        $wildcard = if ($dbType -eq 1) {
            "%"
        }
        else {
            "*"
        }

        $where = "M.MasterType = 2"

        if ($EnforceAllAccountAccess) {
            if (-not $AllowAllAccounts) {
                $groupRows = Get-PartyPermissionGroupRows -fi $fi
                $expandedAllowedGroups = @(
                    Resolve-AllowedAccountGroupCodes `
                        -Groups $groupRows `
                        -SelectedGroupCodes @($AllowedGroupCodes)
                )

                $validAccountCodes = @(
                    @($AllowedAccountCodes) |
                    ForEach-Object {
                        $parsed = 0
                        if ([int]::TryParse([string]$_, [ref]$parsed) -and $parsed -gt 0) { $parsed }
                    } |
                    Select-Object -Unique
                )

                $accessConditions = @()
                if ($expandedAllowedGroups.Count -gt 0) {
                    $allowedGroupSql = ($expandedAllowedGroups | ForEach-Object { [string][int]$_ }) -join ","
                    $accessConditions += "M.ParentGrp IN ($allowedGroupSql)"
                }
                if ($validAccountCodes.Count -gt 0) {
                    $allowedAccountSql = ($validAccountCodes | ForEach-Object { [string][int]$_ }) -join ","
                    $accessConditions += "M.Code IN ($allowedAccountSql)"
                }

                if ($accessConditions.Count -eq 0) {
                    return @{ success=$true; total=0; page=$Page; pageSize=$PageSize; totalPages=1; data=@() }
                }

                $where += " AND (" + ($accessConditions -join " OR ") + ")"
            }
        }
        elseif ($EnforceGroupAccess) {
            $groupRows = Get-PartyPermissionGroupRows -fi $fi
            $selectedCodes = @($AllowedGroupCodes)

            if ($AllowAllEligibleRoots) {
                $tree = Get-PartyPermissionTreeInfo -Groups $groupRows
                $selectedCodes = @($tree.roots | ForEach-Object { [int]$_.code })
            }

            $expandedAllowedGroups = @(
                Resolve-AllowedPartyGroupCodes `
                    -Groups $groupRows `
                    -SelectedGroupCodes $selectedCodes
            )

            if ($expandedAllowedGroups.Count -eq 0) {
                return @{
                    success    = $true
                    total      = 0
                    page       = $Page
                    pageSize   = $PageSize
                    totalPages = 1
                    data       = @()
                }
            }

            $allowedGroupSql = ($expandedAllowedGroups | ForEach-Object { [string][int]$_ }) -join ","
            $where += " AND M.ParentGrp IN ($allowedGroupSql)"
        }

        if ($CashBankOnly) {
            $where += @"
 AND M.ParentGrp IN (
    SELECT Code
    FROM Master1
    WHERE Name IN ('Cash-in-hand', 'Bank Accounts')
)
"@
        }

        if (-not [string]::IsNullOrWhiteSpace($Search)) {
            $safeSearch =
                $Search.Trim() -replace "'", "''"

            $where += @"
 AND (
    M.Name LIKE '$wildcard$safeSearch$wildcard'
    OR M.Alias LIKE '$wildcard$safeSearch$wildcard'
    OR A.TelNo LIKE '$wildcard$safeSearch$wildcard'
    OR A.Mobile LIKE '$wildcard$safeSearch$wildcard'
    OR A.TINNo LIKE '$wildcard$safeSearch$wildcard'
 )
"@
        }

        # Count distinct accounts because the address table is joined.
        $countQry = @"
SELECT COUNT(*) AS TotalCount
FROM Master1 AS M
LEFT JOIN MasterAddressInfo AS A
ON A.MasterCode = M.Code
WHERE $where
"@

        $countRst =
            $fi.GetRecordset($countQry)

        $totalRecords = 0

        if ($countRst -and -not $countRst.EOF) {
            $countValue =
                $countRst.Fields.Item(
                    "TotalCount"
                ).Value

            if (
                $countValue -ne
                [System.DBNull]::Value
            ) {
                $totalRecords =
                    [int][string]$countValue
            }

            try {
                $countRst.Close()
            }
            catch {}
        }

        $qry = @"
SELECT
    M.Code,
    M.Name,
    M.Alias,
    M.ParentGrp AS ParentGroupCode,

    (
        SELECT G.Name
        FROM Master1 AS G
        WHERE G.Code = M.ParentGrp
    ) AS ParentGrpName,

    A.Address1,
    A.Address2,
    A.Address3,
    A.Address4,
    A.TelNo,
    A.Mobile,
    A.Email,
    A.TINNo

FROM
    Master1 AS M

LEFT JOIN
    MasterAddressInfo AS A
ON
    A.MasterCode = M.Code

WHERE
    $where

ORDER BY
    M.Name
"@

        $rst =
            $fi.GetRecordset($qry)

        $parties = @()

        $startIndex =
            ($Page - 1) * $PageSize

        $endIndex =
            $startIndex + $PageSize - 1

        $currentIndex = 0

        if ($rst -and -not $rst.EOF) {
            $rst.MoveFirst()

            while (-not $rst.EOF) {
                if ($currentIndex -lt $startIndex) {
                    $currentIndex++
                    $rst.MoveNext()
                    continue
                }

                if ($currentIndex -gt $endIndex) {
                    break
                }

                $codeValue =
                    $rst.Fields.Item("Code").Value

                $nameValue =
                    $rst.Fields.Item("Name").Value

                $aliasValue =
                    $rst.Fields.Item("Alias").Value

                $groupValue =
                    $rst.Fields.Item(
                        "ParentGrpName"
                    ).Value

                $parentGroupCodeValue =
                    $rst.Fields.Item(
                        "ParentGroupCode"
                    ).Value

                $address1Value =
                    $rst.Fields.Item(
                        "Address1"
                    ).Value

                $address2Value =
                    $rst.Fields.Item(
                        "Address2"
                    ).Value

                $address3Value =
                    $rst.Fields.Item(
                        "Address3"
                    ).Value

                $address4Value =
                    $rst.Fields.Item(
                        "Address4"
                    ).Value

                $telValue =
                    $rst.Fields.Item(
                        "TelNo"
                    ).Value

                $mobileValue =
                    $rst.Fields.Item(
                        "Mobile"
                    ).Value

                $emailValue =
                    $rst.Fields.Item(
                        "Email"
                    ).Value

                $tinValue =
                    $rst.Fields.Item(
                        "TINNo"
                    ).Value

                $code = if (
                    $codeValue -ne
                    [System.DBNull]::Value
                ) {
                    [int][string]$codeValue
                }
                else {
                    0
                }

                $name = if (
                    $nameValue -ne
                    [System.DBNull]::Value
                ) {
                    [string]$nameValue
                }
                else {
                    ""
                }

                $alias = if (
                    $aliasValue -ne
                    [System.DBNull]::Value
                ) {
                    [string]$aliasValue
                }
                else {
                    ""
                }

                $parentGrp = if (
                    $groupValue -ne
                    [System.DBNull]::Value
                ) {
                    [string]$groupValue
                }
                else {
                    ""
                }

                $parentGroupCode = if (
                    $parentGroupCodeValue -ne
                    [System.DBNull]::Value
                ) {
                    [int][string]$parentGroupCodeValue
                }
                else {
                    0
                }

                $addressLines = @()

                foreach ($addressValue in @(
                    $address1Value,
                    $address2Value,
                    $address3Value,
                    $address4Value
                )) {
                    if (
                        $addressValue -ne
                            [System.DBNull]::Value -and
                        -not [string]::IsNullOrWhiteSpace(
                            [string]$addressValue
                        )
                    ) {
                        $addressLines +=
                            ([string]$addressValue).Trim()
                    }
                }

                $address =
                    $addressLines -join ", "

                $telNo = if (
                    $telValue -ne
                    [System.DBNull]::Value
                ) {
                    ([string]$telValue).Trim()
                }
                else {
                    ""
                }

                $mobileNo = if (
                    $mobileValue -ne
                    [System.DBNull]::Value
                ) {
                    ([string]$mobileValue).Trim()
                }
                else {
                    ""
                }

                $phoneParts = @()

                if (
                    -not [string]::IsNullOrWhiteSpace(
                        $telNo
                    )
                ) {
                    $phoneParts += $telNo
                }

                if (
                    -not [string]::IsNullOrWhiteSpace(
                        $mobileNo
                    ) -and
                    $mobileNo -ne $telNo
                ) {
                    $phoneParts += $mobileNo
                }

                $phone =
                    $phoneParts -join ", "

                $email = if (
                    $emailValue -ne
                    [System.DBNull]::Value
                ) {
                    ([string]$emailValue).Trim()
                }
                else {
                    ""
                }

                $taxNo = if (
                    $tinValue -ne
                    [System.DBNull]::Value
                ) {
                    ([string]$tinValue).Trim()
                }
                else {
                    ""
                }

                $partyType = "Other"

                if (
                    $parentGrp -match
                    "Debtor|Customer|Receivable"
                ) {
                    $partyType = "Customer"
                }
                elseif (
                    $parentGrp -match
                    "Creditor|Supplier|Payable"
                ) {
                    $partyType = "Supplier"
                }
                elseif (
                    $name -eq "Cash" -or
                    $parentGrp -match "Cash"
                ) {
                    $partyType = "Cash"
                }

                $parties += @{
                    code     = $code
                    name     = $name
                    alias           = $alias
                    group           = $parentGrp
                    parentGroupCode = $parentGroupCode
                    type            = $partyType

                    address  = $address
                    address1 = if ($address1Value -ne [System.DBNull]::Value) {
                        [string]$address1Value
                    } else {
                        ""
                    }
                    address2 = if ($address2Value -ne [System.DBNull]::Value) {
                        [string]$address2Value
                    } else {
                        ""
                    }
                    address3 = if ($address3Value -ne [System.DBNull]::Value) {
                        [string]$address3Value
                    } else {
                        ""
                    }
                    address4 = if ($address4Value -ne [System.DBNull]::Value) {
                        [string]$address4Value
                    } else {
                        ""
                    }

                    telNo    = $telNo
                    mobileNo = $mobileNo
                    phone    = $phone
                    email    = $email

                    # VAT field from the account master.
                    taxNo    = $taxNo
                    vat      = $taxNo
                    tinNo    = $taxNo
                }

                $currentIndex++
                $rst.MoveNext()
            }

            try {
                $rst.Close()
            }
            catch {}
        }

        $totalPages =
            [Math]::Ceiling(
                $totalRecords /
                [double]$PageSize
            )

        if ($totalPages -lt 1) {
            $totalPages = 1
        }

        return @{
            success    = $true
            total      = $totalRecords
            page       = $Page
            pageSize   = $PageSize
            totalPages = $totalPages
            data       = @($parties)
        }
    }
    catch {
        return @{
            success = $false
            error   = $_.Exception.Message
        }
    }
    finally {
        Disconnect-BUSY $fi
    }
}

function Get-BusyCloudFastSimpleMasterRows {
    param(
        [int[]]$MasterTypes,
        [hashtable]$KindByMasterType = @{},
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        [string]$LogName = "MASTER"
    )

    $ctx = $null
    $reader = $null
    $startedAt = Get-Date

    try {
        $resolver = Get-Command `
            Get-BusyCloudFastConfigDbContext `
            -ErrorAction SilentlyContinue

        if ($null -eq $resolver) {
            throw (
                "Fast fiscal database resolver is unavailable. " +
                "Make sure the optimized vch_setting.ps1 is installed."
            )
        }

        $ctx = Get-BusyCloudFastConfigDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if ($null -eq $ctx -or $null -eq $ctx.connection) {
            throw "Direct fiscal database connection is unavailable."
        }

        $typeList = @(
            $MasterTypes |
            ForEach-Object { [int]$_ } |
            Where-Object { $_ -gt 0 } |
            Select-Object -Unique
        )

        if ($typeList.Count -eq 0) {
            return @{
                success = $true
                count = 0
                data = @()
                database = [string]$ctx.database
                elapsedMs = 0
            }
        }

        $typeSql = [string]::Join(',', $typeList)

        $cmd = $ctx.connection.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}

        $cmd.CommandText = @"
SELECT
    MasterType,
    Code,
    Name
FROM Master1
WHERE MasterType IN ($typeSql)
ORDER BY MasterType, Name
"@

        $reader = $cmd.ExecuteReader()
        $rows = @()

        while ($reader.Read()) {
            $masterType = 0
            $code = 0
            $name = ""

            try {
                if (-not $reader.IsDBNull(0)) {
                    $masterType = [int]$reader.GetValue(0)
                }
            }
            catch { $masterType = 0 }

            try {
                if (-not $reader.IsDBNull(1)) {
                    $code = [int]$reader.GetValue(1)
                }
            }
            catch { $code = 0 }

            try {
                if (-not $reader.IsDBNull(2)) {
                    $name = ([string]$reader.GetValue(2)).Trim()
                }
            }
            catch { $name = "" }

            if ($code -le 0 -or [string]::IsNullOrWhiteSpace($name)) {
                continue
            }

            $row = @{
                code = $code
                name = $name
            }

            if ($KindByMasterType.ContainsKey($masterType)) {
                $row.kind = [string]$KindByMasterType[$masterType]
            }

            $rows += $row
        }

        $elapsedMs = [int](((Get-Date) - $startedAt).TotalMilliseconds)

        Write-Host (
            "  [{0}-FAST] {1}/{2} db={3} rows={4} elapsedMs={5}" -f
            $LogName,
            $InstanceId,
            $CompanyCode,
            [string]$ctx.database,
            @($rows).Count,
            $elapsedMs
        ) -ForegroundColor DarkCyan

        return @{
            success = $true
            count = @($rows).Count
            data = @($rows)
            database = [string]$ctx.database
            elapsedMs = $elapsedMs
        }
    }
    catch {
        $elapsedMs = [int](((Get-Date) - $startedAt).TotalMilliseconds)

        Write-Host (
            "  [{0}-FAST FAIL] {1}/{2} elapsedMs={3} error={4}" -f
            $LogName,
            $InstanceId,
            $CompanyCode,
            $elapsedMs,
            $_.Exception.Message
        ) -ForegroundColor Red

        return @{
            success = $false
            error = $_.Exception.Message
        }
    }
    finally {
        if ($reader) {
            try { $reader.Close() } catch {}
            try { $reader.Dispose() } catch {}
        }

        if ($ctx -and $ctx.connection) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}

function Get-BillSundries {
    param(
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )

    $cacheKey = "$InstanceId|$CompanyCode|bill-sundries"
    $cached = Get-Cache $cacheKey
    if ($cached) {
        Write-Host (
            "  [BILL-SUNDRIES-FAST] cache HIT {0}/{1} rows={2}" -f
            $InstanceId,
            $CompanyCode,
            @($cached.data).Count
        ) -ForegroundColor DarkCyan
        return $cached
    }

    $direct = Get-BusyCloudFastSimpleMasterRows `
        -MasterTypes @(9) `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode `
        -LogName "BILL-SUNDRIES"

    if (-not $direct.success) {
        return @{ success = $false; error = $direct.error }
    }

    $result = @{
        success = $true
        count = @($direct.data).Count
        data = @($direct.data)
    }

    Set-Cache $cacheKey $result
    return $result
}

function Get-MaterialCenters {
    param(
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )

    $cacheKey = "$InstanceId|$CompanyCode|mat-centers"
    $cached = Get-Cache $cacheKey
    if ($cached) {
        Write-Host (
            "  [MAT-CENTERS-FAST] cache HIT {0}/{1} rows={2}" -f
            $InstanceId,
            $CompanyCode,
            @($cached.data).Count
        ) -ForegroundColor DarkCyan
        return $cached
    }

    $direct = Get-BusyCloudFastSimpleMasterRows `
        -MasterTypes @(11) `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode `
        -LogName "MAT-CENTERS"

    if (-not $direct.success) {
        return @{ success = $false; error = $direct.error }
    }

    $result = @{
        success = $true
        count = @($direct.data).Count
        data = @($direct.data)
    }

    Set-Cache $cacheKey $result
    return $result
}

function Get-VoucherSeries {
    param(
        [int]$VchType = 0,
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )
    $cacheKey = "$InstanceId|$CompanyCode|vch-series"
    $cached = Get-Cache $cacheKey
    if ($cached) { return $cached }
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
    try {
        $series = @()
        try {
            $rst = $fi.GetRecordset("SELECT Name, Code FROM VoucherSeries ORDER BY Name")
            $series = Read-Recordset $rst { param($r) @{ code=[int][string]$r.Fields.Item("Code").Value; name=[string]$r.Fields.Item("Name").Value } }
        } catch { }
        if ($series.Count -eq 0) {
            try {
                $rst2 = $fi.GetRecordset("SELECT Name, Code FROM Master1 WHERE MasterType = 21 ORDER BY Name")
                $series = Read-Recordset $rst2 { param($r) $rawName = [string]$r.Fields.Item("Name").Value; @{ code=[int][string]$r.Fields.Item("Code").Value; name=($rawName -replace '^\d+', ''); originalName=$rawName } }
            } catch { }
        }
        $seen = [System.Collections.Generic.HashSet[string]]::new()
        $series = $series | Where-Object { $seen.Add($_.name) }
        $result = @{ success = $true; count = $series.Count; data = $series }
        if ($series.Count -gt 0) { Set-Cache $cacheKey $result }
        return $result
    } finally { Disconnect-BUSY $fi }
}

function Get-VoucherTypes {
    param(
        [string]$Type = "All",
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )

    $normalizedType = ([string]$Type).Trim()
    if ([string]::IsNullOrWhiteSpace($normalizedType)) {
        $normalizedType = "All"
    }

    $cacheKey = "$InstanceId|$CompanyCode|vch-types|$normalizedType"
    $cached = Get-Cache $cacheKey
    if ($cached) {
        Write-Host (
            "  [VCH-TYPES-FAST] cache HIT {0}/{1} type={2} rows={3}" -f
            $InstanceId,
            $CompanyCode,
            $normalizedType,
            @($cached.data).Count
        ) -ForegroundColor DarkCyan
        return $cached
    }

    $masterTypes = @()
    $kindMap = @{}

    if ($normalizedType -eq "All" -or $normalizedType -eq "Sale") {
        $masterTypes += 13
        $kindMap[13] = "sale"
    }

    if ($normalizedType -eq "All" -or $normalizedType -eq "Purchase") {
        $masterTypes += 14
        $kindMap[14] = "purchase"
    }

    if ($masterTypes.Count -eq 0) {
        return @{
            success = $true
            count = 0
            data = @()
        }
    }

    $direct = Get-BusyCloudFastSimpleMasterRows `
        -MasterTypes $masterTypes `
        -KindByMasterType $kindMap `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode `
        -LogName "VCH-TYPES"

    if (-not $direct.success) {
        return @{ success = $false; error = $direct.error }
    }

    $result = @{
        success = $true
        count = @($direct.data).Count
        data = @($direct.data)
    }

    if (@($direct.data).Count -gt 0) {
        Set-Cache $cacheKey $result
    }

    return $result
}