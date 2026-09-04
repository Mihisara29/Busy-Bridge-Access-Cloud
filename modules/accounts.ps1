# modules/accounts.ps1
# Account & Master Data Management (PowerShell 5.1 Safe) - Multi-Instance Version

. "$PSScriptRoot\connection.ps1"
. "$PSScriptRoot\utils.ps1"

# ═══════════════════════════════════════════════════════
#  INTERNAL HELPER - Escape XML special characters
# ═══════════════════════════════════════════════════════
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

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
    try {
        $wildcard = "%"
        $where = "Master1.MasterType = 2 AND Master1.ParentGrp IN (SELECT Code FROM Master1 WHERE Name IN ('Cash-in-hand', 'Bank Accounts'))"

        if ($Search -and $Search -ne "") {
            $safeSearch = $Search -replace "'", "''"
            $where += " AND (Master1.Name LIKE '$wildcard$safeSearch$wildcard' OR Master1.Alias LIKE '$wildcard$safeSearch$wildcard')"
        }

        $qry = "SELECT Master1.Code, Master1.Name, Master1.Alias,
                    (SELECT M1.Name FROM Master1 M1 WHERE M1.Code = Master1.ParentGrp) AS ParentGrpName
                FROM Master1
                WHERE $where
                ORDER BY Master1.Name"

        $rst = $fi.GetRecordset($qry)
        $accounts = Read-Recordset $rst {
            param($r)
            $parentGrp = [string]$r.Fields.Item("ParentGrpName").Value
            $kind = if ($parentGrp -eq "Bank Accounts") { "Bank" } else { "Cash" }
            @{
                code  = [int][string]$r.Fields.Item("Code").Value
                name  = [string]$r.Fields.Item("Name").Value
                alias = [string]$r.Fields.Item("Alias").Value
                group = $parentGrp
                type  = $kind
            }
        }

        return @{ success = $true; count = @($accounts).Count; data = @($accounts) }
    } finally {
        Disconnect-BUSY $fi
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
# Only the three exact BUSY roots below are eligible. Similar sibling groups
# (for example "Sundry Debtors-Priyanwadh") are intentionally excluded.

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

    $acceptedRootNames = @(
        "sundry debtor",
        "sundry debtors",
        "sundry creditor",
        "sundry creditors",
        "cash in hand",
        "cash in hands"
    )

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
            $acceptedRootNames -contains (Get-PartyGroupNormalizedName ([string]$_.name))
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

    $fi = Connect-BUSY `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $fi) {
        return @{ success = $false; error = "BUSY connection failed" }
    }

    try {
        $groups = Get-PartyPermissionGroupRows -fi $fi
        $tree = Get-PartyPermissionTreeInfo -Groups $groups
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

        return @{
            success = $true
            count   = $output.Count
            data    = @($output)
        }
    }
    catch {
        return @{ success = $false; error = $_.Exception.Message }
    }
    finally {
        Disconnect-BUSY $fi
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

    $fi = Connect-BUSY `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $fi) {
        return @{ success = $false; error = "BUSY connection failed" }
    }

    try {
        $groups = Get-PartyPermissionGroupRows -fi $fi
        $groupByCode = @{}

        foreach ($group in @($groups)) {
            $groupByCode["$([int]$group.code)"] = $group
        }

        $resolveGroupPath = {
            param([int]$StartCode)

            $path = @()
            $visited = @{}
            $currentCode = $StartCode

            for ($guard = 0; $guard -lt 100; $guard++) {
                if ($currentCode -le 0 -or -not $groupByCode.ContainsKey("$currentCode")) {
                    break
                }
                if ($visited.ContainsKey("$currentCode")) {
                    break
                }

                $visited["$currentCode"] = $true
                $group = $groupByCode["$currentCode"]
                $path += $group
                $currentCode = [int]$group.parentGrp
            }

            $ordered = @()
            for ($i = $path.Count - 1; $i -ge 0; $i--) {
                $ordered += $path[$i]
            }
            return @($ordered)
        }

        $nodes = @()

        foreach ($group in @($groups)) {
            $path = @(& $resolveGroupPath ([int]$group.code))
            $root = if ($path.Count -gt 0) { $path[0] } else { $group }
            $pathNames = @($path | ForEach-Object { [string]$_.name })

            $nodes += @{
                rootCode   = [int]$root.code
                rootName   = [string]$root.name
                nodeType   = "GROUP"
                code       = [int]$group.code
                name       = [string]$group.name
                alias      = ""
                parentCode = [int]$group.parentGrp
                level      = [Math]::Max(0, $path.Count - 1)
                pathText   = ($pathNames -join " > ")
            }
        }

        $qry = @"
SELECT
    Code,
    Name,
    Alias,
    ParentGrp
FROM Master1
WHERE MasterType = 2
ORDER BY Name
"@

        $rst = $fi.GetRecordset($qry)

        if ($rst -and -not $rst.EOF) {
            try { $rst.MoveFirst() } catch {}

            while (-not $rst.EOF) {
                $codeValue = $rst.Fields.Item("Code").Value
                $nameValue = $rst.Fields.Item("Name").Value
                $aliasValue = $rst.Fields.Item("Alias").Value
                $parentValue = $rst.Fields.Item("ParentGrp").Value

                $code = if ($codeValue -ne [System.DBNull]::Value) { [int][string]$codeValue } else { 0 }
                $name = if ($nameValue -ne [System.DBNull]::Value) { ([string]$nameValue).Trim() } else { "" }
                $alias = if ($aliasValue -ne [System.DBNull]::Value) { ([string]$aliasValue).Trim() } else { "" }
                $parentCode = if ($parentValue -ne [System.DBNull]::Value) { [int][string]$parentValue } else { 0 }

                if ($code -gt 0) {
                    $groupPath = @(& $resolveGroupPath $parentCode)
                    $root = if ($groupPath.Count -gt 0) { $groupPath[0] } else { $null }
                    $pathNames = @($groupPath | ForEach-Object { [string]$_.name }) + @($name)

                    $nodes += @{
                        rootCode   = if ($null -ne $root) { [int]$root.code } else { $parentCode }
                        rootName   = if ($null -ne $root) { [string]$root.name } else { "" }
                        nodeType   = "ACCOUNT"
                        code       = $code
                        name       = $name
                        alias      = $alias
                        parentCode = $parentCode
                        level      = $groupPath.Count
                        pathText   = ($pathNames -join " > ")
                    }
                }

                $rst.MoveNext()
            }

            try { $rst.Close() } catch {}
        }

        $nodes = @($nodes | Sort-Object pathText, nodeType, name)

        return @{
            success = $true
            count   = $nodes.Count
            data    = @($nodes)
        }
    }
    catch {
        return @{ success = $false; error = $_.Exception.Message }
    }
    finally {
        Disconnect-BUSY $fi
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

function Get-BillSundries {
    param(
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )
    $cacheKey = "$InstanceId|$CompanyCode|bill-sundries"
    $cached = Get-Cache $cacheKey
    if ($cached) { return $cached }
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
    try {
        $qry = "SELECT Name, Code FROM Master1 WHERE MasterType = 9 ORDER BY Name"
        $rst = $fi.GetRecordset($qry)
        $bs  = Read-Recordset $rst { param($r) @{ code=[int][string]$r.Fields.Item("Code").Value; name=[string]$r.Fields.Item("Name").Value } }
        $result = @{ success = $true; count = $bs.Count; data = $bs }
        Set-Cache $cacheKey $result
        return $result
    } finally { Disconnect-BUSY $fi }
}

function Get-MaterialCenters {
    param(
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )
    $cacheKey = "$InstanceId|$CompanyCode|mat-centers"
    $cached = Get-Cache $cacheKey
    if ($cached) { return $cached }
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
    try {
        $qry = "SELECT Name, Code FROM Master1 WHERE MasterType = 11 ORDER BY Name"
        $rst = $fi.GetRecordset($qry)
        $centers = Read-Recordset $rst { param($r) @{ code=[int][string]$r.Fields.Item("Code").Value; name=[string]$r.Fields.Item("Name").Value } }
        $result = @{ success = $true; count = $centers.Count; data = $centers }
        Set-Cache $cacheKey $result
        return $result
    } finally { Disconnect-BUSY $fi }
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
    $cacheKey = "$InstanceId|$CompanyCode|vch-types|$Type"
    $cached = Get-Cache $cacheKey
    if ($cached) { return $cached }
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }
    try {
        $types = @()
        if ($Type -eq "All" -or $Type -eq "Sale") {
            try { $types += Read-Recordset ($fi.GetRecordset("SELECT Name, Code FROM Master1 WHERE MasterType = 13 ORDER BY Name")) { param($r) @{ code=[int][string]$r.Fields.Item("Code").Value; name=[string]$r.Fields.Item("Name").Value; kind="sale" } } } catch {}
        }
        if ($Type -eq "All" -or $Type -eq "Purchase") {
            try { $types += Read-Recordset ($fi.GetRecordset("SELECT Name, Code FROM Master1 WHERE MasterType = 14 ORDER BY Name")) { param($r) @{ code=[int][string]$r.Fields.Item("Code").Value; name=[string]$r.Fields.Item("Name").Value; kind="purchase" } } } catch {}
        }
        $result = @{ success = $true; count = $types.Count; data = $types }
        if ($types.Count -gt 0) { Set-Cache $cacheKey $result }
        return $result
    } finally { Disconnect-BUSY $fi }
}