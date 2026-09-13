# modules/items.ps1
# Item & Item Group Management (Access & SQL Server Dialect Compatible)

. "$PSScriptRoot\connection.ps1"
. "$PSScriptRoot\utils.ps1"

function ConvertTo-ItemXmlSafe {
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return "" }
    $Value = $Value -replace '&',  '&amp;' -replace '<',  '&lt;' -replace '>',  '&gt;' -replace '"',  '&quot;' -replace "'",  '&apos;'
    return $Value
}

function Clear-ItemCaches {
    param([string]$InstanceId = "", [string]$CompanyCode = "")
    $prefix = "$InstanceId|$CompanyCode|"
    Clear-Cache "${prefix}item-groups"; Clear-Cache "${prefix}items-all"
    if ($null -ne $script:_cache) {
        $keysToRemove = @(); foreach ($k in $script:_cache.Keys) { if ($k -like "${prefix}items|*" -or $k -like "${prefix}item|*") { $keysToRemove += $k } }
        foreach ($k in $keysToRemove) { $script:_cache.Remove($k) }
    }

    # Item-group permissions use a lightweight hierarchy cache so voucher
    # searches do not have to reload the Item Group tree on every keystroke.
    if ($null -ne $script:VoucherItemGroupHierarchyCache) {
        $cacheKey = "$InstanceId|$CompanyCode".ToLowerInvariant()
        if ($script:VoucherItemGroupHierarchyCache.ContainsKey($cacheKey)) {
            $script:VoucherItemGroupHierarchyCache.Remove($cacheKey)
        }
    }
}

function Get-Items {
    param(
        [string]$Category    = "",
        [string]$Search      = "",
        [string]$InstanceId  = "",
        [string]$CompanyCode = "",
        [int]$Page           = 1,
        [int]$PageSize       = 30
    )

    if ($Page -lt 1)     { $Page = 1 }
    if ($PageSize -lt 1) { $PageSize = 30 }

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY connection failed" } }

    try {
        # Dynamically resolve dialect wildcard matching (% for SQL, * for Access)
        $dbType = 0
        if ($script:ActiveConnection -ne $null) {
            $instance = Get-InstanceConfig -InstanceId $script:ActiveInstanceId
            if ($null -ne $instance -and $null -ne $instance.dbType) {
                $dbType = [int]$instance.dbType
            }
        }
        $wildcard = if ($dbType -eq 1) { "%" } else { "*" }

        $where = "MasterType = 6"
        if ($Category -ne "") {
            $safeCat = $Category -replace "'", "''"
            $where += " AND ParentGrp = (SELECT Code FROM Master1 WHERE Name = '$safeCat' AND MasterType = 5)"
        }
        if ($Search -ne "") {
            $safeSearch = $Search -replace "'", "''"
            $where += " AND (Name LIKE '$wildcard$safeSearch$wildcard' OR Alias LIKE '$wildcard$safeSearch$wildcard')"
        }

        $countRst = $fi.GetRecordset("SELECT COUNT(*) AS TotalCount FROM Master1 WHERE $where")
        $totalRecords = 0
        if ($countRst -and -not $countRst.EOF) {
            $totalRecords = [int]$countRst.Fields.Item("TotalCount").Value
            try { $countRst.Close() } catch {}
        }
        $totalPages = [Math]::Max(1, [Math]::Ceiling($totalRecords / $PageSize))

        if ($totalRecords -eq 0) {
            return @{ success = $true; total = 0; page = $Page; pageSize = $PageSize; totalPages = 1; data = @() }
        }

        $offset  = ($Page - 1) * $PageSize
        $pageQry = "SELECT TOP $($offset + $PageSize) Master1.Code FROM Master1 WHERE $where ORDER BY Master1.Name"
        $pageRst = $fi.GetRecordset($pageQry)

        $pageCodes = [System.Collections.Generic.List[int]]::new()
        if ($pageRst -and $pageRst.RecordCount -gt 0) {
            $pageRst.MoveFirst()
            $skip = 0
            while (-not $pageRst.EOF) {
                if ($skip -lt $offset) { $skip++; $pageRst.MoveNext(); continue }
                $cRaw = $pageRst.Fields.Item("Code").Value
                if ($cRaw -ne [System.DBNull]::Value) { $pageCodes.Add([int][string]$cRaw) }
                $pageRst.MoveNext()
            }
            try { $pageRst.Close() } catch {}
        }

        if ($pageCodes.Count -eq 0) {
            return @{ success = $true; total = $totalRecords; page = $Page; pageSize = $PageSize; totalPages = $totalPages; data = @() }
        }

        $inList = $pageCodes -join ","

        $mcNameMap = @{}
        try {
            $mcRst = $fi.GetRecordset("SELECT Code, Name FROM Master1 WHERE MasterType = 11")
            if ($mcRst -and $mcRst.RecordCount -gt 0) {
                $mcRst.MoveFirst()
                while (-not $mcRst.EOF) {
                    $c = $mcRst.Fields.Item("Code").Value
                    $n = $mcRst.Fields.Item("Name").Value
                    if ($c -ne [System.DBNull]::Value) {
                        $mcNameMap[$c.ToString().Trim()] = if ($n -ne [System.DBNull]::Value) { $n.ToString().Trim() } else { "" }
                    }
                    $mcRst.MoveNext()
                }
                try { $mcRst.Close() } catch {}
            }
        } catch {}

        $opStockMap = @{}
        try {
            $opRst = $fi.GetRecordset(
                "SELECT MasterCode1 AS ItemCode, MasterCode2 AS MCCode, D1 AS OpQty FROM Tran4 WHERE RecType = 0 AND MasterCode1 IN ($inList)"
            )
            if ($opRst -and $opRst.RecordCount -gt 0) {
                $opRst.MoveFirst()
                while (-not $opRst.EOF) {
                    $iRaw = $opRst.Fields.Item("ItemCode").Value
                    if ($iRaw -ne [System.DBNull]::Value) {
                        $iCode  = [int][string]$iRaw
                        $mcCode = ""
                        $mcRaw  = $opRst.Fields.Item("MCCode").Value
                        if ($mcRaw -ne [System.DBNull]::Value) { $mcCode = $mcRaw.ToString().Trim() }
                        $mcName = if ($mcNameMap.ContainsKey($mcCode)) { $mcNameMap[$mcCode] } else { "Default" }
                        if ([string]::IsNullOrEmpty($mcName)) { $mcName = "Default" }
                        $qty = 0.0
                        $qRaw = $opRst.Fields.Item("OpQty").Value
                        if ($qRaw -ne [System.DBNull]::Value) {
                            [double]::TryParse($qRaw.ToString(), [System.Globalization.NumberStyles]::Any,
                                [System.Globalization.CultureInfo]::InvariantCulture, [ref]$qty) | Out-Null
                        }
                        if (-not $opStockMap.ContainsKey($iCode)) { $opStockMap[$iCode] = @{ total = 0.0; mcMap = @{} } }
                        $opStockMap[$iCode].total += $qty
                        if (-not $opStockMap[$iCode].mcMap.ContainsKey($mcName)) { $opStockMap[$iCode].mcMap[$mcName] = 0.0 }
                        $opStockMap[$iCode].mcMap[$mcName] += $qty
                    }
                    $opRst.MoveNext()
                }
                try { $opRst.Close() } catch {}
            }
        } catch {}

        $txnMap = @{}
        try {
            $txnRst = $fi.GetRecordset(
                "SELECT MasterCode1 AS ItemCode, MasterCode2 AS MCCode, SUM(Value1) AS NetQty FROM Tran2 WHERE RecType = 2 AND MasterCode1 IN ($inList) GROUP BY MasterCode1, MasterCode2"
            )
            if ($txnRst -and $txnRst.RecordCount -gt 0) {
                $txnRst.MoveFirst()
                while (-not $txnRst.EOF) {
                    $iRaw = $txnRst.Fields.Item("ItemCode").Value
                    if ($iRaw -ne [System.DBNull]::Value) {
                        $iCode  = [int][string]$iRaw
                        $mcCode = ""
                        $mcRaw  = $txnRst.Fields.Item("MCCode").Value
                        if ($mcRaw -ne [System.DBNull]::Value) { $mcCode = $mcRaw.ToString().Trim() }
                        $mcName = if ($mcNameMap.ContainsKey($mcCode)) { $mcNameMap[$mcCode] } else { "Unknown" }
                        if ([string]::IsNullOrEmpty($mcName)) { $mcName = "Unknown" }
                        $qty = 0.0
                        $qRaw = $txnRst.Fields.Item("NetQty").Value
                        if ($qRaw -ne [System.DBNull]::Value) {
                            [double]::TryParse($qRaw.ToString(), [System.Globalization.NumberStyles]::Any,
                                [System.Globalization.CultureInfo]::InvariantCulture, [ref]$qty) | Out-Null
                        }
                        if (-not $txnMap.ContainsKey($iCode)) { $txnMap[$iCode] = @{} }
                        if (-not $txnMap[$iCode].ContainsKey($mcName)) { $txnMap[$iCode][$mcName] = 0.0 }
                        $txnMap[$iCode][$mcName] += $qty
                    }
                    $txnRst.MoveNext()
                }
                try { $txnRst.Close() } catch {}
            }
        } catch {}

        $altUnitCodes = [System.Collections.Generic.HashSet[int]]::new()
        try {
            $auRst = $fi.GetRecordset(
                "SELECT Code FROM Master1 WHERE MasterType = 6 AND CM2 IS NOT NULL AND CM2 <> 0 AND Code IN ($inList)"
            )
            if ($auRst -and $auRst.RecordCount -gt 0) {
                $auRst.MoveFirst()
                while (-not $auRst.EOF) {
                    $cRaw = $auRst.Fields.Item("Code").Value
                    if ($cRaw -ne [System.DBNull]::Value) { [void]$altUnitCodes.Add([int][string]$cRaw) }
                    $auRst.MoveNext()
                }
                try { $auRst.Close() } catch {}
            }
        } catch {}

        $detailQry = "SELECT Master1.Code, Master1.Name, Master1.Alias,
                        Master1.D2 AS MRP,
                        Master1.D3 AS SalePrice,
                        Master1.D4 AS PurchasePrice,
                        (SELECT M1.Name FROM Master1 M1 WHERE M1.Code = Master1.ParentGrp) AS GroupName,
                        (SELECT M1.Name FROM Master1 M1 WHERE M1.Code = Master1.CM1) AS UnitName
                      FROM Master1
                      WHERE Master1.Code IN ($inList)
                      ORDER BY Master1.Name"

        $rst   = $fi.GetRecordset($detailQry)
        $items = [System.Collections.Generic.List[object]]::new()

        if ($rst -and $rst.RecordCount -gt 0) {
            $rst.MoveFirst()
            while (-not $rst.EOF) {
                $codeRaw = $rst.Fields.Item("Code").Value
                $codeOut = [int][string]$codeRaw

                $altUnitReq      = $false
                $altUnit         = ""
                $convFactor      = 1.0
                $convType        = 1
                $salePriceDU     = 0.0
                $purchasePriceDU = 0.0

                if ($altUnitCodes.Contains($codeOut)) {
                    try {
                        $xmlStr = $fi.GetMasterXML($codeOut)
                        if ($xmlStr) {
                            $xmlDoc = [xml]$xmlStr
                            if ([string]$xmlDoc.Item.AltUnitReq -eq "True") {
                                $altUnitReq  = $true
                                $altUnit     = [string]$xmlDoc.Item.AltUnit
                                $convType    = [int]$xmlDoc.Item.ConFactorType
                                $convFactor  = [Convert]::ToDouble(
                                    $xmlDoc.Item.ConversionFactor.ToString(),
                                    [System.Globalization.CultureInfo]::InvariantCulture
                                )
                                if ($xmlDoc.Item.SalePriceDU -and [string]$xmlDoc.Item.SalePriceDU -ne "") {
                                    $salePriceDU = [double]$xmlDoc.Item.SalePriceDU
                                }
                                if ($xmlDoc.Item.PurchasePriceDU -and [string]$xmlDoc.Item.PurchasePriceDU -ne "") {
                                    $purchasePriceDU = [double]$xmlDoc.Item.PurchasePriceDU
                                }
                            }
                        }
                    } catch {}
                }

                $itemMcMap = @{}
                $opTotal   = 0.0
                if ($opStockMap.ContainsKey($codeOut)) {
                    $opTotal = $opStockMap[$codeOut].total
                    foreach ($k in $opStockMap[$codeOut].mcMap.Keys) {
                        $itemMcMap[$k] = $opStockMap[$codeOut].mcMap[$k]
                    }
                }
                $netTxn = 0.0
                if ($txnMap.ContainsKey($codeOut)) {
                    foreach ($mName in $txnMap[$codeOut].Keys) {
                        $q = $txnMap[$codeOut][$mName]
                        $netTxn += $q
                        if (-not $itemMcMap.ContainsKey($mName)) { $itemMcMap[$mName] = 0.0 }
                        $itemMcMap[$mName] += $q
                    }
                }

                $mcStockArr = foreach ($k in $itemMcMap.Keys) {
                    @{ mcName = $k; stock = [Math]::Round($itemMcMap[$k], 3) }
                }

                $mrpOut = 0.0; $spOut = 0.0; $ppOut = 0.0

                $mrpVal = $rst.Fields.Item("MRP").Value
                $spVal  = $rst.Fields.Item("SalePrice").Value
                $ppVal  = $rst.Fields.Item("PurchasePrice").Value

                if ($mrpVal -ne [System.DBNull]::Value) {
                    [double]::TryParse($mrpVal.ToString(), [System.Globalization.NumberStyles]::Any,
                        [System.Globalization.CultureInfo]::InvariantCulture, [ref]$mrpOut) | Out-Null
                }
                if ($spVal -ne [System.DBNull]::Value) {
                    [double]::TryParse($spVal.ToString(), [System.Globalization.NumberStyles]::Any,
                        [System.Globalization.CultureInfo]::InvariantCulture, [ref]$spOut) | Out-Null
                }
                if ($ppVal -ne [System.DBNull]::Value) {
                    [double]::TryParse($ppVal.ToString(), [System.Globalization.NumberStyles]::Any,
                        [System.Globalization.CultureInfo]::InvariantCulture, [ref]$ppOut) | Out-Null
                }

                $nameVal  = $rst.Fields.Item("Name").Value
                $aliasVal = $rst.Fields.Item("Alias").Value
                $grpVal   = $rst.Fields.Item("GroupName").Value
                $unitVal  = $rst.Fields.Item("UnitName").Value

                $items.Add(@{
                    code             = $codeOut
                    name             = if ($nameVal  -ne [System.DBNull]::Value) { $nameVal.ToString()  } else { "" }
                    alias            = if ($aliasVal -ne [System.DBNull]::Value) { $aliasVal.ToString() } else { "" }
                    group            = if ($grpVal   -ne [System.DBNull]::Value) { $grpVal.ToString()   } else { "" }
                    unit             = if ($unitVal  -ne [System.DBNull]::Value) { $unitVal.ToString()  } else { "Pcs." }
                    salePrice        = $spOut
                    mrp              = $mrpOut
                    purchasePrice    = $ppOut
                    stock            = [Math]::Round($opTotal + $netTxn, 3)
                    mcStock          = @($mcStockArr)
                    altUnitReq       = $altUnitReq
                    altUnit          = $altUnit
                    conversionFactor = $convFactor
                    conversionType   = $convType
                    salePriceDU      = $salePriceDU
                    purchasePriceDU  = $purchasePriceDU
                })

                $rst.MoveNext()
            }
            try { $rst.Close() } catch {}
        }

        return @{
            success    = $true
            total      = $totalRecords
            page       = $Page
            pageSize   = $PageSize
            totalPages = $totalPages
            data       = $items.ToArray()
        }

    } catch {
        return @{ success = $false; error = $_.Exception.Message }
    } finally {
        Disconnect-BUSY $fi
    }
}


# ═══════════════════════════════════════════════════════
# VOUCHER ITEM GROUP PERMISSION HIERARCHY
#
# BUSY item groups are MasterType=5 and items are MasterType=6.
# This cache contains ONLY the small Item Group tree (not all items).
# It lets a selected parent group grant every descendant branch without
# loading thousands of Item rows into PowerShell for each search.
# ═══════════════════════════════════════════════════════
if ($null -eq $script:VoucherItemGroupHierarchyCache) {
    $script:VoucherItemGroupHierarchyCache = @{}
}

function Get-VoucherItemGroupHierarchy {
    param(
        $Fi,
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )

    $cacheKey = "$InstanceId|$CompanyCode".ToLowerInvariant()
    $now = [DateTime]::UtcNow

    if ($script:VoucherItemGroupHierarchyCache.ContainsKey($cacheKey)) {
        $cached = $script:VoucherItemGroupHierarchyCache[$cacheKey]

        if (
            $null -ne $cached -and
            $null -ne $cached.expiresAt -and
            $cached.expiresAt -gt $now
        ) {
            return $cached
        }

        $script:VoucherItemGroupHierarchyCache.Remove($cacheKey)
    }

    $validCodes = @{}
    $childrenByParent = @{}

    $rst = $Fi.GetRecordset(@"
SELECT
    Code,
    ParentGrp
FROM Master1
WHERE MasterType = 5
"@)

    if ($rst -and -not $rst.EOF) {
        try { $rst.MoveFirst() } catch {}

        while (-not $rst.EOF) {
            $code = 0
            $parentCode = 0

            try {
                $rawCode = $rst.Fields.Item("Code").Value
                if ($null -ne $rawCode -and $rawCode -ne [System.DBNull]::Value) {
                    $code = [int][string]$rawCode
                }
            }
            catch {}

            try {
                $rawParent = $rst.Fields.Item("ParentGrp").Value
                if ($null -ne $rawParent -and $rawParent -ne [System.DBNull]::Value) {
                    $parentCode = [int][string]$rawParent
                }
            }
            catch {}

            if ($code -gt 0) {
                $validCodes[$code] = $true

                if (-not $childrenByParent.ContainsKey($parentCode)) {
                    $childrenByParent[$parentCode] = [System.Collections.Generic.List[int]]::new()
                }

                $childrenByParent[$parentCode].Add($code)
            }

            $rst.MoveNext()
        }

        try { $rst.Close() } catch {}
    }

    $entry = @{
        validCodes       = $validCodes
        childrenByParent = $childrenByParent
        expiresAt        = $now.AddMinutes(5)
    }

    $script:VoucherItemGroupHierarchyCache[$cacheKey] = $entry
    return $entry
}

function Resolve-VoucherAllowedItemGroupCodes {
    param(
        $Fi,
        [int[]]$AllowedGroupCodes,
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )

    $hierarchy = Get-VoucherItemGroupHierarchy `
        -Fi $Fi `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    $resolved = @{}
    $queue = [System.Collections.Generic.Queue[int]]::new()

    foreach ($rawCode in @($AllowedGroupCodes)) {
        $code = 0

        if (
            [int]::TryParse([string]$rawCode, [ref]$code) -and
            $code -gt 0 -and
            $hierarchy.validCodes.ContainsKey($code) -and
            -not $resolved.ContainsKey($code)
        ) {
            $resolved[$code] = $true
            $queue.Enqueue($code)
        }
    }

    while ($queue.Count -gt 0) {
        $parentCode = $queue.Dequeue()

        if (-not $hierarchy.childrenByParent.ContainsKey($parentCode)) {
            continue
        }

        foreach ($childCode in $hierarchy.childrenByParent[$parentCode]) {
            $child = [int]$childCode

            if (-not $resolved.ContainsKey($child)) {
                $resolved[$child] = $true
                $queue.Enqueue($child)
            }
        }
    }

    return @(
        $resolved.Keys |
        ForEach-Object { [int]$_ } |
        Sort-Object
    )
}

# ═══════════════════════════════════════════════════════
# VOUCHER ITEM SEARCH
# Returns only the information required by the dropdown:
# code, name, total stock, and material-centre stock.
# ═══════════════════════════════════════════════════════

# Direct-DB equivalent of the existing voucher item-group hierarchy loader.
# It deliberately keeps the SAME 5-minute hierarchy TTL as the original logic.
function Get-VoucherItemGroupHierarchyDirect {
    param(
        $Connection,
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )

    $cacheKey = "$InstanceId|$CompanyCode".ToLowerInvariant()
    $now = [DateTime]::UtcNow

    if ($script:VoucherItemGroupHierarchyCache.ContainsKey($cacheKey)) {
        $cached = $script:VoucherItemGroupHierarchyCache[$cacheKey]

        if (
            $null -ne $cached -and
            $null -ne $cached.expiresAt -and
            $cached.expiresAt -gt $now
        ) {
            return $cached
        }

        $script:VoucherItemGroupHierarchyCache.Remove($cacheKey)
    }

    $validCodes = @{}
    $childrenByParent = @{}
    $reader = $null
    $cmd = $null

    try {
        $cmd = $Connection.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}

        $cmd.CommandText = @"
SELECT
    Code,
    ParentGrp
FROM Master1
WHERE MasterType = 5
"@

        $reader = $cmd.ExecuteReader()

        while ($reader.Read()) {
            $code = 0
            $parentCode = 0

            try {
                if (-not $reader.IsDBNull(0)) {
                    $code = [int][string]$reader.GetValue(0)
                }
            }
            catch {}

            try {
                if (-not $reader.IsDBNull(1)) {
                    $parentCode = [int][string]$reader.GetValue(1)
                }
            }
            catch {}

            if ($code -gt 0) {
                $validCodes[$code] = $true

                if (-not $childrenByParent.ContainsKey($parentCode)) {
                    $childrenByParent[$parentCode] =
                        [System.Collections.Generic.List[int]]::new()
                }

                $childrenByParent[$parentCode].Add($code)
            }
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
    }

    $entry = @{
        validCodes       = $validCodes
        childrenByParent = $childrenByParent
        expiresAt        = $now.AddMinutes(5)
    }

    $script:VoucherItemGroupHierarchyCache[$cacheKey] = $entry
    return $entry
}

function Resolve-VoucherAllowedItemGroupCodesDirect {
    param(
        $Connection,
        [int[]]$AllowedGroupCodes,
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )

    $hierarchy = Get-VoucherItemGroupHierarchyDirect `
        -Connection $Connection `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    $resolved = @{}
    $queue = [System.Collections.Generic.Queue[int]]::new()

    foreach ($rawCode in @($AllowedGroupCodes)) {
        $code = 0

        if (
            [int]::TryParse([string]$rawCode, [ref]$code) -and
            $code -gt 0 -and
            $hierarchy.validCodes.ContainsKey($code) -and
            -not $resolved.ContainsKey($code)
        ) {
            $resolved[$code] = $true
            $queue.Enqueue($code)
        }
    }

    while ($queue.Count -gt 0) {
        $parentCode = $queue.Dequeue()

        if (-not $hierarchy.childrenByParent.ContainsKey($parentCode)) {
            continue
        }

        foreach ($childCode in $hierarchy.childrenByParent[$parentCode]) {
            $child = [int]$childCode

            if (-not $resolved.ContainsKey($child)) {
                $resolved[$child] = $true
                $queue.Enqueue($child)
            }
        }
    }

    return @(
        $resolved.Keys |
        ForEach-Object { [int]$_ } |
        Sort-Object
    )
}

function Get-ItemsForVoucher {
    param(
        [string]$Search      = "",
        [string]$InstanceId  = "",
        [string]$CompanyCode = "",
        [int[]]$AllowedGroupCodes = @(),
        [bool]$EnforceGroupAccess = $false,
        [int]$MaxResults = 0
    )

    # Preserve the exact original dropdown limits.
    $limit = if ($MaxResults -gt 0) {
        [Math]::Max(1, [Math]::Min(30, $MaxResults))
    }
    elseif ([string]::IsNullOrWhiteSpace($Search)) {
        30
    }
    else {
        20
    }

    $startedAt = [System.Diagnostics.Stopwatch]::StartNew()
    $ctx = $null
    $reader = $null
    $cmd = $null

    try {
        # PERFORMANCE CHANGE ONLY:
        # Replace Connect-BUSY/GetRecordset with a direct connection to the
        # same active fiscal database. The item/stock SQL below is intentionally
        # kept equivalent to the original function.
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

        $where = "Master1.MasterType = 6"

        if ($EnforceGroupAccess) {
            $resolvedGroupCodes = @(
                Resolve-VoucherAllowedItemGroupCodesDirect `
                    -Connection $conn `
                    -AllowedGroupCodes @($AllowedGroupCodes) `
                    -InstanceId $InstanceId `
                    -CompanyCode $CompanyCode
            )

            if ($resolvedGroupCodes.Count -eq 0) {
                return @{
                    success = $true
                    total   = 0
                    data    = @()
                }
            }

            $groupInList = (
                $resolvedGroupCodes |
                ForEach-Object { [string][int]$_ }
            ) -join ","

            $where += " AND Master1.ParentGrp IN ($groupInList)"
        }

        if (-not [string]::IsNullOrWhiteSpace($Search)) {
            $safeSearch = $Search.Trim() -replace "'", "''"

            $where += @"
 AND (
    Master1.Name LIKE '$wildcard$safeSearch$wildcard'
    OR Master1.Alias LIKE '$wildcard$safeSearch$wildcard'
 )
"@
        }

        # Same item query as the original function.
        $itemQuery = @"
SELECT TOP $limit
    Master1.Code,
    Master1.Name,
    Master1.Alias
FROM Master1
WHERE $where
ORDER BY Master1.Name
"@

        $cmd = $conn.CreateCommand()
        try { $cmd.CommandTimeout = 5 } catch {}
        $cmd.CommandText = $itemQuery
        $reader = $cmd.ExecuteReader()

        $items = [System.Collections.Generic.List[object]]::new()

        while (-not $reader.IsClosed -and $reader.Read() -and $items.Count -lt $limit) {
            $codeRaw = $reader.GetValue(0)
            $nameRaw = $reader.GetValue(1)
            $aliasRaw = $reader.GetValue(2)

            if ($codeRaw -ne [System.DBNull]::Value) {
                $items.Add(@{
                    code = [int][string]$codeRaw
                    name = if ($nameRaw -ne [System.DBNull]::Value) {
                        $nameRaw.ToString()
                    } else {
                        ""
                    }
                    alias = if ($aliasRaw -ne [System.DBNull]::Value) {
                        $aliasRaw.ToString()
                    } else {
                        ""
                    }
                    stock   = 0.0
                    mcStock = @()
                })
            }
        }

        try { $reader.Close() } catch {}
        try { $reader.Dispose() } catch {}
        $reader = $null
        try { $cmd.Dispose() } catch {}
        $cmd = $null

        if ($items.Count -eq 0) {
            return @{
                success = $true
                total   = 0
                data    = @()
            }
        }

        $itemCodes = $items | ForEach-Object { [int]$_.code }
        $inList = $itemCodes -join ","

        # Same material-centre lookup and same error behavior as the original:
        # failure here does NOT fail the item dropdown.
        $mcNameMap = @{}

        try {
            $cmd = $conn.CreateCommand()
            try { $cmd.CommandTimeout = 5 } catch {}
            $cmd.CommandText =
                "SELECT Code, Name FROM Master1 WHERE MasterType = 11"
            $reader = $cmd.ExecuteReader()

            while ($reader.Read()) {
                $mcCodeRaw = $reader.GetValue(0)
                $mcNameRaw = $reader.GetValue(1)

                if ($mcCodeRaw -ne [System.DBNull]::Value) {
                    $mcCode = $mcCodeRaw.ToString().Trim()
                    $mcNameMap[$mcCode] = if (
                        $mcNameRaw -ne [System.DBNull]::Value
                    ) {
                        $mcNameRaw.ToString().Trim()
                    } else {
                        ""
                    }
                }
            }
        }
        catch {
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

        $stockMap = @{}

        foreach ($code in $itemCodes) {
            $stockMap[$code] = @{
                total = 0.0
                byMc  = @{}
            }
        }

        # Same opening-stock query and same swallowed-error behavior.
        try {
            $openingQuery = @"
SELECT
    MasterCode1 AS ItemCode,
    MasterCode2 AS MCCode,
    D1 AS Quantity
FROM Tran4
WHERE RecType = 0
  AND MasterCode1 IN ($inList)
"@

            $cmd = $conn.CreateCommand()
            try { $cmd.CommandTimeout = 5 } catch {}
            $cmd.CommandText = $openingQuery
            $reader = $cmd.ExecuteReader()

            while ($reader.Read()) {
                $itemCodeRaw = $reader.GetValue(0)

                if ($itemCodeRaw -ne [System.DBNull]::Value) {
                    $itemCode = [int][string]$itemCodeRaw

                    $mcCode = ""
                    $mcCodeRaw = $reader.GetValue(1)
                    if ($mcCodeRaw -ne [System.DBNull]::Value) {
                        $mcCode = $mcCodeRaw.ToString().Trim()
                    }

                    $mcName = if ($mcNameMap.ContainsKey($mcCode)) {
                        $mcNameMap[$mcCode]
                    } else {
                        "Default"
                    }

                    if ([string]::IsNullOrWhiteSpace($mcName)) {
                        $mcName = "Default"
                    }

                    $quantity = 0.0
                    $quantityRaw = $reader.GetValue(2)

                    if ($quantityRaw -ne [System.DBNull]::Value) {
                        [double]::TryParse(
                            $quantityRaw.ToString(),
                            [System.Globalization.NumberStyles]::Any,
                            [System.Globalization.CultureInfo]::InvariantCulture,
                            [ref]$quantity
                        ) | Out-Null
                    }

                    if ($stockMap.ContainsKey($itemCode)) {
                        $stockMap[$itemCode].total += $quantity

                        if (
                            -not $stockMap[$itemCode].byMc.ContainsKey($mcName)
                        ) {
                            $stockMap[$itemCode].byMc[$mcName] = 0.0
                        }

                        $stockMap[$itemCode].byMc[$mcName] += $quantity
                    }
                }
            }
        }
        catch {
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

        # Same transaction-stock query and same swallowed-error behavior.
        try {
            $transactionQuery = @"
SELECT
    MasterCode1 AS ItemCode,
    MasterCode2 AS MCCode,
    SUM(Value1) AS Quantity
FROM Tran2
WHERE RecType = 2
  AND MasterCode1 IN ($inList)
GROUP BY
    MasterCode1,
    MasterCode2
"@

            $cmd = $conn.CreateCommand()
            try { $cmd.CommandTimeout = 5 } catch {}
            $cmd.CommandText = $transactionQuery
            $reader = $cmd.ExecuteReader()

            while ($reader.Read()) {
                $itemCodeRaw = $reader.GetValue(0)

                if ($itemCodeRaw -ne [System.DBNull]::Value) {
                    $itemCode = [int][string]$itemCodeRaw

                    $mcCode = ""
                    $mcCodeRaw = $reader.GetValue(1)
                    if ($mcCodeRaw -ne [System.DBNull]::Value) {
                        $mcCode = $mcCodeRaw.ToString().Trim()
                    }

                    $mcName = if ($mcNameMap.ContainsKey($mcCode)) {
                        $mcNameMap[$mcCode]
                    } else {
                        "Unknown"
                    }

                    if ([string]::IsNullOrWhiteSpace($mcName)) {
                        $mcName = "Unknown"
                    }

                    $quantity = 0.0
                    $quantityRaw = $reader.GetValue(2)

                    if ($quantityRaw -ne [System.DBNull]::Value) {
                        [double]::TryParse(
                            $quantityRaw.ToString(),
                            [System.Globalization.NumberStyles]::Any,
                            [System.Globalization.CultureInfo]::InvariantCulture,
                            [ref]$quantity
                        ) | Out-Null
                    }

                    if ($stockMap.ContainsKey($itemCode)) {
                        $stockMap[$itemCode].total += $quantity

                        if (
                            -not $stockMap[$itemCode].byMc.ContainsKey($mcName)
                        ) {
                            $stockMap[$itemCode].byMc[$mcName] = 0.0
                        }

                        $stockMap[$itemCode].byMc[$mcName] += $quantity
                    }
                }
            }
        }
        catch {
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

        # Exact original response attachment.
        foreach ($item in $items) {
            $itemCode = [int]$item.code
            $mcStock = @()

            if ($stockMap.ContainsKey($itemCode)) {
                foreach ($mcName in $stockMap[$itemCode].byMc.Keys) {
                    $mcStock += @{
                        mcName = $mcName
                        stock  = [Math]::Round(
                            $stockMap[$itemCode].byMc[$mcName],
                            3
                        )
                    }
                }

                $item.stock = [Math]::Round(
                    $stockMap[$itemCode].total,
                    3
                )

                $item.mcStock = @($mcStock)
            }
        }

        $startedAt.Stop()

        Write-Host (
            "  [ITEMS-DIRECT-SAFE] {0}/{1} search='{2}' rows={3} elapsedMs={4}" -f
            $InstanceId,
            $CompanyCode,
            $Search,
            $items.Count,
            [int]$startedAt.ElapsedMilliseconds
        ) -ForegroundColor DarkCyan

        return @{
            success = $true
            total   = $items.Count
            data    = $items.ToArray()
        }
    }
    catch {
        if ($startedAt.IsRunning) {
            $startedAt.Stop()
        }

        Write-Host (
            "  [ITEMS-DIRECT-SAFE FAIL] {0}/{1} search='{2}' elapsedMs={3} error={4}" -f
            $InstanceId,
            $CompanyCode,
            $Search,
            [int]$startedAt.ElapsedMilliseconds,
            $_.Exception.Message
        ) -ForegroundColor Red

        return @{
            success = $false
            error   = $_.Exception.Message
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

        if ($ctx -and $ctx.connection) {
            try { $ctx.connection.Close() } catch {}
            try { $ctx.connection.Dispose() } catch {}
        }
    }
}


function Get-VoucherItemDetail {
    param(
        [int]$Code,
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )

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
        $xmlString = $fi.GetMasterXML($Code)

        if (-not $xmlString) {
            return @{
                success = $false
                error   = "Item not found"
            }
        }

        $xml = [xml]$xmlString

        $altUnitRequired =
            [string]$xml.Item.AltUnitReq -eq "True"

        $conversionFactor = 1.0
        $conversionType = 1

        if ($altUnitRequired) {
            $conversionType =
                [int]$xml.Item.ConFactorType

            $rawFactor =
                [double]$xml.Item.ConversionFactor

            # Keep backend's real factor for voucher calculations.
            $conversionFactor = $rawFactor
        }

        $stock = 0.0
        $mcStockMap = @{}
        $mcCodeNameMap = @{}

        # Load material-centre names.
        try {
            $mcRst = $fi.GetRecordset(
                "SELECT Code, Name FROM Master1 WHERE MasterType = 11"
            )

            if ($mcRst -and $mcRst.RecordCount -gt 0) {
                $mcRst.MoveFirst()

                while (-not $mcRst.EOF) {
                    $mcCodeRaw =
                        $mcRst.Fields.Item("Code").Value

                    $mcNameRaw =
                        $mcRst.Fields.Item("Name").Value

                    if ($mcCodeRaw -ne [System.DBNull]::Value) {
                        $mcCodeNameMap[
                            $mcCodeRaw.ToString().Trim()
                        ] = if (
                            $mcNameRaw -ne [System.DBNull]::Value
                        ) {
                            $mcNameRaw.ToString().Trim()
                        } else {
                            ""
                        }
                    }

                    $mcRst.MoveNext()
                }

                try { $mcRst.Close() } catch {}
            }
        } catch {}

        # Opening stock for the selected item.
        try {
            $openingQuery = @"
SELECT
    MasterCode2 AS MCCode,
    D1 AS Quantity
FROM Tran4
WHERE RecType = 0
  AND MasterCode1 = $Code
"@

            $opRst = $fi.GetRecordset($openingQuery)

            if ($opRst -and $opRst.RecordCount -gt 0) {
                $opRst.MoveFirst()

                while (-not $opRst.EOF) {
                    $mcCode = ""
                    $mcCodeRaw =
                        $opRst.Fields.Item("MCCode").Value

                    if ($mcCodeRaw -ne [System.DBNull]::Value) {
                        $mcCode = $mcCodeRaw.ToString().Trim()
                    }

                    $mcName = if (
                        $mcCodeNameMap.ContainsKey($mcCode)
                    ) {
                        $mcCodeNameMap[$mcCode]
                    } else {
                        "Default"
                    }

                    $quantity = 0.0
                    $quantityRaw =
                        $opRst.Fields.Item("Quantity").Value

                    if ($quantityRaw -ne [System.DBNull]::Value) {
                        [double]::TryParse(
                            $quantityRaw.ToString(),
                            [System.Globalization.NumberStyles]::Any,
                            [System.Globalization.CultureInfo]::InvariantCulture,
                            [ref]$quantity
                        ) | Out-Null
                    }

                    $stock += $quantity

                    if (-not $mcStockMap.ContainsKey($mcName)) {
                        $mcStockMap[$mcName] = 0.0
                    }

                    $mcStockMap[$mcName] += $quantity
                    $opRst.MoveNext()
                }

                try { $opRst.Close() } catch {}
            }
        } catch {}

        # Transactions for only the selected item.
        try {
            $transactionQuery = @"
SELECT
    MasterCode2 AS MCCode,
    SUM(Value1) AS Quantity
FROM Tran2
WHERE RecType = 2
  AND MasterCode1 = $Code
GROUP BY MasterCode2
"@

            $txnRst = $fi.GetRecordset($transactionQuery)

            if ($txnRst -and $txnRst.RecordCount -gt 0) {
                $txnRst.MoveFirst()

                while (-not $txnRst.EOF) {
                    $mcCode = ""
                    $mcCodeRaw =
                        $txnRst.Fields.Item("MCCode").Value

                    if ($mcCodeRaw -ne [System.DBNull]::Value) {
                        $mcCode = $mcCodeRaw.ToString().Trim()
                    }

                    $mcName = if (
                        $mcCodeNameMap.ContainsKey($mcCode)
                    ) {
                        $mcCodeNameMap[$mcCode]
                    } else {
                        "Unknown"
                    }

                    $quantity = 0.0
                    $quantityRaw =
                        $txnRst.Fields.Item("Quantity").Value

                    if ($quantityRaw -ne [System.DBNull]::Value) {
                        [double]::TryParse(
                            $quantityRaw.ToString(),
                            [System.Globalization.NumberStyles]::Any,
                            [System.Globalization.CultureInfo]::InvariantCulture,
                            [ref]$quantity
                        ) | Out-Null
                    }

                    $stock += $quantity

                    if (-not $mcStockMap.ContainsKey($mcName)) {
                        $mcStockMap[$mcName] = 0.0
                    }

                    $mcStockMap[$mcName] += $quantity
                    $txnRst.MoveNext()
                }

                try { $txnRst.Close() } catch {}
            }
        } catch {}

        $mcStock = foreach ($mcName in $mcStockMap.Keys) {
            @{
                mcName = $mcName
                stock  = [Math]::Round(
                    $mcStockMap[$mcName],
                    3
                )
            }
        }

        $data = @{
            code              = $Code
            name              = [string]$xml.Item.Name
            alias             = [string]$xml.Item.Alias
            group             = [string]$xml.Item.ParentGroup
            unit              = [string]$xml.Item.MainUnit

            salePrice         = [double]$xml.Item.SalePrice
            purchasePrice     = [double]$xml.Item.PurchasePrice
            mrp               = [double]$xml.Item.MRP
            minSalePrice      = [double]$xml.Item.MinSalePrice
            salesDiscount     = [double]$xml.Item.SalesDiscount
            purchaseDiscount  = [double]$xml.Item.PurcDiscount

            stock             = [Math]::Round($stock, 3)
            mcStock           = @($mcStock)

            altUnitReq        = $altUnitRequired
            altUnit           = [string]$xml.Item.AltUnit
            conversionFactor  = $conversionFactor
            conversionType    = $conversionType
            salePriceDU       = [double]$xml.Item.SalePriceDU
            purchasePriceDU   = [double]$xml.Item.PurchasePriceDU

            taxCategory       = [string]$xml.Item.TaxCategory
            taxRateLocal      = [double]$xml.Item.TaxRateLocal
        }

        return @{
            success = $true
            data    = $data
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

function Get-ItemDetail {
    param([int]$Code, [string]$InstanceId = "",[string]$CompanyCode = "")
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode; if (-not $fi) { return @{ success = $false; error = "BUSY failed" } }
    try {
        $xmlStr = $fi.GetMasterXML($Code); if (-not $xmlStr) { return @{ success = $false; error = "Not found" } }
        $xml = [xml]$xmlStr

        $altReq = $(if ([string]$xml.Item.AltUnitReq -eq "True") { $true } else { $false })
        $convF = 1.0; $convT = 1
        if ($altReq) {
            $convT = [int]$xml.Item.ConFactorType; $rawF =[double]$xml.Item.ConversionFactor
            $convF = [Math]::Round($(if ($convT -eq 2 -and $rawF -gt 0) { 1 / $rawF } else { $rawF }), 3)
        }

        $mcLvlMap = @{}
        if ($xml.Item.MCWiseCriticalLevel -and $xml.Item.MCWiseCriticalLevel.MCWiseCriticalInfo) {
            foreach ($mci in $xml.Item.MCWiseCriticalLevel.MCWiseCriticalInfo) {
                $mcLvlMap[[string]$mci.MCNameCL] = @{ minL = $mci.MinimumLevelMCWise; minD = $mci.MinimumDaysMCWise; reoL = $mci.ReorderLevelMCWise; reoD = $mci.ReorderDaysMCWise; maxL = $mci.MaximumLevelMCWise; maxD = $mci.MaximumDaysMCWise }
            }
        }

        $mcs = @()
        if ($xml.Item.OPStockDetails) {
            foreach ($d in $xml.Item.OPStockDetails.OPMCStockDetail) {
                $mname = [string]$d.MCName; $lvls = $mcLvlMap[$mname]
                $mcs += @{ name = $mname; opStock = [double]$d.OPStockMainUnit; opAmount = [double]$d.Amount; minimumLevel = $lvls.minL; minimumDays = $lvls.minD; reorderLevel = $lvls.reoL; reorderDays = $lvls.reoD; maximumLevel = $lvls.maxL; maximumDays = $lvls.maxD }
            }
        }

        $data = @{
            code = $Code; name =[string]$xml.Item.Name; printName = [string]$xml.Item.PrintName; alias =[string]$xml.Item.Alias; group = [string]$xml.Item.ParentGroup; unit =[string]$xml.Item.MainUnit
            opStock = [double]$xml.Item.OPStockInMainUnit; opAmount = [double]$xml.Item.OPAmount
            salePrice = [double]$xml.Item.SalePrice; purchasePrice = [double]$xml.Item.PurchasePrice; mrp = [double]$xml.Item.MRP
            minSalePrice = [double]$xml.Item.MinSalePrice; salesDiscount = [double]$xml.Item.SalesDiscount; purcDiscount = [double]$xml.Item.PurcDiscount; stockValPrice =[double]$xml.Item.StockValPrice
            altUnitReq = $altReq; altUnit =[string]$xml.Item.AltUnit; conversionFactor = $convF; conversionType = $convT; salePriceDU =[double]$xml.Item.SalePriceDU; purchasePriceDU = [double]$xml.Item.PurchasePriceDU
            taxCategory = [string]$xml.Item.TaxCategory; taxRateLocal = [double]$xml.Item.TaxRateLocal; taxInclSalePrice = ([string]$xml.Item.TaxInclSalePrice -eq "True"); taxInclPurcPrice = ([string]$xml.Item.TaxInclPurcPrice -eq "True")
            hsnCode = [string]$xml.Item.HSNCode; description = "$($xml.Item.Address.Address1)`n$($xml.Item.Address.Address2)`n$($xml.Item.Address.Address3)".Trim()
            purchaseAccount = [string]$xml.Item.PurchaseAccount; salesAccount = [string]$xml.Item.SalesAccount
            minimumLevel = $xml.Item.MinimumLevel; minimumDays = $xml.Item.MinimumDays; reorderLevel = $xml.Item.ReorderLevel; reorderDays = $xml.Item.ReorderDays; maximumLevel = $xml.Item.MaximumLevel; maximumDays = $xml.Item.MaximumDays
            materialCenters = @($mcs)
        }
        return @{ success = $true; data = $data }
    } catch { return @{ success = $false; error = $_.Exception.Message } } finally { Disconnect-BUSY $fi }
}

function Build-ItemXml {
    param($Data, [string]$OriginalName = "")
    $itemName = ConvertTo-ItemXmlSafe $(if ($OriginalName -ne "") { $OriginalName } else { [string]$Data.name })
    $altReq = $(if ($Data.altUnitReq -eq $true -or [string]$Data.altUnitReq -eq "True") { $true } else { $false })
    $apiF = 1.0; $cType = 1
    if ($altReq) { $cType = [int]$Data.conversionType; $uiF = [double]$Data.conversionFactor; $apiF = $(if ($cType -eq 2 -and $uiF -gt 0) { 1 / $uiF } else { $uiF }) }
    $gStock = [double]$Data.opStock; $gAlt = [Math]::Round($gStock * $apiF, 3)

    $xml = "<Item><Name>$itemName</Name><PrintName>$(ConvertTo-ItemXmlSafe $Data.printName)</PrintName><Alias>$(ConvertTo-ItemXmlSafe $Data.alias)</Alias><ParentGroup>$(ConvertTo-ItemXmlSafe $Data.group)</ParentGroup><MainUnit>$(ConvertTo-ItemXmlSafe $Data.unit)</MainUnit><OPStockInMainUnit>$gStock</OPStockInMainUnit><OPAmount>$([double]$Data.opAmount)</OPAmount>"
    if ($altReq) { $xml += "<AltUnitReq>True</AltUnitReq><AltUnit>$(ConvertTo-ItemXmlSafe $Data.altUnit)</AltUnit><OPStockInAltUnit>$gAlt</OPStockInAltUnit><ConversionFactor>$apiF</ConversionFactor><ConFactorType>$cType</ConFactorType><SalePriceDU>$([double]$Data.salePriceDU)</SalePriceDU><PurchasePriceDU>$([double]$Data.purchasePriceDU)</PurchasePriceDU><DefaultPricesAppliedOnSales>2</DefaultPricesAppliedOnSales><DefaultPricesAppliedOnPurc>2</DefaultPricesAppliedOnPurc>" }
    $xml += "<SalePrice>$([double]$Data.salePrice)</SalePrice><PurchasePrice>$([double]$Data.purchasePrice)</PurchasePrice><MRP>$([double]$Data.mrp)</MRP><MinSalePrice>$([double]$Data.minSalePrice)</MinSalePrice><SalesDiscount>$([double]$Data.salesDiscount)</SalesDiscount><PurcDiscount>$([double]$Data.purcDiscount)</PurcDiscount><StockValPrice>$([double]$Data.stockValPrice)</StockValPrice><PackingUnitName>$(ConvertTo-ItemXmlSafe $Data.unit)</PackingUnitName><ConFactorPU>1</ConFactorPU><StockValMethod>5</StockValMethod><ItemSrNoType>1</ItemSrNoType><OPStockDetails>"
    
    $sr = 1; $sumAlt = 0.0; $mcCount = $Data.materialCenters.Count
    foreach ($mc in $Data.materialCenters) {
        $s = [double]$mc.opStock; $altVal = [Math]::Round($s * $apiF, 3)
        if ($sr -eq $mcCount) { $altVal =[Math]::Round($gAlt - $sumAlt, 3) } else { $sumAlt += $altVal }
        $xml += "<OPMCStockDetail><SrNo>$sr</SrNo><MCName>$($mc.name)</MCName><OPStockMainUnit>$s</OPStockMainUnit><OPStockAltUnit>$altVal</OPStockAltUnit><Amount>$([double]$mc.opAmount)</Amount></OPMCStockDetail>"
        $sr++
    }
    $xml += "</OPStockDetails>"

    if ($Data.minimumLevel -ne $null -and $Data.minimumLevel -ne "") { $xml += "<CriticalLevels>True</CriticalLevels><MinimumLevel>$([double]$Data.minimumLevel)</MinimumLevel><MinimumDays>$([int]$Data.minimumDays)</MinimumDays><ReorderLevel>$([double]$Data.reorderLevel)</ReorderLevel><ReorderDays>$([int]$Data.reorderDays)</ReorderDays><MaximumLevel>$([double]$Data.maximumLevel)</MaximumLevel><MaximumDays>$([int]$Data.maximumDays)</MaximumDays>" }
    $xml += "<MCWiseCriticalLevel>"
    foreach ($mc in $Data.materialCenters) {
        if ($mc.minimumLevel -ne $null -and $mc.minimumLevel -ne "") { $xml += "<MCWiseCriticalInfo><MCNameCL>$($mc.name)</MCNameCL><MinimumLevelMCWise>$([double]$mc.minimumLevel)</MinimumLevelMCWise><MinimumDaysMCWise>$([int]$mc.minimumDays)</MinimumDaysMCWise><ReorderLevelMCWise>$([double]$mc.reorderLevel)</ReorderLevelMCWise><ReorderDaysMCWise>$([int]$mc.reorderDays)</ReorderDaysMCWise><MaximumLevelMCWise>$([double]$mc.maximumLevel)</MaximumLevelMCWise><MaximumDaysMCWise>$([int]$mc.maximumDays)</MaximumDaysMCWise></MCWiseCriticalInfo>" }
    }
    $xml += "</MCWiseCriticalLevel><TaxCategory>$(ConvertTo-ItemXmlSafe $Data.taxCategory)</TaxCategory><TaxRateLocal>$([double]$Data.taxRateLocal)</TaxRateLocal><PercentOfAmount>100</PercentOfAmount><TaxInclSalePrice>$(if($Data.taxInclSalePrice){'True'}else{'False'})</TaxInclSalePrice><TaxInclPurcPrice>$(if($Data.taxInclPurcPrice){'True'}else{'False'})</TaxInclPurcPrice><HSNCode>$(ConvertTo-ItemXmlSafe $Data.hsnCode)</HSNCode><SpecifySalesAcc>True</SpecifySalesAcc><SalesAccount>$(ConvertTo-ItemXmlSafe $Data.salesAccount)</SalesAccount><SpecifyPurcAcc>True</SpecifyPurcAcc><PurchaseAccount>$(ConvertTo-ItemXmlSafe $Data.purchaseAccount)</PurchaseAccount></Item>"
    return $xml
}

function Create-Item {
    param($Data,[string]$InstanceId = "", [string]$CompanyCode = "")
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode; if (-not $fi) { return @{ success = $false; error = "Connect Fail" } }
    try { $xml = Build-ItemXml -Data $Data; $err = ""; $saved = $fi.SaveMasterFromXML(6, $xml,[ref]$err, $false)
        if ($saved) { Clear-ItemCaches -InstanceId $InstanceId -CompanyCode $CompanyCode; return @{ success = $true } }
        return @{ success = $false; error = $(if($err){$err}else{"Save Fail"}) }
    } catch {
        $script:ActiveConnection = $null 
        try {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($fi) | Out-Null } catch {}
        return @{success = $false; error = "Database error. Connection reset."}
    } finally { Disconnect-BUSY $fi }
}

function Update-Item {
    param($Data, [string]$InstanceId = "", [string]$CompanyCode = "")
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode; if (-not $fi) { return @{ success = $false; error = "Connect Fail" } }
    try { $orig = $(if($Data._originalName){$Data._originalName}else{$Data.name}); $xml = Build-ItemXml -Data $Data -OriginalName $orig; $err = ""; $saved = $fi.SaveMasterFromXML(6, $xml,[ref]$err, $true)
        if ($saved) { Clear-ItemCaches -InstanceId $InstanceId -CompanyCode $CompanyCode; return @{ success = $true } }
        return @{ success = $false; error = $(if($err){$err}else{"Update Fail"}) }
    } catch {
        $script:ActiveConnection = $null 
        try {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($fi) | Out-Null } catch {}
        return @{success = $false; error = "Database error. Connection reset."}
    } finally { Disconnect-BUSY $fi }
}

function Get-Units {
    param([string]$InstanceId = "", [string]$CompanyCode = "")
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode; if (-not $fi) { return @{ success = $false; error = "BUSY failed" } }
    try {
        $rst = $fi.GetRecordset("SELECT Name, Code FROM Master1 WHERE MasterType = 8 ORDER BY Name")
        $units = @()
        if ($rst -and $rst.RecordCount -ne 0) {
            $rst.MoveFirst()
            while (-not $rst.EOF) { $units += @{ code = [int][string]$rst.Fields.Item("Code").Value; name = [string]$rst.Fields.Item("Name").Value }; $rst.MoveNext() }
            try { $rst.Close() } catch {} 
        }
        return @{ success = $true; data = @($units) }
    } finally { Disconnect-BUSY $fi }
}

function Get-ItemGroups {
    param([string]$InstanceId = "",[string]$CompanyCode = "")
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode; if (-not $fi) { return @{ success = $false; error = "BUSY failed" } }
    try {
        $rst = $fi.GetRecordset("SELECT M1.Name, M1.Code, M1.Alias, (SELECT P.Name FROM Master1 P WHERE P.Code = M1.ParentGrp) AS ParentName FROM Master1 M1 WHERE M1.MasterType = 5 ORDER BY M1.Name")
        $groups = @()
        if ($rst -and $rst.RecordCount -gt 0) {
            $rst.MoveFirst()
            while (-not $rst.EOF) { 
                $aliasVal = $rst.Fields.Item("Alias").Value
                $parentVal = $rst.Fields.Item("ParentName").Value
                
                $groups += @{ 
                    code = [int][string]$rst.Fields.Item("Code").Value; 
                    name = [string]$rst.Fields.Item("Name").Value;
                    alias = if ($aliasVal -ne [System.DBNull]::Value -and $null -ne $aliasVal) { $aliasVal.ToString() } else { "" };
                    parent = if ($parentVal -ne [System.DBNull]::Value -and $null -ne $parentVal) { $parentVal.ToString() } else { "General" }
                }
                $rst.MoveNext() 
            }
            try { $rst.Close() } catch {} 
        }
        return @{ success = $true; data = @($groups) }
    } finally { Disconnect-BUSY $fi }
}

function Build-ItemGroupXml {
    param($Data, [string]$OriginalName = "")
    $name = if ($OriginalName -ne "") { $OriginalName } else { [string]$Data.name }
    $parent = $Data.parentGroup
    if ([string]::IsNullOrEmpty($parent) -or $parent.ToLower() -eq "general" -or $parent -eq "__GENERAL__") {
        return "<ItemGroup><Name>$(ConvertTo-ItemXmlSafe $name)</Name><Alias>$(ConvertTo-ItemXmlSafe $Data.alias)</Alias><PrimaryGroup>True</PrimaryGroup></ItemGroup>"
    } else {
        return "<ItemGroup><Name>$(ConvertTo-ItemXmlSafe $name)</Name><Alias>$(ConvertTo-ItemXmlSafe $Data.alias)</Alias><ParentGroupName>$(ConvertTo-ItemXmlSafe $parent)</ParentGroupName></ItemGroup>"
    }
}

function Get-TaxCategories {
    param([string]$InstanceId = "", [string]$CompanyCode = "")
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode; if (-not $fi) { return @{ success = $false; error = "BUSY failed" } }
    try {
        $rst = $fi.GetRecordset("SELECT Name, Code FROM Master1 WHERE MasterType = 25 ORDER BY Name")
        $cats = @()
        if ($rst -and $rst.RecordCount -ne 0) {
            $rst.MoveFirst()
            while (-not $rst.EOF) { $cats += @{ code = [int][string]$rst.Fields.Item("Code").Value; name =[string]$rst.Fields.Item("Name").Value }; $rst.MoveNext() }
            try { $rst.Close() } catch {}
        }
        return @{ success = $true; data = @($cats) }
    } finally { Disconnect-BUSY $fi }
}

function Create-ItemGroup {
    param($Data, [string]$InstanceId = "", [string]$CompanyCode = "")
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode; if (-not $fi) { return @{ success = $false; error = "Fail" } }
    try {
        $xml = Build-ItemGroupXml -Data $Data
        $err = ""; $saved = $fi.SaveMasterFromXML(5, $xml,[ref]$err, $false)
        if ($saved) { Clear-ItemCaches -InstanceId $InstanceId -CompanyCode $CompanyCode; return @{ success = $true } }
        return @{ success = $false; error = $err }
    } catch {
        $script:ActiveConnection = $null 
        try {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($fi) | Out-Null } catch {}
        return @{success = $false; error = "Database error. Connection reset."}
    } finally { Disconnect-BUSY $fi }
}

function Update-ItemGroup {
    param($Data, [string]$InstanceId = "", [string]$CompanyCode = "")
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode; if (-not $fi) { return @{ success = $false; error = "Fail" } }
    try {
        $orig = if ($Data._originalName) { $Data._originalName } else { $Data.name }
        $xml = Build-ItemGroupXml -Data $Data -OriginalName $orig
        $err = ""; $saved = $fi.SaveMasterFromXML(5, $xml, [ref]$err, $true)
        if ($saved) { Clear-ItemCaches -InstanceId $InstanceId -CompanyCode $CompanyCode; return @{ success = $true } }
        return @{ success = $false; error = $err }
    } catch {
        $script:ActiveConnection = $null 
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($fi) | Out-Null } catch {}
        return @{success = $false; error = "Database error. Connection reset."}
    } finally { Disconnect-BUSY $fi }
}