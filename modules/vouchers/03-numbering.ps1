# 03-numbering.ps1
# Extracted from vouchers.ps1. Keep functions behavior-compatible with the monolith.

function Get-VoucherAutoSequence {
    param(
        $fi,
        [int]$VchType,
        [string]$SeriesName,
        [string]$VchNo,
        [string]$VchDateStr
    )

    if ($null -eq $fi) { return 0 }

    try {
        $prefixStr = "{0:D2}" -f $VchType
        $prefixedSeriesName = $SeriesName
        if (-not $SeriesName.StartsWith($prefixStr)) {
            $prefixedSeriesName = "$prefixStr$SeriesName"
        }
        $safeSeriesName = $SeriesName -replace "'", "''"
        $safePrefixedName = $prefixedSeriesName -replace "'", "''"

        $seriesCode = 0
        $sRst = $fi.GetRecordset("SELECT Code FROM Master1 WHERE MasterType=21 AND (Name='$safeSeriesName' OR Name='$safePrefixedName')")
        if ($sRst -and -not $sRst.EOF) {
            $seriesCode = [int]$sRst.Fields.Item("Code").Value
            $sRst.Close()
        }

        if ($seriesCode -eq 0) { return 0 }

        $isAuto = $false
        $startNo = 1
        $frequency = 0
        $cfgRst = $fi.GetRecordset("SELECT I2, L2 FROM Config WHERE RecType=6 AND L1=$seriesCode")
        if ($cfgRst -and -not $cfgRst.EOF) {
            if ($cfgRst.Fields.Item("L2").Value -ne [System.DBNull]::Value) {
                $l2val = [long]$cfgRst.Fields.Item("L2").Value
                $isAuto = ($l2val -ge 1)
                if ($l2val -gt 1) { $startNo = $l2val }
            }
            if ($cfgRst.Fields.Item("I2").Value -ne [System.DBNull]::Value) {
                $frequency = [int]$cfgRst.Fields.Item("I2").Value
            }
            $cfgRst.Close()
        }

        if (-not $isAuto) { return 0 }

        $existingAutoNo = 0
        $safeNo = $VchNo.Trim()
        $lookupNo = Pad-VchNo $safeNo
        $lookupNoExact = $safeNo -replace "'", "''"
        $lookupNoPadded = $lookupNo -replace "'", "''"

        $eRst = $fi.GetRecordset("SELECT AutoVchNo FROM Tran1 WHERE VchType=$VchType AND (VchNo='$lookupNoExact' OR VchNo='$lookupNoPadded')")
        if ($eRst -and -not $eRst.EOF) {
            if ($eRst.Fields.Item("AutoVchNo").Value -ne [System.DBNull]::Value) {
                $existingAutoNo = [int]$eRst.Fields.Item("AutoVchNo").Value
            }
            $eRst.Close()
        }

        if ($existingAutoNo -gt 0) {
            return $existingAutoNo
        }

        $vchDate = [datetime]::Now
        if ($VchDateStr -match "(\d{2})-(\d{2})-(\d{4})") {
            $vchDate = [datetime]::ParseExact($VchDateStr, "dd-MM-yyyy", $null)
        } elseif ($VchDateStr -match "(\d{4})-(\d{2})-(\d{2})") {
            $vchDate = [datetime]::ParseExact($VchDateStr, "yyyy-MM-dd", $null)
        }

        $dateFilter = ""
        if ($frequency -eq 1) {
            $dateFilter = " AND [Date] = #$($vchDate.ToString('MM/dd/yyyy'))#"
        }
        elseif ($frequency -eq 2) {
            $startDate = New-Object DateTime $vchDate.Year, $vchDate.Month, 1
            $endDate = $startDate.AddMonths(1).AddDays(-1)
            $dateFilter = " AND [Date] >= #$($startDate.ToString('MM/dd/yyyy'))# AND [Date] <= #$($endDate.ToString('MM/dd/yyyy'))#"
        }
        elseif ($frequency -eq 3) {
            $fyStartYear = if ($vchDate.Month -ge 4) { $vchDate.Year } else { $vchDate.Year - 1 }
            $startDate = New-Object DateTime $fyStartYear, 4, 1
            $endDate = $startDate.AddYears(1).AddDays(-1)
            $dateFilter = " AND [Date] >= #$($startDate.ToString('MM/dd/yyyy'))# AND [Date] <= #$($endDate.ToString('MM/dd/yyyy'))#"
        }

        $lastSeq = 0
        $tranRst = $fi.GetRecordset("SELECT AutoVchNo FROM Tran1 WHERE VchSeriesCode=$seriesCode AND Cancelled=0 AND VchCancelled=0" + $dateFilter)
        if ($tranRst -and -not $tranRst.EOF) {
            $tranRst.MoveFirst()
            while (-not $tranRst.EOF) {
                $dbAutoNo = 0
                try {
                    if ($tranRst.Fields.Item("AutoVchNo").Value -ne [System.DBNull]::Value) {
                        $dbAutoNo = [int]$tranRst.Fields.Item("AutoVchNo").Value
                    }
                } catch {}
                if ($dbAutoNo -gt $lastSeq) { $lastSeq = $dbAutoNo }
                $tranRst.MoveNext()
            }
            $tranRst.Close()
        }

        $nextSeq = if ($lastSeq -gt 0) { $lastSeq + 1 } else { $startNo }
        return $nextSeq
    } catch {
        return 0
    }
}

function Get-BusyAutoVchNoFromVoucherNo {
    param(
        $fi,
        [int]$VchType,
        [string]$SeriesName,
        [string]$VchNo
    )

    if ($null -eq $fi) {
        throw "BUSY connection is unavailable."
    }

    if ([string]::IsNullOrWhiteSpace($VchNo)) {
        throw "Voucher number is empty."
    }

    $cleanSeriesName = ([string]$SeriesName).Trim()
    $cleanVchNo = ([string]$VchNo).Trim()

    $typePrefix = "{0:D2}" -f $VchType
    $prefixedSeriesName = $cleanSeriesName

    if (
        -not $cleanSeriesName.StartsWith(
            $typePrefix,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        $prefixedSeriesName = "$typePrefix$cleanSeriesName"
    }

    $safeSeriesName =
        $cleanSeriesName.Replace("'", "''")

    $safePrefixedSeriesName =
        $prefixedSeriesName.Replace("'", "''")

    $seriesCode = 0
    $seriesRst = $null

    try {
        $seriesQuery = @"
SELECT Code
FROM Master1
WHERE MasterType = 21
  AND (
        Name = '$safeSeriesName'
        OR Name = '$safePrefixedSeriesName'
      )
"@

        $seriesRst = $fi.GetRecordset($seriesQuery)

        if ($seriesRst -and -not $seriesRst.EOF) {
            $value =
                $seriesRst.Fields.Item("Code").Value

            if (
                $null -ne $value -and
                $value -ne [System.DBNull]::Value
            ) {
                $seriesCode = [int]$value
            }
        }
    }
    finally {
        if ($seriesRst) {
            try {
                $seriesRst.Close()
            }
            catch {
            }
        }
    }

    if ($seriesCode -le 0) {
        throw "BUSY series '$cleanSeriesName' was not found."
    }

    $prefix = ""
    $suffix = ""
    $separator = ""
    $dateEnabled = 0
    $frequency = 0
    $datePosition = 0
    $dateFormat = 0

    $cfgRst = $null

    try {
        $configQuery = @"
SELECT C1, C2, C4, I1, I2, I7, I8
FROM Config
WHERE RecType = 6
  AND L1 = $seriesCode
"@

        $cfgRst = $fi.GetRecordset($configQuery)

        if ($cfgRst -and -not $cfgRst.EOF) {
            $readString = {
                param([string]$Name)

                $value =
                    $cfgRst.Fields.Item($Name).Value

                if (
                    $null -eq $value -or
                    $value -eq [System.DBNull]::Value
                ) {
                    return ""
                }

                return ([string]$value).Trim()
            }

            $readInt = {
                param([string]$Name)

                $value =
                    $cfgRst.Fields.Item($Name).Value

                if (
                    $null -eq $value -or
                    $value -eq [System.DBNull]::Value
                ) {
                    return 0
                }

                return [int]$value
            }

            $suffix = & $readString "C1"
            $prefix = & $readString "C2"
            $separator = & $readString "C4"

            $dateEnabled = & $readInt "I1"
            $frequency = & $readInt "I2"
            $datePosition = & $readInt "I7"
            $dateFormat = & $readInt "I8"
        }
    }
    finally {
        if ($cfgRst) {
            try {
                $cfgRst.Close()
            }
            catch {
            }
        }
    }

    $working = $cleanVchNo

    if (
        $prefix -ne "" -and
        $working.StartsWith(
            $prefix,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        $working =
            $working.Substring($prefix.Length)
    }

    if (
        $suffix -ne "" -and
        $working.EndsWith(
            $suffix,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        $working =
            $working.Substring(
                0,
                $working.Length - $suffix.Length
            )
    }

    if ($separator -ne "") {
        $parts = $working.Split(
            @($separator),
            [System.StringSplitOptions]::RemoveEmptyEntries
        )

        $numericParts = @(
            $parts |
            Where-Object {
                $_ -match '^\d+$'
            }
        )

        if ($numericParts.Count -gt 0) {
            # In BUSY formats such as 2026-27-8AZQ,
            # the running number is the final numeric component.
            $candidate =
                $numericParts[$numericParts.Count - 1]

            $parsed = 0L

            if (
                [long]::TryParse(
                    $candidate,
                    [ref]$parsed
                ) -and
                $parsed -gt 0
            ) {
                return $parsed
            }
        }
    }

    # Fallback: use the last numeric block before trailing letters.
    #
    # Example:
    # 2026-27-8AZQ -> 8
    # ABC-001XYZ   -> 001
    $match = [regex]::Match(
        $cleanVchNo,
        '(\d+)(?=[A-Za-z_-]*$)'
    )

    if ($match.Success) {
        $parsed = 0L

        if (
            [long]::TryParse(
                $match.Groups[1].Value,
                [ref]$parsed
            ) -and
            $parsed -gt 0
        ) {
            return $parsed
        }
    }

    throw (
        "Could not extract BUSY AutoVchNo from voucher number " +
        "'$cleanVchNo'."
    )
}

# -----------------------------------------------------------------------------
# Direct database helpers for fast read-only numbering configuration
# -----------------------------------------------------------------------------

function Get-NumberingDbValue {
    param(
        $Reader,
        [string]$Field,
        $Default = $null
    )

    try {
        $ordinal = $Reader.GetOrdinal($Field)

        if (
            $ordinal -ge 0 -and
            -not $Reader.IsDBNull($ordinal)
        ) {
            return $Reader.GetValue($ordinal)
        }
    }
    catch {
    }

    return $Default
}

function Get-NumberingDirectDbContext {
    param(
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    # Preferred resolver. This is provided by the optimized vch_setting.ps1
    # and already supports:
    #   SQL    -> active fiscal SQL database
    #   Access -> fiscal BDS file containing Config
    #
    # 03-numbering.ps1 is loaded before vch_setting.ps1, but this function is
    # invoked only after all API modules have finished loading.
    $fastContextCommand = Get-Command `
        Get-BusyCloudFastConfigDbContext `
        -ErrorAction SilentlyContinue

    if ($null -ne $fastContextCommand) {
        return Get-BusyCloudFastConfigDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }

    # Compatibility fallback when the combined helper is not present but the
    # individual fiscal resolvers are available.
    $found = Get-InstanceForCompany `
        -CompanyCode $CompanyCode `
        -InstanceId $InstanceId

    if (-not $found) {
        throw "Company not found in instances.json"
    }

    $inst = $found.instance
    $comp = $found.company

    $dbType = 0
    if ($null -ne $inst.dbType) {
        $dbType = [int]$inst.dbType
    }

    if ($dbType -eq 1) {
        $sqlResolver = Get-Command `
            Resolve-BusyCloudFiscalSqlDatabaseName `
            -ErrorAction SilentlyContinue

        if ($null -eq $sqlResolver) {
            throw (
                "Fast fiscal SQL database resolver is not loaded. " +
                "Use the optimized vch_setting.ps1 before this numbering module."
            )
        }

        $dbName = Resolve-BusyCloudFiscalSqlDatabaseName `
            -Instance $inst `
            -CompanyCode $CompanyCode `
            -InstanceId ([string]$inst.id)

        $conn = Open-SqlConnection `
            -SqlServer $inst.sqlServer `
            -Database $dbName `
            -SqlUser $inst.sqlUser `
            -SqlPassword $inst.sqlPassword

        return @{
            dbType = 1
            connection = $conn
            instance = $inst
            company = $comp
            database = $dbName
        }
    }

    $accessResolver = Get-Command `
        Resolve-BusyCloudFiscalAccessDatabasePath `
        -ErrorAction SilentlyContinue

    if ($null -eq $accessResolver) {
        throw (
            "Fast fiscal Access database resolver is not loaded. " +
            "Use the optimized vch_setting.ps1 before this numbering module."
        )
    }

    $dbFile = Resolve-BusyCloudFiscalAccessDatabasePath `
        -Instance $inst `
        -Company $comp `
        -InstanceId ([string]$inst.id) `
        -CompanyCode ([string]$comp.code)

    $conn = Open-BdsConnection -DbFile $dbFile

    return @{
        dbType = 0
        connection = $conn
        instance = $inst
        company = $comp
        database = $dbFile
    }
}

function Read-NumberingConfigDbRow {
    param(
        $Reader
    )

    $l2Value = 0L

    try {
        $l2Value = [long](
            Get-NumberingDbValue `
                -Reader $Reader `
                -Field "L2" `
                -Default 0
        )
    }
    catch {
        $l2Value = 0L
    }

    $startNo = 1L

    if ($l2Value -gt 1) {
        $startNo = $l2Value
    }

    return @{
        c1 = (
            [string](
                Get-NumberingDbValue `
                    -Reader $Reader `
                    -Field "C1" `
                    -Default ""
            )
        ).Trim()

        c2 = (
            [string](
                Get-NumberingDbValue `
                    -Reader $Reader `
                    -Field "C2" `
                    -Default ""
            )
        ).Trim()

        c3 = (
            [string](
                Get-NumberingDbValue `
                    -Reader $Reader `
                    -Field "C3" `
                    -Default ""
            )
        ).Trim()

        c4 = (
            [string](
                Get-NumberingDbValue `
                    -Reader $Reader `
                    -Field "C4" `
                    -Default ""
            )
        ).Trim()

        i1 = [int](
            Get-NumberingDbValue `
                -Reader $Reader `
                -Field "I1" `
                -Default 0
        )

        frequency = [int](
            Get-NumberingDbValue `
                -Reader $Reader `
                -Field "I2" `
                -Default 0
        )

        paddingLength = [int](
            Get-NumberingDbValue `
                -Reader $Reader `
                -Field "I3" `
                -Default 0
        )

        i6 = [int](
            Get-NumberingDbValue `
                -Reader $Reader `
                -Field "I6" `
                -Default 0
        )

        i7 = [int](
            Get-NumberingDbValue `
                -Reader $Reader `
                -Field "I7" `
                -Default 0
        )

        i8 = [int](
            Get-NumberingDbValue `
                -Reader $Reader `
                -Field "I8" `
                -Default 0
        )

        isAuto = ($l2Value -ge 1)
        startNo = $startNo

        endingNo = [long](
            Get-NumberingDbValue `
                -Reader $Reader `
                -Field "L3" `
                -Default 0
        )
    }
}

function Get-NumberingVisibleSequence {
    param(
        [string]$VoucherNumber,
        [string]$Prefix,
        [string]$Suffix,
        [string]$Separator,
        [string]$DateText
    )

    if ([string]::IsNullOrWhiteSpace($VoucherNumber)) {
        return 0L
    }

    $temp = $VoucherNumber.Trim()

    if (
        -not [string]::IsNullOrEmpty($Prefix) -and
        $temp.StartsWith(
            $Prefix,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        $temp = $temp.Substring($Prefix.Length)
    }

    if (
        -not [string]::IsNullOrEmpty($Suffix) -and
        $temp.EndsWith(
            $Suffix,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        $temp = $temp.Substring(
            0,
            $temp.Length - $Suffix.Length
        )
    }

    if (
        -not [string]::IsNullOrEmpty($DateText) -and
        $temp.IndexOf(
            $DateText,
            [System.StringComparison]::OrdinalIgnoreCase
        ) -ge 0
    ) {
        $temp = [regex]::Replace(
            $temp,
            [regex]::Escape($DateText),
            "",
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        )
    }

    if (-not [string]::IsNullOrEmpty($Separator)) {
        $temp = $temp.Replace($Separator, "")
    }

    $temp = $temp.Trim()

    $sequence = 0L

    if (
        [long]::TryParse(
            $temp,
            [ref]$sequence
        )
    ) {
        return $sequence
    }

    return 0L
}

function Format-NumberingVoucherNo {
    param(
        [long]$Sequence,
        [int]$PaddingLength,
        [char]$PaddingCharacter,
        [string]$Prefix,
        [string]$Suffix,
        [string]$Separator,
        [string]$DateText,
        [string]$EmbedPosition
    )

    $numberText = [string]$Sequence

    if ($PaddingLength -gt 0) {
        $numberText = $numberText.PadLeft(
            $PaddingLength,
            $PaddingCharacter
        )
    }

    switch ($EmbedPosition) {
        "suffix" {
            return (
                "$Prefix$numberText$Separator$DateText"
            ).Trim()
        }

        "prefix" {
            return (
                "$DateText$Separator$numberText$Suffix"
            ).Trim()
        }

        default {
            return (
                "$Prefix$numberText$Suffix"
            ).Trim()
        }
    }
}

# -----------------------------------------------------------------------------
# STABILITY RESTORE: exact Get-NumberingConfig behavior from the proven monolith.
# Uses BUSY CFixedInterface/GetRecordset instead of short-timeout direct SQL.
# -----------------------------------------------------------------------------
function Get-NumberingConfig {
    param(
        [int]$VchType,
        [string]$SeriesName,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $fi = $null

    try {
        # Detect database type for correct date syntax.
        $targetInst = Get-InstanceConfig -InstanceId $InstanceId
        $dbType = 0

        if (
            $null -ne $targetInst -and
            $null -ne $targetInst.dbType
        ) {
            $dbType = [int]$targetInst.dbType
        }

        $isSql = ($dbType -eq 1)

        $fi = Connect-BUSY `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (-not $fi) {
            return @{
                success = $false
                error   = "BUSY connection failed"
            }
        }

        if ([string]::IsNullOrWhiteSpace($SeriesName)) {
            return @{
                success = $false
                error   = "Series name is required"
            }
        }

        $SeriesName = $SeriesName.Trim()

        # ------------------------------------------------------------
        # Resolve voucher-series code
        # ------------------------------------------------------------

        $prefixStr = "{0:D2}" -f $VchType
        $prefixedSeriesName = $SeriesName

        if (
            -not $SeriesName.StartsWith(
                $prefixStr,
                [System.StringComparison]::OrdinalIgnoreCase
            )
        ) {
            $prefixedSeriesName = "$prefixStr$SeriesName"
        }

        $safeSeriesName =
            $SeriesName.Replace("'", "''")

        $safePrefixedName =
            $prefixedSeriesName.Replace("'", "''")

        $seriesCode = 0
        $sRst = $null

        try {
            $seriesQuery = @"
SELECT Code
FROM Master1
WHERE MasterType = 21
  AND (
        Name = '$safeSeriesName'
        OR Name = '$safePrefixedName'
      )
"@

            $sRst = $fi.GetRecordset($seriesQuery)

            if ($sRst -and -not $sRst.EOF) {
                $codeValue =
                    $sRst.Fields.Item("Code").Value

                if (
                    $null -ne $codeValue -and
                    $codeValue -ne [System.DBNull]::Value
                ) {
                    $seriesCode = [int]$codeValue
                }
            }
        }
        finally {
            if ($sRst) {
                try {
                    $sRst.Close()
                }
                catch {
                }

                try {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $sRst
                    ) | Out-Null
                }
                catch {
                }
            }
        }

        if ($seriesCode -le 0) {
            return @{
                success = $false
                error   = "Series '$SeriesName' not found"
            }
        }

        # ------------------------------------------------------------
        # Default numbering configuration
        # ------------------------------------------------------------

        $c1 = ""
        $c2 = ""
        $c3 = ""
        $c4 = ""

        $i1 = 0
        $i6 = 0
        $i7 = 0
        $i8 = 0

        $prefix = ""
        $suffix = ""
        $paddingLength = 0
        $isAuto = $false
        $frequency = 0
        $endingNo = 0L
        $startNo = 1L
        $gotConfig = $false

        # Helper for safely reading one numbering-config row.
        function Read-NumberingConfigRow {
            param(
                $Recordset
            )

            $result = @{
                c1            = ""
                c2            = ""
                c3            = ""
                c4            = ""
                i1            = 0
                frequency     = 0
                paddingLength = 0
                i6            = 0
                i7            = 0
                i8            = 0
                isAuto        = $false
                startNo       = 1L
                endingNo      = 0L
            }

            foreach ($fieldName in @("C1", "C2", "C3", "C4")) {
                try {
                    $value =
                        $Recordset.Fields.Item($fieldName).Value

                    if (
                        $null -ne $value -and
                        $value -ne [System.DBNull]::Value
                    ) {
                        $result[$fieldName.ToLower()] =
                            ([string]$value).Trim()
                    }
                }
                catch {
                }
            }

            foreach ($fieldName in @("I1", "I2", "I3", "I6", "I7", "I8")) {
                try {
                    $value =
                        $Recordset.Fields.Item($fieldName).Value

                    if (
                        $null -ne $value -and
                        $value -ne [System.DBNull]::Value
                    ) {
                        switch ($fieldName) {
                            "I1" {
                                $result.i1 = [int]$value
                            }

                            "I2" {
                                $result.frequency = [int]$value
                            }

                            "I3" {
                                $result.paddingLength = [int]$value
                            }

                            "I6" {
                                $result.i6 = [int]$value
                            }

                            "I7" {
                                $result.i7 = [int]$value
                            }

                            "I8" {
                                $result.i8 = [int]$value
                            }
                        }
                    }
                }
                catch {
                }
            }

            try {
                $value =
                    $Recordset.Fields.Item("L2").Value

                if (
                    $null -ne $value -and
                    $value -ne [System.DBNull]::Value
                ) {
                    $l2Value = [long]$value

                    $result.isAuto = ($l2Value -ge 1)

                    if ($l2Value -gt 1) {
                        $result.startNo = $l2Value
                    }
                }
            }
            catch {
            }

            try {
                $value =
                    $Recordset.Fields.Item("L3").Value

                if (
                    $null -ne $value -and
                    $value -ne [System.DBNull]::Value
                ) {
                    $result.endingNo = [long]$value
                }
            }
            catch {
            }

            return $result
        }

        # ------------------------------------------------------------
        # Read primary numbering configuration
        # ------------------------------------------------------------

        $cfgRst = $null

        try {
            $configQuery = @"
SELECT
    C1,
    C2,
    C3,
    C4,
    I1,
    I2,
    I3,
    I6,
    I7,
    I8,
    L2,
    L3
FROM Config
WHERE RecType = 6
  AND L1 = $seriesCode
"@

            $cfgRst = $fi.GetRecordset($configQuery)

            if ($cfgRst -and -not $cfgRst.EOF) {
                $gotConfig = $true
                $configValues =
                    Read-NumberingConfigRow -Recordset $cfgRst
            }
        }
        finally {
            if ($cfgRst) {
                try {
                    $cfgRst.Close()
                }
                catch {
                }

                try {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $cfgRst
                    ) | Out-Null
                }
                catch {
                }
            }
        }

        # ------------------------------------------------------------
        # Fallback numbering configuration
        # ------------------------------------------------------------

        if (-not $gotConfig) {
            $cfgRst2 = $null

            try {
                $fallbackQuery = @"
SELECT
    C1,
    C2,
    C3,
    C4,
    I1,
    I2,
    I3,
    I6,
    I7,
    I8,
    L2,
    L3
FROM Config
WHERE L1 = $seriesCode
  AND (I2 = 2 OR I2 = 3)
  AND L2 = 1
"@

                $cfgRst2 = $fi.GetRecordset($fallbackQuery)

                if ($cfgRst2 -and -not $cfgRst2.EOF) {
                    $gotConfig = $true
                    $configValues =
                        Read-NumberingConfigRow -Recordset $cfgRst2
                }
            }
            finally {
                if ($cfgRst2) {
                    try {
                        $cfgRst2.Close()
                    }
                    catch {
                    }

                    try {
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                            $cfgRst2
                        ) | Out-Null
                    }
                    catch {
                    }
                }
            }
        }

        if (-not $gotConfig) {
            return @{
                success = $false
                error   = "Numbering configuration not found for series '$SeriesName'"
            }
        }

        $c1 = [string]$configValues.c1
        $c2 = [string]$configValues.c2
        $c3 = [string]$configValues.c3
        $c4 = [string]$configValues.c4

        $i1 = [int]$configValues.i1
        $frequency = [int]$configValues.frequency
        $paddingLength = [int]$configValues.paddingLength
        $i6 = [int]$configValues.i6
        $i7 = [int]$configValues.i7
        $i8 = [int]$configValues.i8

        $isAuto = [bool]$configValues.isAuto
        $startNo = [long]$configValues.startNo
        $endingNo = [long]$configValues.endingNo

        # ------------------------------------------------------------
        # Resolve visible voucher-number format
        # ------------------------------------------------------------

        $padChar = if (-not [string]::IsNullOrEmpty($c3)) {
            [char]$c3[0]
        }
        else {
            [char]" "
        }

        $sep = $c4
        $prefix = $c2
        $suffix = $c1
        $embedPos = "none"

        if ($i1 -gt 0) {
            switch ($i7) {
                1 {
                    $embedPos = "prefix"
                }

                2 {
                    $embedPos = "suffix"
                }

                0 {
                    $embedPos = "none"
                }

                default {
                    if ($i6 -eq 2) {
                        $embedPos = "prefix"
                    }
                    elseif ($i6 -eq 1) {
                        $embedPos = "suffix"
                    }
                    elseif ($c1 -ne "") {
                        $embedPos = "prefix"
                    }
                    elseif ($c2 -ne "") {
                        $embedPos = "suffix"
                    }
                    else {
                        $embedPos = "suffix"
                    }
                }
            }
        }

        # ------------------------------------------------------------
        # Build date text used inside visible voucher numbers
        # ------------------------------------------------------------

        $now = Get-Date

        $monthAbbr =
            $now.ToString(
                "MMM",
                [System.Globalization.CultureInfo]::InvariantCulture
            ).ToUpperInvariant()

        $financialYear = if ($now.Month -ge 4) {
            $now.Year
        }
        else {
            $now.Year - 1
        }

        $financialYearNext =
            $financialYear + 1

        $fyShort =
            ($financialYear % 100).ToString("D2")

        $fyNextShort =
            ($financialYearNext % 100).ToString("D2")

        $dateStr = ""

        if ($isAuto -and $i1 -gt 0) {
            if ($frequency -eq 3) {
                $dateStr = if ($i8 -eq 1) {
                    "$fyShort-$fyNextShort"
                }
                else {
                    "$financialYear-$fyNextShort"
                }
            }
            elseif ($frequency -eq 2) {
                switch ($i8) {
                    2 {
                        $dateStr = $now.ToString("MMyy")
                    }

                    3 {
                        $dateStr = $now.ToString("MM-yyyy")
                    }

                    4 {
                        $dateStr = $now.ToString("MMyyyy")
                    }

                    11 {
                        $dateStr =
                            "$monthAbbr-$($now.ToString('yyyy'))"
                    }

                    12 {
                        $dateStr =
                            "$monthAbbr-$($now.ToString('yy'))"
                    }

                    13 {
                        $dateStr = $now.ToString("MM-yyyy")
                    }

                    14 {
                        $dateStr = $now.ToString("MM-yy")
                    }

                    15 {
                        $dateStr = $now.ToString("MMyy")
                    }

                    16 {
                        $dateStr = $now.ToString("yyMM")
                    }

                    default {
                        $dateStr = $now.ToString("MMyy")
                    }
                }
            }
        }

        # ------------------------------------------------------------
        # Build date filter for current frequency period
        # ------------------------------------------------------------

        $dateFilter = ""

        if ($isAuto -and $frequency -gt 0) {
            $periodStart = $null
            $periodEndExclusive = $null

            if ($frequency -eq 1) {
                $periodStart = $now.Date
                $periodEndExclusive =
                    $periodStart.AddDays(1)
            }
            elseif ($frequency -eq 2) {
                $periodStart = New-Object DateTime `
                    $now.Year,
                    $now.Month,
                    1

                $periodEndExclusive =
                    $periodStart.AddMonths(1)
            }
            elseif ($frequency -eq 3) {
                $fyStartYear = if ($now.Month -ge 4) {
                    $now.Year
                }
                else {
                    $now.Year - 1
                }

                $periodStart = New-Object DateTime `
                    $fyStartYear,
                    4,
                    1

                $periodEndExclusive =
                    $periodStart.AddYears(1)
            }

            if (
                $null -ne $periodStart -and
                $null -ne $periodEndExclusive
            ) {
                if ($isSql) {
                    $sqlStart =
                        $periodStart.ToString(
                            "yyyy-MM-ddTHH:mm:ss",
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )

                    $sqlEnd =
                        $periodEndExclusive.ToString(
                            "yyyy-MM-ddTHH:mm:ss",
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )

                    $dateFilter = (
                        " AND [Date] >= '$sqlStart'" +
                        " AND [Date] < '$sqlEnd'"
                    )
                }
                else {
                    $accessStart =
                        $periodStart.ToString(
                            "MM/dd/yyyy",
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )

                    $accessEnd =
                        $periodEndExclusive.ToString(
                            "MM/dd/yyyy",
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )

                    $dateFilter = (
                        " AND [Date] >= #$accessStart#" +
                        " AND [Date] < #$accessEnd#"
                    )
                }
            }
        }

        # ------------------------------------------------------------
        # Find highest visible voucher sequence
        # ------------------------------------------------------------

        $lastVisibleSeq = 0L
        $lastAutoSeq = 0L
        $tranRst = $null

        try {
            $tranQuery = (
                "SELECT AutoVchNo, VchNo " +
                "FROM Tran1 " +
                "WHERE VchType = $VchType " +
                "AND VchSeriesCode = $seriesCode " +
                "AND Cancelled = 0 " +
                "AND VchCancelled = 0" +
                $dateFilter
            )

            $tranRst = $fi.GetRecordset($tranQuery)

            if ($tranRst -and -not $tranRst.EOF) {
                $tranRst.MoveFirst()

                while (-not $tranRst.EOF) {
                    $dbAutoNo = 0L
                    $dbVchNo = ""

                    try {
                        $autoValue =
                            $tranRst.Fields.Item("AutoVchNo").Value

                        if (
                            $null -ne $autoValue -and
                            $autoValue -ne [System.DBNull]::Value
                        ) {
                            $dbAutoNo = [long]$autoValue
                        }
                    }
                    catch {
                    }

                    try {
                        $vchNoValue =
                            $tranRst.Fields.Item("VchNo").Value

                        if (
                            $null -ne $vchNoValue -and
                            $vchNoValue -ne [System.DBNull]::Value
                        ) {
                            $dbVchNo =
                                ([string]$vchNoValue).Trim()
                        }
                    }
                    catch {
                    }

                    if ($dbAutoNo -gt $lastAutoSeq) {
                        $lastAutoSeq = $dbAutoNo
                    }

                    if (
                        -not [string]::IsNullOrWhiteSpace($dbVchNo) -and
                        $isAuto
                    ) {
                        $temp = $dbVchNo.Trim()

                        if (
                            $prefix -ne "" -and
                            $temp.StartsWith(
                                $prefix,
                                [System.StringComparison]::OrdinalIgnoreCase
                            )
                        ) {
                            $temp =
                                $temp.Substring($prefix.Length)
                        }

                        if (
                            $suffix -ne "" -and
                            $temp.EndsWith(
                                $suffix,
                                [System.StringComparison]::OrdinalIgnoreCase
                            )
                        ) {
                            $temp =
                                $temp.Substring(
                                    0,
                                    $temp.Length - $suffix.Length
                                )
                        }

                        if (
                            $dateStr -ne "" -and
                            $temp.IndexOf(
                                $dateStr,
                                [System.StringComparison]::OrdinalIgnoreCase
                            ) -ge 0
                        ) {
                            $temp =
                                [regex]::Replace(
                                    $temp,
                                    [regex]::Escape($dateStr),
                                    "",
                                    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
                                )
                        }

                        # Remove the configured separator.
                        if ($sep -ne "") {
                            $temp =
                                $temp.Replace($sep, "")
                        }

                        # Do not remove the padding character.
                        # Leading zeroes are valid and TryParse handles them.
                        $temp = $temp.Trim()

                        $parsedVisible = 0L

                        if (
                            [long]::TryParse(
                                $temp,
                                [ref]$parsedVisible
                            ) -and
                            $parsedVisible -gt $lastVisibleSeq
                        ) {
                            $lastVisibleSeq =
                                $parsedVisible
                        }
                    }
                    elseif (
                        -not [string]::IsNullOrWhiteSpace($dbVchNo) -and
                        -not $isAuto
                    ) {
                        $parsedManual = 0L

                        if (
                            [long]::TryParse(
                                $dbVchNo,
                                [ref]$parsedManual
                            ) -and
                            $parsedManual -gt $lastVisibleSeq
                        ) {
                            $lastVisibleSeq =
                                $parsedManual
                        }
                    }

                    $tranRst.MoveNext()
                }
            }
        }
        finally {
            if ($tranRst) {
                try {
                    $tranRst.Close()
                }
                catch {
                }

                try {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $tranRst
                    ) | Out-Null
                }
                catch {
                }
            }
        }

        # Prefer visible voucher number because AutoVchNo can use a
        # different internal sequence.
        $lastSeq = if ($lastVisibleSeq -gt 0) {
            $lastVisibleSeq
        }
        elseif ($lastAutoSeq -gt 0) {
            $lastAutoSeq
        }
        else {
            0L
        }

        $currentNo = if ($lastSeq -gt 0) {
            $lastSeq + 1
        }
        else {
            $startNo
        }

        if (
            $endingNo -gt 0 -and
            $currentNo -gt $endingNo
        ) {
            return @{
                success = $false
                error   = (
                    "Next voucher number $currentNo exceeds configured " +
                    "ending number $endingNo"
                )
            }
        }

        function Build-LocalVchNo {
            param(
                [long]$Sequence
            )

            $numberText =
                [string]$Sequence

            if ($paddingLength -gt 0) {
                $numberText =
                    $numberText.PadLeft(
                        $paddingLength,
                        $padChar
                    )
            }

            switch ($embedPos) {
                "suffix" {
                    return (
                        "$prefix$numberText$sep$dateStr"
                    ).Trim()
                }

                "prefix" {
                    return (
                        "$dateStr$sep$numberText$suffix"
                    ).Trim()
                }

                default {
                    return (
                        "$prefix$numberText$suffix"
                    ).Trim()
                }
            }
        }

        $nextVchNo = if ($isAuto) {
            Build-LocalVchNo -Sequence $currentNo
        }
        else {
            [string]$currentNo
        }

        $lastVchNo = if ($lastSeq -gt 0) {
            Build-LocalVchNo -Sequence $lastSeq
        }
        else {
            "(none yet)"
        }

        return @{
            success = $true

            data = @{
                vch_type          = $VchType
                series_name       = $SeriesName
                source            = "BUSY"
                prefix            = $prefix
                suffix            = $suffix
                separator         = $sep
                date_text         = $dateStr
                embed_position    = $embedPos
                padding_length    = $paddingLength
                padding_character = [string]$padChar
                start_no          = $startNo
                current_no        = $currentNo
                is_auto           = $isAuto
                frequency         = $frequency
                ending_no         = $endingNo
                next_vch_no       = $nextVchNo
                last_vch_no       = $lastVchNo
            }
        }
    }
    catch {
        return @{
            success = $false
            error   = $_.Exception.Message
        }
    }
    finally {
        if ($fi) {
            try {
                Disconnect-BUSY $fi
            }
            catch {
            }
        }
    }
}

function Test-VoucherNumberExists {
    param(
        $fi,
        [int]$VchType,
        [string]$SeriesName,
        [string]$VchNo
    )

    if ($null -eq $fi) {
        throw "BUSY connection is not available."
    }

    if ([string]::IsNullOrWhiteSpace($SeriesName)) {
        throw "Voucher series is required for duplicate checking."
    }

    if ([string]::IsNullOrWhiteSpace($VchNo)) {
        throw "Voucher number is required for duplicate checking."
    }

    $cleanSeriesName = $SeriesName.Trim()
    $cleanVchNo = $VchNo.Trim()

    try {
        # BUSY sometimes stores the series master name with the voucher
        # type prefix, for example LP-1518 or 09LP-1518.
        $typePrefix = "{0:D2}" -f $VchType
        $prefixedSeriesName = $cleanSeriesName

        if (
            -not $cleanSeriesName.StartsWith(
                $typePrefix,
                [System.StringComparison]::OrdinalIgnoreCase
            )
        ) {
            $prefixedSeriesName = "$typePrefix$cleanSeriesName"
        }

        $safeSeriesName =
            $cleanSeriesName.Replace("'", "''")

        $safePrefixedSeriesName =
            $prefixedSeriesName.Replace("'", "''")

        # ------------------------------------------------------------
        # Resolve the voucher-series code
        # ------------------------------------------------------------

        $seriesCode = 0
        $seriesRst = $null

        try {
            $seriesQuery = @"
SELECT Code
FROM Master1
WHERE MasterType = 21
  AND (
        Name = '$safeSeriesName'
        OR Name = '$safePrefixedSeriesName'
      )
"@

            $seriesRst = $fi.GetRecordset($seriesQuery)

            if ($seriesRst -and -not $seriesRst.EOF) {
                $codeValue =
                    $seriesRst.Fields.Item("Code").Value

                if (
                    $null -ne $codeValue -and
                    $codeValue -ne [System.DBNull]::Value
                ) {
                    $seriesCode = [int]$codeValue
                }
            }
        }
        finally {
            if ($seriesRst) {
                try {
                    $seriesRst.Close()
                }
                catch {
                }

                try {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $seriesRst
                    ) | Out-Null
                }
                catch {
                }
            }
        }

        if ($seriesCode -le 0) {
            throw "Voucher series '$cleanSeriesName' was not found."
        }

        # BUSY can store VchNo as a fixed-width field with leading spaces.
        # LTrim/RTrim is supported through the BUSY recordset provider for
        # both SQL and Access companies.
        $safeVchNo =
            $cleanVchNo.Replace("'", "''")

        $paddedVchNo = $cleanVchNo

        try {
            $paddedVchNo = Pad-VchNo $cleanVchNo
        }
        catch {
            # Continue with the exact number when Pad-VchNo is unavailable.
            $paddedVchNo = $cleanVchNo
        }

        $safePaddedVchNo =
            ([string]$paddedVchNo).Replace("'", "''")

        # ------------------------------------------------------------
        # Check for the same voucher type, series and visible number
        # ------------------------------------------------------------

        $duplicateQuery = @"
SELECT VchCode, VchNo
FROM Tran1
WHERE VchType = $VchType
  AND VchSeriesCode = $seriesCode
  AND Cancelled = 0
  AND VchCancelled = 0
  AND (
        LTRIM(RTRIM(VchNo)) = '$safeVchNo'
        OR VchNo = '$safePaddedVchNo'
      )
"@

        $duplicateRst = $null

        try {
            $duplicateRst =
                $fi.GetRecordset($duplicateQuery)

            return (
                $duplicateRst -and
                -not $duplicateRst.EOF
            )
        }
        finally {
            if ($duplicateRst) {
                try {
                    $duplicateRst.Close()
                }
                catch {
                }

                try {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                        $duplicateRst
                    ) | Out-Null
                }
                catch {
                }
            }
        }
    }
    catch {
        throw (
            "Duplicate voucher check failed: " +
            $_.Exception.Message
        )
    }
}

