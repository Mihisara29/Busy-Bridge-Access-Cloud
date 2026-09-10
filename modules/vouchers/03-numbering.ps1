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

function Get-NumberingConfig {
    param(
        [int]$VchType,
        [string]$SeriesName,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $startedAt = Get-Date
    $ctx = $null
    $reader = $null

    try {
        # ------------------------------------------------------------
        # Validate request
        # ------------------------------------------------------------

        if ($VchType -le 0) {
            return @{
                success = $false
                error = "Voucher type is required"
            }
        }

        if ([string]::IsNullOrWhiteSpace($SeriesName)) {
            return @{
                success = $false
                error = "Series name is required"
            }
        }

        $SeriesName = $SeriesName.Trim()

        # ------------------------------------------------------------
        # Open DIRECT fiscal database connection
        #
        # IMPORTANT:
        # No Connect-BUSY / OpenCSDB / BUSY COM is used here.
        # ------------------------------------------------------------

        $ctx = Get-NumberingDirectDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (
            $null -eq $ctx -or
            $null -eq $ctx.connection
        ) {
            throw "Direct fiscal database connection is unavailable."
        }

        $conn = $ctx.connection
        $isSql = ([int]$ctx.dbType -eq 1)

        # ------------------------------------------------------------
        # Resolve voucher-series code
        # ------------------------------------------------------------

        $typePrefix = "{0:D2}" -f $VchType
        $prefixedSeriesName = $SeriesName

        if (
            -not $SeriesName.StartsWith(
                $typePrefix,
                [System.StringComparison]::OrdinalIgnoreCase
            )
        ) {
            $prefixedSeriesName = "$typePrefix$SeriesName"
        }

        $safeSeriesName =
            $SeriesName.Replace("'", "''")

        $safePrefixedName =
            $prefixedSeriesName.Replace("'", "''")

        $seriesCode = 0

        $cmd = $conn.CreateCommand()
        try {
            $cmd.CommandTimeout = 5
        }
        catch {
        }

        $cmd.CommandText = @"
SELECT TOP 1 Code
FROM Master1
WHERE MasterType = 21
  AND (
        Name = '$safeSeriesName'
        OR Name = '$safePrefixedName'
      )
"@

        $seriesValue = $cmd.ExecuteScalar()

        if (
            $null -ne $seriesValue -and
            $seriesValue -ne [System.DBNull]::Value
        ) {
            $seriesCode = [int]$seriesValue
        }

        if ($seriesCode -le 0) {
            return @{
                success = $false
                error = "Series '$SeriesName' not found"
            }
        }

        # ------------------------------------------------------------
        # Read numbering configuration
        # ------------------------------------------------------------

        $configValues = $null

        $cmd = $conn.CreateCommand()
        try {
            $cmd.CommandTimeout = 5
        }
        catch {
        }

        $cmd.CommandText = @"
SELECT TOP 1
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

        $reader = $cmd.ExecuteReader()

        try {
            if ($reader.Read()) {
                $configValues =
                    Read-NumberingConfigDbRow `
                        -Reader $reader
            }
        }
        finally {
            if ($reader) {
                try {
                    $reader.Close()
                }
                catch {
                }

                try {
                    $reader.Dispose()
                }
                catch {
                }

                $reader = $null
            }
        }

        # Preserve the original fallback configuration lookup.
        if ($null -eq $configValues) {
            $cmd = $conn.CreateCommand()

            try {
                $cmd.CommandTimeout = 5
            }
            catch {
            }

            $cmd.CommandText = @"
SELECT TOP 1
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

            $reader = $cmd.ExecuteReader()

            try {
                if ($reader.Read()) {
                    $configValues =
                        Read-NumberingConfigDbRow `
                            -Reader $reader
                }
            }
            finally {
                if ($reader) {
                    try {
                        $reader.Close()
                    }
                    catch {
                    }

                    try {
                        $reader.Dispose()
                    }
                    catch {
                    }

                    $reader = $null
                }
            }
        }

        if ($null -eq $configValues) {
            return @{
                success = $false
                error = (
                    "Numbering configuration not found for " +
                    "series '$SeriesName'"
                )
            }
        }

        $prefix = [string]$configValues.c2
        $suffix = [string]$configValues.c1
        $separator = [string]$configValues.c4

        $i1 = [int]$configValues.i1
        $frequency = [int]$configValues.frequency
        $paddingLength = [int]$configValues.paddingLength
        $i6 = [int]$configValues.i6
        $i7 = [int]$configValues.i7
        $i8 = [int]$configValues.i8

        $isAuto = [bool]$configValues.isAuto
        $startNo = [long]$configValues.startNo
        $endingNo = [long]$configValues.endingNo

        $padChar = [char]" "

        if (
            -not [string]::IsNullOrEmpty(
                [string]$configValues.c3
            )
        ) {
            $padChar = [char](
                [string]$configValues.c3
            )[0]
        }

        # ------------------------------------------------------------
        # Resolve date embedding position
        # ------------------------------------------------------------

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
                    elseif ($suffix -ne "") {
                        $embedPos = "prefix"
                    }
                    elseif ($prefix -ne "") {
                        $embedPos = "suffix"
                    }
                    else {
                        $embedPos = "suffix"
                    }
                }
            }
        }

        # ------------------------------------------------------------
        # Build the BUSY date text for the current numbering period
        # ------------------------------------------------------------

        $now = Get-Date

        $monthAbbr = $now.ToString(
            "MMM",
            [System.Globalization.CultureInfo]::InvariantCulture
        ).ToUpperInvariant()

        $financialYear = if ($now.Month -ge 4) {
            $now.Year
        }
        else {
            $now.Year - 1
        }

        $financialYearNext = $financialYear + 1

        $fyShort =
            ($financialYear % 100).ToString("D2")

        $fyNextShort =
            ($financialYearNext % 100).ToString("D2")

        $dateText = ""

        if ($isAuto -and $i1 -gt 0) {
            if ($frequency -eq 3) {
                if ($i8 -eq 1) {
                    $dateText =
                        "$fyShort-$fyNextShort"
                }
                else {
                    $dateText =
                        "$financialYear-$fyNextShort"
                }
            }
            elseif ($frequency -eq 2) {
                switch ($i8) {
                    2 {
                        $dateText = $now.ToString("MMyy")
                    }

                    3 {
                        $dateText = $now.ToString("MM-yyyy")
                    }

                    4 {
                        $dateText = $now.ToString("MMyyyy")
                    }

                    11 {
                        $dateText =
                            "$monthAbbr-$($now.ToString('yyyy'))"
                    }

                    12 {
                        $dateText =
                            "$monthAbbr-$($now.ToString('yy'))"
                    }

                    13 {
                        $dateText = $now.ToString("MM-yyyy")
                    }

                    14 {
                        $dateText = $now.ToString("MM-yy")
                    }

                    15 {
                        $dateText = $now.ToString("MMyy")
                    }

                    16 {
                        $dateText = $now.ToString("yyMM")
                    }

                    default {
                        $dateText = $now.ToString("MMyy")
                    }
                }
            }
        }

        # ------------------------------------------------------------
        # Limit the sequence scan to the current numbering period
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
                    $now.Year, `
                    $now.Month, `
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
                    $fyStartYear, `
                    4, `
                    1

                $periodEndExclusive =
                    $periodStart.AddYears(1)
            }

            if (
                $null -ne $periodStart -and
                $null -ne $periodEndExclusive
            ) {
                if ($isSql) {
                    $startText =
                        $periodStart.ToString(
                            "yyyy-MM-ddTHH:mm:ss",
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )

                    $endText =
                        $periodEndExclusive.ToString(
                            "yyyy-MM-ddTHH:mm:ss",
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )

                    $dateFilter = (
                        " AND [Date] >= '$startText'" +
                        " AND [Date] < '$endText'"
                    )
                }
                else {
                    $startText =
                        $periodStart.ToString(
                            "MM/dd/yyyy",
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )

                    $endText =
                        $periodEndExclusive.ToString(
                            "MM/dd/yyyy",
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )

                    $dateFilter = (
                        " AND [Date] >= #$startText#" +
                        " AND [Date] < #$endText#"
                    )
                }
            }
        }

        # ------------------------------------------------------------
        # Determine the current sequence directly from Tran1
        #
        # We preserve the previous behavior:
        # visible VchNo sequence is preferred over AutoVchNo because BUSY can
        # use a different internal auto sequence.
        #
        # Only the selected voucher type, series, active rows and current
        # numbering period are scanned.
        # ------------------------------------------------------------

        $lastVisibleSeq = 0L
        $lastAutoSeq = 0L

        $cmd = $conn.CreateCommand()

        try {
            $cmd.CommandTimeout = 5
        }
        catch {
        }

        $cmd.CommandText = (
            "SELECT AutoVchNo, VchNo " +
            "FROM Tran1 " +
            "WHERE VchType = $VchType " +
            "AND VchSeriesCode = $seriesCode " +
            "AND Cancelled = 0 " +
            "AND VchCancelled = 0" +
            $dateFilter
        )

        $reader = $cmd.ExecuteReader()

        try {
            while ($reader.Read()) {
                $autoNo = 0L

                try {
                    $autoNo = [long](
                        Get-NumberingDbValue `
                            -Reader $reader `
                            -Field "AutoVchNo" `
                            -Default 0
                    )
                }
                catch {
                    $autoNo = 0L
                }

                if ($autoNo -gt $lastAutoSeq) {
                    $lastAutoSeq = $autoNo
                }

                $dbVoucherNo = [string](
                    Get-NumberingDbValue `
                        -Reader $reader `
                        -Field "VchNo" `
                        -Default ""
                )

                if ([string]::IsNullOrWhiteSpace($dbVoucherNo)) {
                    continue
                }

                if ($isAuto) {
                    $visibleSeq =
                        Get-NumberingVisibleSequence `
                            -VoucherNumber $dbVoucherNo `
                            -Prefix $prefix `
                            -Suffix $suffix `
                            -Separator $separator `
                            -DateText $dateText

                    if ($visibleSeq -gt $lastVisibleSeq) {
                        $lastVisibleSeq = $visibleSeq
                    }
                }
                else {
                    $manualSeq = 0L

                    if (
                        [long]::TryParse(
                            $dbVoucherNo.Trim(),
                            [ref]$manualSeq
                        ) -and
                        $manualSeq -gt $lastVisibleSeq
                    ) {
                        $lastVisibleSeq = $manualSeq
                    }
                }
            }
        }
        finally {
            if ($reader) {
                try {
                    $reader.Close()
                }
                catch {
                }

                try {
                    $reader.Dispose()
                }
                catch {
                }

                $reader = $null
            }
        }

        $lastSeq = 0L

        if ($lastVisibleSeq -gt 0) {
            $lastSeq = $lastVisibleSeq
        }
        elseif ($lastAutoSeq -gt 0) {
            $lastSeq = $lastAutoSeq
        }

        $currentNo = $startNo

        if ($lastSeq -gt 0) {
            $currentNo = $lastSeq + 1
        }

        if (
            $endingNo -gt 0 -and
            $currentNo -gt $endingNo
        ) {
            return @{
                success = $false
                error = (
                    "Next voucher number $currentNo exceeds " +
                    "configured ending number $endingNo"
                )
            }
        }

        $nextVchNo = ""

        if ($isAuto) {
            $nextVchNo =
                Format-NumberingVoucherNo `
                    -Sequence $currentNo `
                    -PaddingLength $paddingLength `
                    -PaddingCharacter $padChar `
                    -Prefix $prefix `
                    -Suffix $suffix `
                    -Separator $separator `
                    -DateText $dateText `
                    -EmbedPosition $embedPos
        }
        else {
            $nextVchNo = [string]$currentNo
        }

        $lastVchNo = "(none yet)"

        if ($lastSeq -gt 0) {
            $lastVchNo =
                Format-NumberingVoucherNo `
                    -Sequence $lastSeq `
                    -PaddingLength $paddingLength `
                    -PaddingCharacter $padChar `
                    -Prefix $prefix `
                    -Suffix $suffix `
                    -Separator $separator `
                    -DateText $dateText `
                    -EmbedPosition $embedPos
        }

        $elapsedMs = [int](
            ((Get-Date) - $startedAt).TotalMilliseconds
        )

        $databaseLabel = [string]$ctx.database

        Write-Host (
            (
                "  [NUMBERING-FAST] {0}/{1} type={2} " +
                "series='{3}' db={4} next='{5}' elapsedMs={6}"
            ) -f
                $InstanceId,
                $CompanyCode,
                $VchType,
                $SeriesName,
                $databaseLabel,
                $nextVchNo,
                $elapsedMs
        ) -ForegroundColor DarkCyan

        return @{
            success = $true

            data = @{
                vch_type = $VchType
                series_name = $SeriesName
                prefix = $prefix
                suffix = $suffix
                padding_length = $paddingLength
                current_no = $currentNo
                is_auto = $isAuto
                frequency = $frequency
                ending_no = $endingNo
                next_vch_no = $nextVchNo
                last_vch_no = $lastVchNo
            }
        }
    }
    catch {
        $elapsedMs = [int](
            ((Get-Date) - $startedAt).TotalMilliseconds
        )

        Write-Host (
            (
                "  [NUMBERING-FAST FAIL] {0}/{1} type={2} " +
                "series='{3}' elapsedMs={4} error={5}"
            ) -f
                $InstanceId,
                $CompanyCode,
                $VchType,
                $SeriesName,
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
            try {
                $reader.Close()
            }
            catch {
            }

            try {
                $reader.Dispose()
            }
            catch {
            }
        }

        if (
            $ctx -and
            $ctx.connection
        ) {
            try {
                $ctx.connection.Close()
            }
            catch {
            }

            try {
                $ctx.connection.Dispose()
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

