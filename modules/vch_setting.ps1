# modules/vch_setting.ps1
# Voucher UI settings + Web/Busy voucher numbering configuration.

if ($null -eq $script:Config) {
    . "$PSScriptRoot\connection.ps1"
    . "$PSScriptRoot\utils.ps1"
}

# -----------------------------------------------------------------------------
# Existing column/POS configuration
# -----------------------------------------------------------------------------
function Get-ColumnConfig {
    param(
        [int]$VchType,
        [int]$DeviceType,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY database connection failed" } }

    try {
        $qry = "SELECT * FROM Config WHERE RecType=201 AND [Type]=$VchType AND D15=$DeviceType"
        $rst = $fi.GetRecordset($qry)

        function Map-IntToBehavior {
            param([int]$val)
            if ($val -eq 2) { return "semi_variable" }
            if ($val -eq 3) { return "fixed" }
            return "variable"
        }

        function Get-SafeVal {
            param($Recordset, [string]$Field, $Default)
            try {
                $value = $Recordset.Fields.Item($Field).Value
                if ($null -ne $value -and $value -ne [System.DBNull]::Value) { return $value }
            } catch {}
            return $Default
        }

        $config = $null
        if ($rst -and -not $rst.EOF) {
            $config = @{
                vch_type             = $VchType
                device_type          = $DeviceType
                enable_item_discount = ([int](Get-SafeVal $rst "I1" 1) -eq 1)
                enable_alt_units     = ([int](Get-SafeVal $rst "I2" 1) -eq 1)
                col_qty              = Map-IntToBehavior ([int](Get-SafeVal $rst "I3" 1))
                col_unit             = Map-IntToBehavior ([int](Get-SafeVal $rst "I4" 1))
                col_price            = Map-IntToBehavior ([int](Get-SafeVal $rst "I5" 1))
                col_amount           = Map-IntToBehavior ([int](Get-SafeVal $rst "I6" 1))
                col_discount         = Map-IntToBehavior ([int](Get-SafeVal $rst "I7" 1))
                col_cfact            = Map-IntToBehavior ([int](Get-SafeVal $rst "I8" 1))
                col_alt_qty          = Map-IntToBehavior ([int](Get-SafeVal $rst "I9" 1))
                col_alt_price        = Map-IntToBehavior ([int](Get-SafeVal $rst "I10" 1))
                enable_pos           = ([int](Get-SafeVal $rst "I11" 0) -eq 1)
                def_card_acc         = [string](Get-SafeVal $rst "C1" "")
                def_gift_acc         = [string](Get-SafeVal $rst "C2" "")
            }
        }

        if ($rst) { try { $rst.Close() } catch {} }

        if ($null -eq $config) {
            $config = @{
                vch_type=$VchType; device_type=$DeviceType
                enable_item_discount=$true; enable_alt_units=$true
                col_qty="variable"; col_unit="variable"; col_price="variable"; col_amount="variable"
                col_discount="variable"; col_cfact="variable"; col_alt_qty="variable"; col_alt_price="variable"
                enable_pos=$false; def_card_acc=""; def_gift_acc=""
            }
        }

        return @{ success = $true; data = $config }
    } catch {
        return @{ success = $false; error = $_.Exception.Message }
    } finally {
        Disconnect-BUSY $fi
    }
}

function Save-ColumnConfig {
    param($Data, [string]$InstanceId = "", [string]$CompanyCode = "")

    $vchType = [int]$Data.vch_type
    $deviceType = if ($null -ne $Data.device_type) { [int]$Data.device_type } else { 0 }
    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "BUSY database connection failed" } }

    try {
        function Map-BehaviorToInt {
            param([string]$val)
            if ($val -eq "semi_variable") { return 2 }
            if ($val -eq "fixed") { return 3 }
            return 1
        }

        $i1  = if ($Data.enable_item_discount -eq $true -or $Data.enable_item_discount -eq "true") { 1 } else { 0 }
        $i2  = if ($Data.enable_alt_units -eq $true -or $Data.enable_alt_units -eq "true") { 1 } else { 0 }
        $i3  = Map-BehaviorToInt $Data.col_qty
        $i4  = Map-BehaviorToInt $Data.col_unit
        $i5  = Map-BehaviorToInt $Data.col_price
        $i6  = Map-BehaviorToInt $Data.col_amount
        $i7  = Map-BehaviorToInt $Data.col_discount
        $i8  = Map-BehaviorToInt $Data.col_cfact
        $i9  = Map-BehaviorToInt $Data.col_alt_qty
        $i10 = Map-BehaviorToInt $Data.col_alt_price
        $i11 = if ($Data.enable_pos -eq $true -or $Data.enable_pos -eq "true") { 1 } else { 0 }
        $c1 = if ($Data.def_card_acc) { ([string]$Data.def_card_acc).Replace("'", "''") } else { "" }
        $c2 = if ($Data.def_gift_acc) { ([string]$Data.def_gift_acc).Replace("'", "''") } else { "" }

        $exists = $false
        $rst = $fi.GetRecordset("SELECT RecType FROM Config WHERE RecType=201 AND [Type]=$vchType AND D15=$deviceType")
        if ($rst -and -not $rst.EOF) { $exists = $true }
        if ($rst) { try { $rst.Close() } catch {} }

        if ($exists) {
            $sql = "UPDATE Config SET I1=$i1,I2=$i2,I3=$i3,I4=$i4,I5=$i5,I6=$i6,I7=$i7,I8=$i8,I9=$i9,I10=$i10,I11=$i11,C1='$c1',C2='$c2' WHERE RecType=201 AND [Type]=$vchType AND D15=$deviceType"
        } else {
            $sql = "INSERT INTO Config (RecType,[Type],D15,I1,I2,I3,I4,I5,I6,I7,I8,I9,I10,I11,C1,C2) VALUES (201,$vchType,$deviceType,$i1,$i2,$i3,$i4,$i5,$i6,$i7,$i8,$i9,$i10,$i11,'$c1','$c2')"
        }
        $fi.ExecuteQuery($sql)
        return @{ success = $true; message = "Voucher configuration updated successfully" }
    } catch {
        return @{ success = $false; error = $_.Exception.Message }
    } finally {
        Disconnect-BUSY $fi
    }
}

# -----------------------------------------------------------------------------
# Web voucher numbering configuration
# Uses BUSY Config table with custom RecType 202.
# One row per voucher type + voucher series.
#
# Mapping:
#   RecType = 202
#   Type    = voucher type
#   D15     = voucher series Master1.Code
#   I1      = source: 0 BUSY, 1 WEB
#   I2      = reset: 0 NEVER, 1 DAILY, 2 MONTHLY, 3 YEARLY, 4 FINANCIAL_YEAR
#   I3      = padding length
#   I4      = year format: 0 NONE, 1 YY, 2 YYYY
#   I5      = month format: 0 NONE, 1 MMM, 2 MM
#   I6      = active: 0/1
#   I7      = numbering date basis: 0 VOUCHER_DATE, 1 REAL_TIME
#   L2      = starting number
#   L3      = ending number (0 = unlimited)
#   C1      = organisation code
#   C2      = static prefix
#   C3      = static suffix
#   C4      = separator
# -----------------------------------------------------------------------------

function Resolve-VoucherSeriesCode {
    param($fi, [int]$VchType, [string]$SeriesName)

    if ([string]::IsNullOrWhiteSpace($SeriesName)) { throw "Voucher series is required." }
    $clean = $SeriesName.Trim()
    $typePrefix = "{0:D2}" -f $VchType
    $prefixed = if ($clean.StartsWith($typePrefix, [System.StringComparison]::OrdinalIgnoreCase)) { $clean } else { "$typePrefix$clean" }
    $safe1 = $clean.Replace("'", "''")
    $safe2 = $prefixed.Replace("'", "''")

    $rst = $fi.GetRecordset("SELECT Code FROM Master1 WHERE MasterType=21 AND (Name='$safe1' OR Name='$safe2')")
    try {
        if ($rst -and -not $rst.EOF) { return [int]$rst.Fields.Item("Code").Value }
    } finally {
        if ($rst) { try { $rst.Close() } catch {} }
    }
    throw "Voucher series '$clean' was not found for voucher type $VchType."
}

function Get-ConfigSafeValue {
    param($Recordset, [string]$Field, $Default)
    try {
        $value = $Recordset.Fields.Item($Field).Value
        if ($null -ne $value -and $value -ne [System.DBNull]::Value) { return $value }
    } catch {}
    return $Default
}

function Convert-ResetNameToInt {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { $Value = "NEVER" }
    switch ($Value.ToUpperInvariant()) {
        "DAILY" { return 1 }
        "MONTHLY" { return 2 }
        "YEARLY" { return 3 }
        "FINANCIAL_YEAR" { return 4 }
        default { return 0 }
    }
}

function Convert-ResetIntToName {
    param([int]$Value)
    switch ($Value) {
        1 { return "DAILY" }
        2 { return "MONTHLY" }
        3 { return "YEARLY" }
        4 { return "FINANCIAL_YEAR" }
        default { return "NEVER" }
    }
}

function Convert-YearFormatToInt {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { $Value = "YY" }
    switch ($Value.ToUpperInvariant()) { "YYYY" {2}; "NONE" {0}; default {1} }
}

function Convert-YearFormatIntToName {
    param([int]$Value)
    switch ($Value) { 2 {"YYYY"}; 0 {"NONE"}; default {"YY"} }
}

function Convert-MonthFormatToInt {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { $Value = "MMM" }
    switch ($Value.ToUpperInvariant()) { "MM" {2}; "NONE" {0}; default {1} }
}

function Convert-MonthFormatIntToName {
    param([int]$Value)
    switch ($Value) { 2 {"MM"}; 0 {"NONE"}; default {"MMM"} }
}

function Convert-DateBasisToInt {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        $Value = "VOUCHER_DATE"
    }

    switch ($Value.Trim().ToUpperInvariant()) {
        "REAL_TIME" { return 1 }
        default { return 0 }
    }
}

function Convert-DateBasisIntToName {
    param([int]$Value)

    if ($Value -eq 1) {
        return "REAL_TIME"
    }

    return "VOUCHER_DATE"
}

function Resolve-NumberingDate {
    param(
        [string]$DateBasis,
        [datetime]$VoucherDate
    )

    if (
        -not [string]::IsNullOrWhiteSpace($DateBasis) -and
        $DateBasis.Trim().ToUpperInvariant() -eq "REAL_TIME"
    ) {
        return (Get-Date).Date
    }

    return $VoucherDate.Date
}

function Build-WebVoucherNumber {
    param($Config, [datetime]$VoucherDate, [long]$Sequence)

    $parts = New-Object System.Collections.Generic.List[string]
    $separator = [string]$Config.separator
    if ($null -eq $separator) { $separator = "_" }

    $prefix = ([string]$Config.prefix).Trim()
    $suffix = ([string]$Config.suffix).Trim()
    $orgCode = (([string]$Config.organisation_code) -replace '\s+', '').ToUpperInvariant()

    $datePart = ""
    if ($Config.year_format -eq "YY") { $datePart += $VoucherDate.ToString("yy") }
    elseif ($Config.year_format -eq "YYYY") { $datePart += $VoucherDate.ToString("yyyy") }

    if ($Config.month_format -eq "MMM") {
        $datePart += $VoucherDate.ToString("MMM", [System.Globalization.CultureInfo]::InvariantCulture).ToUpperInvariant()
    } elseif ($Config.month_format -eq "MM") {
        $datePart += $VoucherDate.ToString("MM")
    }

    if ($prefix) { $parts.Add($prefix) }
    if ($datePart) { $parts.Add($datePart) }
    if ($orgCode) { $parts.Add($orgCode) }

    $padding = [Math]::Max(1, [int]$Config.padding_length)
    $parts.Add($Sequence.ToString().PadLeft($padding, '0'))
    if ($suffix) { $parts.Add($suffix) }

    $number = [string]::Join($separator, $parts)
    if ($number.Length -gt 40) { throw "Generated voucher number '$number' exceeds 40 characters." }
    if ($number -match '\s') { throw "Generated voucher number cannot contain spaces." }
    if ($number -notmatch '\d') { throw "Generated voucher number must contain a numeric sequence." }
    return $number
}

function Get-NumberingPeriod {
    param([datetime]$VoucherDate, [string]$ResetFrequency)

    switch ($ResetFrequency) {
        "DAILY" { return @{ Start=$VoucherDate.Date; End=$VoucherDate.Date.AddDays(1) } }
        "MONTHLY" {
            $s = New-Object DateTime $VoucherDate.Year, $VoucherDate.Month, 1
            return @{ Start=$s; End=$s.AddMonths(1) }
        }
        "YEARLY" {
            $s = New-Object DateTime $VoucherDate.Year, 1, 1
            return @{ Start=$s; End=$s.AddYears(1) }
        }
        "FINANCIAL_YEAR" {
            $year = if ($VoucherDate.Month -ge 4) { $VoucherDate.Year } else { $VoucherDate.Year - 1 }
            $s = New-Object DateTime $year, 4, 1
            return @{ Start=$s; End=$s.AddYears(1) }
        }
        default { return $null }
    }
}

function Get-WebNumberRegex {
    param(
        $Config,
        [datetime]$VoucherDate
    )

    $sequenceMarker = 987654321L

    $effectiveDate = Resolve-NumberingDate `
        -DateBasis ([string]$Config.date_basis) `
        -VoucherDate $VoucherDate

    $sampleNumber = Build-WebVoucherNumber `
        -Config $Config `
        -VoucherDate $effectiveDate `
        -Sequence $sequenceMarker

    $escapedSample = [regex]::Escape(
        ([string]$sampleNumber).Trim()
    )

    $escapedMarker = [regex]::Escape(
        $sequenceMarker.ToString()
    )

    $pattern = '^' + $escapedSample.Replace(
        $escapedMarker,
        '(\d+)'
    ) + '$'

    return New-Object System.Text.RegularExpressions.Regex(
        $pattern,
        (
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase `
            -bor
            [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
        )
    )
}

function Get-WebConfigBySeriesCode {
    param(
        $fi,
        [int]$VchType,
        [int]$SeriesCode,
        [string]$SeriesName = ""
    )

    $rst = $null
    try {
        $rst = $fi.GetRecordset(
            "SELECT * FROM Config WHERE RecType=202 AND [Type]=$VchType AND D15=$SeriesCode"
        )

        if (-not $rst -or $rst.EOF) { return $null }

        return @{
            vch_type          = $VchType
            series_name       = $SeriesName
            series_code       = $SeriesCode
            source            = if ([int](Get-ConfigSafeValue $rst "I1" 0) -eq 1) { "WEB" } else { "BUSY" }
            reset_frequency   = Convert-ResetIntToName ([int](Get-ConfigSafeValue $rst "I2" 3))
            padding_length    = [int](Get-ConfigSafeValue $rst "I3" 1)
            year_format       = Convert-YearFormatIntToName ([int](Get-ConfigSafeValue $rst "I4" 1))
            month_format      = Convert-MonthFormatIntToName ([int](Get-ConfigSafeValue $rst "I5" 1))
            is_active         = ([int](Get-ConfigSafeValue $rst "I6" 1) -eq 1)
            date_basis       = Convert-DateBasisIntToName ([int](Get-ConfigSafeValue $rst "I7" 0))
            starting_number   = [long](Get-ConfigSafeValue $rst "L2" 1)
            ending_number     = [long](Get-ConfigSafeValue $rst "L3" 0)
            organisation_code = ([string](Get-ConfigSafeValue $rst "C1" "")).Trim()
            prefix            = ([string](Get-ConfigSafeValue $rst "C2" "")).Trim()
            suffix            = ([string](Get-ConfigSafeValue $rst "C3" "")).Trim()
            separator         = [string](Get-ConfigSafeValue $rst "C4" "_")
        }
    }
    finally {
        if ($rst) { try { $rst.Close() } catch {} }
    }
}

function Get-ExistingWebVoucherMigrationPlan {
    param(
        $fi,
        $OldConfig,
        $NewConfig,
        [int]$VchType,
        [int]$SeriesCode
    )

    $sourceRows = @()
    $rst = $null

    try {
        $qry = @"
SELECT
    VchCode,
    VchNo,
    [Date]

FROM Tran1

WHERE
    VchType=$VchType
    AND VchSeriesCode=$SeriesCode
    AND Cancelled=0
    AND VchCancelled=0

ORDER BY
    VchCode
"@

        $rst = $fi.GetRecordset($qry)

        if ($rst -and -not $rst.EOF) {
            $rst.MoveFirst()

            while (-not $rst.EOF) {
                $vchCode = 0
                $oldVchNo = ""
                $voucherDate = Get-Date

                try {
                    $rawCode =
                        $rst.Fields.Item(
                            "VchCode"
                        ).Value

                    if (
                        $null -ne $rawCode -and
                        $rawCode -ne
                            [System.DBNull]::Value
                    ) {
                        $vchCode =
                            [int]$rawCode
                    }
                }
                catch {}

                try {
                    $rawNo =
                        $rst.Fields.Item(
                            "VchNo"
                        ).Value

                    if (
                        $null -ne $rawNo -and
                        $rawNo -ne
                            [System.DBNull]::Value
                    ) {
                        $oldVchNo =
                            ([string]$rawNo).Trim()
                    }
                }
                catch {}

                try {
                    $rawDate =
                        $rst.Fields.Item(
                            "Date"
                        ).Value

                    if (
                        $null -ne $rawDate -and
                        $rawDate -ne
                            [System.DBNull]::Value
                    ) {
                        $voucherDate =
                            [datetime]$rawDate
                    }
                }
                catch {}

                if (
                    $vchCode -gt 0 -and
                    -not [string]::IsNullOrWhiteSpace(
                        $oldVchNo
                    )
                ) {
                    $oldNumberingDate =
                        Resolve-NumberingDate `
                            -DateBasis (
                                [string]$OldConfig.date_basis
                            ) `
                            -VoucherDate $voucherDate

                    $oldRegex =
                        Get-WebNumberRegex `
                            -Config $OldConfig `
                            -VoucherDate $oldNumberingDate

                    $match =
                        $oldRegex.Match(
                            $oldVchNo
                        )

                    if ($match.Success) {
                        $sourceRows +=
                            [pscustomobject]@{
                                VchCode =
                                    $vchCode

                                OldVchNo =
                                    $oldVchNo

                                VoucherDate =
                                    $voucherDate
                            }
                    }
                }

                $rst.MoveNext()
            }
        }
    }
    finally {
        if ($rst) {
            try {
                $rst.Close()
            }
            catch {}
        }
    }

    # Use one fixed server date for the whole migration.
    # Otherwise midnight changing during migration could split the results.
    $migrationRealTimeDate =
        (Get-Date).Date

    $startingNumber =
        [Math]::Max(
            1L,
            [long]$NewConfig.starting_number
        )

    $endingNumber =
        [Math]::Max(
            0L,
            [long]$NewConfig.ending_number
        )

    $nextSequenceByPeriod = @{}
    $rows = @()

    foreach (
        $sourceRow in (
            $sourceRows |
            Sort-Object VchCode
        )
    ) {
        $effectiveDate = if (
            (
                [string]$NewConfig.date_basis
            ).Trim().ToUpperInvariant() -eq
                "REAL_TIME"
        ) {
            $migrationRealTimeDate
        }
        else {
            ([datetime]$sourceRow.VoucherDate).Date
        }

        $periodKey = switch (
            (
                [string]$NewConfig.reset_frequency
            ).Trim().ToUpperInvariant()
        ) {
            "DAILY" {
                $effectiveDate.ToString(
                    "yyyyMMdd",
                    [System.Globalization.CultureInfo]::InvariantCulture
                )
            }

            "MONTHLY" {
                $effectiveDate.ToString(
                    "yyyyMM",
                    [System.Globalization.CultureInfo]::InvariantCulture
                )
            }

            "YEARLY" {
                $effectiveDate.ToString(
                    "yyyy",
                    [System.Globalization.CultureInfo]::InvariantCulture
                )
            }

            "FINANCIAL_YEAR" {
                $financialYearStart = if (
                    $effectiveDate.Month -ge 4
                ) {
                    $effectiveDate.Year
                }
                else {
                    $effectiveDate.Year - 1
                }

                "FY$financialYearStart"
            }

            default {
                "NEVER"
            }
        }

        if (
            -not $nextSequenceByPeriod.ContainsKey(
                $periodKey
            )
        ) {
            $nextSequenceByPeriod[$periodKey] =
                $startingNumber
        }

        $sequence =
            [long]$nextSequenceByPeriod[
                $periodKey
            ]

        if (
            $endingNumber -gt 0 -and
            $sequence -gt $endingNumber
        ) {
            throw (
                "Voucher sequence has reached the configured " +
                "ending number $endingNumber for period " +
                "'$periodKey'. No changes were applied."
            )
        }

        $newVchNo =
            Build-WebVoucherNumber `
                -Config $NewConfig `
                -VoucherDate $effectiveDate `
                -Sequence $sequence

        $rows +=
            [pscustomobject]@{
                VchCode =
                    [int]$sourceRow.VchCode

                OldVchNo =
                    [string]$sourceRow.OldVchNo

                NewVchNo =
                    $newVchNo

                Sequence =
                    $sequence

                VoucherDate =
                    [datetime]$sourceRow.VoucherDate

                NumberingDate =
                    $effectiveDate

                PeriodKey =
                    $periodKey
            }

        $nextSequenceByPeriod[$periodKey] =
            $sequence + 1
    }

    $duplicates = @(
        $rows |
        Group-Object NewVchNo |
        Where-Object {
            $_.Count -gt 1
        }
    )

    if ($duplicates.Count -gt 0) {
        $duplicateNumbers =
            $duplicates |
            ForEach-Object {
                $_.Name
            }

        throw (
            "The new configuration would generate duplicate " +
            "voucher numbers: " +
            (
                $duplicateNumbers -join ", "
            ) +
            ". No changes were applied."
        )
    }

    foreach ($row in $rows) {
        $safeNew =
            ([string]$row.NewVchNo).Replace(
                "'",
                "''"
            )

        $dupRst = $null

        try {
            $dupRst =
                $fi.GetRecordset(@"
SELECT VchCode

FROM Tran1

WHERE
    VchType=$VchType
    AND VchSeriesCode=$SeriesCode
    AND LTRIM(RTRIM(VchNo))='$safeNew'
    AND VchCode<>$($row.VchCode)
"@)

            if (
                $dupRst -and
                -not $dupRst.EOF
            ) {
                $conflictingCode =
                    [int]$dupRst.Fields.Item(
                        "VchCode"
                    ).Value

                $includedInMigration =
                    @(
                        $rows |
                        Where-Object {
                            [int]$_.VchCode -eq
                                $conflictingCode
                        }
                    ).Count -gt 0

                if (-not $includedInMigration) {
                    throw (
                        "Voucher number '$($row.NewVchNo)' " +
                        "already exists outside the migration set. " +
                        "No changes were applied."
                    )
                }
            }
        }
        finally {
            if ($dupRst) {
                try {
                    $dupRst.Close()
                }
                catch {}
            }
        }
    }

    return ,$rows
}

function Get-DbVoucherNumberValue {
    param(
        $fi,
        [string]$TableName,
        [string]$ColumnName,
        [int]$VchCode,
        [int]$VchType
    )

    $rst = $null

    try {
        $query = @"
SELECT $ColumnName
FROM $TableName
WHERE VchCode=$VchCode
  AND VchType=$VchType
"@

        $rst = $fi.GetRecordset($query)

        if (-not $rst -or $rst.EOF) {
            return $null
        }

        $value = $rst.Fields.Item($ColumnName).Value

        if (
            $null -eq $value -or
            $value -eq [System.DBNull]::Value
        ) {
            return ""
        }

        return [string]$value
    }
    finally {
        if ($rst) {
            try { $rst.Close() } catch {}
        }
    }
}

function Get-SqlBusyPaddedVoucherNo {
    param(
        [string]$ExistingValue,
        [string]$NewVoucherNo
    )

    if ($null -eq $ExistingValue) {
        return $NewVoucherNo
    }

    # SQL BUSY may return fixed-width CHAR values with leading spaces.
    # Keep the total stored width, not the old number of spaces.
    #
    # Example:
    # Old field : "              26JUL_QQQ_1"
    # New value : "26JUL_QQQP_1"
    # Result    : one fewer leading space, same total field width.
    $fieldWidth = ([string]$ExistingValue).Length
    $cleanNew = ([string]$NewVoucherNo).Trim()

    if ($fieldWidth -le 0) {
        return $cleanNew
    }

    if ($cleanNew.Length -gt $fieldWidth) {
        throw (
            "The new voucher number '$cleanNew' has $($cleanNew.Length) " +
            "characters, but the BUSY SQL field currently allows only " +
            "$fieldWidth characters. Shorten the numbering format."
        )
    }

    return $cleanNew.PadLeft($fieldWidth, [char]' ')
}

function Set-BusyVoucherNumberInTable {
    param(
        $fi,
        [string]$TableName,
        [string]$ColumnName,
        [int]$VchCode,
        [int]$VchType,
        [string]$NewVoucherNo,
        [bool]$IsSql
    )

    $currentValue = Get-DbVoucherNumberValue `
        -fi $fi `
        -TableName $TableName `
        -ColumnName $ColumnName `
        -VchCode $VchCode `
        -VchType $VchType

    # Some voucher types may not create a row in every table.
    if ($null -eq $currentValue) {
        Write-Host (
            "  [WEB-NUMBER MIGRATION] " +
            "$TableName.$ColumnName has no row for " +
            "VchCode=$VchCode, VchType=$VchType. Skipped."
        ) -ForegroundColor DarkYellow

        return $false
    }

    $storedValue = if ($IsSql) {
        Get-SqlBusyPaddedVoucherNo `
            -ExistingValue $currentValue `
            -NewVoucherNo $NewVoucherNo
    }
    else {
        $NewVoucherNo
    }

    $safeStoredValue = $storedValue.Replace("'", "''")

    $updateQuery = @"
UPDATE $TableName
SET $ColumnName='$safeStoredValue'
WHERE VchCode=$VchCode
  AND VchType=$VchType
"@

    $currentValueForLog = ([string]$currentValue) -replace "[\r\n]", ""
    $storedValueForLog = ([string]$storedValue) -replace "[\r\n]", ""

    Write-Host (
        "  [WEB-NUMBER MIGRATION] Updating " +
        "$TableName.$ColumnName | VchCode=$VchCode | " +
        "Old='$currentValueForLog' | " +
        "New='$storedValueForLog'"
    ) -ForegroundColor Cyan

    $fi.ExecuteQuery($updateQuery)

    $savedValue = Get-DbVoucherNumberValue `
        -fi $fi `
        -TableName $TableName `
        -ColumnName $ColumnName `
        -VchCode $VchCode `
        -VchType $VchType

    if ($null -eq $savedValue) {
        throw (
            "Could not verify $TableName.$ColumnName after updating " +
            "voucher code $VchCode."
        )
    }

    if (
        -not $savedValue.Trim().Equals(
            $NewVoucherNo.Trim(),
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        throw (
            "$TableName.$ColumnName was not updated correctly for " +
            "voucher code $VchCode. Expected '$NewVoucherNo', " +
            "database returned '$savedValue'."
        )
    }

    return $true
}

function Update-ExistingWebVoucherNumbers {
    param(
        $fi,
        $MigrationPlan,
        [int]$VchType,
        [int]$SeriesCode,
        [bool]$IsSql = $false
    )

    $updatedVoucherCount = 0

    foreach ($row in @($MigrationPlan)) {
        $vchCode = [int]$row.VchCode
        $newVchNo = [string]$row.NewVchNo
        $oldVchNo = [string]$row.OldVchNo

        Write-Host (
            "[WEB-NUMBER MIGRATION] Voucher " +
            "VchCode=$vchCode | '$oldVchNo' -> '$newVchNo'"
        ) -ForegroundColor Green

        # BUSY stores the visible voucher number in multiple transaction tables.
        #
        # Tran1 : Voucher header
        # Tran2 : Voucher transaction/detail rows
        # Tran3 : Bill/reference rows; the voucher number field is named [No]
        #
        # AutoVchNo is intentionally NOT updated for WEB-numbered vouchers.
        $tran1Updated = Set-BusyVoucherNumberInTable `
            -fi $fi `
            -TableName "Tran1" `
            -ColumnName "VchNo" `
            -VchCode $vchCode `
            -VchType $VchType `
            -NewVoucherNo $newVchNo `
            -IsSql $IsSql

        $tran2Updated = Set-BusyVoucherNumberInTable `
            -fi $fi `
            -TableName "Tran2" `
            -ColumnName "VchNo" `
            -VchCode $vchCode `
            -VchType $VchType `
            -NewVoucherNo $newVchNo `
            -IsSql $IsSql

        $tran3Updated = Set-BusyVoucherNumberInTable `
            -fi $fi `
            -TableName "Tran3" `
            -ColumnName "[No]" `
            -VchCode $vchCode `
            -VchType $VchType `
            -NewVoucherNo $newVchNo `
            -IsSql $IsSql

        if (-not $tran1Updated) {
            throw (
                "Tran1 voucher header was not found for VchCode=$vchCode. " +
                "Migration was stopped."
            )
        }

        Write-Host (
            "[WEB-NUMBER MIGRATION] Verified VchCode=$vchCode | " +
            "Tran1=$tran1Updated, Tran2=$tran2Updated, Tran3=$tran3Updated"
        ) -ForegroundColor Green

        $updatedVoucherCount++
    }

    return $updatedVoucherCount
}

function Get-WebNumberingConfig {
    param(
        [int]$VchType,
        [string]$SeriesName,
        [string]$VoucherDate = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success=$false; error="BUSY database connection failed" } }

    try {
        $seriesCode = Resolve-VoucherSeriesCode -fi $fi -VchType $VchType -SeriesName $SeriesName
        $rst = $fi.GetRecordset("SELECT * FROM Config WHERE RecType=202 AND [Type]=$VchType AND D15=$seriesCode")

        $cfg = @{
            vch_type=$VchType; series_name=$SeriesName; series_code=$seriesCode
            source="BUSY"; reset_frequency="YEARLY"; padding_length=1
            year_format="YY"; month_format="MMM"; is_active=$true
            date_basis="VOUCHER_DATE"
            starting_number=1L; ending_number=0L
            organisation_code=""; prefix=""; suffix=""; separator="_"
        }

        if ($rst -and -not $rst.EOF) {
            $cfg.source = if ([int](Get-ConfigSafeValue $rst "I1" 0) -eq 1) { "WEB" } else { "BUSY" }
            $cfg.reset_frequency = Convert-ResetIntToName ([int](Get-ConfigSafeValue $rst "I2" 3))
            $cfg.padding_length = [int](Get-ConfigSafeValue $rst "I3" 1)
            $cfg.year_format = Convert-YearFormatIntToName ([int](Get-ConfigSafeValue $rst "I4" 1))
            $cfg.month_format = Convert-MonthFormatIntToName ([int](Get-ConfigSafeValue $rst "I5" 1))
            $cfg.is_active = ([int](Get-ConfigSafeValue $rst "I6" 1) -eq 1)
            $cfg.date_basis = Convert-DateBasisIntToName ([int](Get-ConfigSafeValue $rst "I7" 0))
            $cfg.starting_number = [long](Get-ConfigSafeValue $rst "L2" 1)
            $cfg.ending_number = [long](Get-ConfigSafeValue $rst "L3" 0)
            $cfg.organisation_code = ([string](Get-ConfigSafeValue $rst "C1" "")).Trim()
            $cfg.prefix = ([string](Get-ConfigSafeValue $rst "C2" "")).Trim()
            $cfg.suffix = ([string](Get-ConfigSafeValue $rst "C3" "")).Trim()
            $cfg.separator = [string](Get-ConfigSafeValue $rst "C4" "_")
        }
        if ($rst) { try { $rst.Close() } catch {} }

        $date = if (-not [string]::IsNullOrWhiteSpace($VoucherDate)) {
            [datetime]::Parse($VoucherDate, [System.Globalization.CultureInfo]::InvariantCulture)
        } else {
            Get-Date
        }

        $effectiveDate = Resolve-NumberingDate `
            -DateBasis ([string]$cfg.date_basis) `
            -VoucherDate $date

        if ($cfg.source -eq "WEB" -and $cfg.is_active) {
            $targetInstance = Get-InstanceConfig -InstanceId $InstanceId
            $dbType = 0
            if ($null -ne $targetInstance -and $null -ne $targetInstance.dbType) {
                $dbType = [int]$targetInstance.dbType
            }
            $isSql = ($dbType -eq 1)

            $preview = Get-NextWebVoucherNumberInternal `
                -fi $fi `
                -Config $cfg `
                -VoucherDate $effectiveDate `
                -IsSql $isSql

            $cfg.last_sequence = $preview.last_sequence
            $cfg.next_sequence = $preview.next_sequence
            $cfg.next_vch_no = $preview.next_vch_no
        } else {
            $cfg.last_sequence = 0
            $cfg.next_sequence = 0
            $cfg.next_vch_no = ""
        }

        return @{ success=$true; data=$cfg }
    } catch {
        return @{ success=$false; error=$_.Exception.Message }
    } finally {
        Disconnect-BUSY $fi
    }
}

function Save-WebNumberingConfig {
    param(
        $Data,
        [string]$UpdatedBy = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) {
        return @{ success = $false; error = "BUSY database connection failed" }
    }

    try {
        $vchType = [int]$Data.vch_type
        $seriesName = ([string]$Data.series_name).Trim()

        if ($vchType -le 0) { throw "Voucher type is required." }
        if ([string]::IsNullOrWhiteSpace($seriesName)) { throw "Voucher series is required." }

        $seriesCode = Resolve-VoucherSeriesCode -fi $fi -VchType $vchType -SeriesName $seriesName
        $oldConfig = Get-WebConfigBySeriesCode -fi $fi -VchType $vchType -SeriesCode $seriesCode -SeriesName $seriesName

        $source = ([string]$Data.source).Trim().ToUpperInvariant()
        if ($source -notin @("BUSY", "WEB")) { throw "Source must be BUSY or WEB." }

        $org = (([string]$Data.organisation_code) -replace '\s+', '').ToUpperInvariant()
        if ($source -eq "WEB" -and [string]::IsNullOrWhiteSpace($org)) {
            throw "Organisation code is required for WEB numbering."
        }
        if ($org -and $org -notmatch '^[A-Z0-9_-]+$') {
            throw "Organisation code may contain only letters, numbers, underscore and hyphen."
        }

        $prefixRaw = (([string]$Data.prefix) -replace '\s+', '')
        $suffixRaw = (([string]$Data.suffix) -replace '\s+', '')
        $separatorRaw = [string]$Data.separator
        if ($null -eq $separatorRaw) { $separatorRaw = "_" }
        if ($separatorRaw.Length -gt 1) { throw "Separator must be one character or empty." }

        $sourceInt = if ($source -eq "WEB") { 1 } else { 0 }
        $resetInt = Convert-ResetNameToInt ([string]$Data.reset_frequency)
        $yearInt = Convert-YearFormatToInt ([string]$Data.year_format)
        $monthInt = Convert-MonthFormatToInt ([string]$Data.month_format)
        $dateBasisInt = Convert-DateBasisToInt ([string]$Data.date_basis)
        $dateBasis = Convert-DateBasisIntToName $dateBasisInt
        $padding = [Math]::Min(15, [Math]::Max(1, [int]$Data.padding_length))

        $active = 1
        $isActiveProperty = $Data.PSObject.Properties["is_active"]
        if ($null -ne $isActiveProperty) {
            $rawActive = $isActiveProperty.Value
            if ($rawActive -is [bool]) {
                $active = if ([bool]$rawActive) { 1 } else { 0 }
            }
            else {
                $activeText = ([string]$rawActive).Trim().ToLowerInvariant()
                $active = if ($activeText -in @("true", "1", "yes", "on")) { 1 } else { 0 }
            }
        }

        $updateExisting = $false
        $updateProperty = $Data.PSObject.Properties["update_existing_web_vouchers"]
        if ($null -ne $updateProperty) {
            $rawUpdate = $updateProperty.Value
            if ($rawUpdate -is [bool]) {
                $updateExisting = [bool]$rawUpdate
            }
            else {
                $updateExisting = (([string]$rawUpdate).Trim().ToLowerInvariant() -in @("true", "1", "yes", "on"))
            }
        }

        $startNo = [Math]::Max(1L, [long]$Data.starting_number)
        $endNo = [Math]::Max(0L, [long]$Data.ending_number)
        if ($endNo -gt 0 -and $endNo -lt $startNo) {
            throw "Ending number cannot be lower than starting number."
        }

        # REAL_TIME numbering cannot safely use voucher transaction dates to
        # separate historical periods. Require a visible date token so each
        # real-time period can be identified from VchNo itself.
        if ($dateBasis -eq "REAL_TIME") {
            $resetName = Convert-ResetIntToName $resetInt
            $yearName = Convert-YearFormatIntToName $yearInt
            $monthName = Convert-MonthFormatIntToName $monthInt

            if ($resetName -eq "DAILY") {
                throw "Real-Time Date does not support Daily reset because the format has no day token."
            }

            if ($resetName -eq "MONTHLY" -and (
                $yearName -eq "NONE" -or
                $monthName -eq "NONE"
            )) {
                throw "Real-Time Date with Monthly reset requires both year and month in the voucher number."
            }

            if ($resetName -in @("YEARLY", "FINANCIAL_YEAR") -and $yearName -eq "NONE") {
                throw "Real-Time Date with yearly reset requires the year in the voucher number."
            }

            if ($resetName -eq "FINANCIAL_YEAR") {
                throw "Real-Time Date does not support Financial Year reset safely. Use Voucher Form Date for financial-year numbering."
            }
        }

        $newConfig = @{
            vch_type          = $vchType
            series_name       = $seriesName
            series_code       = $seriesCode
            source            = $source
            reset_frequency   = Convert-ResetIntToName $resetInt
            padding_length    = $padding
            year_format       = Convert-YearFormatIntToName $yearInt
            month_format      = Convert-MonthFormatIntToName $monthInt
            is_active         = ($active -eq 1)
            date_basis       = $dateBasis
            starting_number   = $startNo
            ending_number     = $endNo
            organisation_code = $org
            prefix            = $prefixRaw
            suffix            = $suffixRaw
            separator         = $separatorRaw
        }

        $validationDate = Resolve-NumberingDate `
            -DateBasis $dateBasis `
            -VoucherDate (Get-Date)

        [void](Build-WebVoucherNumber `
            -Config $newConfig `
            -VoucherDate $validationDate `
            -Sequence $startNo)

        # Build and validate the complete migration plan BEFORE changing anything.
        $migrationPlan = @()
        if (
            $updateExisting -and
            $null -ne $oldConfig -and
            $oldConfig.source -eq "WEB" -and
            $source -eq "WEB"
        ) {
            $migrationPlan = Get-ExistingWebVoucherMigrationPlan `
                -fi $fi `
                -OldConfig $oldConfig `
                -NewConfig $newConfig `
                -VchType $vchType `
                -SeriesCode $seriesCode
        }

        $safeOrg = $org.Replace("'", "''")
        $safePrefix = $prefixRaw.Replace("'", "''")
        $safeSuffix = $suffixRaw.Replace("'", "''")
        $safeSeparator = $separatorRaw.Replace("'", "''")

        $exists = ($null -ne $oldConfig)
        if ($exists) {
            $sql = "UPDATE Config SET I1=$sourceInt,I2=$resetInt,I3=$padding,I4=$yearInt,I5=$monthInt,I6=$active,I7=$dateBasisInt,L2=$startNo,L3=$endNo,C1='$safeOrg',C2='$safePrefix',C3='$safeSuffix',C4='$safeSeparator' WHERE RecType=202 AND [Type]=$vchType AND D15=$seriesCode"
        }
        else {
            $sql = "INSERT INTO Config (RecType,[Type],D15,I1,I2,I3,I4,I5,I6,I7,L2,L3,C1,C2,C3,C4) VALUES (202,$vchType,$seriesCode,$sourceInt,$resetInt,$padding,$yearInt,$monthInt,$active,$dateBasisInt,$startNo,$endNo,'$safeOrg','$safePrefix','$safeSuffix','$safeSeparator')"
        }
        $fi.ExecuteQuery($sql)

        $updatedCount = 0
        if ($migrationPlan.Count -gt 0) {
            $targetInstance = Get-InstanceConfig `
                -InstanceId $InstanceId

            $dbType = 0

            if (
                $null -ne $targetInstance -and
                $null -ne $targetInstance.dbType
            ) {
                $dbType = [int]$targetInstance.dbType
            }

            $isSql = ($dbType -eq 1)

            Write-Host (
                "[WEB-NUMBER MIGRATION] Database=" +
                $(if ($isSql) { "SQL Server" } else { "Access" }) +
                " | VoucherType=$vchType | SeriesCode=$seriesCode | " +
                "Count=$($migrationPlan.Count)"
            ) -ForegroundColor Magenta

            $updatedCount = Update-ExistingWebVoucherNumbers `
                -fi $fi `
                -MigrationPlan $migrationPlan `
                -VchType $vchType `
                -SeriesCode $seriesCode `
                -IsSql $isSql
        }

        $verifyRst = $null
        $savedActive = -1
        try {
            $verifyRst = $fi.GetRecordset("SELECT I6 FROM Config WHERE RecType=202 AND [Type]=$vchType AND D15=$seriesCode")
            if ($verifyRst -and -not $verifyRst.EOF) {
                $savedValue = $verifyRst.Fields.Item("I6").Value
                if ($null -ne $savedValue -and $savedValue -ne [System.DBNull]::Value) {
                    $savedActive = [int]$savedValue
                }
            }
        }
        finally {
            if ($verifyRst) { try { $verifyRst.Close() } catch {} }
        }

        if ($savedActive -ne $active) {
            throw "The Active setting was not saved correctly. Expected I6=$active but database returned I6=$savedActive."
        }

        $result = Get-WebNumberingConfig `
            -VchType $vchType `
            -SeriesName $seriesName `
            -VoucherDate ([string]$Data.voucher_date) `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if ($result.success) {
            $result.data.updated_existing_count = $updatedCount
        }
        return $result
    }
    catch {
        return @{ success = $false; error = $_.Exception.Message }
    }
    finally {
        if ($fi) { Disconnect-BUSY $fi }
    }
}

function Get-NextWebVoucherNumberInternal {
    param(
        $fi,
        $Config,
        [datetime]$VoucherDate,
        [bool]$IsSql = $false
    )

    if ($null -eq $fi) {
        throw "BUSY connection is unavailable."
    }

    # VoucherDate passed to this function is already the effective numbering
    # date resolved from either VOUCHER_DATE or REAL_TIME.
    $effectiveDate = $VoucherDate.Date
    $dateBasis = ([string]$Config.date_basis).Trim().ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($dateBasis)) {
        $dateBasis = "VOUCHER_DATE"
    }

    # Voucher-date mode can safely filter Tran1.[Date], because that column is
    # the voucher form date. Real-time mode must not filter Tran1.[Date]:
    # backdated vouchers may have a different form date. In real-time mode the
    # exact year/month token in VchNo identifies the active period.
    $dateFilter = ""

    if ($dateBasis -eq "VOUCHER_DATE") {
        $period = Get-NumberingPeriod `
            -VoucherDate $effectiveDate `
            -ResetFrequency ([string]$Config.reset_frequency)

        if ($null -ne $period) {
            if ($IsSql) {
                $startText = $period.Start.ToString(
                    "yyyy-MM-ddTHH:mm:ss",
                    [System.Globalization.CultureInfo]::InvariantCulture
                )
                $endText = $period.End.ToString(
                    "yyyy-MM-ddTHH:mm:ss",
                    [System.Globalization.CultureInfo]::InvariantCulture
                )
                $dateFilter = " AND [Date] >= '$startText' AND [Date] < '$endText'"
            }
            else {
                $startText = $period.Start.ToString(
                    "MM/dd/yyyy",
                    [System.Globalization.CultureInfo]::InvariantCulture
                )
                $endText = $period.End.ToString(
                    "MM/dd/yyyy",
                    [System.Globalization.CultureInfo]::InvariantCulture
                )
                $dateFilter = " AND [Date] >= #$startText# AND [Date] < #$endText#"
            }
        }
    }

    # Build an exact regex for the active date token and WEB format.
    # This ignores old BUSY-generated numbers and numbers from other periods.
    $webNumberRegex = Get-WebNumberRegex `
        -Config $Config `
        -VoucherDate $effectiveDate

    $qry = (
        "SELECT VchNo FROM Tran1 " +
        "WHERE VchType=$([int]$Config.vch_type) " +
        "AND VchSeriesCode=$([int]$Config.series_code) " +
        "AND Cancelled=0 AND VchCancelled=0" +
        $dateFilter
    )

    $rst = $null
    $highest = 0L

    try {
        $rst = $fi.GetRecordset($qry)

        if ($rst -and -not $rst.EOF) {
            $rst.MoveFirst()

            while (-not $rst.EOF) {
                $raw = ([string](Get-ConfigSafeValue $rst "VchNo" "")).Trim()

                if (-not [string]::IsNullOrWhiteSpace($raw)) {
                    $match = $webNumberRegex.Match($raw)

                    if ($match.Success) {
                        $seq = 0L

                        if (
                            [long]::TryParse(
                                $match.Groups[1].Value,
                                [ref]$seq
                            ) -and
                            $seq -gt $highest
                        ) {
                            $highest = $seq
                        }
                    }
                }

                $rst.MoveNext()
            }
        }
    }
    finally {
        if ($rst) {
            try { $rst.Close() } catch {}

            try {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $rst
                ) | Out-Null
            }
            catch {}
        }
    }

    $startingNumber =
        [Math]::Max(1L, [long]$Config.starting_number)

    $next = if ($highest -gt 0) {
        $highest + 1
    }
    else {
        $startingNumber
    }

    $endingNumber =
        [Math]::Max(0L, [long]$Config.ending_number)

    if (
        $endingNumber -gt 0 -and
        $next -gt $endingNumber
    ) {
        throw (
            "Voucher sequence has reached the configured ending " +
            "number $endingNumber."
        )
    }

    return @{
        last_sequence = $highest
        next_sequence = $next
        next_vch_no   = Build-WebVoucherNumber `
            -Config $Config `
            -VoucherDate $effectiveDate `
            -Sequence $next
    }
}

function Get-EffectiveVoucherNumber {
    param(
        $fi,
        [int]$VchType,
        [string]$SeriesName,
        [datetime]$VoucherDate,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $cfgResult = Get-WebNumberingConfig -VchType $VchType -SeriesName $SeriesName -VoucherDate $VoucherDate.ToString("yyyy-MM-dd") -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $cfgResult.success) { throw $cfgResult.error }
    $cfg = $cfgResult.data

    if ($cfg.source -eq "WEB" -and $cfg.is_active) {
        return @{ source="WEB"; vch_no=$cfg.next_vch_no; sequence=$cfg.next_sequence }
    }

    return @{ source="BUSY"; vch_no=""; sequence=0 }
}

# Returns the same response shape already consumed by useNumberingConfig.
# WEB configuration overrides BUSY only when source=WEB and active=true.
function Get-EffectiveNumberingConfig {
    param(
        [int]$VchType,
        [string]$SeriesName,
        [string]$VoucherDate = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $webResult = Get-WebNumberingConfig -VchType $VchType -SeriesName $SeriesName -VoucherDate $VoucherDate -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $webResult.success) { return $webResult }

    $web = $webResult.data
    if ($web.source -eq "WEB" -and $web.is_active) {
        return @{
            success = $true
            data = @{
                vch_type       = $VchType
                series_name    = $SeriesName
                source         = "WEB"
                date_basis     = $web.date_basis
                prefix         = $web.prefix
                suffix         = $web.suffix
                padding_length = $web.padding_length
                current_no     = $web.next_sequence
                is_auto        = $true
                frequency      = switch ($web.reset_frequency) { "DAILY" {1}; "MONTHLY" {2}; "YEARLY" {3}; "FINANCIAL_YEAR" {3}; default {0} }
                ending_no      = $web.ending_number
                next_vch_no    = $web.next_vch_no
                last_vch_no    = if ($web.last_sequence -gt 0) {
                    $rawVoucherDate = if ($VoucherDate) {
                        [datetime]::Parse(
                            $VoucherDate,
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )
                    }
                    else {
                        Get-Date
                    }

                    $effectiveLastDate = Resolve-NumberingDate `
                        -DateBasis ([string]$web.date_basis) `
                        -VoucherDate $rawVoucherDate

                    Build-WebVoucherNumber `
                        -Config $web `
                        -VoucherDate $effectiveLastDate `
                        -Sequence $web.last_sequence
                }
                else {
                    "(none yet)"
                }
            }
        }
    }

    # Existing BUSY function from vouchers.ps1.
    return Get-NumberingConfig -VchType $VchType -SeriesName $SeriesName -InstanceId $InstanceId -CompanyCode $CompanyCode
}