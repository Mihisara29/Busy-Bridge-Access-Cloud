Set-StrictMode -Version Latest

function Assert-OfflineVoucherRequest {
    param(
        [Parameter(Mandatory = $true)] $Data
    )

    $requiredFields = @(
        'localId',
        'instanceId',
        'companyCode',
        'userName',
        'vchType',
        'seriesName',
        'voucherDate',
        'payload'
    )

    foreach ($field in $requiredFields) {
        if (-not $Data.PSObject.Properties.Name.Contains($field)) {
            throw "Missing required offline voucher field: $field"
        }
    }

    if ([string]::IsNullOrWhiteSpace([string]$Data.localId)) {
        throw 'localId is required.'
    }

    if ([string]::IsNullOrWhiteSpace([string]$Data.instanceId)) {
        throw 'instanceId is required.'
    }

    if ([string]::IsNullOrWhiteSpace([string]$Data.companyCode)) {
        throw 'companyCode is required.'
    }

    if ([string]::IsNullOrWhiteSpace([string]$Data.userName)) {
        throw 'userName is required.'
    }

    if ([string]::IsNullOrWhiteSpace([string]$Data.seriesName)) {
        throw 'Voucher series is required.'
    }

    $vchType = 0
    if (-not [int]::TryParse(([string]$Data.vchType), [ref]$vchType) -or $vchType -le 0) {
        throw 'Invalid voucher type.'
    }

    $parsedDate = [datetime]::MinValue
    if (-not [datetime]::TryParse(([string]$Data.voucherDate), [ref]$parsedDate)) {
        throw 'Invalid voucher date.'
    }

    if ($null -eq $Data.payload) {
        throw 'Voucher payload is required.'
    }

    # Negative stock is allowed by business rule.
    # Do not reject a voucher only because stock becomes negative.
}

function Assert-OfflineVoucherPermission {
    param(
        [Parameter(Mandatory = $true)] $CurrentUser,
        [Parameter(Mandatory = $true)] $Data
    )

    if ($null -eq $CurrentUser) {
        throw 'Authentication is required.'
    }

    $currentUserName = [string]$CurrentUser.name

    if ([string]::IsNullOrWhiteSpace($currentUserName)) {
        throw 'Authenticated user name is missing.'
    }

    # The route overwrites Data.userName with the authenticated user name.
    if (
        -not [string]::IsNullOrWhiteSpace([string]$Data.userName) -and
        ([string]$Data.userName).Trim().ToLowerInvariant() -ne
        $currentUserName.Trim().ToLowerInvariant()
    ) {
        throw 'Offline voucher user does not match the authenticated user.'
    }

    # Hook your existing voucher-create permission function here if available.
    if (Get-Command Test-UserVoucherPermission -ErrorAction SilentlyContinue) {
        $allowed = Test-UserVoucherPermission `
            -User $CurrentUser `
            -VchType ([int]$Data.vchType) `
            -Action 'create'

        if (-not $allowed) {
            throw 'The current user does not have permission to create this voucher type.'
        }
    }
}
