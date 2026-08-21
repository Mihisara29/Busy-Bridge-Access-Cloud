Set-StrictMode -Version Latest

function Test-OfflineVoucherAlreadyProcessed {
    param(
        [Parameter(Mandatory = $true)] $DbConnection,
        [Parameter(Mandatory = $true)] [bool] $IsSqlServer,
        [Parameter(Mandatory = $true)] $Data
    )

    $record = Get-OfflineSyncRecord `
        -DbConnection $DbConnection `
        -IsSqlServer $IsSqlServer `
        -LocalId ([string]$Data.localId) `
        -InstanceId ([string]$Data.instanceId) `
        -CompanyCode ([string]$Data.companyCode)

    if ($null -eq $record) {
        return $null
    }

    $status = ([string]$record.Status).Trim().ToUpperInvariant()

    if (
        $status -eq 'SYNCED' -and
        -not [string]::IsNullOrWhiteSpace([string]$record.FinalVchNo)
    ) {
        $syncedAt = (Get-Date).ToString('o')

        if ($null -ne $record.SyncedAt) {
            try {
                $syncedAt = ([datetime]$record.SyncedAt).ToString('o')
            } catch { }
        }

        return @{
            success = $true
            alreadyProcessed = $true
            localId = [string]$Data.localId
            finalVchNo = [string]$record.FinalVchNo
            syncedAt = $syncedAt
            negativeStockWarnings = @()
        }
    }

    if ($status -eq 'PROCESSING') {
        throw 'This offline voucher is already being processed.'
    }

    # FAILED records may be retried.
    return $null
}
