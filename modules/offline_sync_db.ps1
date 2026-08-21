Set-StrictMode -Version Latest

function New-OfflineDbCommand {
    param(
        [Parameter(Mandatory = $true)] $DbConnection,
        [Parameter(Mandatory = $true)] [string] $Sql,
        [hashtable] $Parameters = @{}
    )

    $command = $DbConnection.CreateCommand()
    $command.CommandText = $Sql

    foreach ($key in $Parameters.Keys) {
        $parameter = $command.CreateParameter()
        $parameter.ParameterName = "@$key"
        $parameter.Value = if ($null -eq $Parameters[$key]) {
            [DBNull]::Value
        } else {
            $Parameters[$key]
        }

        [void]$command.Parameters.Add($parameter)
    }

    return $command
}

function Invoke-OfflineNonQuery {
    param(
        [Parameter(Mandatory = $true)] $DbConnection,
        [Parameter(Mandatory = $true)] [string] $Sql,
        [hashtable] $Parameters = @{}
    )

    $command = New-OfflineDbCommand `
        -DbConnection $DbConnection `
        -Sql $Sql `
        -Parameters $Parameters

    try {
        return $command.ExecuteNonQuery()
    }
    finally {
        $command.Dispose()
    }
}

function Invoke-OfflineQuerySingle {
    param(
        [Parameter(Mandatory = $true)] $DbConnection,
        [Parameter(Mandatory = $true)] [string] $Sql,
        [hashtable] $Parameters = @{}
    )

    $command = New-OfflineDbCommand `
        -DbConnection $DbConnection `
        -Sql $Sql `
        -Parameters $Parameters

    $reader = $null

    try {
        $reader = $command.ExecuteReader()

        if (-not $reader.Read()) {
            return $null
        }

        $row = [ordered]@{}

        for ($index = 0; $index -lt $reader.FieldCount; $index++) {
            $name = $reader.GetName($index)
            $value = if ($reader.IsDBNull($index)) {
                $null
            } else {
                $reader.GetValue($index)
            }

            $row[$name] = $value
        }

        return [pscustomobject]$row
    }
    finally {
        if ($null -ne $reader) {
            $reader.Dispose()
        }

        $command.Dispose()
    }
}
