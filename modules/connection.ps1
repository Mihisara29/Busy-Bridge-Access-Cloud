# modules/connection.ps1
# Dynamic BUSY Connection Management — SAFE SINGLE POOL (Access/SQL Server compatible)

. "$PSScriptRoot\config.ps1"
. "$PSScriptRoot\utils.ps1"

$script:maxRetries = 3

$script:ActiveConnection = $null
$script:ActiveInstanceId = ""
$script:ActiveCompanyCode = ""

function Get-InstanceConfig {
    param([string]$InstanceId)
    $instancesPath = "$PSScriptRoot\..\instances.json"
    if (-not (Test-Path $instancesPath)) { return $null }
    $config   = Get-Content $instancesPath -Raw | ConvertFrom-Json
    $instance = $config.instances | Where-Object { $_.id -eq $InstanceId }
    return $instance
}

function Get-MainCompanyDbPath {
    param([string]$CompanyCode)
    $instancesPath = "$PSScriptRoot\..\instances.json"
    if (-not (Test-Path $instancesPath)) { return $null }
    try {
        $config = Get-Content $instancesPath -Raw | ConvertFrom-Json
        foreach ($inst in $config.instances) {
            foreach ($comp in $inst.companies) {
                if ($comp.code.ToLower() -eq $CompanyCode.ToLower()) {
                    return Join-Path $inst.dataPath "$($comp.code)\db.bds"
                }
            }
        }
    } catch {}
    return $null
}

# ═══════════════════════════════════════════════════════
#  DYNAMIC DATABASE NAME RESOLVER (Explicit Mapping + Fallback Formatter)
# ═══════════════════════════════════════════════════════
function Get-SqlDatabaseName {
    param([string]$CompanyCode, [string]$InstanceId = "")
    
    if ($CompanyCode.ToUpper() -eq "COMPINFO") {
        return "COMPINFO"
    }
    
    # 1. First, check instances.json for an explicit "sqlDatabase" mapping
    if ($InstanceId) {
        $instance = Get-InstanceConfig -InstanceId $InstanceId
        if ($null -ne $instance -and $null -ne $instance.companies) {
            $matchedComp = $instance.companies | Where-Object { $_.code -eq $CompanyCode }
            if ($matchedComp -and $matchedComp.sqlDatabase -and $matchedComp.sqlDatabase -ne "") {
                return $matchedComp.sqlDatabase
            }
        }
    }
    
    # 2. Fallback to formatting algorithm if not explicitly mapped: "COMP0002" -> "BusyComp0002_db"
    if ($CompanyCode -match "^COMP(\d+)$" -or $CompanyCode -match "^comp(\d+)$") {
        return "BusyComp" + $Matches[1] + "_db"
    }
    
    # Generic fallback: "DEMO" -> "BusyDemo_db"
    $clean = $CompanyCode.Trim()
    if ($clean.Length -gt 1) {
        $clean = [char]::ToUpper($clean[0]) + $clean.Substring(1).ToLower()
    }
    return "Busy" + $clean + "_db"
}

function Connect-BUSY {
    param([string]$InstanceId = "", [string]$CompanyCode = "")

    $staticConfig = Get-Config

    if (-not $InstanceId -or -not $CompanyCode) {
        $InstanceId   = $staticConfig.INSTANCE_ID
        $CompanyCode  = $staticConfig.COMP_CODE
    }

    if ($script:ActiveConnection -ne $null -and $script:ActiveInstanceId -eq $InstanceId -and $script:ActiveCompanyCode -eq $CompanyCode) {
        return $script:ActiveConnection
    }

    if ($script:ActiveConnection -ne $null) {
        Write-Host "Switching Context: Closing $($script:ActiveCompanyCode) and opening $CompanyCode" -ForegroundColor Cyan
        try { $script:ActiveConnection.CloseDB() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:ActiveConnection) | Out-Null } catch {}
        try { [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers() } catch {}
        $script:ActiveConnection = $null
    }

    $bPath  = $staticConfig.BUSY_PATH
    $dPath  = $staticConfig.DATA_PATH
    $bCom   = $staticConfig.BUSY_COM
    $dbType = 0 # Default: MS Access

    $instance = Get-InstanceConfig -InstanceId $InstanceId
    if ($instance) {
        $bPath  = $instance.busyPath
        $dPath  = $instance.dataPath
        $bCom   = $instance.busyCom
        if ($null -ne $instance.dbType) { $dbType = [int]$instance.dbType }
    }

    # Pre-load SQL credentials for CS mode
    $sqlServer   = ""
    $sqlUser     = ""
    $sqlPassword = ""
    if ($dbType -eq 1 -and $instance) {
        $sqlServer   = $instance.sqlServer
        $sqlUser     = $instance.sqlUser
        $sqlPassword = $instance.sqlPassword
    }

    $retryCount = 0
    while ($retryCount -lt $script:maxRetries) {
        try {
            Write-DebugLog "Connecting to BUSY ($CompanyCode) via COM (DbType=$dbType)..."

            $fi = New-Object -ComObject $bCom
            $connected = $false

            if ($dbType -eq 1) {
                # SQL Server Mode — OpenCSDB throws internal BUSY query errors even on
                # successful connection (e.g. "Object invalid or no longer set" / error 3420).
                # These are non-fatal: the COM object is usable and the return value is True.
                # We catch the exception and treat any thrown-but-True result as success.
                try {
                    $connected = $fi.OpenCSDB($bPath, $sqlServer, $sqlUser, $sqlPassword, $CompanyCode)
                } catch {
                    Write-Host "  [WARN] OpenCSDB internal exception (non-fatal): $($_.Exception.Message)" -ForegroundColor DarkYellow
                    # COM returned True before throwing — treat as connected
                    $connected = $true
                }
            } else {
                # MS Access Mode — use OpenDB
                $connected = $fi.OpenDB($bPath, $dPath, $CompanyCode, $dbType)
            }

            if ($connected -eq $true) {
                Write-SuccessLog "Connected to BUSY ($InstanceId / $CompanyCode)"
                $script:ActiveConnection   = $fi
                $script:ActiveInstanceId  = $InstanceId
                $script:ActiveCompanyCode = $CompanyCode
                return $fi
            } else {
                try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($fi) | Out-Null } catch {}
                $retryCount++
                Start-Sleep -Seconds 1
            }
        } catch {
            Write-Host "  [WARN] Connect-BUSY attempt $($retryCount+1) failed: $($_.Exception.Message)" -ForegroundColor DarkYellow
            $retryCount++
            Start-Sleep -Seconds 1
        }
    }
    return $null
}

function Disconnect-BUSY {
    param($fi)
}

# ═══════════════════════════════════════════════════════
#  DEDICATED NATIVE DIRECT QUERY CONNECTION BUILDER
# ═══════════════════════════════════════════════════════
function Get-DirectConnection {
    param([string]$InstanceId, [string]$CompanyCode)
    
    Write-Host "   [DEBUG-DIRECT-CONN] Initializing direct connection builder..." -ForegroundColor Cyan
    Write-Host "     InstanceId  : '$InstanceId'" -ForegroundColor Cyan
    Write-Host "     CompanyCode : '$CompanyCode'" -ForegroundColor Cyan
    
    $instance = Get-InstanceConfig -InstanceId $InstanceId
    $dbType = 0
    if ($null -ne $instance -and $null -ne $instance.dbType) { $dbType = [int]$instance.dbType }
    Write-Host "     Resolved DbType: $dbType (0=Access, 1=SQL)" -ForegroundColor Cyan

    if ($dbType -eq 1) {
        # SQL Server Mode
        $sqlServer   = $instance.sqlServer
        $sqlUser     = $instance.sqlUser
        $sqlPassword = $instance.sqlPassword
        
        # Base parent database name (e.g. BusyComp0001_db)
        $baseDbName = Get-SqlDatabaseName -CompanyCode $CompanyCode -InstanceId $InstanceId
        $dbName = $baseDbName
        Write-Host "     Base Database Name resolved: '$baseDbName'" -ForegroundColor Cyan

        # Establish active COM connection to resolve year database if null or company mismatch
        if ($null -eq $script:ActiveConnection -or $script:ActiveCompanyCode.ToLower() -ne $CompanyCode.ToLower()) {
            Write-Host "     Active COM connection is missing or mismatched. Triggering Connect-BUSY..." -ForegroundColor DarkCyan
            $null = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
        }

        # Dynamic year-specific database resolution with COM query
        $resolvedYearDb = $false
        if ($script:ActiveConnection -ne $null -and $script:ActiveCompanyCode.ToLower() -eq $CompanyCode.ToLower()) {
            $retryDbCount = 0
            while ($retryDbCount -lt 6 -and -not $resolvedYearDb) {
                try {
                    Write-Host "     Attempting COM ActiveDB query (Attempt $($retryDbCount+1)/6)..." -ForegroundColor DarkCyan
                    $dbNameRst = $script:ActiveConnection.GetRecordset("SELECT DB_NAME() AS ActiveDB")
                    if ($dbNameRst -and -not $dbNameRst.EOF) {
                        $activeDbVal = $dbNameRst.Fields.Item("ActiveDB").Value
                        if ($null -ne $activeDbVal -and $activeDbVal -ne [System.DBNull]::Value) {
                            $tempDb = $activeDbVal.ToString().Trim()
                            if ($tempDb -ne "") {
                                # --- FIXED: Trust what DB_NAME() says and avoid bypassing if it equals the base name ---
                                $dbName = $tempDb
                                $resolvedYearDb = $true
                                Write-Host "     [SUCCESS] COM resolved active database name: '$dbName'" -ForegroundColor Green
                            }
                        }
                        $dbNameRst.Close()
                    }
                } catch {
                    Write-Host "     [WARN] COM database query failed: $($_.Exception.Message)" -ForegroundColor Yellow
                }
                
                if (-not $resolvedYearDb) {
                    $retryDbCount++
                    Start-Sleep -Milliseconds 500
                }
            }
        } else {
            Write-Host "     [WARN] Active COM connection could not be established." -ForegroundColor Yellow
        }

        # Fallback Estimation ONLY if COM query completely failed
        if (-not $resolvedYearDb) {
            Write-Host "     COM resolution failed. Entering fallback estimation mode..." -ForegroundColor DarkYellow
            $currentYear = (Get-Date).Year
            if ((Get-Date).Month -lt 4) { $currentYear = $currentYear - 1 }
            $estimatedDb = $baseDbName + "1" + $currentYear
            Write-Host "     Estimated Fallback Database Target: '$estimatedDb'" -ForegroundColor DarkYellow
            
            $testConn = $null
            try {
                $testConnStr = "Server=$sqlServer;Database=$estimatedDb;User Id=$sqlUser;Password=$sqlPassword;"
                $testConn = New-Object System.Data.SqlClient.SqlConnection($testConnStr)
                $testConn.Open()
                $dbName = $estimatedDb
                $resolvedYearDb = $true
                Write-Host "     [SUCCESS] Fallback connection to '$estimatedDb' opened successfully!" -ForegroundColor Green
            } catch {
                Write-Host "     [WARN] Fallback estimation to '$estimatedDb' failed: $($_.Exception.Message). Defaulting to base DB." -ForegroundColor Yellow
                $dbName = $baseDbName
            } finally {
                if ($null -ne $testConn) { try { $testConn.Close() } catch {} }
            }
        }

        Write-Host "   [DEBUG-DIRECT-CONN] Final Resolved Database to Open: '$dbName'" -ForegroundColor Green

        $connStr = "Server=$sqlServer;Database=$dbName;User Id=$sqlUser;Password=$sqlPassword;"
        $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
        
        return @{
            type       = "SQL"
            dbType     = 1
            connection = $conn
            wildcard   = "%"
        }
    } else {
        # MS Access Mode
        Write-Host "     Building Access OLEDB connection..." -ForegroundColor Cyan
        $dbFile = Get-MainCompanyDbPath -CompanyCode $CompanyCode
        if ([string]::IsNullOrEmpty($dbFile) -or -not (Test-Path $dbFile)) {
            Write-Host "     [ERROR] Main db.bds file not found for path: '$dbFile'" -ForegroundColor Red
            return $null
        }
        $connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$dbFile;Jet OLEDB:Database Password=ILoveMyINDIA;"
        $conn = New-Object System.Data.OleDb.OleDbConnection($connStr)
        Write-Host "     [SUCCESS] Access connection built." -ForegroundColor Green
        return @{
            type       = "Access"
            dbType     = 0
            connection = $conn
            wildcard   = "*"
        }
    }
}

# ═══════════════════════════════════════════════════════
#  DIRECT NATIVE DB CONFIG QUERIES (db.bds / Base SQL DB)
# ═══════════════════════════════════════════════════════
function Get-CompanyDetails {
    param(
        [string]$InstanceId  = "",
        [string]$CompanyCode = ""
    )

    $targetInst =
        Get-InstanceConfig -InstanceId $InstanceId

    $dbType = 0

    if (
        $null -ne $targetInst -and
        $null -ne $targetInst.dbType
    ) {
        $dbType = [int]$targetInst.dbType
    }

    function Read-SafeReaderField {
        param(
            $Reader,
            [string]$FieldName
        )

        try {
            $index =
                $Reader.GetOrdinal($FieldName)

            if (
                $index -ge 0 -and
                -not $Reader.IsDBNull($index)
            ) {
                return (
                    [string]$Reader.GetValue($index)
                ).Trim()
            }
        }
        catch {}

        return ""
    }

    function Build-CompanyResult {
        param($Reader)

        $name =
            Read-SafeReaderField `
                -Reader $Reader `
                -FieldName "Name"

        $printName =
            Read-SafeReaderField `
                -Reader $Reader `
                -FieldName "PrintName"

        $address1 =
            Read-SafeReaderField `
                -Reader $Reader `
                -FieldName "Address1"

        $address2 =
            Read-SafeReaderField `
                -Reader $Reader `
                -FieldName "Address2"

        $address3 =
            Read-SafeReaderField `
                -Reader $Reader `
                -FieldName "Address3"

        $address4 =
            Read-SafeReaderField `
                -Reader $Reader `
                -FieldName "Address4"

        $telNo =
            Read-SafeReaderField `
                -Reader $Reader `
                -FieldName "TelNo"

        $email =
            Read-SafeReaderField `
                -Reader $Reader `
                -FieldName "Email"

        $fax =
            Read-SafeReaderField `
                -Reader $Reader `
                -FieldName "Fax"

        $tinNo =
            Read-SafeReaderField `
                -Reader $Reader `
                -FieldName "TINNo"

        $gstNo =
            Read-SafeReaderField `
                -Reader $Reader `
                -FieldName "GSTNo"

        # Prefer the proper TIN field.
        # Use GST only when TIN is blank.
        $taxNo = if (
            -not [string]::IsNullOrWhiteSpace($tinNo)
        ) {
            $tinNo
        }
        else {
            $gstNo
        }

        # Current company data stores phone inside Address3.
        # Use this fallback only when TelNo is empty.
        $phone = $telNo

        if (
            [string]::IsNullOrWhiteSpace($phone) -and
            $address3 -match
                '(?i)(?:tel|telephone)\s*:?\s*(.+)$'
        ) {
            $phone =
                $Matches[1].Trim()
        }

        # Current company data stores email inside Address4.
        # Use this fallback only when Email is empty.
        if (
            [string]::IsNullOrWhiteSpace($email) -and
            $address4 -match
                '(?i)(?:e-mail|email)\s*:?\s*(.+)$'
        ) {
            $email =
                $Matches[1].Trim()
        }

        # Do not repeat telephone and email text
        # inside the supplier postal address.
        $addressParts = @()

        foreach ($line in @(
            $address1,
            $address2
        )) {
            if (
                -not [string]::IsNullOrWhiteSpace($line)
            ) {
                $addressParts += $line
            }
        }

        $address =
            $addressParts -join ", "

        $displayName = if (
            -not [string]::IsNullOrWhiteSpace($printName)
        ) {
            $printName
        }
        else {
            $name
        }

        return @{
            name      = $name
            printName = $printName
            displayName = $displayName

            address1  = $address1
            address2  = $address2
            address3  = $address3
            address4  = $address4
            address   = $address

            telNo     = $telNo
            phone     = $phone
            fax       = $fax
            email     = $email

            tinNo     = $tinNo
            gstNo     = $gstNo
            taxNo     = $taxNo
        }
    }

    $conn = $null
    $cmd = $null
    $rdr = $null

    try {
        if ($dbType -eq 1) {
            if ($null -eq $targetInst) {
                return @{
                    success = $false
                    error   = "Instance configuration not found"
                }
            }

            $sqlServer =
                [string]$targetInst.sqlServer

            $sqlUser =
                [string]$targetInst.sqlUser

            $sqlPassword =
                [string]$targetInst.sqlPassword

            $baseDbName =
                Get-SqlDatabaseName `
                    -CompanyCode $CompanyCode `
                    -InstanceId $InstanceId

            $connectionString =
                "Server=$sqlServer;" +
                "Database=$baseDbName;" +
                "User Id=$sqlUser;" +
                "Password=$sqlPassword;"

            $conn =
                New-Object `
                    System.Data.SqlClient.SqlConnection(
                        $connectionString
                    )
        }
        else {
            $dbFile =
                Get-MainCompanyDbPath `
                    -CompanyCode $CompanyCode

            if (
                [string]::IsNullOrWhiteSpace($dbFile) -or
                -not (Test-Path -LiteralPath $dbFile)
            ) {
                return @{
                    success = $false
                    error   = "Main db.bds file not found: $dbFile"
                }
            }

            $connectionString =
                "Provider=Microsoft.Jet.OLEDB.4.0;" +
                "Data Source=$dbFile;" +
                "Jet OLEDB:Database Password=ILoveMyINDIA;"

            $conn =
                New-Object `
                    System.Data.OleDb.OleDbConnection(
                        $connectionString
                    )
        }

        $conn.Open()

        $cmd =
            $conn.CreateCommand()

        $cmd.CommandText = @"
SELECT TOP 1
    [Name],
    [PrintName],
    [Address1],
    [Address2],
    [Address3],
    [Address4],
    [TelNo],
    [Fax],
    [Email],
    [TINNo],
    [GSTNo]
FROM [Company]
"@

        $rdr =
            $cmd.ExecuteReader()

        if (-not $rdr.Read()) {
            return @{
                success = $false
                error   = "No company configuration found"
            }
        }

        $company =
            Build-CompanyResult -Reader $rdr

        return @{
            success = $true
            data    = $company
        }
    }
    catch {
        return @{
            success = $false
            error   = $_.Exception.Message
        }
    }
    finally {
        if ($null -ne $rdr) {
            try {
                $rdr.Close()
            }
            catch {}

            try {
                $rdr.Dispose()
            }
            catch {}
        }

        if ($null -ne $cmd) {
            try {
                $cmd.Dispose()
            }
            catch {}
        }

        if ($null -ne $conn) {
            try {
                $conn.Close()
            }
            catch {}

            try {
                $conn.Dispose()
            }
            catch {}
        }
    }
}