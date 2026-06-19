# modules/auth_native.ps1
# Unified Authentication & Parallel Scan Module (ASCII-Safe, Tracing Enabled)
# Supports SQL (DbType=1) and Access (DbType=0) databases

Write-Host "=========================================" -ForegroundColor Magenta
Write-Host ">>> UNIFIED AUTHENTICATION MODULE LOADING <<<" -ForegroundColor Magenta
Write-Host "=========================================" -ForegroundColor Magenta


$script:_authCache = @{}
$script:_AUTH_CACHE_TTL = 86400
$script:KEY = @(82,107,117,123,101,115,87,97,117,119,98,115,80,98,117,128,107,121,86,100)

# ===============================================================
#  DECRYPT BUSY PASSWORD
# ===============================================================
function Decrypt-BusyPassword {
    param([string]$CipherText)
    Write-Host "[DEBUG-AUTH] Decrypt-BusyPassword called with ciphertext length: $($CipherText.Length)" -ForegroundColor DarkGray
    if ([string]::IsNullOrEmpty($CipherText)) { return "" }
    try {
        $encoding = [System.Text.Encoding]::GetEncoding(1252)
        $bytes = $encoding.GetBytes($CipherText)
        $pwdLen = $bytes.Length - 5
        if ($pwdLen -le 0) { return $CipherText }
        $result = ""
        for ($i = 0; $i -lt $pwdLen; $i++) {
            $plain = ($bytes[$i] - $script:KEY[$i % 20] + 256) % 256
            $result += [char]$plain
        }
        return $result
    } catch {
        Write-Host "[DEBUG-AUTH-ERR] Decrypt-BusyPassword failed: $($_.Exception.Message)" -ForegroundColor Red
        return $CipherText
    }
}

function Get-InstanceForCompany {
    param([string]$CompanyCode, [string]$InstanceId = "")
    Write-Host "[DEBUG-AUTH] Resolving instance for CompanyCode: '$CompanyCode' (InstanceId: '$InstanceId')" -ForegroundColor DarkGray
    $instancesPath = "$PSScriptRoot\..\instances.json"
    if (-not (Test-Path $instancesPath)) { return $null }
    try {
        $config = Get-Content $instancesPath -Raw | ConvertFrom-Json
        $firstMatch = $null
        foreach ($inst in $config.instances) {
            if ($InstanceId -ne "" -and $inst.id -ne $InstanceId) { continue }
            foreach ($comp in $inst.companies) {
                if ($comp.code.ToLower() -eq $CompanyCode.ToLower()) {
                    if ($InstanceId -ne "") { return @{ instance = $inst; company = $comp } }
                    if ($null -eq $firstMatch) { $firstMatch = @{ instance = $inst; company = $comp } }
                }
            }
        }
        return $firstMatch
    } catch {
        Write-Host "[DEBUG-AUTH-ERR] Failed to read instances.json: $($_.Exception.Message)" -ForegroundColor Red
    }
    return $null
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

function Open-BdsConnection {
    param([string]$DbFile)
    Write-Host "[DEBUG-AUTH] Opening direct Access OLEDB connection to: $DbFile" -ForegroundColor DarkGray
    $connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$DbFile;Jet OLEDB:Database Password=ILoveMyINDIA;"
    $conn = New-Object System.Data.OleDb.OleDbConnection($connStr)
    $conn.Open()
    return $conn
}

function Open-SqlConnection {
    param([string]$SqlServer, [string]$Database, [string]$SqlUser, [string]$SqlPassword)
    Write-Host "[DEBUG-AUTH] Opening direct SQL connection to: Server=$SqlServer, Database=$Database" -ForegroundColor DarkGray
    $connStr = "Server=$SqlServer;Database=$Database;User Id=$SqlUser;Password=$SqlPassword;"
    $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
    $conn.Open()
    return $conn
}

function Ensure-MobileUserPreferenceTable {
    param([System.Data.SqlClient.SqlConnection]$Conn)
    try {
        $cmd = $Conn.CreateCommand()
        $cmd.CommandText = @"
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='MobileUserPreference' AND xtype='U')
CREATE TABLE MobileUserPreference (
    [Name] NVARCHAR(50) PRIMARY KEY,
    C1  INT DEFAULT 0, C2  INT DEFAULT 0, C3  INT DEFAULT 0, C4  INT DEFAULT 0,
    C5  INT DEFAULT 0, C6  INT DEFAULT 0, C7  INT DEFAULT 0, C8  INT DEFAULT 0,
    C9  INT DEFAULT 0, C10 INT DEFAULT 0,
    I1  INT DEFAULT 0, I2  INT DEFAULT 0, I3  INT DEFAULT 0, I4  INT DEFAULT 0,
    I5  INT DEFAULT 0, I6  INT DEFAULT 0, I7  INT DEFAULT 0, I8  INT DEFAULT 0,
    I9  INT DEFAULT 0, I10 INT DEFAULT 0, I11 INT DEFAULT 0, I12 INT DEFAULT 0,
    I13 INT DEFAULT 0, I14 INT DEFAULT 0,
    M1  NVARCHAR(MAX) DEFAULT '{}',
    M2  NVARCHAR(MAX) DEFAULT '{}'
)
"@
        $cmd.ExecuteNonQuery() | Out-Null
    } catch {
        Write-Host "  [WARN] Ensure-MobileUserPreferenceTable: $($_.Exception.Message)" -ForegroundColor DarkYellow
    }
}

function Read-PermissionRow {
    param($Rdr)
    $userObj = @{
        name = ""
        C1=0; C2=0; C3=0; C4=0; C5=0; C6=0; C7=0; C8=0; C9=0; C10=0
        I1=0; I2=0; I3=0; I4=0; I5=0; I6=0; I7=0; I8=0; I9=0; I10=0
        I11=0; I12=0; I13=0; I14=0
        M1="{}"; M2="{}"
    }
    foreach ($col in @($userObj.Keys)) {
        try {
            $idx = $Rdr.GetOrdinal($col)
            if ($idx -ge 0) {
                $v = $Rdr.GetValue($idx)
                if ($null -ne $v -and "$v" -ne "") { $userObj[$col] = $v.ToString().Trim() }
            }
        } catch {}
    }
    foreach ($col in @("C1","C2","C3","C4","C5","C6","C7","C8","C9","C10",
                       "I1","I2","I3","I4","I5","I6","I7","I8","I9","I10","I11","I12","I13","I14")) {
        $userObj[$col] = [int]($userObj[$col])
    }
    return $userObj
}

function Get-IsSuperUser-OLEDB {
    param([System.Data.OleDb.OleDbConnection]$Conn, [string]$Username)
    $safe = $Username -replace "'", "''"
    try {
        $cmd = $Conn.CreateCommand()
        $cmd.CommandText = "SELECT [SuperUser] FROM UserPreferences WHERE [Name]='$safe'"
        $rdr = $cmd.ExecuteReader()
        $result = $false
        if ($rdr.Read()) {
            $raw = $rdr.GetValue(0)
            $rawStr = "$raw".Trim().ToUpper()
            if ($rawStr -eq "TRUE" -or $rawStr -eq "1" -or $rawStr -eq "-1" -or $rawStr -eq "YES") { $result = $true }
        }
        $rdr.Close()
        return $result
    } catch { return $false }
}

function Get-HasMobileAccess-OLEDB {
    param([System.Data.OleDb.OleDbConnection]$Conn, [string]$Username)
    # B20 permission check removed — mobile access is always granted.
    return $true
}

function Get-IsSuperUser-SQL {
    param([System.Data.SqlClient.SqlConnection]$Conn, [string]$Username)
    try {
        $cmd = $Conn.CreateCommand()
        $cmd.CommandText = "SELECT [SuperUser] FROM UserPreferences WHERE [Name] = @user"
        $cmd.Parameters.AddWithValue("@user", $Username) | Out-Null
        $raw = $cmd.ExecuteScalar()
        if ($null -ne $raw -and "$raw" -ne "") {
            $rawStr = $raw.ToString().Trim().ToUpper()
            return ($rawStr -eq "TRUE" -or $rawStr -eq "1" -or $rawStr -eq "-1" -or $rawStr -eq "YES")
        }
    } catch {}
    return $false
}

function Get-HasMobileAccess-SQL {
    param([System.Data.SqlClient.SqlConnection]$Conn, [string]$Username)
    # B20 permission check removed — mobile access is always granted.
    return $true
}

function Get-CompanyUsers {
    param([string]$InstanceId = "", [string]$CompanyCode = "")
    $found = Get-InstanceForCompany -CompanyCode $CompanyCode -InstanceId $InstanceId
    if (-not $found) { return @{ success = $false; error = "Company not found in instances.json" } }
    $inst   = $found.instance
    $dbType = if ($null -ne $inst.dbType) { [int]$inst.dbType } else { 0 }
    $users  = @()
    if ($dbType -eq 1) {
        $sqlDb = Get-SqlDatabaseName -CompanyCode $CompanyCode -InstanceId $inst.id
        $conn  = $null
        try {
            $conn = Open-SqlConnection -SqlServer $inst.sqlServer -Database $sqlDb -SqlUser $inst.sqlUser -SqlPassword $inst.sqlPassword
            $cmd  = $conn.CreateCommand()
            $cmd.CommandText = "SELECT DISTINCT [Name] FROM UserPreferences ORDER BY [Name]"
            $rdr  = $cmd.ExecuteReader()
            while ($rdr.Read()) {
                $v = $rdr.GetValue(0)
                if ($null -ne $v -and "$v" -ne "") {
                    $name = $v.ToString().Trim()
                    if ($name -ne "" -and $name.ToLower() -ne "sa" -and $users -notcontains $name) { $users += $name }
                }
            }
            $rdr.Close()
            return @{ success = $true; data = $users }
        } catch { return @{ success = $false; error = $_.Exception.Message } } finally { if ($null -ne $conn) { try { $conn.Close() } catch {} } }
    } else {
        $dbFile = Get-MainCompanyDbPath -CompanyCode $CompanyCode
        if ([string]::IsNullOrEmpty($dbFile) -or -not (Test-Path $dbFile)) { return @{ success = $false; error = "Main database not found" } }
        $conn = $null
        try {
            $conn = Open-BdsConnection -DbFile $dbFile
            $cmd  = $conn.CreateCommand()
            $cmd.CommandText = "SELECT DISTINCT [Name] FROM UserPreferences"
            $rdr  = $cmd.ExecuteReader()
            while ($rdr.Read()) {
                $v = $rdr.GetValue(0)
                if ($null -ne $v -and "$v" -ne "") {
                    $name = $v.ToString().Trim()
                    if ($name -ne "" -and $name.ToLower() -ne "sa" -and $users -notcontains $name) { $users += $name }
                }
            }
            $rdr.Close()
            return @{ success = $true; data = $users }
        } catch { return @{ success = $false; error = $_.Exception.Message } } finally { if ($null -ne $conn) { try { $conn.Close() } catch {} } }
    }
}

function Get-UserPermissions {
    param([string]$InstanceId = "", [string]$CompanyCode = "")
    $found = Get-InstanceForCompany -CompanyCode $CompanyCode -InstanceId $InstanceId
    if (-not $found) { return @{ success = $false; error = "Company not found in instances.json" } }
    $inst   = $found.instance
    $dbType = if ($null -ne $inst.dbType) { [int]$inst.dbType } else { 0 }
    $items  = @()
    if ($dbType -eq 1) {
        $sqlDb = Get-SqlDatabaseName -CompanyCode $CompanyCode -InstanceId $inst.id
        $conn  = $null
        try {
            $conn = Open-SqlConnection -SqlServer $inst.sqlServer -Database $sqlDb -SqlUser $inst.sqlUser -SqlPassword $inst.sqlPassword
            Ensure-MobileUserPreferenceTable -Conn $conn
            $cmd = $conn.CreateCommand()
            $cmd.CommandText = "SELECT * FROM MobileUserPreference"
            $rdr = $cmd.ExecuteReader()
            while ($rdr.Read()) { $items += (Read-PermissionRow -Rdr $rdr) }
            $rdr.Close()
            return @{ success = $true; data = $items }
        } catch { return @{ success = $true; data = @() } } finally { if ($null -ne $conn) { try { $conn.Close() } catch {} } }
    } else {
        $dbFile = Get-MainCompanyDbPath -CompanyCode $CompanyCode
        if ([string]::IsNullOrEmpty($dbFile) -or -not (Test-Path $dbFile)) { return @{ success = $false; error = "Main database not found" } }
        $conn = $null
        try {
            $conn = Open-BdsConnection -DbFile $dbFile
            $cmd  = $conn.CreateCommand()
            $cmd.CommandText = "SELECT * FROM MobileUserPreference"
            $rdr = $cmd.ExecuteReader()
            while ($rdr.Read()) { $items += (Read-PermissionRow -Rdr $rdr) }
            $rdr.Close()
            return @{ success = $true; data = $items }
        } catch { return @{ success = $true; data = @() } } finally { if ($null -ne $conn) { try { $conn.Close() } catch {} } }
    }
}

function Save-UserPermissions {
    param($Data, [string]$InstanceId = "", [string]$CompanyCode = "")
    $found = Get-InstanceForCompany -CompanyCode $CompanyCode -InstanceId $InstanceId
    if (-not $found) { return @{ success = $false; error = "Company not found in instances.json" } }
    $inst   = $found.instance
    $dbType = if ($null -ne $inst.dbType) { [int]$inst.dbType } else { 0 }
    $userName = $Data.name
    if ([string]::IsNullOrEmpty($userName)) { return @{ success = $false; error = "User Name is required" } }
    $intCols = @("C1","C2","C3","C4","C5","C6","C7","C8","C9","C10","I1","I2","I3","I4","I5","I6","I7","I8","I9","I10","I11","I12","I13","I14")
    $m1Val = if ($null -ne $Data.M1) { $Data.M1 } else { "{}" }
    $m2Val = if ($null -ne $Data.M2) { $Data.M2 } else { "{}" }

    if ($dbType -eq 1) {
        $sqlDb = Get-SqlDatabaseName -CompanyCode $CompanyCode -InstanceId $inst.id
        $conn  = $null
        try {
            $conn = Open-SqlConnection -SqlServer $inst.sqlServer -Database $sqlDb -SqlUser $inst.sqlUser -SqlPassword $inst.sqlPassword
            Ensure-MobileUserPreferenceTable -Conn $conn
            $chkCmd = $conn.CreateCommand()
            $chkCmd.CommandText = "SELECT COUNT(*) FROM MobileUserPreference WHERE [Name]=@u"
            $chkCmd.Parameters.AddWithValue("@u", $userName) | Out-Null
            $exists = ([int]$chkCmd.ExecuteScalar()) -gt 0
            $writeCmd = $conn.CreateCommand()
            if ($exists) {
                $sets = ($intCols | ForEach-Object { "[$_]=@$_" }) -join ", "
                $writeCmd.CommandText = "UPDATE MobileUserPreference SET $sets, M1=@M1, M2=@M2 WHERE [Name]=@u"
            } else {
                $colList = "[Name]," + (($intCols | ForEach-Object { "[$_]" }) -join ",") + ",M1,M2"
                $valList = "@u,"    + (($intCols | ForEach-Object { "@$_"  }) -join ",") + ",@M1,@M2"
                $writeCmd.CommandText = "INSERT INTO MobileUserPreference ($colList) VALUES ($valList)"
            }
            $writeCmd.Parameters.AddWithValue("@u", $userName) | Out-Null
            foreach ($col in $intCols) {
                $val = if ($null -ne $Data.$col) { [int]$Data.$col } else { 0 }
                $writeCmd.Parameters.AddWithValue("@$col", $val) | Out-Null
            }
            $writeCmd.Parameters.AddWithValue("@M1", $m1Val) | Out-Null
            $writeCmd.Parameters.AddWithValue("@M2", $m2Val) | Out-Null
            $writeCmd.ExecuteNonQuery() | Out-Null
            return @{ success = $true; message = "User permissions updated successfully" }
        } catch { return @{ success = $false; error = $_.Exception.Message } } finally { if ($null -ne $conn) { try { $conn.Close() } catch {} } }
    } else {
        $dbFile = Get-MainCompanyDbPath -CompanyCode $CompanyCode
        if ([string]::IsNullOrEmpty($dbFile) -or -not (Test-Path $dbFile)) { return @{ success = $false; error = "Main database not found" } }
        $conn = $null
        try {
            $conn = Open-BdsConnection -DbFile $dbFile
            $safeName = $userName -replace "'", "''"
            $exists      = $false
            $tableExists = $true
            try {
                $cmd = $conn.CreateCommand()
                $cmd.CommandText = "SELECT [Name] FROM MobileUserPreference WHERE [Name]='$safeName'"
                $rdr = $cmd.ExecuteReader()
                $exists = $rdr.Read()
                $rdr.Close()
            } catch { $tableExists = $false }

            if (-not $tableExists) {
                $cmd = $conn.CreateCommand()
                $cmd.CommandText = "CREATE TABLE MobileUserPreference ([Name] TEXT(50) PRIMARY KEY, C1 INTEGER, C2 INTEGER, C3 INTEGER, C4 INTEGER, C5 INTEGER, C6 INTEGER, C7 INTEGER, C8 INTEGER, C9 INTEGER, C10 INTEGER, I1 INTEGER, I2 INTEGER, I3 INTEGER, I4 INTEGER, I5 INTEGER, I6 INTEGER, I7 INTEGER, I8 INTEGER, I9 INTEGER, I10 INTEGER, I11 INTEGER, I12 INTEGER, I13 INTEGER, I14 INTEGER, M1 MEMO, M2 MEMO)"
                $cmd.ExecuteNonQuery() | Out-Null
            }

            $cmdSchema = $conn.CreateCommand()
            $cmdSchema.CommandText = "SELECT * FROM MobileUserPreference"
            $rdrSchema = $cmdSchema.ExecuteReader([System.Data.CommandBehavior]::SchemaOnly)
            $schemaTable = $rdrSchema.GetSchemaTable()
            $existingCols = @()
            if ($null -ne $schemaTable) {
                foreach ($row in $schemaTable.Rows) { $existingCols += $row["ColumnName"].ToString().ToLower() }
            }
            $rdrSchema.Close()

            foreach ($f in @("I1","I2","I3","I4","I5","I6","I7","I8","I9","I10","I11","I12","I13","I14","M1","M2")) {
                if ($existingCols -notcontains $f.ToLower()) {
                    $colType = if ($f -like "M*") { "MEMO" } else { "INTEGER" }
                    $cmd = $conn.CreateCommand()
                    $cmd.CommandText = "ALTER TABLE MobileUserPreference ADD COLUMN $f $colType"
                    $cmd.ExecuteNonQuery() | Out-Null
                }
            }

            $updateAssignments = @()
            $insertFields      = @("[Name]")
            $insertValues      = @("'$safeName'")
            foreach ($f in $intCols) {
                $val = if ($null -ne $Data.$f) { [int]$Data.$f } else { 0 }
                $updateAssignments += "$f = $val"
                $insertFields      += $f
                $insertValues      += "$val"
            }
            $safeM1 = $m1Val -replace "'", "''"
            $safeM2 = $m2Val -replace "'", "''"
            $updateAssignments += "M1 = '$safeM1'"; $updateAssignments += "M2 = '$safeM2'"
            $insertFields += "M1"; $insertValues += "'$safeM1'"
            $insertFields += "M2"; $insertValues += "'$safeM2'"

            $cmd = $conn.CreateCommand()
            if ($exists) {
                $cmd.CommandText = "UPDATE MobileUserPreference SET " + ($updateAssignments -join ", ") + " WHERE [Name]='$safeName'"
            } else {
                $cmd.CommandText = "INSERT INTO MobileUserPreference (" + ($insertFields -join ", ") + ") VALUES (" + ($insertValues -join ", ") + ")"
            }
            $cmd.ExecuteNonQuery() | Out-Null
            return @{ success = $true; message = "User permissions updated successfully" }
        } catch { return @{ success = $false; error = $_.Exception.Message } } finally { if ($null -ne $conn) { try { $conn.Close() } catch {} } }
    }
}

# ===============================================================
#  MAIN LOGIN FUNCTION
#  InstanceId param is REQUIRED when the same company code exists
#  in multiple instances (e.g. Comp0001 in both Access + SQL).
# ===============================================================
function Invoke-BusyLogin {
    param(
        [string]$CompanyIdentifier,
        [string]$Username,
        [string]$Password,
        [string]$InstanceId = ""   # <-- NEW: pin to a specific instance
    )

    Write-Host " [BUSY-LOGIN] Executing login request..." -ForegroundColor Yellow
    Write-Host "   CompanyID  : '$CompanyIdentifier'" -ForegroundColor Yellow
    Write-Host "   User       : '$Username'" -ForegroundColor Yellow
    Write-Host "   InstanceId : '$InstanceId'" -ForegroundColor Yellow

    $instancesPath = "$PSScriptRoot\..\instances.json"
    if (-not (Test-Path $instancesPath)) { return @{ success = $false; error = "instances.json not found on server." } }

    $config     = Get-Content $instancesPath -Raw | ConvertFrom-Json
    $targetInst = $null
    $targetComp = $null

    foreach ($inst in $config.instances) {
        # If caller pinned an instanceId, skip non-matching instances
        if ($InstanceId -ne "" -and $inst.id -ne $InstanceId) { continue }
        foreach ($comp in $inst.companies) {
            if ($comp.code.ToLower() -eq $CompanyIdentifier.ToLower() -or $comp.name.ToLower() -eq $CompanyIdentifier.ToLower()) {
                $targetInst = $inst; $targetComp = $comp; break
            }
        }
        if ($targetInst) { break }
    }

    if (-not $targetInst) { return @{ success = $false; error = "Company '$CompanyIdentifier' not found." } }

    $InstanceId  = $targetInst.id
    $CompanyCode = $targetComp.code
    $CompanyName = $targetComp.name
    $dPath       = $targetInst.dataPath
    $dbType      = if ($null -ne $targetInst.dbType) { [int]$targetInst.dbType } else { 0 }
    $dbFile      = Join-Path $dPath "$CompanyCode\db.bds"
    $isAccessDb  = ($dbType -eq 0)

    Write-Host " [BUSY-LOGIN] Resolved instance '$InstanceId', dbType=$dbType, isAccessDb=$isAccessDb" -ForegroundColor Cyan

    $fi = Connect-BUSY -InstanceId $InstanceId -CompanyCode $CompanyCode
    if (-not $fi) { return @{ success = $false; error = "Could not connect to BUSY Database." } }

    try {
        $dbCipher    = $null
        $matchedUser = ""

        if ($isAccessDb) {
            $uConn = $null
            try {
                $uConn = Open-BdsConnection -DbFile $dbFile
                foreach ($table in @("UserPreferences", "Users")) {
                    $pwdCol = if ($table -eq "Users") { "Password" } else { "UserPassword" }
                    try {
                        $uCmd = $uConn.CreateCommand()
                        $uCmd.CommandText = "SELECT [Name], [$pwdCol] FROM [$table]"
                        $uRdr = $uCmd.ExecuteReader()
                        while ($uRdr.Read()) {
                            $uName = ""
                            $v = $uRdr.GetValue(0)
                            if ($null -ne $v -and "$v" -ne "") { $uName = $v.ToString().Trim() }
                            if ($uName -and $uName.ToLower() -eq $Username.ToLower()) {
                                $matchedUser = $uName
                                $vp = $uRdr.GetValue(1)
                                $dbCipher = if ($null -ne $vp -and "$vp" -ne "") { $vp.ToString() } else { "" }
                                break
                            }
                        }
                        $uRdr.Close()
                        if ($matchedUser) { break }
                    } catch {}
                }
            } finally { if ($null -ne $uConn) { try { $uConn.Close() } catch {} } }
        } else {
            $sqlConn = $null
            try {
                $dbName  = Get-SqlDatabaseName -CompanyCode $CompanyCode -InstanceId $InstanceId
                $sqlConn = Open-SqlConnection -SqlServer $targetInst.sqlServer -Database $dbName -SqlUser $targetInst.sqlUser -SqlPassword $targetInst.sqlPassword
                foreach ($table in @("UserPreferences", "Users")) {
                    $pwdCol = if ($table -eq "Users") { "Password" } else { "UserPassword" }
                    try {
                        $cmd = $sqlConn.CreateCommand()
                        $cmd.CommandText = "SELECT [Name], [$pwdCol] FROM [$table]"
                        $rdr = $cmd.ExecuteReader()
                        while ($rdr.Read()) {
                            $uName = ""
                            $v = $rdr.GetValue(0)
                            if ($null -ne $v -and "$v" -ne "") { $uName = $v.ToString().Trim() }
                            if ($uName -and $uName.ToLower() -eq $Username.ToLower()) {
                                $matchedUser = $uName
                                $vp = $rdr.GetValue(1)
                                $dbCipher = if ($null -ne $vp -and "$vp" -ne "") { $vp.ToString() } else { "" }
                                break
                            }
                        }
                        $rdr.Close()
                        if ($matchedUser) { break }
                    } catch {}
                }
            } finally { if ($null -ne $sqlConn) { try { $sqlConn.Close() } catch {} } }
        }

        if (-not $matchedUser) {
            if ($Username.ToLower() -eq "sa" -and $Password -eq "") {
                $matchedUser = "sa"; $dbCipher = ""
            } else { return @{ success = $false; error = "Invalid credentials." } }
        }

        $decryptedDbPassword = Decrypt-BusyPassword -CipherText $dbCipher
        if ($decryptedDbPassword -ne $Password -and $dbCipher -ne $Password) {
            return @{ success = $false; error = "Invalid credentials." }
        }

        $hasMobileAccess = $false
        $assignedRole    = "user"

        if ($matchedUser.ToLower() -eq "sa") {
            $hasMobileAccess = $true; $assignedRole = "superadmin"
        } elseif ($isAccessDb) {
            $bdsConn = $null
            try {
                $bdsConn = Open-BdsConnection -DbFile $dbFile
                $isSuperUser = Get-IsSuperUser-OLEDB -Conn $bdsConn -Username $matchedUser
                if ($isSuperUser) {
                    $hasMobileAccess = $true; $assignedRole = "superadmin"
                } else { $hasMobileAccess = Get-HasMobileAccess-OLEDB -Conn $bdsConn -Username $matchedUser }
            } finally { if ($null -ne $bdsConn) { try { $bdsConn.Close() } catch {} } }
        } else {
            $sqlConn = $null
            try {
                $dbName  = Get-SqlDatabaseName -CompanyCode $CompanyCode -InstanceId $InstanceId
                $sqlConn = Open-SqlConnection -SqlServer $targetInst.sqlServer -Database $dbName -SqlUser $targetInst.sqlUser -SqlPassword $targetInst.sqlPassword
                $isSuperUser = Get-IsSuperUser-SQL -Conn $sqlConn -Username $matchedUser
                if ($isSuperUser) {
                    $hasMobileAccess = $true; $assignedRole = "superadmin"
                } else { $hasMobileAccess = Get-HasMobileAccess-SQL -Conn $sqlConn -Username $matchedUser }
            } finally { if ($null -ne $sqlConn) { try { $sqlConn.Close() } catch {} } }
        }

        if (-not $hasMobileAccess) { return @{ success = $false; error = "MOBILE_ACCESS_DENIED" } }

# Default: NO transaction access. Only superadmins get full access by
# default; everyone else must be granted access explicitly via
# MobileUserPreference (set by an admin on the Permissions page).
if ($assignedRole -eq "superadmin") {
    $permissions = @{ name = $matchedUser; C1=1;C2=1;C3=1;C4=1;C5=1;C6=1;C7=1;C8=1;C9=1;C10=1;I1=1;I2=1;I3=1;I4=1;I5=1;I6=1;I7=1;I8=1;I9=1;I10=1;I11=1;I12=1;I13=1;I14=1;M1="{}"; M2="{}" }
} else {
    $permissions = @{ name = $matchedUser; C1=0;C2=0;C3=0;C4=0;C5=0;C6=0;C7=0;C8=0;C9=0;C10=0;I1=0;I2=0;I3=0;I4=0;I5=0;I6=0;I7=0;I8=0;I9=0;I10=0;I11=0;I12=0;I13=0;I14=0;M1="{}"; M2="{}" }
}

if ($assignedRole -ne "superadmin") {
            if ($isAccessDb) {
                try {
                    $safeUser = $matchedUser -replace "'", "''"
                    $prefConn = Open-BdsConnection -DbFile $dbFile
                    $prefCmd  = $prefConn.CreateCommand()
                    $prefCmd.CommandText = "SELECT * FROM MobileUserPreference WHERE [Name]='$safeUser'"
                    $prefRdr  = $prefCmd.ExecuteReader()
                    if ($prefRdr.Read()) {
                        $loaded = Read-PermissionRow -Rdr $prefRdr
                        foreach ($k in $loaded.Keys) { $permissions[$k] = $loaded[$k] }
                    }
                    $prefRdr.Close(); $prefConn.Close()
                } catch {}
            } else {
                $prefConn = $null
                try {
                    $dbName   = Get-SqlDatabaseName -CompanyCode $CompanyCode -InstanceId $InstanceId
                    $prefConn = Open-SqlConnection -SqlServer $targetInst.sqlServer -Database $dbName -SqlUser $targetInst.sqlUser -SqlPassword $targetInst.sqlPassword
                    Ensure-MobileUserPreferenceTable -Conn $prefConn
                    $prefCmd  = $prefConn.CreateCommand()
                    $prefCmd.CommandText = "SELECT * FROM MobileUserPreference WHERE [Name]=@u"
                    $prefCmd.Parameters.AddWithValue("@u", $matchedUser) | Out-Null
                    $prefRdr  = $prefCmd.ExecuteReader()
                    if ($prefRdr.Read()) {
                        $loaded = Read-PermissionRow -Rdr $prefRdr
                        foreach ($k in $loaded.Keys) { $permissions[$k] = $loaded[$k] }
                    }
                    $prefRdr.Close()
                } catch {} finally { if ($null -ne $prefConn) { try { $prefConn.Close() } catch {} } }
            }
        }

        $token   = [guid]::NewGuid().ToString()
        $userDoc = @{ uid = $matchedUser; name = $matchedUser; role = $assignedRole; companyCode = $CompanyCode; companyName = $CompanyName; instanceId = $InstanceId; instanceName = $targetInst.displayName; permissions = $permissions }
        $script:_authCache[$token] = @{ user = $userDoc; expiry = (Get-Date).AddSeconds($script:_AUTH_CACHE_TTL) }

        Write-Host " [BUSY-LOGIN] Login successful for user: $matchedUser on company: $CompanyName ($InstanceId)" -ForegroundColor Green
        return @{ success = $true; token = $token; user = $userDoc }
    } catch {
        Write-Host "[BUSY-LOGIN-ERR] General error inside Invoke-BusyLogin: $($_.Exception.Message)" -ForegroundColor Red
        return @{ success = $false; error = "Database error: $($_.Exception.Message)" }
    } finally { Disconnect-BUSY $fi }
}

# ===============================================================
#  AUTH CHECK
# ===============================================================
function Invoke-AuthCheck {
    param($Request, [string]$InstanceId = "", [string]$CompanyCode = "")
    $authHeader = $Request.Headers["Authorization"]
    if (-not $authHeader -or -not $authHeader.StartsWith("Bearer ")) { return @{ allowed = $false; reason = "Missing or invalid token" } }
    $token = $authHeader.Substring(7)
    if ($script:_authCache.ContainsKey($token)) {
        $entry = $script:_authCache[$token]
        if ((Get-Date) -lt $entry.expiry) {
            if ($CompanyCode -and $entry.user.companyCode -ne $CompanyCode) { return @{ allowed = $false; reason = "Token is not valid for company" } }
            return @{ allowed = $true; user = $entry.user }
        } else { $script:_authCache.Remove($token); return @{ allowed = $false; reason = "Token expired" } }
    }
    return @{ allowed = $false; reason = "Invalid token" }
}

# ===============================================================
#  PARALLEL SCAN LOGIN
# ===============================================================
function Invoke-ScanLogin {
    param(
        [string]$Username,
        [string]$Password
    )

    Write-Host " [SCAN-LOGIN] Initiating scan for user: '$Username'..." -ForegroundColor Cyan

    $instancesPath = "$PSScriptRoot\..\instances.json"
    if (-not (Test-Path $instancesPath)) {
        Write-Host "[SCAN-LOGIN-ERR] instances.json not found on disk" -ForegroundColor Red
        return @{ success = $false; error = "instances.json not found on server." }
    }

    $config = Get-Content $instancesPath -Raw | ConvertFrom-Json
    $accessTargets = [System.Collections.ArrayList]@()
    $sqlTargets    = [System.Collections.ArrayList]@()

    foreach ($inst in $config.instances) {
        $dbType = if ($null -ne $inst.dbType) { [int]$inst.dbType } else { 0 }
        foreach ($comp in $inst.companies) {
            $cc     = [string]$comp.code
            $dp     = [string]$inst.dataPath
            $dbFile = Join-Path $dp "$cc\db.bds"

            if ($dbType -eq 0) {
                [void]$accessTargets.Add(@{
                    instanceId  = [string]$inst.id
                    displayName = [string]$inst.displayName
                    dataPath    = $dp
                    companyCode = $cc
                    companyName = [string]$comp.name
                    dbFile      = $dbFile
                })
            } else {
                $sqlDb = ""
                if ($comp.sqlDatabase -and $comp.sqlDatabase -ne "") {
                    $sqlDb = [string]$comp.sqlDatabase
                } elseif ($cc.ToUpper().StartsWith("COMP")) {
                    $sqlDb = "BusyComp" + $cc.Substring(4) + "_db"
                } else {
                    $clean = $cc.Trim()
                    if ($clean.Length -gt 1) { $clean = [char]::ToUpper($clean[0]) + $clean.Substring(1).ToLower() }
                    $sqlDb = "Busy" + $clean + "_db"
                }
                [void]$sqlTargets.Add(@{
                    instanceId  = [string]$inst.id
                    displayName = [string]$inst.displayName
                    sqlServer   = [string]$inst.sqlServer
                    sqlUser     = [string]$inst.sqlUser
                    sqlPassword = [string]$inst.sqlPassword
                    sqlDb       = $sqlDb
                    companyCode = $cc
                    companyName = [string]$comp.name
                })
            }
        }
    }

    Write-Host " [SCAN-LOGIN] Targets Resolved: Access=$($accessTargets.Count), SQL=$($sqlTargets.Count)" -ForegroundColor Cyan
    if ($accessTargets.Count -eq 0 -and $sqlTargets.Count -eq 0) {
        return @{ success = $false; error = "No companies configured on this server." }
    }

    $KEY = @(82,107,117,123,101,115,87,97,117,119,98,115,80,98,117,128,107,121,86,100)
    function Local-Decrypt([string]$C) {
        if ([string]::IsNullOrEmpty($C)) { return "" }
        try {
            $enc = [System.Text.Encoding]::GetEncoding(1252)
            $b   = $enc.GetBytes($C)
            $len = $b.Length - 5
            if ($len -le 0) { return $C }
            $r = ""
            for ($i = 0; $i -lt $len; $i++) {
                $r += [char](($b[$i] - $KEY[$i % 20] + 256) % 256)
            }
            return $r
        } catch { return $C }
    }

    $hitList    = [System.Collections.ArrayList]@()
    $deniedList = [System.Collections.ArrayList]@()

    # ── Phase 1: Access companies on main thread ──────────────────────────────
    foreach ($target in $accessTargets) {
        $cc     = $target.companyCode
        $dbFile = $target.dbFile
        Write-Host "   [SCAN-ACCESS] Checking $($target.instanceId)/$cc" -ForegroundColor DarkGray

        if (-not (Test-Path $dbFile)) {
            Write-Host "   [SCAN-ACCESS-WARN] File missing: $dbFile" -ForegroundColor Yellow
            continue
        }

        $mUser  = ""
        $cipher = ""
        try {
            $cs   = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$dbFile;Jet OLEDB:Database Password=ILoveMyINDIA;"
            $conn = New-Object System.Data.OleDb.OleDbConnection($cs)
            $conn.Open()
            try {
                foreach ($tbl in @("UserPreferences", "Users")) {
                    $pc = if ($tbl -eq "Users") { "Password" } else { "UserPassword" }
                    try {
                        $cmd = $conn.CreateCommand()
                        $cmd.CommandText = "SELECT [Name], [$pc] FROM [$tbl]"
                        $rdr = $cmd.ExecuteReader()
                        while ($rdr.Read()) {
                            $v = $rdr.GetValue(0)
                            $n = if ($null -ne $v -and "$v" -ne "") { $v.ToString().Trim() } else { "" }
                            if ($n -and $n.ToLower() -eq $Username.ToLower()) {
                                $mUser  = $n
                                $vp     = $rdr.GetValue(1)
                                $cipher = if ($null -ne $vp -and "$vp" -ne "") { $vp.ToString() } else { "" }
                                break
                            }
                        }
                        $rdr.Close()
                        if ($mUser) { break }
                    } catch {}
                }
            } finally { $conn.Close() }
        } catch {
            Write-Host "   [SCAN-ACCESS-WARN] Database connection failed for $cc" -ForegroundColor Yellow
            continue
        }

        if (-not $mUser) { continue }

        if ($Username.ToLower() -eq "sa" -and $Password -eq "" -and $mUser.ToLower() -eq "sa") {
            [void]$hitList.Add(@{ instanceId = $target.instanceId; companyCode = $cc; companyName = $target.companyName; matchedUser = $mUser; isAccessDb = $true })
            continue
        }

        $dec = Local-Decrypt $cipher
        if ($dec -ne $Password -and $cipher -ne $Password) { continue }

        # B20 permission check removed — mobile access is always granted
        # for any user whose credentials matched above. SuperUser status
        # is still checked separately later (in Invoke-BusyLogin) for role
        # assignment, but no longer gates mobile access itself.
        $hasMobile = $true

        $entry = @{ instanceId = $target.instanceId; companyCode = $cc; companyName = $target.companyName; matchedUser = $mUser; isAccessDb = $true }
        if ($hasMobile) { [void]$hitList.Add($entry) } else { [void]$deniedList.Add($entry) }
    }

    # ── Phase 2: SQL companies in parallel runspaces ──────────────────────────
    if ($sqlTargets.Count -gt 0) {
        Write-Host "   [SCAN-SQL] Launching parallel workers..." -ForegroundColor Yellow
        $pool = [RunspaceFactory]::CreateRunspacePool(1, [Math]::Max($sqlTargets.Count, 1))
        $pool.Open()
        $jobs = [System.Collections.ArrayList]@()

        $sqlWorker = {
            param($Target, $Username, $Password)

            $KEY = @(82,107,117,123,101,115,87,97,117,119,98,115,80,98,117,128,107,121,86,100)
            function Local-Decrypt([string]$C) {
                if ([string]::IsNullOrEmpty($C)) { return "" }
                try {
                    $enc = [System.Text.Encoding]::GetEncoding(1252)
                    $b   = $enc.GetBytes($C)
                    $len = $b.Length - 5
                    if ($len -le 0) { return $C }
                    $r = ""
                    for ($i = 0; $i -lt $len; $i++) {
                        $r += [char](($b[$i] - $KEY[$i % 20] + 256) % 256)
                    }
                    return $r
                } catch { return $C }
            }

            $iid = [string]$Target.instanceId
            $cc  = [string]$Target.companyCode
            $cn  = [string]$Target.companyName
            $sv  = [string]$Target.sqlServer
            $su  = [string]$Target.sqlUser
            $sp  = [string]$Target.sqlPassword
            $db  = [string]$Target.sqlDb

            if ([string]::IsNullOrEmpty($db)) { return $null }

            $mUser  = ""
            $cipher = ""
            try {
                $conn = New-Object System.Data.SqlClient.SqlConnection("Server=$sv;Database=$db;User Id=$su;Password=$sp;Connect Timeout=5;")
                $conn.Open()
                try {
                    foreach ($tbl in @("UserPreferences", "Users")) {
                        $pc = if ($tbl -eq "Users") { "Password" } else { "UserPassword" }
                        try {
                            $cmd = $conn.CreateCommand()
                            $cmd.CommandText = "SELECT [Name], [$pc] FROM [$tbl]"
                            $rdr = $cmd.ExecuteReader()
                            while ($rdr.Read()) {
                                $v = $rdr.GetValue(0)
                                $n = if ($null -ne $v -and "$v" -ne "") { $v.ToString().Trim() } else { "" }
                                if ($n -and $n.ToLower() -eq $Username.ToLower()) {
                                    $mUser  = $n
                                    $vp     = $rdr.GetValue(1)
                                    $cipher = if ($null -ne $vp -and "$vp" -ne "") { $vp.ToString() } else { "" }
                                    break
                                }
                            }
                            $rdr.Close()
                            if ($mUser) { break }
                        } catch {}
                    }
                } finally { $conn.Close() }
            } catch { return $null }

            if (-not $mUser) { return $null }

            if ($Username.ToLower() -eq "sa" -and $Password -eq "" -and $mUser.ToLower() -eq "sa") {
                return @{ matched=$true; accessDenied=$false; instanceId=$iid; companyCode=$cc; companyName=$cn; matchedUser=$mUser; isAccessDb=$false }
            }

            $dec = Local-Decrypt $cipher
            if ($dec -ne $Password -and $cipher -ne $Password) { return $null }

            # B20 permission check removed — mobile access is always granted.
            $hasMobile = $true

            return @{ matched=$true; accessDenied=(-not $hasMobile); instanceId=$iid; companyCode=$cc; companyName=$cn; matchedUser=$mUser; isAccessDb=$false }
        }

        foreach ($target in $sqlTargets) {
            $ps = [PowerShell]::Create()
            $ps.RunspacePool = $pool
            [void]$ps.AddScript($sqlWorker)
            [void]$ps.AddArgument($target)
            [void]$ps.AddArgument($Username)
            [void]$ps.AddArgument($Password)
            [void]$jobs.Add(@{ PS = $ps; Handle = $ps.BeginInvoke() })
        }

        foreach ($job in $jobs) {
            try {
                $rs = $job.PS.EndInvoke($job.Handle)
                if ($null -ne $rs -and $rs.Count -gt 0) {
                    $r = $rs[0]
                    if ($null -ne $r -and $r.matched -eq $true) {
                        if ($r.accessDenied -eq $true) { [void]$deniedList.Add($r) } else { [void]$hitList.Add($r) }
                    }
                }
            } catch {
                Write-Host "   [SCAN-SQL-ERR] Job failed: $($_.Exception.Message)" -ForegroundColor Red
            } finally { $job.PS.Dispose() }
        }
        $pool.Close(); $pool.Dispose()
    }

    Write-Host " [SCAN-LOGIN] Results: hitList=$($hitList.Count), deniedList=$($deniedList.Count)" -ForegroundColor Cyan

    if ($hitList.Count -eq 0 -and $deniedList.Count -eq 0) {
        return @{ success = $false; error = "Invalid username or password." }
    }
    if ($hitList.Count -eq 0 -and $deniedList.Count -gt 0) {
        return @{ success = $false; error = "MOBILE_ACCESS_DENIED" }
    }
    if ($hitList.Count -gt 1) {
        $nlist = [System.Collections.ArrayList]@()
        foreach ($h in $hitList) { [void]$nlist.Add([string]$h.companyName) }
        $names = $nlist -join ", "
        return @{ success = $false; error = "AMBIGUOUS_CREDENTIALS"; detail = "Credentials match multiple companies: $names" }
    }

    $hit     = [hashtable]$hitList[0]
    $hitCode = [string]$hit.companyCode
    $hitInst = [string]$hit.instanceId
    $hitUser = [string]$hit.matchedUser

    Write-Host " [SCAN-LOGIN] Target resolved: Instance=$hitInst, Company=$($hit.companyName) ($hitCode), User='$hitUser'. Executing direct logon..." -ForegroundColor Green

    # Pass instanceId so Invoke-BusyLogin pins to the correct instance
    return Invoke-BusyLogin -CompanyIdentifier $hitCode -Username $hitUser -Password $Password -InstanceId $hitInst
}

Write-Host ">>> UNIFIED AUTHENTICATION MODULE FULLY LOADED <<<" -ForegroundColor Green
