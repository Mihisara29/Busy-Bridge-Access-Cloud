# modules/web_approval.ps1
# ============================================================================
# BUSYCLOUD WEB APPROVAL - PHASE 1 FOUNDATION
# ============================================================================
# Purpose
# -------
# 1. Automatically creates the complete Web Approval support tables when missing.
#    This baseline build does NOT modify existing table structures.
# 2. Works with BOTH BUSY SQL Server and Access/BDS companies.
# 3. Stores sales workflow role in two existing Boolean-style fields:
#       B35 = Salesman flag
#       B36 = Sales Manager flag
#
#    Valid combinations:
#       B35=0, B36=0 -> NONE
#       B35=1, B36=0 -> SALESMAN
#       B35=0, B36=1 -> SALES_MANAGER
#
#    B35=1, B36=1 is treated as invalid and is never written by this module.
# 4. Stores Salesman <-> Sales Manager relationships in the PERMANENT company DB.
# 5. Stores Web Approval transactions/audit/sync/notifications in the ACTIVE
#    FINANCIAL-YEAR database.
#
# IMPORTANT
# ---------
# This module NEVER creates accounting vouchers by writing Tran1/Tran2.
# When Phase 3 adds synchronization, the real BUSY voucher will still be
# created only through the existing BUSY COM Create-Voucher pipeline.
# ============================================================================

# ============================================================================
# BUSYCLOUD WEB APPROVAL MODULE VERSION
# ============================================================================
$script:BusyCloudWebApprovalModuleVersion = "6.2-auto-push-bootstrap-npm-fix"
Write-Host "  [WEB-APPROVAL] Module version $script:BusyCloudWebApprovalModuleVersion loaded." -ForegroundColor DarkCyan

$script:WebApprovalSchemaVersion = 1

$script:WebApprovalRoleNone         = 0
$script:WebApprovalRoleSalesman     = 1
$script:WebApprovalRoleSalesManager = 2

$script:WebApprovalSupportedVoucherTypes = @(26, 12, 9, 3, 14, 11)

function ConvertTo-WebApprovalRoleValue {
    param($Value)

    if ($null -eq $Value) { return 0 }

    $text = ([string]$Value).Trim().ToUpperInvariant()

    switch ($text) {
        "0"             { return 0 }
        "NONE"          { return 0 }
        "1"             { return 1 }
        "SALESMAN"      { return 1 }
        "SALES_MAN"     { return 1 }
        "2"             { return 2 }
        "SALES_MANAGER" { return 2 }
        "SALESMANAGER"  { return 2 }
        "SALES MANAGER" { return 2 }
        default         { return -1 }
    }
}

function Get-WebApprovalRoleName {
    param([int]$Role)

    switch ($Role) {
        1 { return "SALESMAN" }
        2 { return "SALES_MANAGER" }
        default { return "NONE" }
    }
}

function Test-WebApprovalSupportedVoucherType {
    param([int]$VchType)
    return @($script:WebApprovalSupportedVoucherTypes) -contains $VchType
}

function Close-WebApprovalDbContext {
    param($Context)

    if ($null -eq $Context) { return }

    try {
        if ($Context.connection) {
            $Context.connection.Close()
        }
    } catch {}

    try {
        if ($Context.connection) {
            $Context.connection.Dispose()
        }
    } catch {}
}

function Get-WebApprovalPermanentDbContext {
    param(
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $found = Get-InstanceForCompany -CompanyCode $CompanyCode -InstanceId $InstanceId
    if (-not $found) {
        throw "Company not found in instances.json: $InstanceId/$CompanyCode"
    }

    $inst = $found.instance
    $comp = $found.company
    $dbType = if ($null -ne $inst.dbType) { [int]$inst.dbType } else { 0 }

    if ($dbType -eq 1) {
        $dbName = Get-SqlDatabaseName `
            -CompanyCode ([string]$comp.code) `
            -InstanceId ([string]$inst.id)

        if ([string]::IsNullOrWhiteSpace($dbName)) {
            throw "Could not resolve permanent SQL database for $InstanceId/$CompanyCode."
        }

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
            scope = "PERMANENT"
        }
    }

    $companyFolder = Join-Path ([string]$inst.dataPath) ([string]$comp.code)
    $dbFile = Join-Path $companyFolder "db.bds"

    if (-not (Test-Path -LiteralPath $dbFile)) {
        throw "Permanent Access/BDS database not found: $dbFile"
    }

    $conn = Open-BdsConnection -DbFile $dbFile

    return @{
        dbType = 0
        connection = $conn
        instance = $inst
        company = $comp
        database = $dbFile
        scope = "PERMANENT"
    }
}

function Get-WebApprovalFiscalDbContext {
    param(
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    # vch_setting.ps1 is loaded before this module. It already owns the tested
    # active-fiscal-database resolver for both SQL and Access/BDS.
    if (-not (Get-Command Get-BusyCloudFastConfigDbContext -ErrorAction SilentlyContinue)) {
        throw "Get-BusyCloudFastConfigDbContext is not loaded. Load vch_setting.ps1 before web_approval.ps1."
    }

    $ctx = Get-BusyCloudFastConfigDbContext `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    $ctx["scope"] = "FISCAL"
    return $ctx
}

function Test-WebApprovalTableExists {
    param(
        $Connection,
        [int]$DbType,
        [string]$TableName
    )

    if ($DbType -eq 1) {
        $cmd = $Connection.CreateCommand()
        $cmd.CommandText = @"
SELECT COUNT(*)
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_SCHEMA='dbo' AND TABLE_NAME=@name
"@
        [void]$cmd.Parameters.AddWithValue("@name", $TableName)
        return ([int]$cmd.ExecuteScalar() -gt 0)
    }

    try {
        $schema = $Connection.GetSchema("Tables")
        foreach ($row in $schema.Rows) {
            $name = [string]$row["TABLE_NAME"]
            $type = [string]$row["TABLE_TYPE"]
            if (
                $name.Trim().Equals($TableName, [System.StringComparison]::OrdinalIgnoreCase) -and
                ($type -eq "TABLE" -or $type -eq "BASE TABLE")
            ) {
                return $true
            }
        }
    } catch {}

    # Safe fallback for Jet/OLEDB providers that expose a limited schema.
    try {
        $cmd = $Connection.CreateCommand()
        $cmd.CommandText = "SELECT TOP 1 * FROM [$TableName]"
        $rdr = $cmd.ExecuteReader()
        try { $rdr.Close() } catch {}
        return $true
    } catch {
        return $false
    }
}

function Test-WebApprovalColumnExists {
    param(
        $Connection,
        [int]$DbType,
        [string]$TableName,
        [string]$ColumnName
    )

    if ($DbType -eq 1) {
        $cmd = $Connection.CreateCommand()
        $cmd.CommandText = @"
SELECT COUNT(*)
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA='dbo' AND TABLE_NAME=@tableName AND COLUMN_NAME=@columnName
"@
        [void]$cmd.Parameters.AddWithValue("@tableName", $TableName)
        [void]$cmd.Parameters.AddWithValue("@columnName", $ColumnName)
        return ([int]$cmd.ExecuteScalar() -gt 0)
    }

    try {
        $cmd = $Connection.CreateCommand()
        $cmd.CommandText = "SELECT * FROM [$TableName]"
        $rdr = $cmd.ExecuteReader([System.Data.CommandBehavior]::SchemaOnly)
        $schema = $rdr.GetSchemaTable()

        if ($null -ne $schema) {
            foreach ($row in $schema.Rows) {
                $name = [string]$row["ColumnName"]
                if ($name.Trim().Equals($ColumnName, [System.StringComparison]::OrdinalIgnoreCase)) {
                    try { $rdr.Close() } catch {}
                    return $true
                }
            }
        }

        try { $rdr.Close() } catch {}
    } catch {}

    return $false
}

function Invoke-WebApprovalNonQuery {
    param(
        $Connection,
        [string]$Sql,
        $Transaction = $null
    )

    $cmd = $Connection.CreateCommand()
    if ($null -ne $Transaction) {
        $cmd.Transaction = $Transaction
    }
    $cmd.CommandText = $Sql
    return [int]$cmd.ExecuteNonQuery()
}

function Ensure-WebApprovalRoleColumn {
    param($Context)

    $conn = $Context.connection
    $dbType = [int]$Context.dbType

    if (-not (Test-WebApprovalTableExists -Connection $conn -DbType $dbType -TableName "MobileUserPreference")) {
        throw "MobileUserPreference table does not exist in permanent database '$($Context.database)'."
    }

    foreach ($columnName in @("B35", "B36")) {
        if (-not (Test-WebApprovalColumnExists `
            -Connection $conn `
            -DbType $dbType `
            -TableName "MobileUserPreference" `
            -ColumnName $columnName)) {

            throw (
                "MobileUserPreference.{0} is required for Web Approval workflow roles " +
                "but does not exist in permanent database '{1}'."
            ) -f $columnName, $Context.database
        }
    }

    return $true
}

function Get-WebApprovalRoleFromFlags {
    param(
        $SalesmanFlag,
        $ManagerFlag
    )

    $isSalesman = $false
    $isManager = $false

    try { $isSalesman = ([int]$SalesmanFlag -ne 0) } catch {}
    try { $isManager = ([int]$ManagerFlag -ne 0) } catch {}

    if ($isSalesman -and -not $isManager) {
        return $script:WebApprovalRoleSalesman
    }

    if ($isManager -and -not $isSalesman) {
        return $script:WebApprovalRoleSalesManager
    }

    # NONE, or invalid legacy state where both flags are ON.
    return $script:WebApprovalRoleNone
}

function Ensure-WebApprovalPermanentTables {
    param($Context)

    $conn = $Context.connection
    $dbType = [int]$Context.dbType

    Ensure-WebApprovalRoleColumn -Context $Context

    if (-not (Test-WebApprovalTableExists -Connection $conn -DbType $dbType -TableName "BusyCloudSalesmanManager")) {
        if ($dbType -eq 1) {
            $sql = @"
CREATE TABLE dbo.BusyCloudSalesmanManager (
    Id NVARCHAR(36) NOT NULL PRIMARY KEY,
    SalesmanUserName NVARCHAR(100) NOT NULL,
    ManagerUserName NVARCHAR(100) NOT NULL,
    IsActive BIT NOT NULL CONSTRAINT DF_BCSM_IsActive DEFAULT(1),
    AssignedBy NVARCHAR(100) NULL,
    AssignedAt DATETIME2 NOT NULL CONSTRAINT DF_BCSM_AssignedAt DEFAULT(SYSUTCDATETIME()),
    UpdatedAt DATETIME2 NOT NULL CONSTRAINT DF_BCSM_UpdatedAt DEFAULT(SYSUTCDATETIME())
)
"@
        }
        else {
            $sql = @"
CREATE TABLE [BusyCloudSalesmanManager] (
    [Id] TEXT(36) NOT NULL,
    [SalesmanUserName] TEXT(100) NOT NULL,
    [ManagerUserName] TEXT(100) NOT NULL,
    [IsActive] BYTE,
    [AssignedBy] TEXT(100),
    [AssignedAt] DATETIME,
    [UpdatedAt] DATETIME,
    CONSTRAINT [PK_BusyCloudSalesmanManager] PRIMARY KEY ([Id])
)
"@
        }

        [void](Invoke-WebApprovalNonQuery -Connection $conn -Sql $sql)
        Write-Host "  [WEB-APPROVAL MIGRATION] Created BusyCloudSalesmanManager in $($Context.database)" -ForegroundColor Yellow
    }

    # Idempotent indexes. Failures are harmless when they already exist.
    try {
        if ($dbType -eq 1) {
            $sql = @"
IF NOT EXISTS (
    SELECT 1 FROM sys.indexes
    WHERE name='UX_BusyCloudSalesmanManager_Pair'
      AND object_id=OBJECT_ID('dbo.BusyCloudSalesmanManager')
)
CREATE UNIQUE INDEX UX_BusyCloudSalesmanManager_Pair
ON dbo.BusyCloudSalesmanManager(SalesmanUserName, ManagerUserName)
"@
            [void](Invoke-WebApprovalNonQuery -Connection $conn -Sql $sql)
        }
        else {
            [void](Invoke-WebApprovalNonQuery -Connection $conn -Sql "CREATE UNIQUE INDEX UX_BusyCloudSalesmanManager_Pair ON BusyCloudSalesmanManager (SalesmanUserName, ManagerUserName)")
        }
    } catch {}

    if (-not (Test-WebApprovalTableExists -Connection $conn -DbType $dbType -TableName "BusyCloudPushSubscription")) {
        if ($dbType -eq 1) {
            $sql = @"
CREATE TABLE dbo.BusyCloudPushSubscription (
    Id NVARCHAR(36) NOT NULL PRIMARY KEY,
    UserName NVARCHAR(100) NOT NULL,
    EndpointHash CHAR(64) NOT NULL,
    Endpoint NVARCHAR(MAX) NOT NULL,
    P256dh NVARCHAR(255) NOT NULL,
    AuthKey NVARCHAR(255) NOT NULL,
    DeviceName NVARCHAR(150) NULL,
    UserAgent NVARCHAR(500) NULL,
    IsActive BIT NOT NULL CONSTRAINT DF_BCPS_IsActive DEFAULT(1),
    CreatedAt DATETIME2 NOT NULL,
    UpdatedAt DATETIME2 NOT NULL,
    LastSeenAt DATETIME2 NULL
)
"@
        }
        else {
            $sql = @"
CREATE TABLE [BusyCloudPushSubscription] (
    [Id] TEXT(36) NOT NULL,
    [UserName] TEXT(100) NOT NULL,
    [EndpointHash] TEXT(64) NOT NULL,
    [Endpoint] MEMO,
    [P256dh] TEXT(255) NOT NULL,
    [AuthKey] TEXT(255) NOT NULL,
    [DeviceName] TEXT(150),
    [UserAgent] MEMO,
    [IsActive] BYTE,
    [CreatedAt] DATETIME,
    [UpdatedAt] DATETIME,
    [LastSeenAt] DATETIME,
    CONSTRAINT [PK_BusyCloudPushSubscription] PRIMARY KEY ([Id])
)
"@
        }

        [void](Invoke-WebApprovalNonQuery -Connection $conn -Sql $sql)
        Write-Host "  [WEB-APPROVAL MIGRATION] Created BusyCloudPushSubscription in $($Context.database)" -ForegroundColor Yellow
    }

    try {
        if ($dbType -eq 1) {
            $sql = @"
IF NOT EXISTS (
    SELECT 1 FROM sys.indexes
    WHERE name='UX_BusyCloudPushSubscription_EndpointHash'
      AND object_id=OBJECT_ID('dbo.BusyCloudPushSubscription')
)
CREATE UNIQUE INDEX UX_BusyCloudPushSubscription_EndpointHash
ON dbo.BusyCloudPushSubscription(EndpointHash)
"@
            [void](Invoke-WebApprovalNonQuery -Connection $conn -Sql $sql)

            $sql = @"
IF NOT EXISTS (
    SELECT 1 FROM sys.indexes
    WHERE name='IX_BusyCloudPushSubscription_User'
      AND object_id=OBJECT_ID('dbo.BusyCloudPushSubscription')
)
CREATE INDEX IX_BusyCloudPushSubscription_User
ON dbo.BusyCloudPushSubscription(UserName, IsActive)
"@
            [void](Invoke-WebApprovalNonQuery -Connection $conn -Sql $sql)
        }
        else {
            try {
                [void](Invoke-WebApprovalNonQuery -Connection $conn -Sql "CREATE UNIQUE INDEX [UX_BusyCloudPushSubscription_EndpointHash] ON [BusyCloudPushSubscription] ([EndpointHash])")
            } catch {}

            try {
                [void](Invoke-WebApprovalNonQuery -Connection $conn -Sql "CREATE INDEX [IX_BusyCloudPushSubscription_User] ON [BusyCloudPushSubscription] ([UserName], [IsActive])")
            } catch {}
        }
    }
    catch {}

    return $true
}

# ============================================================================
# BASELINE SCHEMA POLICY
# ============================================================================
# Fresh databases:
#   The CREATE TABLE statements below already include the complete current
#   Phase 3.1 number-audit schema.
#
# Existing databases:
#   Tables are only checked for existence and are NOT altered by this baseline
#   module. Current databases should therefore already have the Phase 3.1
#   columns before switching to this file.
# ============================================================================

function Ensure-WebApprovalFiscalTables {
    param($Context)

    $conn = $Context.connection
    $dbType = [int]$Context.dbType

    if (-not (Test-WebApprovalTableExists -Connection $conn -DbType $dbType -TableName "BusyCloudWebApproval")) {
        if ($dbType -eq 1) {
            $sql = @"
CREATE TABLE dbo.BusyCloudWebApproval (
    Id NVARCHAR(36) NOT NULL PRIMARY KEY,
    SubmittedBy NVARCHAR(100) NOT NULL,
    VoucherType INT NOT NULL,
    VoucherDate DATETIME2 NOT NULL,
    RequestedSeries NVARCHAR(100) NULL,
    RequestedVoucherNo NVARCHAR(100) NULL,
    PartyName NVARCHAR(255) NULL,
    Amount DECIMAL(19,4) NOT NULL CONSTRAINT DF_BCWA_Amount DEFAULT(0),
    PayloadJson NVARCHAR(MAX) NOT NULL,
    PayloadHash CHAR(64) NOT NULL,
    ApprovalStatus NVARCHAR(20) NOT NULL,
    SyncStatus NVARCHAR(30) NOT NULL,
    DecisionBy NVARCHAR(100) NULL,
    DecisionAt DATETIME2 NULL,
    RejectionReason NVARCHAR(MAX) NULL,
    BusyVoucherNo NVARCHAR(100) NULL,
    BusyVoucherCode NVARCHAR(50) NULL,
    NumberChangeReason NVARCHAR(MAX) NULL,
    NumberResolvedAt DATETIME2 NULL,
    SyncAttempts INT NOT NULL CONSTRAINT DF_BCWA_SyncAttempts DEFAULT(0),
    LastSyncError NVARCHAR(MAX) NULL,
    SubmittedAt DATETIME2 NOT NULL,
    UpdatedAt DATETIME2 NOT NULL,
    SyncedAt DATETIME2 NULL,
    Version INT NOT NULL CONSTRAINT DF_BCWA_Version DEFAULT(1)
)
"@
        }
        else {
            $sql = @"
CREATE TABLE [BusyCloudWebApproval] (
    [Id] TEXT(36) NOT NULL,
    [SubmittedBy] TEXT(100) NOT NULL,
    [VoucherType] INTEGER NOT NULL,
    [VoucherDate] DATETIME NOT NULL,
    [RequestedSeries] TEXT(100),
    [RequestedVoucherNo] TEXT(100),
    [PartyName] TEXT(255),
    [Amount] DOUBLE,
    [PayloadJson] MEMO,
    [PayloadHash] TEXT(64) NOT NULL,
    [ApprovalStatus] TEXT(20) NOT NULL,
    [SyncStatus] TEXT(30) NOT NULL,
    [DecisionBy] TEXT(100),
    [DecisionAt] DATETIME,
    [RejectionReason] MEMO,
    [BusyVoucherNo] TEXT(100),
    [BusyVoucherCode] TEXT(50),
    [NumberChangeReason] MEMO,
    [NumberResolvedAt] DATETIME,
    [SyncAttempts] LONG,
    [LastSyncError] MEMO,
    [SubmittedAt] DATETIME,
    [UpdatedAt] DATETIME,
    [SyncedAt] DATETIME,
    [Version] LONG,
    CONSTRAINT [PK_BusyCloudWebApproval] PRIMARY KEY ([Id])
)
"@
        }

        [void](Invoke-WebApprovalNonQuery -Connection $conn -Sql $sql)
        Write-Host "  [WEB-APPROVAL MIGRATION] Created BusyCloudWebApproval in $($Context.database)" -ForegroundColor Yellow
    }

    if (-not (Test-WebApprovalTableExists -Connection $conn -DbType $dbType -TableName "BusyCloudWebApprovalAction")) {
        if ($dbType -eq 1) {
            $sql = @"
CREATE TABLE dbo.BusyCloudWebApprovalAction (
    Id NVARCHAR(36) NOT NULL PRIMARY KEY,
    WebApprovalId NVARCHAR(36) NOT NULL,
    Action NVARCHAR(30) NOT NULL,
    ActionBy NVARCHAR(100) NOT NULL,
    Remarks NVARCHAR(MAX) NULL,
    MetadataJson NVARCHAR(MAX) NULL,
    ActionTime DATETIME2 NOT NULL
)
"@
        }
        else {
            $sql = @"
CREATE TABLE [BusyCloudWebApprovalAction] (
    [Id] TEXT(36) NOT NULL,
    [WebApprovalId] TEXT(36) NOT NULL,
    [Action] TEXT(30) NOT NULL,
    [ActionBy] TEXT(100) NOT NULL,
    [Remarks] MEMO,
    [MetadataJson] MEMO,
    [ActionTime] DATETIME,
    CONSTRAINT [PK_BusyCloudWebApprovalAction] PRIMARY KEY ([Id])
)
"@
        }

        [void](Invoke-WebApprovalNonQuery -Connection $conn -Sql $sql)
        Write-Host "  [WEB-APPROVAL MIGRATION] Created BusyCloudWebApprovalAction in $($Context.database)" -ForegroundColor Yellow
    }

    if (-not (Test-WebApprovalTableExists -Connection $conn -DbType $dbType -TableName "BusyCloudWebSyncLedger")) {
        if ($dbType -eq 1) {
            $sql = @"
CREATE TABLE dbo.BusyCloudWebSyncLedger (
    WebApprovalId NVARCHAR(36) NOT NULL PRIMARY KEY,
    PayloadHash CHAR(64) NOT NULL,
    State NVARCHAR(30) NOT NULL,
    RequestedVoucherNo NVARCHAR(100) NULL,
    BusyVoucherNo NVARCHAR(100) NULL,
    BusyVoucherCode NVARCHAR(50) NULL,
    NumberChangeReason NVARCHAR(MAX) NULL,
    NumberResolvedAt DATETIME2 NULL,
    Attempts INT NOT NULL CONSTRAINT DF_BCWSL_Attempts DEFAULT(0),
    LastError NVARCHAR(MAX) NULL,
    CreatedAt DATETIME2 NOT NULL,
    UpdatedAt DATETIME2 NOT NULL
)
"@
        }
        else {
            $sql = @"
CREATE TABLE [BusyCloudWebSyncLedger] (
    [WebApprovalId] TEXT(36) NOT NULL,
    [PayloadHash] TEXT(64) NOT NULL,
    [State] TEXT(30) NOT NULL,
    [RequestedVoucherNo] TEXT(100),
    [BusyVoucherNo] TEXT(100),
    [BusyVoucherCode] TEXT(50),
    [NumberChangeReason] MEMO,
    [NumberResolvedAt] DATETIME,
    [Attempts] LONG,
    [LastError] MEMO,
    [CreatedAt] DATETIME,
    [UpdatedAt] DATETIME,
    CONSTRAINT [PK_BusyCloudWebSyncLedger] PRIMARY KEY ([WebApprovalId])
)
"@
        }

        [void](Invoke-WebApprovalNonQuery -Connection $conn -Sql $sql)
        Write-Host "  [WEB-APPROVAL MIGRATION] Created BusyCloudWebSyncLedger in $($Context.database)" -ForegroundColor Yellow
    }

    if (-not (Test-WebApprovalTableExists -Connection $conn -DbType $dbType -TableName "BusyCloudNotification")) {
        if ($dbType -eq 1) {
            $sql = @"
CREATE TABLE dbo.BusyCloudNotification (
    Id NVARCHAR(36) NOT NULL PRIMARY KEY,
    RecipientUserName NVARCHAR(100) NOT NULL,
    WebApprovalId NVARCHAR(36) NULL,
    NotificationType NVARCHAR(50) NOT NULL,
    Title NVARCHAR(250) NOT NULL,
    Message NVARCHAR(MAX) NOT NULL,
    IsRead BIT NOT NULL CONSTRAINT DF_BCN_IsRead DEFAULT(0),
    CreatedAt DATETIME2 NOT NULL,
    ReadAt DATETIME2 NULL
)
"@
        }
        else {
            $sql = @"
CREATE TABLE [BusyCloudNotification] (
    [Id] TEXT(36) NOT NULL,
    [RecipientUserName] TEXT(100) NOT NULL,
    [WebApprovalId] TEXT(36),
    [NotificationType] TEXT(50) NOT NULL,
    [Title] TEXT(250),
    [Message] MEMO,
    [IsRead] BYTE,
    [CreatedAt] DATETIME,
    [ReadAt] DATETIME,
    CONSTRAINT [PK_BusyCloudNotification] PRIMARY KEY ([Id])
)
"@
        }

        [void](Invoke-WebApprovalNonQuery -Connection $conn -Sql $sql)
        Write-Host "  [WEB-APPROVAL MIGRATION] Created BusyCloudNotification in $($Context.database)" -ForegroundColor Yellow
    }

    if (-not (Test-WebApprovalTableExists -Connection $conn -DbType $dbType -TableName "BusyCloudNotificationDelivery")) {
        if ($dbType -eq 1) {
            $sql = @"
CREATE TABLE dbo.BusyCloudNotificationDelivery (
    Id NVARCHAR(36) NOT NULL PRIMARY KEY,
    NotificationId NVARCHAR(36) NOT NULL,
    SubscriptionId NVARCHAR(36) NOT NULL,
    RecipientUserName NVARCHAR(100) NOT NULL,
    State NVARCHAR(20) NOT NULL,
    Attempts INT NOT NULL CONSTRAINT DF_BCND_Attempts DEFAULT(0),
    LastHttpStatus INT NULL,
    LastError NVARCHAR(MAX) NULL,
    CreatedAt DATETIME2 NOT NULL,
    UpdatedAt DATETIME2 NOT NULL,
    SentAt DATETIME2 NULL,
    NextAttemptAt DATETIME2 NULL
)
"@
        }
        else {
            $sql = @"
CREATE TABLE [BusyCloudNotificationDelivery] (
    [Id] TEXT(36) NOT NULL,
    [NotificationId] TEXT(36) NOT NULL,
    [SubscriptionId] TEXT(36) NOT NULL,
    [RecipientUserName] TEXT(100) NOT NULL,
    [State] TEXT(20) NOT NULL,
    [Attempts] LONG,
    [LastHttpStatus] LONG,
    [LastError] MEMO,
    [CreatedAt] DATETIME,
    [UpdatedAt] DATETIME,
    [SentAt] DATETIME,
    [NextAttemptAt] DATETIME,
    CONSTRAINT [PK_BusyCloudNotificationDelivery] PRIMARY KEY ([Id])
)
"@
        }

        [void](Invoke-WebApprovalNonQuery -Connection $conn -Sql $sql)
        Write-Host "  [WEB-APPROVAL MIGRATION] Created BusyCloudNotificationDelivery in $($Context.database)" -ForegroundColor Yellow
    }

    # Helpful indexes. Access CREATE INDEX throws when the index already exists;
    # that is intentionally ignored so startup remains idempotent.
    $sqlIndexes = @(
        "CREATE INDEX IX_BCWA_SubmittedBy ON dbo.BusyCloudWebApproval(SubmittedBy)",
        "CREATE INDEX IX_BCWA_Status ON dbo.BusyCloudWebApproval(ApprovalStatus, SyncStatus)",
        "CREATE INDEX IX_BCWA_VoucherType ON dbo.BusyCloudWebApproval(VoucherType)",
        "CREATE INDEX IX_BCWAA_WebApprovalId ON dbo.BusyCloudWebApprovalAction(WebApprovalId)",
        "CREATE INDEX IX_BCN_Recipient ON dbo.BusyCloudNotification(RecipientUserName, IsRead)",
        "CREATE UNIQUE INDEX UX_BCND_NotificationSubscription ON dbo.BusyCloudNotificationDelivery(NotificationId, SubscriptionId)",
        "CREATE INDEX IX_BCND_State ON dbo.BusyCloudNotificationDelivery(State, Attempts)"
    )

    $accessIndexes = @(
        "CREATE INDEX [IX_BCWA_SubmittedBy] ON [BusyCloudWebApproval] ([SubmittedBy])",
        "CREATE INDEX [IX_BCWA_Status] ON [BusyCloudWebApproval] ([ApprovalStatus], [SyncStatus])",
        "CREATE INDEX [IX_BCWA_VoucherType] ON [BusyCloudWebApproval] ([VoucherType])",
        "CREATE INDEX [IX_BCWAA_WebApprovalId] ON [BusyCloudWebApprovalAction] ([WebApprovalId])",
        "CREATE INDEX [IX_BCN_Recipient] ON [BusyCloudNotification] ([RecipientUserName], [IsRead])",
        "CREATE UNIQUE INDEX [UX_BCND_NotificationSubscription] ON [BusyCloudNotificationDelivery] ([NotificationId], [SubscriptionId])",
        "CREATE INDEX [IX_BCND_State] ON [BusyCloudNotificationDelivery] ([State], [Attempts])"
    )

    $indexes = if ($dbType -eq 1) { $sqlIndexes } else { $accessIndexes }

    foreach ($indexSql in $indexes) {
        try {
            if ($dbType -eq 1) {
                $indexName = ""
                if ($indexSql -match "CREATE INDEX\s+([A-Za-z0-9_]+)") {
                    $indexName = $Matches[1]
                }

                if ($indexName) {
                    $check = $conn.CreateCommand()
                    $check.CommandText = "SELECT COUNT(*) FROM sys.indexes WHERE name=@n"
                    [void]$check.Parameters.AddWithValue("@n", $indexName)

                    if ([int]$check.ExecuteScalar() -eq 0) {
                        [void](Invoke-WebApprovalNonQuery -Connection $conn -Sql $indexSql)
                    }
                }
            }
            else {
                [void](Invoke-WebApprovalNonQuery -Connection $conn -Sql $indexSql)
            }
        } catch {}
    }

    return $true
}

function Ensure-BusyCloudWebApprovalStorage {
    param(
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $permanent = $null
    $fiscal = $null

    try {
        $permanent = Get-WebApprovalPermanentDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        [void](Ensure-WebApprovalPermanentTables -Context $permanent)

        $fiscal = Get-WebApprovalFiscalDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        [void](Ensure-WebApprovalFiscalTables -Context $fiscal)

        return @{
            success = $true
            instanceId = $InstanceId
            companyCode = $CompanyCode
            permanentDatabase = [string]$permanent.database
            fiscalDatabase = [string]$fiscal.database
            schemaVersion = $script:WebApprovalSchemaVersion
        }
    }
    catch {
        return @{
            success = $false
            instanceId = $InstanceId
            companyCode = $CompanyCode
            error = $_.Exception.Message
        }
    }
    finally {
        Close-WebApprovalDbContext -Context $fiscal
        Close-WebApprovalDbContext -Context $permanent
    }
}

function Initialize-BusyCloudWebApprovalStorage {
    param([bool]$VerboseOutput = $true)

    $instancesPath = Join-Path (Split-Path -Parent $PSScriptRoot) "instances.json"

    if (-not (Test-Path -LiteralPath $instancesPath)) {
        return @{
            success = $false
            initialized = 0
            failed = 1
            results = @()
            error = "instances.json was not found: $instancesPath"
        }
    }

    $config = Get-Content -LiteralPath $instancesPath -Raw | ConvertFrom-Json
    $results = @()
    $ok = 0
    $failed = 0

    foreach ($inst in @($config.instances)) {
        foreach ($comp in @($inst.companies)) {
            $result = Ensure-BusyCloudWebApprovalStorage `
                -InstanceId ([string]$inst.id) `
                -CompanyCode ([string]$comp.code)

            $results += $result

            if ($result.success) {
                $ok++

                if ($VerboseOutput) {
                    Write-Host (
                        "  Web Approval storage ready: {0}/{1} | permanent={2} | fiscal={3}" -f `
                        $inst.id,
                        $comp.code,
                        $result.permanentDatabase,
                        $result.fiscalDatabase
                    ) -ForegroundColor DarkGreen
                }
            }
            else {
                $failed++

                if ($VerboseOutput) {
                    Write-Host (
                        "  [WARN] Web Approval storage unavailable: {0}/{1} | {2}" -f `
                        $inst.id,
                        $comp.code,
                        $result.error
                    ) -ForegroundColor DarkYellow
                }
            }
        }
    }

    return @{
        success = ($failed -eq 0)
        initialized = $ok
        failed = $failed
        results = @($results)
    }
}

function Get-WebApprovalRoleDirectory {
    param(
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $ctx = $null
    $rdr = $null

    try {
        $ctx = Get-WebApprovalPermanentDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        [void](Ensure-WebApprovalPermanentTables -Context $ctx)

        $cmd = $ctx.connection.CreateCommand()
        $cmd.CommandText = "SELECT [Name], B35, B36 FROM MobileUserPreference ORDER BY [Name]"
        $rdr = $cmd.ExecuteReader()

        $items = @()

        while ($rdr.Read()) {
            $name = ""
            $role = 0

            if (-not $rdr.IsDBNull(0)) {
                $name = ([string]$rdr.GetValue(0)).Trim()
            }

            $salesmanFlag = 0
            $managerFlag = 0

            if (-not $rdr.IsDBNull(1)) {
                try { $salesmanFlag = [int]$rdr.GetValue(1) } catch {}
            }

            if (-not $rdr.IsDBNull(2)) {
                try { $managerFlag = [int]$rdr.GetValue(2) } catch {}
            }

            $role = Get-WebApprovalRoleFromFlags `
                -SalesmanFlag $salesmanFlag `
                -ManagerFlag $managerFlag

            if ($name) {
                $items += @{
                    userName = $name
                    role = $role
                    roleName = Get-WebApprovalRoleName -Role $role
                }
            }
        }

        return @{
            success = $true
            data = @($items)
        }
    }
    catch {
        return @{ success = $false; error = $_.Exception.Message }
    }
    finally {
        if ($rdr) {
            try { $rdr.Close() } catch {}
            try { $rdr.Dispose() } catch {}
        }
        Close-WebApprovalDbContext -Context $ctx
    }
}

function Get-WebApprovalUserRole {
    param(
        [string]$UserName,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ([string]::IsNullOrWhiteSpace($UserName)) {
        return @{ success=$false; error="UserName is required." }
    }

    $directory = Get-WebApprovalRoleDirectory `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $directory.success) {
        return $directory
    }

    $target = $UserName.Trim().ToLowerInvariant()

    $row = @($directory.data) |
        Where-Object {
            ([string]$_.userName).Trim().ToLowerInvariant() -eq $target
        } |
        Select-Object -First 1

    if ($null -eq $row) {
        return @{
            success = $true
            found = $false
            data = @{
                userName = $UserName.Trim()
                role = 0
                roleName = "NONE"
            }
        }
    }

    return @{
        success = $true
        found = $true
        data = $row
    }
}

function Set-WebApprovalUserRole {
    param(
        [string]$UserName,
        $Role,
        [string]$ActionBy = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ([string]::IsNullOrWhiteSpace($UserName)) {
        return @{ success=$false; error="UserName is required." }
    }

    $roleValue = ConvertTo-WebApprovalRoleValue -Value $Role

    if ($roleValue -notin @(0,1,2)) {
        return @{ success=$false; error="Role must be NONE, SALESMAN or SALES_MANAGER." }
    }

    $ctx = $null
    $tx = $null

    try {
        $ctx = Get-WebApprovalPermanentDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        [void](Ensure-WebApprovalPermanentTables -Context $ctx)

        $safeUser = $UserName.Trim().Replace("'", "''")
        $safeActionBy = ([string]$ActionBy).Trim().Replace("'", "''")

        $cmd = $ctx.connection.CreateCommand()
        $cmd.CommandText = "SELECT COUNT(*) FROM MobileUserPreference WHERE [Name]='$safeUser'"

        if ([int]$cmd.ExecuteScalar() -le 0) {
            return @{
                success = $false
                httpStatus = 404
                error = "MobileUserPreference profile was not found for '$($UserName.Trim())'. Create the web user profile first."
            }
        }

        $tx = $ctx.connection.BeginTransaction()

        $salesmanFlag = 0
        $managerFlag = 0

        if ($roleValue -eq $script:WebApprovalRoleSalesman) {
            $salesmanFlag = 1
        }
        elseif ($roleValue -eq $script:WebApprovalRoleSalesManager) {
            $managerFlag = 1
        }

        [void](Invoke-WebApprovalNonQuery `
            -Connection $ctx.connection `
            -Transaction $tx `
            -Sql "UPDATE [MobileUserPreference] SET [B35]=$salesmanFlag, [B36]=$managerFlag WHERE [Name]='$safeUser'")

        # If the user is no longer a salesman, their outgoing assignments are disabled.
        if ($roleValue -ne $script:WebApprovalRoleSalesman) {
            [void](Invoke-WebApprovalNonQuery `
                -Connection $ctx.connection `
                -Transaction $tx `
                -Sql "UPDATE [BusyCloudSalesmanManager] SET [IsActive]=0 WHERE [SalesmanUserName]='$safeUser'")
        }

        # If the user is no longer a manager, assignments pointing to that manager
        # are disabled. This prevents an invalid manager relationship from lingering.
        if ($roleValue -ne $script:WebApprovalRoleSalesManager) {
            [void](Invoke-WebApprovalNonQuery `
                -Connection $ctx.connection `
                -Transaction $tx `
                -Sql "UPDATE [BusyCloudSalesmanManager] SET [IsActive]=0 WHERE [ManagerUserName]='$safeUser'")
        }

        $tx.Commit()
        $tx = $null

        return @{
            success = $true
            message = "Sales workflow role updated successfully."
            data = @{
                userName = $UserName.Trim()
                role = $roleValue
                roleName = Get-WebApprovalRoleName -Role $roleValue
                updatedBy = $ActionBy
            }
        }
    }
    catch {
        if ($tx) {
            try { $tx.Rollback() } catch {}
        }

        return @{ success=$false; error=$_.Exception.Message }
    }
    finally {
        Close-WebApprovalDbContext -Context $ctx
    }
}

function Get-WebApprovalManagersForSalesman {
    param(
        [string]$SalesmanUserName,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ([string]::IsNullOrWhiteSpace($SalesmanUserName)) {
        return @{ success=$false; error="SalesmanUserName is required." }
    }

    $ctx = $null
    $rdr = $null

    try {
        $ctx = Get-WebApprovalPermanentDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        [void](Ensure-WebApprovalPermanentTables -Context $ctx)

        $safe = $SalesmanUserName.Trim().Replace("'", "''")

        $cmd = $ctx.connection.CreateCommand()
        $cmd.CommandText = @"
SELECT ManagerUserName
FROM [BusyCloudSalesmanManager]
WHERE [SalesmanUserName]='$safe' AND [IsActive]<>0
ORDER BY [ManagerUserName]
"@

        $rdr = $cmd.ExecuteReader()
        $managers = @()

        while ($rdr.Read()) {
            if (-not $rdr.IsDBNull(0)) {
                $name = ([string]$rdr.GetValue(0)).Trim()
                if ($name -and $managers -notcontains $name) {
                    $managers += $name
                }
            }
        }

        return @{
            success = $true
            data = @{
                salesmanUserName = $SalesmanUserName.Trim()
                managers = @($managers)
            }
        }
    }
    catch {
        return @{ success=$false; error=$_.Exception.Message }
    }
    finally {
        if ($rdr) {
            try { $rdr.Close() } catch {}
            try { $rdr.Dispose() } catch {}
        }

        Close-WebApprovalDbContext -Context $ctx
    }
}

function Set-WebApprovalManagersForSalesman {
    param(
        [string]$SalesmanUserName,
        [object[]]$ManagerUserNames,
        [string]$ActionBy = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ([string]::IsNullOrWhiteSpace($SalesmanUserName)) {
        return @{ success=$false; error="SalesmanUserName is required." }
    }

    $ctx = $null
    $tx = $null

    try {
        $ctx = Get-WebApprovalPermanentDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        [void](Ensure-WebApprovalPermanentTables -Context $ctx)

        $salesmanSafe = $SalesmanUserName.Trim().Replace("'", "''")

        $cmd = $ctx.connection.CreateCommand()
        $cmd.CommandText = "SELECT B35, B36 FROM MobileUserPreference WHERE [Name]='$salesmanSafe'"
        $salesmanReader = $cmd.ExecuteReader()

        if (-not $salesmanReader.Read()) {
            try { $salesmanReader.Close() } catch {}
            return @{ success=$false; httpStatus=404; error="Salesman user profile was not found." }
        }

        $salesmanFlag = 0
        $salesmanManagerFlag = 0

        if (-not $salesmanReader.IsDBNull(0)) {
            try { $salesmanFlag = [int]$salesmanReader.GetValue(0) } catch {}
        }

        if (-not $salesmanReader.IsDBNull(1)) {
            try { $salesmanManagerFlag = [int]$salesmanReader.GetValue(1) } catch {}
        }

        try { $salesmanReader.Close() } catch {}

        $salesmanRole = Get-WebApprovalRoleFromFlags `
            -SalesmanFlag $salesmanFlag `
            -ManagerFlag $salesmanManagerFlag

        if ($salesmanRole -ne $script:WebApprovalRoleSalesman) {
            return @{
                success = $false
                httpStatus = 409
                error = "'$($SalesmanUserName.Trim())' must have the SALESMAN role before managers can be assigned."
            }
        }

        $requested = @(
            @($ManagerUserNames) |
            ForEach-Object { ([string]$_).Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
        )

        $canonicalManagers = @()

        foreach ($requestedName in $requested) {
            if ($requestedName.Equals($SalesmanUserName.Trim(), [System.StringComparison]::OrdinalIgnoreCase)) {
                return @{ success=$false; httpStatus=409; error="A salesman cannot be assigned as their own manager." }
            }

            $managerSafe = $requestedName.Replace("'", "''")
            $cmd = $ctx.connection.CreateCommand()
            $cmd.CommandText = "SELECT [Name], B35, B36 FROM MobileUserPreference WHERE [Name]='$managerSafe'"
            $rdr = $cmd.ExecuteReader()

            $foundManager = $false
            $canonical = ""
            $role = 0

            if ($rdr.Read()) {
                $foundManager = $true
                $canonical = ([string]$rdr.GetValue(0)).Trim()
                $managerSalesmanFlag = 0
                $managerRoleFlag = 0

                if (-not $rdr.IsDBNull(1)) {
                    try { $managerSalesmanFlag = [int]$rdr.GetValue(1) } catch {}
                }

                if (-not $rdr.IsDBNull(2)) {
                    try { $managerRoleFlag = [int]$rdr.GetValue(2) } catch {}
                }

                $role = Get-WebApprovalRoleFromFlags `
                    -SalesmanFlag $managerSalesmanFlag `
                    -ManagerFlag $managerRoleFlag
            }

            try { $rdr.Close() } catch {}

            if (-not $foundManager) {
                return @{ success=$false; httpStatus=404; error="Manager user '$requestedName' was not found." }
            }

            if ($role -ne $script:WebApprovalRoleSalesManager) {
                return @{
                    success = $false
                    httpStatus = 409
                    error = "'$canonical' is not a SALES_MANAGER."
                }
            }

            if ($canonicalManagers -notcontains $canonical) {
                $canonicalManagers += $canonical
            }
        }

        $tx = $ctx.connection.BeginTransaction()
        $now = Get-Date
        $nowSql = if ([int]$ctx.dbType -eq 1) {
            "'" + $now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'"
        }
        else {
            "#" + $now.ToString("MM/dd/yyyy HH:mm:ss") + "#"
        }

        [void](Invoke-WebApprovalNonQuery `
            -Connection $ctx.connection `
            -Transaction $tx `
            -Sql "UPDATE [BusyCloudSalesmanManager] SET [IsActive]=0, [UpdatedAt]=$nowSql WHERE [SalesmanUserName]='$salesmanSafe'")

        foreach ($managerName in $canonicalManagers) {
            $managerSafe = $managerName.Replace("'", "''")
            $actionSafe = ([string]$ActionBy).Trim().Replace("'", "''")

            $cmd = $ctx.connection.CreateCommand()
            $cmd.Transaction = $tx
            $cmd.CommandText = "SELECT COUNT(*) FROM BusyCloudSalesmanManager WHERE [SalesmanUserName]='$salesmanSafe' AND [ManagerUserName]='$managerSafe'"
            $exists = ([int]$cmd.ExecuteScalar() -gt 0)

            if ($exists) {
                [void](Invoke-WebApprovalNonQuery `
                    -Connection $ctx.connection `
                    -Transaction $tx `
                    -Sql "UPDATE [BusyCloudSalesmanManager] SET [IsActive]=1, [AssignedBy]='$actionSafe', [UpdatedAt]=$nowSql WHERE [SalesmanUserName]='$salesmanSafe' AND [ManagerUserName]='$managerSafe'")
            }
            else {
                $id = [guid]::NewGuid().ToString()
                [void](Invoke-WebApprovalNonQuery `
                    -Connection $ctx.connection `
                    -Transaction $tx `
                    -Sql "INSERT INTO [BusyCloudSalesmanManager] ([Id],[SalesmanUserName],[ManagerUserName],[IsActive],[AssignedBy],[AssignedAt],[UpdatedAt]) VALUES ('$id','$salesmanSafe','$managerSafe',1,'$actionSafe',$nowSql,$nowSql)")
            }
        }

        $tx.Commit()
        $tx = $null

        return @{
            success = $true
            message = "Sales manager assignments updated successfully."
            data = @{
                salesmanUserName = $SalesmanUserName.Trim()
                managers = @($canonicalManagers)
            }
        }
    }
    catch {
        if ($tx) {
            try { $tx.Rollback() } catch {}
        }

        return @{ success=$false; error=$_.Exception.Message }
    }
    finally {
        Close-WebApprovalDbContext -Context $ctx
    }
}

function Test-WebApprovalManagerAccess {
    param(
        [string]$SalesmanUserName,
        [string]$ManagerUserName,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if (
        [string]::IsNullOrWhiteSpace($SalesmanUserName) -or
        [string]::IsNullOrWhiteSpace($ManagerUserName)
    ) {
        return @{ success=$true; allowed=$false }
    }

    $ctx = $null

    try {
        $ctx = Get-WebApprovalPermanentDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        [void](Ensure-WebApprovalPermanentTables -Context $ctx)

        $salesmanSafe = $SalesmanUserName.Trim().Replace("'", "''")
        $managerSafe = $ManagerUserName.Trim().Replace("'", "''")

        $cmd = $ctx.connection.CreateCommand()
        $cmd.CommandText = @"
SELECT COUNT(*)
FROM [BusyCloudSalesmanManager]
WHERE [SalesmanUserName]='$salesmanSafe'
  AND [ManagerUserName]='$managerSafe'
  AND [IsActive]<>0
"@

        return @{
            success = $true
            allowed = ([int]$cmd.ExecuteScalar() -gt 0)
        }
    }
    catch {
        return @{ success=$false; allowed=$false; error=$_.Exception.Message }
    }
    finally {
        Close-WebApprovalDbContext -Context $ctx
    }
}

function Get-WebApprovalStorageStatus {
    param(
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    return Ensure-BusyCloudWebApprovalStorage `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode
}



# ============================================================================
# PHASE 3.1 - REQUESTED VS FINAL BUSY NUMBER AUDIT
# ============================================================================

function Get-WebApprovalNumberChangeReason {
    param(
        [string]$RequestedVoucherNo,
        [string]$BusyVoucherNo,
        [string]$Reason = ""
    )

    $requested = ([string]$RequestedVoucherNo).Trim()
    $posted = ([string]$BusyVoucherNo).Trim()
    $explicitReason = ([string]$Reason).Trim()

    # No final BUSY voucher exists yet.
    if ([string]::IsNullOrWhiteSpace($posted)) {
        return $explicitReason
    }

    # Same number means there is nothing exceptional to explain.
    if ($requested.Equals($posted, [System.StringComparison]::OrdinalIgnoreCase)) {
        return ""
    }

    if (-not [string]::IsNullOrWhiteSpace($explicitReason)) {
        return $explicitReason
    }

    return (
        "Requested voucher number '{0}' was provisional and was no longer the current " +
        "BUSY number at synchronization time. The voucher was posted using '{1}' " +
        "after resolving the latest BUSY numbering configuration."
    ) -f $requested, $posted
}


function Set-WebApprovalNumberAudit {
    param(
        [string]$WebApprovalId,
        [string]$RequestedVoucherNo,
        [string]$BusyVoucherNo,
        [string]$BusyVoucherCode = "",
        [string]$Reason = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        $ExistingContext = $null,
        $ExistingTransaction = $null
    )

    if ([string]::IsNullOrWhiteSpace($WebApprovalId)) {
        throw "WebApprovalId is required."
    }

    $ownsContext = ($null -eq $ExistingContext)
    $ctx = $ExistingContext
    $tx = $ExistingTransaction
    $ownsTransaction = $false

    try {
        if ($null -eq $ctx) {
            $ctx = Get-WebApprovalFiscalDbContext `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode

            [void](Ensure-WebApprovalFiscalTables -Context $ctx)
        }

        if ($null -eq $tx) {
            $tx = $ctx.connection.BeginTransaction()
            $ownsTransaction = $true
        }

        $requested = ([string]$RequestedVoucherNo).Trim()
        $posted = ([string]$BusyVoucherNo).Trim()
        $code = ([string]$BusyVoucherCode).Trim()

        $numberReason = Get-WebApprovalNumberChangeReason `
            -RequestedVoucherNo $requested `
            -BusyVoucherNo $posted `
            -Reason $Reason

        $resolvedAt = [datetime]::UtcNow

        $cmd = $ctx.connection.CreateCommand()
        $cmd.Transaction = $tx

        if ([int]$ctx.dbType -eq 1) {
            $cmd.CommandText = @"
UPDATE dbo.BusyCloudWebApproval
SET
    BusyVoucherNo=@busyVoucherNo,
    BusyVoucherCode=@busyVoucherCode,
    NumberChangeReason=@numberChangeReason,
    NumberResolvedAt=@numberResolvedAt,
    UpdatedAt=@updatedAt
WHERE Id=@id
"@

            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@busyVoucherNo" -Value $posted -Kind Text -Size 100)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@busyVoucherCode" -Value $code -Kind Text -Size 50)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@numberChangeReason" -Value $numberReason -Kind LongText)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@numberResolvedAt" -Value $resolvedAt -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@updatedAt" -Value $resolvedAt -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@id" -Value $WebApprovalId -Kind Text -Size 36)
        }
        else {
            $cmd.CommandText = @"
UPDATE [BusyCloudWebApproval]
SET
    [BusyVoucherNo]=?,
    [BusyVoucherCode]=?,
    [NumberChangeReason]=?,
    [NumberResolvedAt]=?,
    [UpdatedAt]=?
WHERE [Id]=?
"@

            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p1" -Value $posted -Kind Text -Size 100)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p2" -Value $code -Kind Text -Size 50)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p3" -Value $numberReason -Kind LongText)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p4" -Value $resolvedAt -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p5" -Value $resolvedAt -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p6" -Value $WebApprovalId -Kind Text -Size 36)
        }

        [void]$cmd.ExecuteNonQuery()

        # Update an existing sync ledger row if Phase 5 has already created it.
        # This helper intentionally does NOT create a ledger row itself.
        $ledgerCmd = $ctx.connection.CreateCommand()
        $ledgerCmd.Transaction = $tx

        if ([int]$ctx.dbType -eq 1) {
            $ledgerCmd.CommandText = @"
UPDATE dbo.BusyCloudWebSyncLedger
SET
    RequestedVoucherNo=@requestedVoucherNo,
    BusyVoucherNo=@busyVoucherNo,
    BusyVoucherCode=@busyVoucherCode,
    NumberChangeReason=@numberChangeReason,
    NumberResolvedAt=@numberResolvedAt,
    UpdatedAt=@updatedAt
WHERE WebApprovalId=@id
"@

            [void](Add-WebApprovalCommandParameter -Command $ledgerCmd -DbType 1 -Name "@requestedVoucherNo" -Value $requested -Kind Text -Size 100)
            [void](Add-WebApprovalCommandParameter -Command $ledgerCmd -DbType 1 -Name "@busyVoucherNo" -Value $posted -Kind Text -Size 100)
            [void](Add-WebApprovalCommandParameter -Command $ledgerCmd -DbType 1 -Name "@busyVoucherCode" -Value $code -Kind Text -Size 50)
            [void](Add-WebApprovalCommandParameter -Command $ledgerCmd -DbType 1 -Name "@numberChangeReason" -Value $numberReason -Kind LongText)
            [void](Add-WebApprovalCommandParameter -Command $ledgerCmd -DbType 1 -Name "@numberResolvedAt" -Value $resolvedAt -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $ledgerCmd -DbType 1 -Name "@updatedAt" -Value $resolvedAt -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $ledgerCmd -DbType 1 -Name "@id" -Value $WebApprovalId -Kind Text -Size 36)
        }
        else {
            $ledgerCmd.CommandText = @"
UPDATE [BusyCloudWebSyncLedger]
SET
    [RequestedVoucherNo]=?,
    [BusyVoucherNo]=?,
    [BusyVoucherCode]=?,
    [NumberChangeReason]=?,
    [NumberResolvedAt]=?,
    [UpdatedAt]=?
WHERE [WebApprovalId]=?
"@

            [void](Add-WebApprovalCommandParameter -Command $ledgerCmd -DbType 0 -Name "@p1" -Value $requested -Kind Text -Size 100)
            [void](Add-WebApprovalCommandParameter -Command $ledgerCmd -DbType 0 -Name "@p2" -Value $posted -Kind Text -Size 100)
            [void](Add-WebApprovalCommandParameter -Command $ledgerCmd -DbType 0 -Name "@p3" -Value $code -Kind Text -Size 50)
            [void](Add-WebApprovalCommandParameter -Command $ledgerCmd -DbType 0 -Name "@p4" -Value $numberReason -Kind LongText)
            [void](Add-WebApprovalCommandParameter -Command $ledgerCmd -DbType 0 -Name "@p5" -Value $resolvedAt -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $ledgerCmd -DbType 0 -Name "@p6" -Value $resolvedAt -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $ledgerCmd -DbType 0 -Name "@p7" -Value $WebApprovalId -Kind Text -Size 36)
        }

        [void]$ledgerCmd.ExecuteNonQuery()

        if ($ownsTransaction) {
            $tx.Commit()
            $tx = $null
        }

        return @{
            requestedVoucherNo = $requested
            busyVoucherNo = $posted
            busyVoucherCode = $code
            numberChangeReason = $numberReason
            numberResolvedAt = $resolvedAt
        }
    }
    catch {
        if ($ownsTransaction -and $tx) {
            try { $tx.Rollback() } catch {}
        }

        throw
    }
    finally {
        if ($ownsContext) {
            Close-WebApprovalDbContext -Context $ctx
        }
    }
}


# ============================================================================
# PHASE 3 - WEB APPROVAL SUBMISSION ROUTING
# ============================================================================
# When Config RecType=203 / I1=2 (WEB), a CREATE request is NOT sent to
# Create-Voucher. Instead the validated payload is stored in
# BusyCloudWebApproval as PENDING and assigned managers receive notifications.
#
# IMPORTANT:
# - Existing BUSY approval (I1=1) remains unchanged.
# - No Approval (I1=0) remains unchanged.
# - Modify requests remain unchanged.
# - Actual BUSY voucher creation is NOT performed in this phase.
# ============================================================================

function Get-WebApprovalPropertyValue {
    param(
        $Object,
        [string[]]$Names,
        $DefaultValue = $null
    )

    if ($null -eq $Object) {
        return $DefaultValue
    }

    foreach ($name in @($Names)) {
        try {
            if ($Object -is [System.Collections.IDictionary]) {
                if ($Object.Contains($name)) {
                    return $Object[$name]
                }
            }

            $prop = $Object.PSObject.Properties[$name]
            if ($null -ne $prop) {
                return $prop.Value
            }
        }
        catch {}
    }

    return $DefaultValue
}

function ConvertTo-WebApprovalVoucherDate {
    param($Value)

    if ($Value -is [datetime]) {
        return [datetime]$Value
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        throw "Voucher date is required."
    }

    $parsed = [datetime]::MinValue
    $formats = @(
        "dd-MM-yyyy",
        "yyyy-MM-dd",
        "dd/MM/yyyy",
        "yyyy/MM/dd",
        "MM/dd/yyyy"
    )

    foreach ($format in $formats) {
        if (
            [datetime]::TryParseExact(
                $text,
                $format,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None,
                [ref]$parsed
            )
        ) {
            return $parsed
        }
    }

    if (
        [datetime]::TryParse(
            $text,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::AllowWhiteSpaces,
            [ref]$parsed
        )
    ) {
        return $parsed
    }

    throw "Invalid voucher date: $text"
}

function Get-WebApprovalSha256 {
    param([string]$Text)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$Text)
        $hash = $sha.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hash)).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Get-WebApprovalRequestId {
    param($Data)

    $raw = [string](Get-WebApprovalPropertyValue `
        -Object $Data `
        -Names @("clientRequestId", "webApprovalRequestId") `
        -DefaultValue "")

    $guid = [guid]::Empty

    if (
        -not [string]::IsNullOrWhiteSpace($raw) -and
        [guid]::TryParse($raw.Trim(), [ref]$guid)
    ) {
        return $guid.ToString()
    }

    return [guid]::NewGuid().ToString()
}

function Get-WebApprovalPayloadSnapshot {
    param($Data)

    # Deep clone using JSON so the stored object is detached from later route
    # mutations. Client-only idempotency metadata is removed from the approved
    # accounting payload.
    $json = $Data | ConvertTo-Json -Depth 50 -Compress
    $copy = $json | ConvertFrom-Json

    foreach ($name in @("clientRequestId", "webApprovalRequestId")) {
        try {
            $copy.PSObject.Properties.Remove($name)
        }
        catch {}
    }

    return ($copy | ConvertTo-Json -Depth 50 -Compress)
}

function Get-WebApprovalPayloadSummary {
    param($Data)

    $vchType = 0
    try {
        $vchType = [int](Get-WebApprovalPropertyValue -Object $Data -Names @("vchType") -DefaultValue 0)
    }
    catch {}

    $partyName = ([string](Get-WebApprovalPropertyValue `
        -Object $Data `
        -Names @("party", "partyName", "customerName") `
        -DefaultValue "")).Trim()

    if ([string]::IsNullOrWhiteSpace($partyName) -and $vchType -eq 14) {
        foreach ($account in @(Get-WebApprovalPropertyValue -Object $Data -Names @("accounts") -DefaultValue @())) {
            if ($null -eq $account) { continue }

            $dc = ([string](Get-WebApprovalPropertyValue -Object $account -Names @("dc") -DefaultValue "")).Trim().ToUpperInvariant()
            $name = ([string](Get-WebApprovalPropertyValue -Object $account -Names @("accountName", "account") -DefaultValue "")).Trim()

            if ($dc -eq "C" -and -not [string]::IsNullOrWhiteSpace($name)) {
                $partyName = $name
                break
            }
        }
    }

    $amount = [decimal]0

    $accounts = @(Get-WebApprovalPropertyValue -Object $Data -Names @("accounts") -DefaultValue @())
    if ($accounts.Count -gt 0) {
        $debitTotal = [decimal]0

        foreach ($account in $accounts) {
            if ($null -eq $account) { continue }

            $dc = ([string](Get-WebApprovalPropertyValue -Object $account -Names @("dc") -DefaultValue "")).Trim().ToUpperInvariant()
            if ($dc -ne "D") { continue }

            try {
                $debitTotal += [decimal](Get-WebApprovalPropertyValue -Object $account -Names @("amount") -DefaultValue 0)
            }
            catch {}
        }

        $amount = $debitTotal
    }
    else {
        foreach ($item in @(Get-WebApprovalPropertyValue -Object $Data -Names @("items") -DefaultValue @())) {
            if ($null -eq $item) { continue }

            try {
                $amount += [decimal](Get-WebApprovalPropertyValue -Object $item -Names @("amount") -DefaultValue 0)
            }
            catch {}
        }

        foreach ($sundry in @(Get-WebApprovalPropertyValue -Object $Data -Names @("billSundries") -DefaultValue @())) {
            if ($null -eq $sundry) { continue }

            try {
                $amount += [decimal](Get-WebApprovalPropertyValue -Object $sundry -Names @("amount") -DefaultValue 0)
            }
            catch {}
        }
    }

    return @{
        partyName = $partyName
        amount = $amount
    }
}

function Add-WebApprovalCommandParameter {
    param(
        $Command,
        [int]$DbType,
        [string]$Name,
        $Value,
        [ValidateSet("Text", "LongText", "Int", "Decimal", "Date", "Bool")]
        [string]$Kind = "Text",
        [int]$Size = 255
    )

    $dbValue = if ($null -eq $Value) { [System.DBNull]::Value } else { $Value }

    if ($DbType -eq 1) {
        switch ($Kind) {
            "LongText" {
                $p = $Command.Parameters.Add($Name, [System.Data.SqlDbType]::NVarChar, -1)
            }
            "Int" {
                $p = $Command.Parameters.Add($Name, [System.Data.SqlDbType]::Int)
            }
            "Decimal" {
                $p = $Command.Parameters.Add($Name, [System.Data.SqlDbType]::Decimal)
                $p.Precision = 19
                $p.Scale = 4
            }
            "Date" {
                $p = $Command.Parameters.Add($Name, [System.Data.SqlDbType]::DateTime2)
            }
            "Bool" {
                $p = $Command.Parameters.Add($Name, [System.Data.SqlDbType]::Bit)
            }
            default {
                $safeSize = if ($Size -gt 0) { $Size } else { 255 }
                $p = $Command.Parameters.Add($Name, [System.Data.SqlDbType]::NVarChar, $safeSize)
            }
        }

        $p.Value = $dbValue
        return $p
    }

    $oleType = [System.Data.OleDb.OleDbType]::VarWChar

    switch ($Kind) {
        "LongText" { $oleType = [System.Data.OleDb.OleDbType]::LongVarWChar }
        "Int"      { $oleType = [System.Data.OleDb.OleDbType]::Integer }
        "Decimal"  { $oleType = [System.Data.OleDb.OleDbType]::Double }
        "Date"     { $oleType = [System.Data.OleDb.OleDbType]::DBTimeStamp }
        "Bool"     { $oleType = [System.Data.OleDb.OleDbType]::UnsignedTinyInt }
        default    { $oleType = [System.Data.OleDb.OleDbType]::VarWChar }
    }

    $p = New-Object System.Data.OleDb.OleDbParameter
    $p.ParameterName = $Name
    $p.OleDbType = $oleType

    if ($Kind -eq "Text" -and $Size -gt 0) {
        $p.Size = $Size
    }

    $p.Value = $dbValue
    [void]$Command.Parameters.Add($p)
    return $p
}

function Get-WebApprovalExistingSubmission {
    param(
        $Context,
        [string]$Id
    )

    $cmd = $Context.connection.CreateCommand()

    if ([int]$Context.dbType -eq 1) {
        $cmd.CommandText = @"
SELECT TOP 1
    Id,
    SubmittedBy,
    PayloadHash,
    ApprovalStatus,
    SyncStatus,
    SubmittedAt
FROM dbo.BusyCloudWebApproval
WHERE Id=@id
"@
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@id" -Value $Id -Kind Text -Size 36)
    }
    else {
        $cmd.CommandText = @"
SELECT TOP 1
    [Id],
    [SubmittedBy],
    [PayloadHash],
    [ApprovalStatus],
    [SyncStatus],
    [SubmittedAt]
FROM [BusyCloudWebApproval]
WHERE [Id]=?
"@
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@id" -Value $Id -Kind Text -Size 36)
    }

    $rdr = $null

    try {
        $rdr = $cmd.ExecuteReader()
        if (-not $rdr.Read()) {
            return $null
        }

        return @{
            id = if ($rdr.IsDBNull(0)) { "" } else { [string]$rdr.GetValue(0) }
            submittedBy = if ($rdr.IsDBNull(1)) { "" } else { [string]$rdr.GetValue(1) }
            payloadHash = if ($rdr.IsDBNull(2)) { "" } else { [string]$rdr.GetValue(2) }
            approvalStatus = if ($rdr.IsDBNull(3)) { "" } else { [string]$rdr.GetValue(3) }
            syncStatus = if ($rdr.IsDBNull(4)) { "" } else { [string]$rdr.GetValue(4) }
            submittedAt = if ($rdr.IsDBNull(5)) { $null } else { $rdr.GetValue(5) }
        }
    }
    finally {
        if ($rdr) {
            try { $rdr.Close() } catch {}
            try { $rdr.Dispose() } catch {}
        }
    }
}

function Add-WebApprovalPendingRow {
    param(
        $Context,
        $Transaction,
        [string]$Id,
        [string]$SubmittedBy,
        [int]$VoucherType,
        [datetime]$VoucherDate,
        [string]$RequestedSeries,
        [string]$RequestedVoucherNo,
        [string]$PartyName,
        [decimal]$Amount,
        [string]$PayloadJson,
        [string]$PayloadHash,
        [datetime]$NowUtc
    )

    $cmd = $Context.connection.CreateCommand()
    $cmd.Transaction = $Transaction
    $dbType = [int]$Context.dbType

    if ($dbType -eq 1) {
        $cmd.CommandText = @"
INSERT INTO dbo.BusyCloudWebApproval
(
    Id,
    SubmittedBy,
    VoucherType,
    VoucherDate,
    RequestedSeries,
    RequestedVoucherNo,
    PartyName,
    Amount,
    PayloadJson,
    PayloadHash,
    ApprovalStatus,
    SyncStatus,
    SyncAttempts,
    SubmittedAt,
    UpdatedAt,
    Version
)
VALUES
(
    @id,
    @submittedBy,
    @voucherType,
    @voucherDate,
    @requestedSeries,
    @requestedVoucherNo,
    @partyName,
    @amount,
    @payloadJson,
    @payloadHash,
    @approvalStatus,
    @syncStatus,
    @syncAttempts,
    @submittedAt,
    @updatedAt,
    @version
)
"@

        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@id" -Value $Id -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@submittedBy" -Value $SubmittedBy -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@voucherType" -Value $VoucherType -Kind Int)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@voucherDate" -Value $VoucherDate -Kind Date)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@requestedSeries" -Value $RequestedSeries -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@requestedVoucherNo" -Value $RequestedVoucherNo -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@partyName" -Value $PartyName -Kind Text -Size 255)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@amount" -Value $Amount -Kind Decimal)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@payloadJson" -Value $PayloadJson -Kind LongText)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@payloadHash" -Value $PayloadHash -Kind Text -Size 64)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@approvalStatus" -Value "PENDING" -Kind Text -Size 20)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@syncStatus" -Value "NOT_READY" -Kind Text -Size 30)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@syncAttempts" -Value 0 -Kind Int)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@submittedAt" -Value $NowUtc -Kind Date)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@updatedAt" -Value $NowUtc -Kind Date)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@version" -Value 1 -Kind Int)
    }
    else {
        $cmd.CommandText = @"
INSERT INTO [BusyCloudWebApproval]
(
    [Id],
    [SubmittedBy],
    [VoucherType],
    [VoucherDate],
    [RequestedSeries],
    [RequestedVoucherNo],
    [PartyName],
    [Amount],
    [PayloadJson],
    [PayloadHash],
    [ApprovalStatus],
    [SyncStatus],
    [SyncAttempts],
    [SubmittedAt],
    [UpdatedAt],
    [Version]
)
VALUES
(
    ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
)
"@

        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p1" -Value $Id -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p2" -Value $SubmittedBy -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p3" -Value $VoucherType -Kind Int)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p4" -Value $VoucherDate -Kind Date)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p5" -Value $RequestedSeries -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p6" -Value $RequestedVoucherNo -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p7" -Value $PartyName -Kind Text -Size 255)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p8" -Value ([double]$Amount) -Kind Decimal)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p9" -Value $PayloadJson -Kind LongText)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p10" -Value $PayloadHash -Kind Text -Size 64)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p11" -Value "PENDING" -Kind Text -Size 20)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p12" -Value "NOT_READY" -Kind Text -Size 30)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p13" -Value 0 -Kind Int)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p14" -Value $NowUtc -Kind Date)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p15" -Value $NowUtc -Kind Date)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p16" -Value 1 -Kind Int)
    }

    [void]$cmd.ExecuteNonQuery()
}

function Add-WebApprovalActionRow {
    param(
        $Context,
        $Transaction,
        [string]$WebApprovalId,
        [string]$Action,
        [string]$ActionBy,
        [string]$Remarks,
        [string]$MetadataJson,
        [datetime]$ActionTime
    )

    $cmd = $Context.connection.CreateCommand()
    $cmd.Transaction = $Transaction
    $dbType = [int]$Context.dbType
    $id = [guid]::NewGuid().ToString()

    if ($dbType -eq 1) {
        $cmd.CommandText = @"
INSERT INTO dbo.BusyCloudWebApprovalAction
(
    Id,
    WebApprovalId,
    Action,
    ActionBy,
    Remarks,
    MetadataJson,
    ActionTime
)
VALUES
(
    @id,
    @webApprovalId,
    @action,
    @actionBy,
    @remarks,
    @metadataJson,
    @actionTime
)
"@

        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@id" -Value $id -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@webApprovalId" -Value $WebApprovalId -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@action" -Value $Action -Kind Text -Size 30)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@actionBy" -Value $ActionBy -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@remarks" -Value $Remarks -Kind LongText)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@metadataJson" -Value $MetadataJson -Kind LongText)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@actionTime" -Value $ActionTime -Kind Date)
    }
    else {
        $cmd.CommandText = @"
INSERT INTO [BusyCloudWebApprovalAction]
(
    [Id],
    [WebApprovalId],
    [Action],
    [ActionBy],
    [Remarks],
    [MetadataJson],
    [ActionTime]
)
VALUES
(
    ?,?,?,?,?,?,?
)
"@

        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p1" -Value $id -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p2" -Value $WebApprovalId -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p3" -Value $Action -Kind Text -Size 30)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p4" -Value $ActionBy -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p5" -Value $Remarks -Kind LongText)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p6" -Value $MetadataJson -Kind LongText)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p7" -Value $ActionTime -Kind Date)
    }

    [void]$cmd.ExecuteNonQuery()
}

function Add-WebApprovalNotificationRow {
    param(
        $Context,
        $Transaction,
        [string]$RecipientUserName,
        [string]$WebApprovalId,
        [string]$NotificationType,
        [string]$Title,
        [string]$Message,
        [datetime]$CreatedAt
    )

    $cmd = $Context.connection.CreateCommand()
    $cmd.Transaction = $Transaction
    $dbType = [int]$Context.dbType
    $id = [guid]::NewGuid().ToString()

    if ($dbType -eq 1) {
        $cmd.CommandText = @"
INSERT INTO dbo.BusyCloudNotification
(
    Id,
    RecipientUserName,
    WebApprovalId,
    NotificationType,
    Title,
    Message,
    IsRead,
    CreatedAt
)
VALUES
(
    @id,
    @recipientUserName,
    @webApprovalId,
    @notificationType,
    @title,
    @message,
    @isRead,
    @createdAt
)
"@

        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@id" -Value $id -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@recipientUserName" -Value $RecipientUserName -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@webApprovalId" -Value $WebApprovalId -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@notificationType" -Value $NotificationType -Kind Text -Size 50)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@title" -Value $Title -Kind Text -Size 250)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@message" -Value $Message -Kind LongText)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@isRead" -Value 0 -Kind Bool)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@createdAt" -Value $CreatedAt -Kind Date)
    }
    else {
        $cmd.CommandText = @"
INSERT INTO [BusyCloudNotification]
(
    [Id],
    [RecipientUserName],
    [WebApprovalId],
    [NotificationType],
    [Title],
    [Message],
    [IsRead],
    [CreatedAt]
)
VALUES
(
    ?,?,?,?,?,?,?,?
)
"@

        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p1" -Value $id -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p2" -Value $RecipientUserName -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p3" -Value $WebApprovalId -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p4" -Value $NotificationType -Kind Text -Size 50)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p5" -Value $Title -Kind Text -Size 250)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p6" -Value $Message -Kind LongText)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p7" -Value 0 -Kind Bool)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType $dbType -Name "@p8" -Value $CreatedAt -Kind Date)
    }

    [void]$cmd.ExecuteNonQuery()
}

function Submit-WebApprovalVoucher {
    param(
        $Data,
        [string]$SubmittedBy,
        [string[]]$ManagerUserNames,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ($null -eq $Data) {
        return @{
            success = $false
            httpStatus = 400
            errorCode = "WEB_APPROVAL_PAYLOAD_REQUIRED"
            error = "Voucher payload is required."
        }
    }

    if ([string]::IsNullOrWhiteSpace($SubmittedBy)) {
        return @{
            success = $false
            httpStatus = 401
            errorCode = "WEB_APPROVAL_AUTH_REQUIRED"
            error = "Authenticated BUSY user is required."
        }
    }

    $vchType = 0
    try {
        $vchType = [int](Get-WebApprovalPropertyValue -Object $Data -Names @("vchType") -DefaultValue 0)
    }
    catch {}

    if (-not (Test-WebApprovalSupportedVoucherType -VchType $vchType)) {
        return @{
            success = $false
            httpStatus = 400
            errorCode = "WEB_APPROVAL_UNSUPPORTED_VOUCHER_TYPE"
            error = "This voucher type is not supported by Web Approval."
        }
    }

    $managers = @(
        @($ManagerUserNames) |
        ForEach-Object { ([string]$_).Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique
    )

    if ($managers.Count -eq 0) {
        return @{
            success = $false
            httpStatus = 409
            errorCode = "WEB_APPROVAL_NO_MANAGER"
            error = "No Sales Manager is assigned to this Salesman. Ask an administrator to assign at least one Sales Manager."
        }
    }

    $id = Get-WebApprovalRequestId -Data $Data
    $payloadJson = Get-WebApprovalPayloadSnapshot -Data $Data
    $payloadHash = Get-WebApprovalSha256 -Text $payloadJson

    $voucherDate = $null
    try {
        $voucherDate = ConvertTo-WebApprovalVoucherDate `
            -Value (Get-WebApprovalPropertyValue -Object $Data -Names @("date") -DefaultValue "")
    }
    catch {
        return @{
            success = $false
            httpStatus = 400
            errorCode = "WEB_APPROVAL_INVALID_DATE"
            error = $_.Exception.Message
        }
    }

    $requestedSeries = ([string](Get-WebApprovalPropertyValue -Object $Data -Names @("vchSeries") -DefaultValue "")).Trim()
    $requestedVoucherNo = ([string](Get-WebApprovalPropertyValue -Object $Data -Names @("vchNo") -DefaultValue "")).Trim()
    $summary = Get-WebApprovalPayloadSummary -Data $Data
    $nowUtc = [datetime]::UtcNow

    $ctx = $null
    $tx = $null

    try {
        $ctx = Get-WebApprovalFiscalDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        [void](Ensure-WebApprovalFiscalTables -Context $ctx)

        $existing = Get-WebApprovalExistingSubmission -Context $ctx -Id $id

        if ($null -ne $existing) {
            $sameUser = ([string]$existing.submittedBy).Trim().Equals(
                $SubmittedBy.Trim(),
                [System.StringComparison]::OrdinalIgnoreCase
            )

            if (-not $sameUser) {
                return @{
                    success = $false
                    httpStatus = 409
                    errorCode = "WEB_APPROVAL_IDEMPOTENCY_CONFLICT"
                    error = "This submission id is already owned by another user."
                }
            }

            if (([string]$existing.payloadHash).Trim().ToLowerInvariant() -ne $payloadHash) {
                return @{
                    success = $false
                    httpStatus = 409
                    errorCode = "WEB_APPROVAL_PAYLOAD_CHANGED"
                    error = "This submission was already received with different voucher data. Refresh the form before submitting again."
                }
            }

            return @{
                success = $true
                queuedForWebApproval = $true
                duplicate = $true
                webApprovalId = $id
                webApprovalReference = ("WA-" + $id.Substring(0, 8).ToUpperInvariant())
                approvalStatus = $existing.approvalStatus
                syncStatus = $existing.syncStatus
                assignedManagers = @($managers)
                message = "This voucher was already submitted for Web Approval."
            }
        }

        $tx = $ctx.connection.BeginTransaction()

        Add-WebApprovalPendingRow `
            -Context $ctx `
            -Transaction $tx `
            -Id $id `
            -SubmittedBy $SubmittedBy.Trim() `
            -VoucherType $vchType `
            -VoucherDate $voucherDate `
            -RequestedSeries $requestedSeries `
            -RequestedVoucherNo $requestedVoucherNo `
            -PartyName ([string]$summary.partyName) `
            -Amount ([decimal]$summary.amount) `
            -PayloadJson $payloadJson `
            -PayloadHash $payloadHash `
            -NowUtc $nowUtc

        $managerMetadata = @{
            assignedManagers = @($managers)
            voucherType = $vchType
            requestedSeries = $requestedSeries
            requestedVoucherNo = $requestedVoucherNo
        } | ConvertTo-Json -Depth 10 -Compress

        Add-WebApprovalActionRow `
            -Context $ctx `
            -Transaction $tx `
            -WebApprovalId $id `
            -Action "SUBMITTED" `
            -ActionBy $SubmittedBy.Trim() `
            -Remarks "Submitted for Web Approval." `
            -MetadataJson $managerMetadata `
            -ActionTime $nowUtc

        $reference = "WA-" + $id.Substring(0, 8).ToUpperInvariant()
        $partyText = ([string]$summary.partyName).Trim()
        $notificationTitle = "Web Approval Required"

        foreach ($manager in $managers) {
            $message = if ([string]::IsNullOrWhiteSpace($partyText)) {
                "$($SubmittedBy.Trim()) submitted voucher $reference for your approval."
            }
            else {
                "$($SubmittedBy.Trim()) submitted voucher $reference for $partyText for your approval."
            }

            Add-WebApprovalNotificationRow `
                -Context $ctx `
                -Transaction $tx `
                -RecipientUserName $manager `
                -WebApprovalId $id `
                -NotificationType "WEB_APPROVAL_SUBMITTED" `
                -Title $notificationTitle `
                -Message $message `
                -CreatedAt $nowUtc
        }

        $tx.Commit()
        $tx = $null

        return @{
            success = $true
            queuedForWebApproval = $true
            duplicate = $false
            webApprovalId = $id
            webApprovalReference = $reference
            approvalStatus = "PENDING"
            syncStatus = "NOT_READY"
            assignedManagers = @($managers)
            message = "Voucher submitted for Web Approval. No BUSY voucher has been created yet."
        }
    }
    catch {
        if ($tx) {
            try { $tx.Rollback() } catch {}
        }

        return @{
            success = $false
            httpStatus = 500
            errorCode = "WEB_APPROVAL_SUBMIT_FAILED"
            error = $_.Exception.Message
        }
    }
    finally {
        Close-WebApprovalDbContext -Context $ctx
    }
}

function Invoke-WebApprovalVoucherCreateRouting {
    param(
        $AuthResult,
        $Data,
        [string]$InstanceId = "",
        [string]$CompanyCode = "",
        [bool]$RequireAuth = $true
    )

    $vchType = 0
    try {
        $vchType = [int](Get-WebApprovalPropertyValue -Object $Data -Names @("vchType") -DefaultValue 0)
    }
    catch {}

    if (-not (Test-WebApprovalSupportedVoucherType -VchType $vchType)) {
        return @{
            success = $true
            handled = $false
        }
    }

    $config = Get-VoucherApprovalConfig `
        -VchType $vchType `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $config.success) {
        return @{
            success = $false
            handled = $true
            result = @{
                success = $false
                httpStatus = 500
                errorCode = "WEB_APPROVAL_CONFIG_READ_FAILED"
                error = if ($config.error) { [string]$config.error } else { "Could not read voucher approval mode." }
            }
        }
    }

    $mode = 0
    try {
        $mode = [int]$config.data.approval_mode_value
    }
    catch {
        if (([string]$config.data.approval_mode).Trim().ToUpperInvariant() -eq "WEB") {
            $mode = 2
        }
        elseif ([bool]$config.data.approval_required) {
            $mode = 1
        }
    }

    if ($mode -ne 2) {
        # NONE and BUSY continue through the existing Create-Voucher path.
        return @{
            success = $true
            handled = $false
        }
    }

    if (
        -not $RequireAuth -or
        $null -eq $AuthResult -or
        $null -eq $AuthResult.user -or
        [string]::IsNullOrWhiteSpace([string]$AuthResult.user.name)
    ) {
        return @{
            success = $false
            handled = $true
            result = @{
                success = $false
                httpStatus = 401
                errorCode = "WEB_APPROVAL_AUTH_REQUIRED"
                error = "Web Approval submissions require an authenticated BUSY user."
            }
        }
    }

    $userName = ([string]$AuthResult.user.name).Trim()

    $roleResult = Get-WebApprovalUserRole `
        -UserName $userName `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $roleResult.success) {
        return @{
            success = $false
            handled = $true
            result = @{
                success = $false
                httpStatus = 500
                errorCode = "WEB_APPROVAL_ROLE_READ_FAILED"
                error = if ($roleResult.error) { [string]$roleResult.error } else { "Could not read the user's Sales Workflow role." }
            }
        }
    }

    $role = 0
    try { $role = [int]$roleResult.data.role } catch {}

    if ($role -ne $script:WebApprovalRoleSalesman) {
        return @{
            success = $false
            handled = $true
            result = @{
                success = $false
                httpStatus = 403
                errorCode = "WEB_APPROVAL_SALESMAN_REQUIRED"
                error = "This voucher type uses Web Approval. The signed-in user must have the SALESMAN workflow role to submit a new voucher."
            }
        }
    }

    $managerResult = Get-WebApprovalManagersForSalesman `
        -SalesmanUserName $userName `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $managerResult.success) {
        return @{
            success = $false
            handled = $true
            result = @{
                success = $false
                httpStatus = 500
                errorCode = "WEB_APPROVAL_MANAGER_READ_FAILED"
                error = if ($managerResult.error) { [string]$managerResult.error } else { "Could not resolve assigned Sales Managers." }
            }
        }
    }

    $managers = @($managerResult.data.managers)

    $submitResult = Submit-WebApprovalVoucher `
        -Data $Data `
        -SubmittedBy $userName `
        -ManagerUserNames $managers `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    return @{
        success = [bool]$submitResult.success
        handled = $true
        result = $submitResult
    }
}

function Get-WebApprovalMySubmissions {
    param(
        [string]$UserName,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ([string]::IsNullOrWhiteSpace($UserName)) {
        return @{
            success = $false
            httpStatus = 401
            error = "Authenticated BUSY user is required."
        }
    }

    $ctx = $null
    $rdr = $null

    try {
        $ctx = Get-WebApprovalFiscalDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        [void](Ensure-WebApprovalFiscalTables -Context $ctx)

        $cmd = $ctx.connection.CreateCommand()

        if ([int]$ctx.dbType -eq 1) {
            $cmd.CommandText = @"
SELECT TOP 100
    Id,
    VoucherType,
    VoucherDate,
    RequestedSeries,
    RequestedVoucherNo,
    PartyName,
    Amount,
    ApprovalStatus,
    SyncStatus,
    DecisionBy,
    DecisionAt,
    RejectionReason,
    BusyVoucherNo,
    BusyVoucherCode,
    NumberChangeReason,
    NumberResolvedAt,
    SubmittedAt,
    UpdatedAt,
    SyncedAt,
    LastSyncError
FROM dbo.BusyCloudWebApproval
WHERE SubmittedBy=@userName
ORDER BY SubmittedAt DESC
"@
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@userName" -Value $UserName.Trim() -Kind Text -Size 100)
        }
        else {
            $cmd.CommandText = @"
SELECT TOP 100
    [Id],
    [VoucherType],
    [VoucherDate],
    [RequestedSeries],
    [RequestedVoucherNo],
    [PartyName],
    [Amount],
    [ApprovalStatus],
    [SyncStatus],
    [DecisionBy],
    [DecisionAt],
    [RejectionReason],
    [BusyVoucherNo],
    [BusyVoucherCode],
    [NumberChangeReason],
    [NumberResolvedAt],
    [SubmittedAt],
    [UpdatedAt],
    [SyncedAt],
    [LastSyncError]
FROM [BusyCloudWebApproval]
WHERE [SubmittedBy]=?
ORDER BY [SubmittedAt] DESC
"@
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@userName" -Value $UserName.Trim() -Kind Text -Size 100)
        }

        $rdr = $cmd.ExecuteReader()
        $items = @()

        while ($rdr.Read()) {
            $id = if ($rdr.IsDBNull(0)) { "" } else { [string]$rdr.GetValue(0) }

            $items += @{
                id = $id
                reference = if ($id.Length -ge 8) { "WA-" + $id.Substring(0, 8).ToUpperInvariant() } else { $id }
                voucherType = if ($rdr.IsDBNull(1)) { 0 } else { [int]$rdr.GetValue(1) }
                voucherDate = if ($rdr.IsDBNull(2)) { $null } else { $rdr.GetValue(2) }
                requestedSeries = if ($rdr.IsDBNull(3)) { "" } else { [string]$rdr.GetValue(3) }
                requestedVoucherNo = if ($rdr.IsDBNull(4)) { "" } else { [string]$rdr.GetValue(4) }
                partyName = if ($rdr.IsDBNull(5)) { "" } else { [string]$rdr.GetValue(5) }
                amount = if ($rdr.IsDBNull(6)) { 0 } else { [double]$rdr.GetValue(6) }
                approvalStatus = if ($rdr.IsDBNull(7)) { "" } else { [string]$rdr.GetValue(7) }
                syncStatus = if ($rdr.IsDBNull(8)) { "" } else { [string]$rdr.GetValue(8) }
                decisionBy = if ($rdr.IsDBNull(9)) { "" } else { [string]$rdr.GetValue(9) }
                decisionAt = if ($rdr.IsDBNull(10)) { $null } else { $rdr.GetValue(10) }
                rejectionReason = if ($rdr.IsDBNull(11)) { "" } else { [string]$rdr.GetValue(11) }
                busyVoucherNo = if ($rdr.IsDBNull(12)) { "" } else { [string]$rdr.GetValue(12) }
                busyVoucherCode = if ($rdr.IsDBNull(13)) { "" } else { [string]$rdr.GetValue(13) }
                numberChangeReason = if ($rdr.IsDBNull(14)) { "" } else { [string]$rdr.GetValue(14) }
                numberResolvedAt = if ($rdr.IsDBNull(15)) { $null } else { $rdr.GetValue(15) }
                submittedAt = if ($rdr.IsDBNull(16)) { $null } else { $rdr.GetValue(16) }
                updatedAt = if ($rdr.IsDBNull(17)) { $null } else { $rdr.GetValue(17) }
                syncedAt = if ($rdr.IsDBNull(18)) { $null } else { $rdr.GetValue(18) }
                lastSyncError = if ($rdr.IsDBNull(19)) { "" } else { [string]$rdr.GetValue(19) }
            }
        }

        return @{
            success = $true
            data = @{
                items = @($items)
                count = @($items).Count
            }
        }
    }
    catch {
        return @{
            success = $false
            httpStatus = 500
            error = $_.Exception.Message
        }
    }
    finally {
        if ($rdr) {
            try { $rdr.Close() } catch {}
            try { $rdr.Dispose() } catch {}
        }

        Close-WebApprovalDbContext -Context $ctx
    }
}

function Get-WebApprovalNotifications {
    param(
        [string]$UserName,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ([string]::IsNullOrWhiteSpace($UserName)) {
        return @{
            success = $false
            httpStatus = 401
            error = "Authenticated BUSY user is required."
        }
    }

    $ctx = $null
    $rdr = $null

    try {
        $ctx = Get-WebApprovalFiscalDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        [void](Ensure-WebApprovalFiscalTables -Context $ctx)

        $cmd = $ctx.connection.CreateCommand()

        if ([int]$ctx.dbType -eq 1) {
            $cmd.CommandText = @"
SELECT TOP 100
    Id,
    WebApprovalId,
    NotificationType,
    Title,
    Message,
    IsRead,
    CreatedAt,
    ReadAt
FROM dbo.BusyCloudNotification
WHERE RecipientUserName=@userName
ORDER BY CreatedAt DESC
"@
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@userName" -Value $UserName.Trim() -Kind Text -Size 100)
        }
        else {
            $cmd.CommandText = @"
SELECT TOP 100
    [Id],
    [WebApprovalId],
    [NotificationType],
    [Title],
    [Message],
    [IsRead],
    [CreatedAt],
    [ReadAt]
FROM [BusyCloudNotification]
WHERE [RecipientUserName]=?
ORDER BY [CreatedAt] DESC
"@
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@userName" -Value $UserName.Trim() -Kind Text -Size 100)
        }

        $rdr = $cmd.ExecuteReader()
        $items = @()
        $unread = 0

        while ($rdr.Read()) {
            $isRead = $false
            if (-not $rdr.IsDBNull(5)) {
                try { $isRead = ([int]$rdr.GetValue(5) -ne 0) } catch {}
            }

            if (-not $isRead) {
                $unread++
            }

            $items += @{
                id = if ($rdr.IsDBNull(0)) { "" } else { [string]$rdr.GetValue(0) }
                webApprovalId = if ($rdr.IsDBNull(1)) { "" } else { [string]$rdr.GetValue(1) }
                type = if ($rdr.IsDBNull(2)) { "" } else { [string]$rdr.GetValue(2) }
                title = if ($rdr.IsDBNull(3)) { "" } else { [string]$rdr.GetValue(3) }
                message = if ($rdr.IsDBNull(4)) { "" } else { [string]$rdr.GetValue(4) }
                isRead = $isRead
                createdAt = if ($rdr.IsDBNull(6)) { $null } else { $rdr.GetValue(6) }
                readAt = if ($rdr.IsDBNull(7)) { $null } else { $rdr.GetValue(7) }
            }
        }

        return @{
            success = $true
            data = @{
                items = @($items)
                unreadCount = $unread
            }
        }
    }
    catch {
        return @{
            success = $false
            httpStatus = 500
            error = $_.Exception.Message
        }
    }
    finally {
        if ($rdr) {
            try { $rdr.Close() } catch {}
            try { $rdr.Dispose() } catch {}
        }

        Close-WebApprovalDbContext -Context $ctx
    }
}


# ============================================================================
# PHASE 4 + PHASE 5
# SALES MANAGER INBOX / APPROVE / REJECT / SAFE BUSY COM SYNC
# ============================================================================

function Read-WebApprovalReaderValue {
    param(
        $Reader,
        [string]$Name,
        $DefaultValue = $null
    )

    try {
        $ordinal = $Reader.GetOrdinal($Name)
        if ($Reader.IsDBNull($ordinal)) {
            return $DefaultValue
        }
        return $Reader.GetValue($ordinal)
    }
    catch {
        return $DefaultValue
    }
}

function Convert-WebApprovalRecordToPublic {
    param(
        $Record,
        [bool]$IncludePayload = $false
    )

    if ($null -eq $Record) { return $null }

    $id = ([string]$Record.Id).Trim()
    $reference = $id
    if ($id.Length -ge 8) {
        $reference = "WA-" + $id.Substring(0, 8).ToUpperInvariant()
    }

    $out = [ordered]@{
        id = $id
        reference = $reference
        submittedBy = ([string]$Record.SubmittedBy).Trim()
        voucherType = [int]$Record.VoucherType
        voucherDate = $Record.VoucherDate
        requestedSeries = ([string]$Record.RequestedSeries).Trim()
        requestedVoucherNo = ([string]$Record.RequestedVoucherNo).Trim()
        partyName = ([string]$Record.PartyName).Trim()
        amount = [double]$Record.Amount
        approvalStatus = ([string]$Record.ApprovalStatus).Trim()
        syncStatus = ([string]$Record.SyncStatus).Trim()
        decisionBy = ([string]$Record.DecisionBy).Trim()
        decisionAt = $Record.DecisionAt
        rejectionReason = ([string]$Record.RejectionReason).Trim()
        busyVoucherNo = ([string]$Record.BusyVoucherNo).Trim()
        busyVoucherCode = ([string]$Record.BusyVoucherCode).Trim()
        numberChangeReason = ([string]$Record.NumberChangeReason).Trim()
        numberResolvedAt = $Record.NumberResolvedAt
        syncAttempts = [int]$Record.SyncAttempts
        lastSyncError = ([string]$Record.LastSyncError).Trim()
        submittedAt = $Record.SubmittedAt
        updatedAt = $Record.UpdatedAt
        syncedAt = $Record.SyncedAt
        version = [int]$Record.Version
    }

    if ($IncludePayload) {
        $payload = $null
        try {
            if (-not [string]::IsNullOrWhiteSpace([string]$Record.PayloadJson)) {
                $payload = ([string]$Record.PayloadJson) | ConvertFrom-Json
            }
        }
        catch {
            $payload = $null
        }

        $out["payload"] = $payload
        $out["payloadJson"] = [string]$Record.PayloadJson
        $out["payloadHash"] = ([string]$Record.PayloadHash).Trim()
    }

    return $out
}

function Get-WebApprovalRecordByIdInternal {
    param(
        $Context,
        [string]$Id
    )

    if ([string]::IsNullOrWhiteSpace($Id)) {
        return $null
    }

    $cmd = $Context.connection.CreateCommand()
    $dbType = [int]$Context.dbType

    if ($dbType -eq 1) {
        $cmd.CommandText = @"
SELECT TOP 1
    Id, SubmittedBy, VoucherType, VoucherDate,
    RequestedSeries, RequestedVoucherNo, PartyName, Amount,
    PayloadJson, PayloadHash,
    ApprovalStatus, SyncStatus,
    DecisionBy, DecisionAt, RejectionReason,
    BusyVoucherNo, BusyVoucherCode,
    NumberChangeReason, NumberResolvedAt,
    SyncAttempts, LastSyncError,
    SubmittedAt, UpdatedAt, SyncedAt, Version
FROM dbo.BusyCloudWebApproval
WHERE Id=@id
"@
        [void](Add-WebApprovalCommandParameter `
            -Command $cmd -DbType 1 -Name "@id" `
            -Value $Id.Trim() -Kind Text -Size 36)
    }
    else {
        $cmd.CommandText = @"
SELECT TOP 1
    [Id], [SubmittedBy], [VoucherType], [VoucherDate],
    [RequestedSeries], [RequestedVoucherNo], [PartyName], [Amount],
    [PayloadJson], [PayloadHash],
    [ApprovalStatus], [SyncStatus],
    [DecisionBy], [DecisionAt], [RejectionReason],
    [BusyVoucherNo], [BusyVoucherCode],
    [NumberChangeReason], [NumberResolvedAt],
    [SyncAttempts], [LastSyncError],
    [SubmittedAt], [UpdatedAt], [SyncedAt], [Version]
FROM [BusyCloudWebApproval]
WHERE [Id]=?
"@
        [void](Add-WebApprovalCommandParameter `
            -Command $cmd -DbType 0 -Name "@p1" `
            -Value $Id.Trim() -Kind Text -Size 36)
    }

    $rdr = $null

    try {
        $rdr = $cmd.ExecuteReader()
        if (-not $rdr.Read()) {
            return $null
        }

        return @{
            Id = Read-WebApprovalReaderValue $rdr "Id" ""
            SubmittedBy = Read-WebApprovalReaderValue $rdr "SubmittedBy" ""
            VoucherType = Read-WebApprovalReaderValue $rdr "VoucherType" 0
            VoucherDate = Read-WebApprovalReaderValue $rdr "VoucherDate" $null
            RequestedSeries = Read-WebApprovalReaderValue $rdr "RequestedSeries" ""
            RequestedVoucherNo = Read-WebApprovalReaderValue $rdr "RequestedVoucherNo" ""
            PartyName = Read-WebApprovalReaderValue $rdr "PartyName" ""
            Amount = Read-WebApprovalReaderValue $rdr "Amount" 0
            PayloadJson = Read-WebApprovalReaderValue $rdr "PayloadJson" ""
            PayloadHash = Read-WebApprovalReaderValue $rdr "PayloadHash" ""
            ApprovalStatus = Read-WebApprovalReaderValue $rdr "ApprovalStatus" ""
            SyncStatus = Read-WebApprovalReaderValue $rdr "SyncStatus" ""
            DecisionBy = Read-WebApprovalReaderValue $rdr "DecisionBy" ""
            DecisionAt = Read-WebApprovalReaderValue $rdr "DecisionAt" $null
            RejectionReason = Read-WebApprovalReaderValue $rdr "RejectionReason" ""
            BusyVoucherNo = Read-WebApprovalReaderValue $rdr "BusyVoucherNo" ""
            BusyVoucherCode = Read-WebApprovalReaderValue $rdr "BusyVoucherCode" ""
            NumberChangeReason = Read-WebApprovalReaderValue $rdr "NumberChangeReason" ""
            NumberResolvedAt = Read-WebApprovalReaderValue $rdr "NumberResolvedAt" $null
            SyncAttempts = Read-WebApprovalReaderValue $rdr "SyncAttempts" 0
            LastSyncError = Read-WebApprovalReaderValue $rdr "LastSyncError" ""
            SubmittedAt = Read-WebApprovalReaderValue $rdr "SubmittedAt" $null
            UpdatedAt = Read-WebApprovalReaderValue $rdr "UpdatedAt" $null
            SyncedAt = Read-WebApprovalReaderValue $rdr "SyncedAt" $null
            Version = Read-WebApprovalReaderValue $rdr "Version" 1
        }
    }
    finally {
        if ($rdr) {
            try { $rdr.Close() } catch {}
            try { $rdr.Dispose() } catch {}
        }
    }
}

function Test-WebApprovalManagerRole {
    param(
        [string]$ManagerUserName,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ([string]::IsNullOrWhiteSpace($ManagerUserName)) {
        return @{
            success = $false
            allowed = $false
            httpStatus = 401
            error = "Authenticated Sales Manager is required."
        }
    }

    $roleResult = Get-WebApprovalUserRole `
        -UserName $ManagerUserName.Trim() `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $roleResult.success) {
        return @{
            success = $false
            allowed = $false
            httpStatus = 500
            error = $roleResult.error
        }
    }

    $role = 0
    try { $role = [int]$roleResult.data.role } catch {}

    if ($role -ne $script:WebApprovalRoleSalesManager) {
        return @{
            success = $true
            allowed = $false
            httpStatus = 403
            error = "The signed-in user does not have the SALES_MANAGER workflow role."
        }
    }

    return @{
        success = $true
        allowed = $true
        role = $role
    }
}

function Get-WebApprovalSalesmenForManager {
    param(
        [string]$ManagerUserName,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $role = Test-WebApprovalManagerRole `
        -ManagerUserName $ManagerUserName `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $role.success -or -not $role.allowed) {
        return $role
    }

    $ctx = $null
    $rdr = $null

    try {
        $ctx = Get-WebApprovalPermanentDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        [void](Ensure-WebApprovalPermanentTables -Context $ctx)

        $cmd = $ctx.connection.CreateCommand()

        if ([int]$ctx.dbType -eq 1) {
            $cmd.CommandText = @"
SELECT SalesmanUserName
FROM dbo.BusyCloudSalesmanManager
WHERE ManagerUserName=@manager
  AND IsActive<>0
ORDER BY SalesmanUserName
"@
            [void](Add-WebApprovalCommandParameter `
                -Command $cmd -DbType 1 -Name "@manager" `
                -Value $ManagerUserName.Trim() -Kind Text -Size 100)
        }
        else {
            $cmd.CommandText = @"
SELECT [SalesmanUserName]
FROM [BusyCloudSalesmanManager]
WHERE [ManagerUserName]=?
  AND [IsActive]<>0
ORDER BY [SalesmanUserName]
"@
            [void](Add-WebApprovalCommandParameter `
                -Command $cmd -DbType 0 -Name "@p1" `
                -Value $ManagerUserName.Trim() -Kind Text -Size 100)
        }

        $rdr = $cmd.ExecuteReader()
        $items = @()

        while ($rdr.Read()) {
            $name = ""
            if (-not $rdr.IsDBNull(0)) {
                $name = ([string]$rdr.GetValue(0)).Trim()
            }

            if ($name -and $items -notcontains $name) {
                $items += $name
            }
        }

        return @{
            success = $true
            allowed = $true
            data = @{
                managerUserName = $ManagerUserName.Trim()
                salesmen = @($items)
            }
        }
    }
    catch {
        return @{
            success = $false
            allowed = $false
            httpStatus = 500
            error = $_.Exception.Message
        }
    }
    finally {
        if ($rdr) {
            try { $rdr.Close() } catch {}
            try { $rdr.Dispose() } catch {}
        }
        Close-WebApprovalDbContext -Context $ctx
    }
}

function Test-WebApprovalManagerCanActOnRecord {
    param(
        [string]$ManagerUserName,
        $Record,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ($null -eq $Record) {
        return @{
            success = $false
            allowed = $false
            httpStatus = 404
            error = "Web Approval transaction was not found."
        }
    }

    $role = Test-WebApprovalManagerRole `
        -ManagerUserName $ManagerUserName `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $role.success -or -not $role.allowed) {
        return $role
    }

    $mapping = Test-WebApprovalManagerAccess `
        -SalesmanUserName ([string]$Record.SubmittedBy) `
        -ManagerUserName $ManagerUserName `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $mapping.success) {
        return @{
            success = $false
            allowed = $false
            httpStatus = 500
            error = $mapping.error
        }
    }

    if (-not $mapping.allowed) {
        return @{
            success = $true
            allowed = $false
            httpStatus = 403
            error = "This transaction belongs to a Salesman who is not assigned to you."
        }
    }

    return @{
        success = $true
        allowed = $true
    }
}

function Test-WebApprovalQueueStatusMatch {
    param(
        $Record,
        [string]$Status
    )

    $statusText = ([string]$Status).Trim().ToUpperInvariant()
    if (-not $statusText -or $statusText -eq "ALL") {
        return $true
    }

    $approval = ([string]$Record.ApprovalStatus).Trim().ToUpperInvariant()
    $sync = ([string]$Record.SyncStatus).Trim().ToUpperInvariant()

    switch ($statusText) {
        "PENDING" { return ($approval -eq "PENDING") }
        "APPROVED" {
            return (
                $approval -eq "APPROVED" -and
                $sync -notin @("SYNCED", "REVIEW_REQUIRED")
            )
        }
        "FAILED" {
            return (
                $approval -eq "APPROVED" -and
                $sync -eq "FAILED"
            )
        }
        "REJECTED" { return ($approval -eq "REJECTED") }
        "SYNCED" { return ($sync -eq "SYNCED") }
        "REVIEW_REQUIRED" { return ($sync -eq "REVIEW_REQUIRED") }
        default { return $true }
    }
}

function Get-WebApprovalAssignedQueue {
    param(
        [string]$ManagerUserName,
        [string]$Status = "PENDING",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $salesmenResult = Get-WebApprovalSalesmenForManager `
        -ManagerUserName $ManagerUserName `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $salesmenResult.success -or -not $salesmenResult.allowed) {
        return $salesmenResult
    }

    $assignedSalesmen = @($salesmenResult.data.salesmen)
    $assignedLookup = @{}

    foreach ($name in $assignedSalesmen) {
        $key = ([string]$name).Trim().ToLowerInvariant()
        if ($key) { $assignedLookup[$key] = $true }
    }

    if ($assignedLookup.Count -eq 0) {
        return @{
            success = $true
            data = @{
                status = $Status
                items = @()
                count = 0
                assignedSalesmen = @()
            }
        }
    }

    $ctx = $null
    $rdr = $null

    try {
        $ctx = Get-WebApprovalFiscalDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        [void](Ensure-WebApprovalFiscalTables -Context $ctx)

        $cmd = $ctx.connection.CreateCommand()

        if ([int]$ctx.dbType -eq 1) {
            $cmd.CommandText = @"
SELECT TOP 500
    Id, SubmittedBy, VoucherType, VoucherDate,
    RequestedSeries, RequestedVoucherNo, PartyName, Amount,
    ApprovalStatus, SyncStatus,
    DecisionBy, DecisionAt, RejectionReason,
    BusyVoucherNo, BusyVoucherCode,
    NumberChangeReason, NumberResolvedAt,
    SyncAttempts, LastSyncError,
    SubmittedAt, UpdatedAt, SyncedAt, Version
FROM dbo.BusyCloudWebApproval
ORDER BY SubmittedAt DESC
"@
        }
        else {
            $cmd.CommandText = @"
SELECT TOP 500
    [Id], [SubmittedBy], [VoucherType], [VoucherDate],
    [RequestedSeries], [RequestedVoucherNo], [PartyName], [Amount],
    [ApprovalStatus], [SyncStatus],
    [DecisionBy], [DecisionAt], [RejectionReason],
    [BusyVoucherNo], [BusyVoucherCode],
    [NumberChangeReason], [NumberResolvedAt],
    [SyncAttempts], [LastSyncError],
    [SubmittedAt], [UpdatedAt], [SyncedAt], [Version]
FROM [BusyCloudWebApproval]
ORDER BY [SubmittedAt] DESC
"@
        }

        $rdr = $cmd.ExecuteReader()
        $items = @()

        while ($rdr.Read()) {
            $record = @{
                Id = Read-WebApprovalReaderValue $rdr "Id" ""
                SubmittedBy = Read-WebApprovalReaderValue $rdr "SubmittedBy" ""
                VoucherType = Read-WebApprovalReaderValue $rdr "VoucherType" 0
                VoucherDate = Read-WebApprovalReaderValue $rdr "VoucherDate" $null
                RequestedSeries = Read-WebApprovalReaderValue $rdr "RequestedSeries" ""
                RequestedVoucherNo = Read-WebApprovalReaderValue $rdr "RequestedVoucherNo" ""
                PartyName = Read-WebApprovalReaderValue $rdr "PartyName" ""
                Amount = Read-WebApprovalReaderValue $rdr "Amount" 0
                ApprovalStatus = Read-WebApprovalReaderValue $rdr "ApprovalStatus" ""
                SyncStatus = Read-WebApprovalReaderValue $rdr "SyncStatus" ""
                DecisionBy = Read-WebApprovalReaderValue $rdr "DecisionBy" ""
                DecisionAt = Read-WebApprovalReaderValue $rdr "DecisionAt" $null
                RejectionReason = Read-WebApprovalReaderValue $rdr "RejectionReason" ""
                BusyVoucherNo = Read-WebApprovalReaderValue $rdr "BusyVoucherNo" ""
                BusyVoucherCode = Read-WebApprovalReaderValue $rdr "BusyVoucherCode" ""
                NumberChangeReason = Read-WebApprovalReaderValue $rdr "NumberChangeReason" ""
                NumberResolvedAt = Read-WebApprovalReaderValue $rdr "NumberResolvedAt" $null
                SyncAttempts = Read-WebApprovalReaderValue $rdr "SyncAttempts" 0
                LastSyncError = Read-WebApprovalReaderValue $rdr "LastSyncError" ""
                SubmittedAt = Read-WebApprovalReaderValue $rdr "SubmittedAt" $null
                UpdatedAt = Read-WebApprovalReaderValue $rdr "UpdatedAt" $null
                SyncedAt = Read-WebApprovalReaderValue $rdr "SyncedAt" $null
                Version = Read-WebApprovalReaderValue $rdr "Version" 1
            }

            $submitterKey = ([string]$record.SubmittedBy).Trim().ToLowerInvariant()
            if (-not $assignedLookup.ContainsKey($submitterKey)) {
                continue
            }

            if (-not (Test-WebApprovalQueueStatusMatch -Record $record -Status $Status)) {
                continue
            }

            $items += Convert-WebApprovalRecordToPublic -Record $record
        }

        return @{
            success = $true
            data = @{
                status = $Status
                items = @($items)
                count = @($items).Count
                assignedSalesmen = @($assignedSalesmen)
            }
        }
    }
    catch {
        return @{
            success = $false
            httpStatus = 500
            error = $_.Exception.Message
        }
    }
    finally {
        if ($rdr) {
            try { $rdr.Close() } catch {}
            try { $rdr.Dispose() } catch {}
        }
        Close-WebApprovalDbContext -Context $ctx
    }
}

function Get-WebApprovalActionHistoryInternal {
    param(
        $Context,
        [string]$WebApprovalId
    )

    $cmd = $Context.connection.CreateCommand()

    if ([int]$Context.dbType -eq 1) {
        $cmd.CommandText = @"
SELECT
    Id, Action, ActionBy, Remarks, MetadataJson, ActionTime
FROM dbo.BusyCloudWebApprovalAction
WHERE WebApprovalId=@id
ORDER BY ActionTime ASC
"@
        [void](Add-WebApprovalCommandParameter `
            -Command $cmd -DbType 1 -Name "@id" `
            -Value $WebApprovalId.Trim() -Kind Text -Size 36)
    }
    else {
        $cmd.CommandText = @"
SELECT
    [Id], [Action], [ActionBy], [Remarks], [MetadataJson], [ActionTime]
FROM [BusyCloudWebApprovalAction]
WHERE [WebApprovalId]=?
ORDER BY [ActionTime] ASC
"@
        [void](Add-WebApprovalCommandParameter `
            -Command $cmd -DbType 0 -Name "@p1" `
            -Value $WebApprovalId.Trim() -Kind Text -Size 36)
    }

    $rdr = $null
    $items = @()

    try {
        $rdr = $cmd.ExecuteReader()

        while ($rdr.Read()) {
            $metadataText = [string](Read-WebApprovalReaderValue $rdr "MetadataJson" "")
            $metadata = $null

            if (-not [string]::IsNullOrWhiteSpace($metadataText)) {
                try { $metadata = $metadataText | ConvertFrom-Json } catch {}
            }

            $items += @{
                id = [string](Read-WebApprovalReaderValue $rdr "Id" "")
                action = [string](Read-WebApprovalReaderValue $rdr "Action" "")
                actionBy = [string](Read-WebApprovalReaderValue $rdr "ActionBy" "")
                remarks = [string](Read-WebApprovalReaderValue $rdr "Remarks" "")
                metadata = $metadata
                actionTime = Read-WebApprovalReaderValue $rdr "ActionTime" $null
            }
        }
    }
    finally {
        if ($rdr) {
            try { $rdr.Close() } catch {}
            try { $rdr.Dispose() } catch {}
        }
    }

    return @($items)
}

function Get-WebApprovalDetailForManager {
    param(
        [string]$Id,
        [string]$ManagerUserName,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $ctx = $null

    try {
        $ctx = Get-WebApprovalFiscalDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        $record = Get-WebApprovalRecordByIdInternal `
            -Context $ctx `
            -Id $Id

        if ($null -eq $record) {
            return @{
                success = $false
                httpStatus = 404
                error = "Web Approval transaction was not found."
            }
        }

        $access = Test-WebApprovalManagerCanActOnRecord `
            -ManagerUserName $ManagerUserName `
            -Record $record `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (-not $access.success -or -not $access.allowed) {
            return $access
        }

        $history = Get-WebApprovalActionHistoryInternal `
            -Context $ctx `
            -WebApprovalId $Id

        return @{
            success = $true
            data = @{
                transaction = Convert-WebApprovalRecordToPublic `
                    -Record $record `
                    -IncludePayload $true
                actions = @($history)
            }
        }
    }
    catch {
        return @{
            success = $false
            httpStatus = 500
            error = $_.Exception.Message
        }
    }
    finally {
        Close-WebApprovalDbContext -Context $ctx
    }
}

function Set-WebApprovalDecision {
    param(
        [string]$Id,
        [string]$ManagerUserName,
        [ValidateSet("APPROVE", "REJECT")]
        [string]$Decision,
        [string]$Remarks = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ($Decision -eq "REJECT" -and [string]::IsNullOrWhiteSpace($Remarks)) {
        return @{
            success = $false
            httpStatus = 400
            errorCode = "REJECTION_REASON_REQUIRED"
            error = "A rejection reason is required."
        }
    }

    $ctx = $null
    $tx = $null

    try {
        $ctx = Get-WebApprovalFiscalDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        $record = Get-WebApprovalRecordByIdInternal `
            -Context $ctx `
            -Id $Id

        if ($null -eq $record) {
            return @{
                success = $false
                httpStatus = 404
                error = "Web Approval transaction was not found."
            }
        }

        $access = Test-WebApprovalManagerCanActOnRecord `
            -ManagerUserName $ManagerUserName `
            -Record $record `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (-not $access.success -or -not $access.allowed) {
            return $access
        }

        if (([string]$record.ApprovalStatus).Trim().ToUpperInvariant() -ne "PENDING") {
            return @{
                success = $false
                httpStatus = 409
                errorCode = "WEB_APPROVAL_ALREADY_DECIDED"
                error = "This transaction has already been approved or rejected by another manager."
            }
        }

        $now = [datetime]::UtcNow
        $newApproval = if ($Decision -eq "APPROVE") { "APPROVED" } else { "REJECTED" }
        $newSync = if ($Decision -eq "APPROVE") { "READY" } else { "NOT_READY" }
        $reason = if ($Decision -eq "REJECT") { $Remarks.Trim() } else { "" }

        $tx = $ctx.connection.BeginTransaction()
        $cmd = $ctx.connection.CreateCommand()
        $cmd.Transaction = $tx

        if ([int]$ctx.dbType -eq 1) {
            $cmd.CommandText = @"
UPDATE dbo.BusyCloudWebApproval
SET
    ApprovalStatus=@approvalStatus,
    SyncStatus=@syncStatus,
    DecisionBy=@decisionBy,
    DecisionAt=@decisionAt,
    RejectionReason=@rejectionReason,
    LastSyncError=NULL,
    UpdatedAt=@updatedAt,
    Version=Version+1
WHERE Id=@id
  AND ApprovalStatus='PENDING'
"@

            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@approvalStatus" -Value $newApproval -Kind Text -Size 20)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@syncStatus" -Value $newSync -Kind Text -Size 30)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@decisionBy" -Value $ManagerUserName.Trim() -Kind Text -Size 100)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@decisionAt" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@rejectionReason" -Value $reason -Kind LongText)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@updatedAt" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@id" -Value $Id.Trim() -Kind Text -Size 36)
        }
        else {
            $cmd.CommandText = @"
UPDATE [BusyCloudWebApproval]
SET
    [ApprovalStatus]=?,
    [SyncStatus]=?,
    [DecisionBy]=?,
    [DecisionAt]=?,
    [RejectionReason]=?,
    [LastSyncError]=NULL,
    [UpdatedAt]=?,
    [Version]=[Version]+1
WHERE [Id]=?
  AND [ApprovalStatus]='PENDING'
"@

            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p1" -Value $newApproval -Kind Text -Size 20)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p2" -Value $newSync -Kind Text -Size 30)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p3" -Value $ManagerUserName.Trim() -Kind Text -Size 100)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p4" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p5" -Value $reason -Kind LongText)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p6" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p7" -Value $Id.Trim() -Kind Text -Size 36)
        }

        $affected = [int]$cmd.ExecuteNonQuery()

        if ($affected -ne 1) {
            try { $tx.Rollback() } catch {}
            $tx = $null

            return @{
                success = $false
                httpStatus = 409
                errorCode = "WEB_APPROVAL_DECISION_CONFLICT"
                error = "Another manager already decided this transaction. Refresh the page."
            }
        }

        $actionName = if ($Decision -eq "APPROVE") { "APPROVED" } else { "REJECTED" }

        Add-WebApprovalActionRow `
            -Context $ctx `
            -Transaction $tx `
            -WebApprovalId $Id.Trim() `
            -Action $actionName `
            -ActionBy $ManagerUserName.Trim() `
            -Remarks $Remarks.Trim() `
            -MetadataJson "" `
            -ActionTime $now

        $reference = "WA-" + $Id.Trim().Substring(0, [Math]::Min(8, $Id.Trim().Length)).ToUpperInvariant()
        $notificationType = if ($Decision -eq "APPROVE") {
            "WEB_APPROVAL_APPROVED"
        } else {
            "WEB_APPROVAL_REJECTED"
        }

        $title = if ($Decision -eq "APPROVE") {
            "Web Approval Approved"
        } else {
            "Web Approval Rejected"
        }

        $message = if ($Decision -eq "APPROVE") {
            "$reference was approved by $($ManagerUserName.Trim()) and is ready to synchronize to BUSY."
        } else {
            "$reference was rejected by $($ManagerUserName.Trim()). Reason: $($Remarks.Trim())"
        }

        Add-WebApprovalNotificationRow `
            -Context $ctx `
            -Transaction $tx `
            -RecipientUserName ([string]$record.SubmittedBy).Trim() `
            -WebApprovalId $Id.Trim() `
            -NotificationType $notificationType `
            -Title $title `
            -Message $message `
            -CreatedAt $now

        $tx.Commit()
        $tx = $null

        return @{
            success = $true
            message = if ($Decision -eq "APPROVE") {
                "Transaction approved. It is now ready to synchronize to BUSY."
            } else {
                "Transaction rejected."
            }
            data = @{
                id = $Id.Trim()
                approvalStatus = $newApproval
                syncStatus = $newSync
                decisionBy = $ManagerUserName.Trim()
                decisionAt = $now
                rejectionReason = $reason
            }
        }
    }
    catch {
        if ($tx) {
            try { $tx.Rollback() } catch {}
        }

        return @{
            success = $false
            httpStatus = 500
            error = $_.Exception.Message
        }
    }
    finally {
        Close-WebApprovalDbContext -Context $ctx
    }
}

function Approve-WebApprovalVoucher {
    param(
        [string]$Id,
        [string]$ManagerUserName,
        [string]$Remarks = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    return Set-WebApprovalDecision `
        -Id $Id `
        -ManagerUserName $ManagerUserName `
        -Decision "APPROVE" `
        -Remarks $Remarks `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode
}

function Reject-WebApprovalVoucher {
    param(
        [string]$Id,
        [string]$ManagerUserName,
        [string]$Reason,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    return Set-WebApprovalDecision `
        -Id $Id `
        -ManagerUserName $ManagerUserName `
        -Decision "REJECT" `
        -Remarks $Reason `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode
}

function Copy-WebApprovalPayloadForPosting {
    param([string]$PayloadJson)

    if ([string]::IsNullOrWhiteSpace($PayloadJson)) {
        throw "Stored approved payload is empty."
    }

    return $PayloadJson | ConvertFrom-Json
}

function Update-WebApprovalSelfNumberReferences {
    param(
        $Payload,
        [string]$RequestedVoucherNo,
        [string]$FinalVoucherNo
    )

    $requested = ([string]$RequestedVoucherNo).Trim()
    $final = ([string]$FinalVoucherNo).Trim()

    if (
        $null -eq $Payload -or
        [string]::IsNullOrWhiteSpace($requested) -or
        [string]::IsNullOrWhiteSpace($final) -or
        $requested.Equals($final, [System.StringComparison]::OrdinalIgnoreCase)
    ) {
        return $Payload
    }

    try {
        foreach ($entry in @($Payload.refEntries)) {
            if ($null -eq $entry) { continue }

            $refType = 0
            try { $refType = [int]$entry.refType } catch {}

            $refNo = ([string]$entry.refNo).Trim()

            if (
                $refType -eq 1 -and
                $refNo.Equals($requested, [System.StringComparison]::OrdinalIgnoreCase)
            ) {
                $entry.refNo = $final
            }
        }
    }
    catch {}

    try {
        foreach ($account in @($Payload.accounts)) {
            if ($null -eq $account) { continue }

            foreach ($entry in @($account.bbaEntries)) {
                if ($null -eq $entry) { continue }

                $refType = 0
                try { $refType = [int]$entry.refType } catch {}

                $refNo = ([string]$entry.refNo).Trim()

                if (
                    $refType -eq 1 -and
                    $refNo.Equals($requested, [System.StringComparison]::OrdinalIgnoreCase)
                ) {
                    $entry.refNo = $final
                }
            }
        }
    }
    catch {}

    return $Payload
}

function Resolve-WebApprovalPostingNumber {
    param(
        $Record,
        [string]$ManualVoucherNo = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if (-not (Get-Command Get-EffectiveNumberingConfig -ErrorAction SilentlyContinue)) {
        return @{
            success = $false
            httpStatus = 500
            errorCode = "NUMBERING_ENGINE_NOT_LOADED"
            error = "Current BUSY voucher numbering engine is not loaded."
        }
    }

    $voucherDate = $null

    try {
        $voucherDate = ConvertTo-WebApprovalVoucherDate -Value $Record.VoucherDate
    }
    catch {
        try {
            $voucherDate = ConvertTo-WebApprovalVoucherDate `
                -Value (
                    Get-WebApprovalPropertyValue `
                        -Object (($Record.PayloadJson | ConvertFrom-Json)) `
                        -Names @("date") `
                        -DefaultValue ""
                )
        }
        catch {
            return @{
                success = $false
                httpStatus = 400
                errorCode = "INVALID_VOUCHER_DATE"
                error = $_.Exception.Message
            }
        }
    }

    $numberingDate = $voucherDate.ToString("yyyy-MM-dd")
    $vchType = [int]$Record.VoucherType
    $series = ([string]$Record.RequestedSeries).Trim()

    $config = Get-EffectiveNumberingConfig `
        -VchType $vchType `
        -SeriesName $series `
        -VoucherDate $numberingDate `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if ($null -eq $config -or $config.success -eq $false) {
        return @{
            success = $false
            httpStatus = 400
            errorCode = "NUMBERING_CONFIG_FAILED"
            error = if ($config -and $config.error) {
                [string]$config.error
            } else {
                "Could not resolve current BUSY voucher numbering."
            }
        }
    }

    $configData = Get-WebApprovalPropertyValue `
        -Object $config `
        -Names @("data") `
        -DefaultValue $config

    $rawAuto = Get-WebApprovalPropertyValue `
        -Object $configData `
        -Names @("is_auto", "isAuto") `
        -DefaultValue $false

    $isAuto = $false
    $autoText = ([string]$rawAuto).Trim().ToLowerInvariant()
    if ($rawAuto -eq $true -or $autoText -in @("1", "true", "yes")) {
        $isAuto = $true
    }

    if ($isAuto) {
        $nextNumber = ([string](
            Get-WebApprovalPropertyValue `
                -Object $configData `
                -Names @("next_vch_no", "nextVchNo") `
                -DefaultValue ""
        )).Trim()

        if ([string]::IsNullOrWhiteSpace($nextNumber)) {
            return @{
                success = $false
                httpStatus = 400
                errorCode = "NEXT_VOUCHER_NUMBER_UNAVAILABLE"
                error = "This series uses automatic numbering, but the latest BUSY voucher number could not be calculated."
            }
        }

        return @{
            success = $true
            isAuto = $true
            vchNo = $nextNumber
            config = $configData
        }
    }

    $manual = ([string]$ManualVoucherNo).Trim()

    if ([string]::IsNullOrWhiteSpace($manual)) {
        $manual = ([string]$Record.RequestedVoucherNo).Trim()
    }

    if ([string]::IsNullOrWhiteSpace($manual)) {
        return @{
            success = $false
            httpStatus = 400
            errorCode = "MANUAL_VOUCHER_NUMBER_REQUIRED"
            error = "This series uses manual numbering. A voucher number is required."
        }
    }

    return @{
        success = $true
        isAuto = $false
        vchNo = $manual
        config = $configData
    }
}

function Find-WebApprovalBusyVoucherCom {
    param(
        [int]$VchType,
        [string]$VchNo = "",
        [string]$SeriesName = "",
        [int]$VchCode = 0,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $fi = $null
    $rst = $null
    $seriesRst = $null

    try {
        $fi = Connect-BUSY `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        if (-not $fi) {
            return @{
                success = $false
                error = "BUSY COM connection failed while verifying voucher number."
            }
        }

        $sql = ""

        if ($VchCode -gt 0) {
            $sql = "SELECT TOP 1 VchCode,VchNo,VchSeriesCode,[Date] FROM Tran1 WHERE VchCode=$VchCode AND VchType=$VchType"
        }
        else {
            $safeNo = ([string]$VchNo).Replace("'", "''")

            if ([string]::IsNullOrWhiteSpace($safeNo)) {
                return @{ success=$true; found=$false }
            }

            $seriesFilter = ""
            $series = ([string]$SeriesName).Trim()

            if (-not [string]::IsNullOrWhiteSpace($series)) {
                $safeSeries = $series.Replace("'", "''")
                $prefixed = ("{0:D2}{1}" -f $VchType, $series).Replace("'", "''")

                $seriesFilter = " AND VchSeriesCode IN (SELECT Code FROM Master1 WHERE MasterType=21 AND (Name='$safeSeries' OR Name='$prefixed'))"
            }

            $sql = "SELECT TOP 1 VchCode,VchNo,VchSeriesCode,[Date] FROM Tran1 WHERE VchType=$VchType AND VchNo='$safeNo'$seriesFilter ORDER BY VchCode DESC"
        }

        $rst = $fi.GetRecordset($sql)

        if ($null -eq $rst -or $rst.EOF) {
            return @{
                success = $true
                found = $false
            }
        }

        $code = 0
        try { $code = [int]$rst.Fields.Item("VchCode").Value } catch {}

        $number = ""
        try { $number = ([string]$rst.Fields.Item("VchNo").Value).Trim() } catch {}

        $seriesCode = 0
        try { $seriesCode = [int]$rst.Fields.Item("VchSeriesCode").Value } catch {}

        $date = $null
        try { $date = $rst.Fields.Item("Date").Value } catch {}

        $seriesResolved = ""

        if ($seriesCode -gt 0) {
            try {
                $seriesRst = $fi.GetRecordset(
                    "SELECT TOP 1 Name FROM Master1 WHERE Code=$seriesCode"
                )

                if ($seriesRst -and -not $seriesRst.EOF) {
                    $seriesResolved = ([string]$seriesRst.Fields.Item("Name").Value).Trim()

                    $prefix = ("{0:D2}" -f $VchType)
                    if (
                        $seriesResolved.StartsWith(
                            $prefix,
                            [System.StringComparison]::OrdinalIgnoreCase
                        )
                    ) {
                        $seriesResolved = $seriesResolved.Substring($prefix.Length).Trim()
                    }
                }
            }
            catch {}
        }

        return @{
            success = $true
            found = $true
            vchCode = $code
            vchNo = $number
            vchSeriesCode = $seriesCode
            vchSeries = $seriesResolved
            date = $date
        }
    }
    catch {
        return @{
            success = $false
            found = $false
            error = $_.Exception.Message
        }
    }
    finally {
        if ($seriesRst) { try { $seriesRst.Close() } catch {} }
        if ($rst) { try { $rst.Close() } catch {} }
        if ($fi) { try { Disconnect-BUSY $fi } catch {} }
    }
}

function Get-WebApprovalCreateResultInfo {
    param($Result)

    if ($null -eq $Result) {
        return @{
            success = $false
            created = $false
            vchNo = ""
            vchCode = 0
            message = ""
            error = "Create-Voucher returned no result."
        }
    }

    $successRaw = Get-WebApprovalPropertyValue `
        -Object $Result `
        -Names @("success") `
        -DefaultValue $false

    $createdRaw = Get-WebApprovalPropertyValue `
        -Object $Result `
        -Names @("created") `
        -DefaultValue $false

    $success = (
        $successRaw -eq $true -or
        ([string]$successRaw).Trim().ToLowerInvariant() -in @("1", "true", "yes")
    )

    $created = (
        $createdRaw -eq $true -or
        ([string]$createdRaw).Trim().ToLowerInvariant() -in @("1", "true", "yes")
    )

    $data = Get-WebApprovalPropertyValue `
        -Object $Result `
        -Names @("data") `
        -DefaultValue $null

    $vchNo = ([string](
        Get-WebApprovalPropertyValue `
            -Object $Result `
            -Names @("finalVchNo", "vchNo", "voucherNo") `
            -DefaultValue ""
    )).Trim()

    if ([string]::IsNullOrWhiteSpace($vchNo) -and $null -ne $data) {
        $vchNo = ([string](
            Get-WebApprovalPropertyValue `
                -Object $data `
                -Names @("finalVchNo", "vchNo", "voucherNo") `
                -DefaultValue ""
        )).Trim()
    }

    $vchCode = 0
    $rawCode = Get-WebApprovalPropertyValue `
        -Object $Result `
        -Names @("finalVchCode", "vchCode", "voucherCode") `
        -DefaultValue 0

    [void][int]::TryParse([string]$rawCode, [ref]$vchCode)

    if ($vchCode -le 0 -and $null -ne $data) {
        $rawCode = Get-WebApprovalPropertyValue `
            -Object $data `
            -Names @("finalVchCode", "vchCode", "voucherCode") `
            -DefaultValue 0

        [void][int]::TryParse([string]$rawCode, [ref]$vchCode)
    }

    $message = ([string](
        Get-WebApprovalPropertyValue `
            -Object $Result `
            -Names @("message") `
            -DefaultValue ""
    )).Trim()

    $error = ([string](
        Get-WebApprovalPropertyValue `
            -Object $Result `
            -Names @("error", "warning") `
            -DefaultValue ""
    )).Trim()

    return @{
        success = $success
        created = $created
        vchNo = $vchNo
        vchCode = $vchCode
        message = $message
        error = $error
    }
}

function Set-WebApprovalSyncLedgerInternal {
    param(
        $Context,
        $Transaction,
        [string]$WebApprovalId,
        [string]$PayloadHash,
        [string]$State,
        [string]$RequestedVoucherNo,
        [string]$BusyVoucherNo = "",
        [string]$BusyVoucherCode = "",
        [string]$NumberChangeReason = "",
        $NumberResolvedAt = $null,
        [int]$Attempts = 0,
        [string]$LastError = ""
    )

    $now = [datetime]::UtcNow
    $dbType = [int]$Context.dbType

    $cmd = $Context.connection.CreateCommand()
    $cmd.Transaction = $Transaction

    if ($dbType -eq 1) {
        $cmd.CommandText = @"
UPDATE dbo.BusyCloudWebSyncLedger
SET
    PayloadHash=@payloadHash,
    State=@state,
    RequestedVoucherNo=@requestedVoucherNo,
    BusyVoucherNo=@busyVoucherNo,
    BusyVoucherCode=@busyVoucherCode,
    NumberChangeReason=@numberChangeReason,
    NumberResolvedAt=@numberResolvedAt,
    Attempts=@attempts,
    LastError=@lastError,
    UpdatedAt=@updatedAt
WHERE WebApprovalId=@id
"@

        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@payloadHash" -Value $PayloadHash -Kind Text -Size 64)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@state" -Value $State -Kind Text -Size 30)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@requestedVoucherNo" -Value $RequestedVoucherNo -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@busyVoucherNo" -Value $BusyVoucherNo -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@busyVoucherCode" -Value $BusyVoucherCode -Kind Text -Size 50)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@numberChangeReason" -Value $NumberChangeReason -Kind LongText)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@numberResolvedAt" -Value $NumberResolvedAt -Kind Date)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@attempts" -Value $Attempts -Kind Int)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@lastError" -Value $LastError -Kind LongText)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@updatedAt" -Value $now -Kind Date)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@id" -Value $WebApprovalId -Kind Text -Size 36)
    }
    else {
        $cmd.CommandText = @"
UPDATE [BusyCloudWebSyncLedger]
SET
    [PayloadHash]=?,
    [State]=?,
    [RequestedVoucherNo]=?,
    [BusyVoucherNo]=?,
    [BusyVoucherCode]=?,
    [NumberChangeReason]=?,
    [NumberResolvedAt]=?,
    [Attempts]=?,
    [LastError]=?,
    [UpdatedAt]=?
WHERE [WebApprovalId]=?
"@

        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p1" -Value $PayloadHash -Kind Text -Size 64)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p2" -Value $State -Kind Text -Size 30)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p3" -Value $RequestedVoucherNo -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p4" -Value $BusyVoucherNo -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p5" -Value $BusyVoucherCode -Kind Text -Size 50)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p6" -Value $NumberChangeReason -Kind LongText)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p7" -Value $NumberResolvedAt -Kind Date)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p8" -Value $Attempts -Kind Int)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p9" -Value $LastError -Kind LongText)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p10" -Value $now -Kind Date)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p11" -Value $WebApprovalId -Kind Text -Size 36)
    }

    $affected = [int]$cmd.ExecuteNonQuery()

    if ($affected -gt 0) {
        return
    }

    $insert = $Context.connection.CreateCommand()
    $insert.Transaction = $Transaction

    if ($dbType -eq 1) {
        $insert.CommandText = @"
INSERT INTO dbo.BusyCloudWebSyncLedger
(
    WebApprovalId, PayloadHash, State,
    RequestedVoucherNo, BusyVoucherNo, BusyVoucherCode,
    NumberChangeReason, NumberResolvedAt,
    Attempts, LastError, CreatedAt, UpdatedAt
)
VALUES
(
    @id, @payloadHash, @state,
    @requestedVoucherNo, @busyVoucherNo, @busyVoucherCode,
    @numberChangeReason, @numberResolvedAt,
    @attempts, @lastError, @createdAt, @updatedAt
)
"@

        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@id" -Value $WebApprovalId -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@payloadHash" -Value $PayloadHash -Kind Text -Size 64)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@state" -Value $State -Kind Text -Size 30)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@requestedVoucherNo" -Value $RequestedVoucherNo -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@busyVoucherNo" -Value $BusyVoucherNo -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@busyVoucherCode" -Value $BusyVoucherCode -Kind Text -Size 50)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@numberChangeReason" -Value $NumberChangeReason -Kind LongText)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@numberResolvedAt" -Value $NumberResolvedAt -Kind Date)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@attempts" -Value $Attempts -Kind Int)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@lastError" -Value $LastError -Kind LongText)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@createdAt" -Value $now -Kind Date)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@updatedAt" -Value $now -Kind Date)
    }
    else {
        $insert.CommandText = @"
INSERT INTO [BusyCloudWebSyncLedger]
(
    [WebApprovalId], [PayloadHash], [State],
    [RequestedVoucherNo], [BusyVoucherNo], [BusyVoucherCode],
    [NumberChangeReason], [NumberResolvedAt],
    [Attempts], [LastError], [CreatedAt], [UpdatedAt]
)
VALUES
(
    ?,?,?,?,?,?,?,?,?,?,?,?
)
"@

        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p1" -Value $WebApprovalId -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p2" -Value $PayloadHash -Kind Text -Size 64)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p3" -Value $State -Kind Text -Size 30)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p4" -Value $RequestedVoucherNo -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p5" -Value $BusyVoucherNo -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p6" -Value $BusyVoucherCode -Kind Text -Size 50)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p7" -Value $NumberChangeReason -Kind LongText)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p8" -Value $NumberResolvedAt -Kind Date)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p9" -Value $Attempts -Kind Int)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p10" -Value $LastError -Kind LongText)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p11" -Value $now -Kind Date)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p12" -Value $now -Kind Date)
    }

    [void]$insert.ExecuteNonQuery()
}

function Set-WebApprovalPreSyncFailure {
    param(
        $Record,
        [string]$ManagerUserName,
        [string]$SyncStatus,
        [string]$ErrorCode,
        [string]$ErrorMessage,
        [string]$NumberReason = "",
        [string]$BusyVoucherNo = "",
        [string]$BusyVoucherCode = "",
        $NumberResolvedAt = $null,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    $ctx = $null
    $tx = $null

    try {
        $ctx = Get-WebApprovalFiscalDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        $tx = $ctx.connection.BeginTransaction()
        $now = [datetime]::UtcNow

        $cmd = $ctx.connection.CreateCommand()
        $cmd.Transaction = $tx

        if ([int]$ctx.dbType -eq 1) {
            $cmd.CommandText = @"
UPDATE dbo.BusyCloudWebApproval
SET
    SyncStatus=@syncStatus,
    LastSyncError=@lastError,
    BusyVoucherNo=@busyVoucherNo,
    BusyVoucherCode=@busyVoucherCode,
    NumberChangeReason=@numberReason,
    NumberResolvedAt=@numberResolvedAt,
    UpdatedAt=@updatedAt,
    Version=Version+1
WHERE Id=@id
  AND ApprovalStatus='APPROVED'
  AND SyncStatus<>'SYNCED'
"@

            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@syncStatus" -Value $SyncStatus -Kind Text -Size 30)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@lastError" -Value $ErrorMessage -Kind LongText)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@busyVoucherNo" -Value $BusyVoucherNo -Kind Text -Size 100)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@busyVoucherCode" -Value $BusyVoucherCode -Kind Text -Size 50)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@numberReason" -Value $NumberReason -Kind LongText)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@numberResolvedAt" -Value $NumberResolvedAt -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@updatedAt" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@id" -Value ([string]$Record.Id) -Kind Text -Size 36)
        }
        else {
            $cmd.CommandText = @"
UPDATE [BusyCloudWebApproval]
SET
    [SyncStatus]=?,
    [LastSyncError]=?,
    [BusyVoucherNo]=?,
    [BusyVoucherCode]=?,
    [NumberChangeReason]=?,
    [NumberResolvedAt]=?,
    [UpdatedAt]=?,
    [Version]=[Version]+1
WHERE [Id]=?
  AND [ApprovalStatus]='APPROVED'
  AND [SyncStatus]<>'SYNCED'
"@

            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p1" -Value $SyncStatus -Kind Text -Size 30)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p2" -Value $ErrorMessage -Kind LongText)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p3" -Value $BusyVoucherNo -Kind Text -Size 100)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p4" -Value $BusyVoucherCode -Kind Text -Size 50)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p5" -Value $NumberReason -Kind LongText)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p6" -Value $NumberResolvedAt -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p7" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p8" -Value ([string]$Record.Id) -Kind Text -Size 36)
        }

        [void]$cmd.ExecuteNonQuery()

        $action = if ($SyncStatus -eq "REVIEW_REQUIRED") {
            "SYNC_REVIEW_REQUIRED"
        } else {
            "SYNC_FAILED"
        }

        Add-WebApprovalActionRow `
            -Context $ctx `
            -Transaction $tx `
            -WebApprovalId ([string]$Record.Id) `
            -Action $action `
            -ActionBy $ManagerUserName.Trim() `
            -Remarks $ErrorMessage `
            -MetadataJson (@{
                errorCode = $ErrorCode
                requestedVoucherNo = [string]$Record.RequestedVoucherNo
            } | ConvertTo-Json -Compress) `
            -ActionTime $now

        Set-WebApprovalSyncLedgerInternal `
            -Context $ctx `
            -Transaction $tx `
            -WebApprovalId ([string]$Record.Id) `
            -PayloadHash ([string]$Record.PayloadHash) `
            -State $SyncStatus `
            -RequestedVoucherNo ([string]$Record.RequestedVoucherNo) `
            -BusyVoucherNo $BusyVoucherNo `
            -BusyVoucherCode $BusyVoucherCode `
            -NumberChangeReason $NumberReason `
            -NumberResolvedAt $NumberResolvedAt `
            -Attempts ([int]$Record.SyncAttempts) `
            -LastError $ErrorMessage

        # Persist a manager-facing notification as well. This is important when
        # a failure/review state occurs and the manager later closes/logs out.
        $reference = "WA-" + ([string]$Record.Id).Substring(
            0,
            [Math]::Min(8, ([string]$Record.Id).Length)
        ).ToUpperInvariant()

        $notificationType = if ($SyncStatus -eq "REVIEW_REQUIRED") {
            "WEB_APPROVAL_SYNC_REVIEW_REQUIRED"
        } else {
            "WEB_APPROVAL_SYNC_FAILED"
        }

        $notificationTitle = if ($SyncStatus -eq "REVIEW_REQUIRED") {
            "Web Approval Needs Review"
        } else {
            "Web Approval Sync Failed"
        }

        Add-WebApprovalNotificationRow `
            -Context $ctx `
            -Transaction $tx `
            -RecipientUserName $ManagerUserName.Trim() `
            -WebApprovalId ([string]$Record.Id) `
            -NotificationType $notificationType `
            -Title $notificationTitle `
            -Message "$reference requires your attention. $ErrorMessage" `
            -CreatedAt $now

        $tx.Commit()
        $tx = $null

        return @{
            success = $false
            httpStatus = 409
            errorCode = $ErrorCode
            reviewRequired = ($SyncStatus -eq "REVIEW_REQUIRED")
            error = $ErrorMessage
        }
    }
    catch {
        if ($tx) { try { $tx.Rollback() } catch {} }

        return @{
            success = $false
            httpStatus = 500
            error = $_.Exception.Message
        }
    }
    finally {
        Close-WebApprovalDbContext -Context $ctx
    }
}

function Sync-WebApprovalVoucher {
    param(
        [string]$Id,
        [string]$ManagerUserName,
        [string]$ManualVoucherNo = "",
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    # -----------------------------------------------------------------
    # 1. Load immutable approved transaction and authorize this manager.
    # -----------------------------------------------------------------
    $ctx = $null
    $record = $null

    try {
        $ctx = Get-WebApprovalFiscalDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        $record = Get-WebApprovalRecordByIdInternal `
            -Context $ctx `
            -Id $Id
    }
    catch {
        return @{
            success = $false
            httpStatus = 500
            error = $_.Exception.Message
        }
    }
    finally {
        Close-WebApprovalDbContext -Context $ctx
    }

    if ($null -eq $record) {
        return @{
            success = $false
            httpStatus = 404
            error = "Web Approval transaction was not found."
        }
    }

    $access = Test-WebApprovalManagerCanActOnRecord `
        -ManagerUserName $ManagerUserName `
        -Record $record `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $access.success -or -not $access.allowed) {
        return $access
    }

    $approvalStatus = ([string]$record.ApprovalStatus).Trim().ToUpperInvariant()
    $syncStatus = ([string]$record.SyncStatus).Trim().ToUpperInvariant()

    if ($approvalStatus -ne "APPROVED") {
        return @{
            success = $false
            httpStatus = 409
            errorCode = "WEB_APPROVAL_NOT_APPROVED"
            error = "Only an APPROVED Web Approval transaction can be synchronized."
        }
    }

    if ($syncStatus -eq "SYNCED") {
        return @{
            success = $true
            alreadySynced = $true
            data = Convert-WebApprovalRecordToPublic -Record $record
            message = "This transaction is already synchronized to BUSY."
        }
    }

    if ($syncStatus -eq "REVIEW_REQUIRED") {
        return @{
            success = $false
            httpStatus = 409
            reviewRequired = $true
            errorCode = "WEB_APPROVAL_REVIEW_REQUIRED"
            error = "This transaction requires reconciliation before another BUSY create can be attempted."
        }
    }

    if ($syncStatus -notin @("READY", "FAILED")) {
        return @{
            success = $false
            httpStatus = 409
            errorCode = "WEB_APPROVAL_SYNC_BUSY"
            error = "This transaction is not currently available for synchronization."
        }
    }

    # -----------------------------------------------------------------
    # 2. Verify the exact approved payload before changing posting copy.
    # -----------------------------------------------------------------
    $currentHash = Get-WebApprovalSha256 -Text ([string]$record.PayloadJson)
    $storedHash = ([string]$record.PayloadHash).Trim().ToLowerInvariant()

    if ($currentHash -ne $storedHash) {
        return Set-WebApprovalPreSyncFailure `
            -Record $record `
            -ManagerUserName $ManagerUserName `
            -SyncStatus "REVIEW_REQUIRED" `
            -ErrorCode "WEB_APPROVAL_PAYLOAD_HASH_MISMATCH" `
            -ErrorMessage "The approved voucher payload no longer matches its stored SHA-256 hash. Synchronization is blocked." `
            -NumberReason "" `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }

    # -----------------------------------------------------------------
    # 3. Resolve CURRENT numbering immediately before COM creation.
    # -----------------------------------------------------------------
    $numbering = Resolve-WebApprovalPostingNumber `
        -Record $record `
        -ManualVoucherNo $ManualVoucherNo `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $numbering.success) {
        return Set-WebApprovalPreSyncFailure `
            -Record $record `
            -ManagerUserName $ManagerUserName `
            -SyncStatus "FAILED" `
            -ErrorCode ([string]$numbering.errorCode) `
            -ErrorMessage ([string]$numbering.error) `
            -NumberReason ([string]$numbering.error) `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }

    $candidateNo = ([string]$numbering.vchNo).Trim()
    $requestedNo = ([string]$record.RequestedVoucherNo).Trim()
    $series = ([string]$record.RequestedSeries).Trim()
    $vchType = [int]$record.VoucherType

    # -----------------------------------------------------------------
    # 4. Pre-check candidate through BUSY COM.
    # -----------------------------------------------------------------
    $existingCandidate = Find-WebApprovalBusyVoucherCom `
        -VchType $vchType `
        -VchNo $candidateNo `
        -SeriesName $series `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $existingCandidate.success) {
        return Set-WebApprovalPreSyncFailure `
            -Record $record `
            -ManagerUserName $ManagerUserName `
            -SyncStatus "FAILED" `
            -ErrorCode "BUSY_NUMBER_CHECK_FAILED" `
            -ErrorMessage ([string]$existingCandidate.error) `
            -NumberReason ([string]$existingCandidate.error) `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }

    if ($existingCandidate.found) {
        if (-not [bool]$numbering.isAuto) {
            $message = "Requested manual voucher number '$candidateNo' already exists in BUSY. Enter another manual voucher number and retry Sync."

            return Set-WebApprovalPreSyncFailure `
                -Record $record `
                -ManagerUserName $ManagerUserName `
                -SyncStatus "FAILED" `
                -ErrorCode "MANUAL_VOUCHER_NUMBER_EXISTS" `
                -ErrorMessage $message `
                -NumberReason $message `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode
        }

        $message = "The current automatic voucher candidate '$candidateNo' was already taken before synchronization. Retry Sync to resolve the latest BUSY number."

        return Set-WebApprovalPreSyncFailure `
            -Record $record `
            -ManagerUserName $ManagerUserName `
            -SyncStatus "FAILED" `
            -ErrorCode "AUTOMATIC_VOUCHER_NUMBER_RACE" `
            -ErrorMessage $message `
            -NumberReason $message `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }

    # -----------------------------------------------------------------
    # 5. Build a temporary posting copy. Approved PayloadJson stays intact.
    # -----------------------------------------------------------------
    $postingPayload = $null

    try {
        $postingPayload = Copy-WebApprovalPayloadForPosting `
            -PayloadJson ([string]$record.PayloadJson)

        $postingPayload | Add-Member `
            -MemberType NoteProperty `
            -Name "vchNo" `
            -Value $candidateNo `
            -Force

        $postingPayload | Add-Member `
            -MemberType NoteProperty `
            -Name "bridgeUserName" `
            -Value $ManagerUserName.Trim() `
            -Force

        $postingPayload = Update-WebApprovalSelfNumberReferences `
            -Payload $postingPayload `
            -RequestedVoucherNo $requestedNo `
            -FinalVoucherNo $candidateNo
    }
    catch {
        return Set-WebApprovalPreSyncFailure `
            -Record $record `
            -ManagerUserName $ManagerUserName `
            -SyncStatus "FAILED" `
            -ErrorCode "POSTING_PAYLOAD_BUILD_FAILED" `
            -ErrorMessage $_.Exception.Message `
            -NumberReason "" `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }

    # -----------------------------------------------------------------
    # 6. Atomically claim READY/FAILED -> SYNCING and create ledger state.
    # -----------------------------------------------------------------
    $claimCtx = $null
    $claimTx = $null
    $attemptNumber = ([int]$record.SyncAttempts) + 1

    try {
        $claimCtx = Get-WebApprovalFiscalDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        $claimTx = $claimCtx.connection.BeginTransaction()
        $now = [datetime]::UtcNow

        $claim = $claimCtx.connection.CreateCommand()
        $claim.Transaction = $claimTx

        if ([int]$claimCtx.dbType -eq 1) {
            $claim.CommandText = @"
UPDATE dbo.BusyCloudWebApproval
SET
    SyncStatus='SYNCING',
    SyncAttempts=SyncAttempts+1,
    LastSyncError=NULL,
    UpdatedAt=@updatedAt,
    Version=Version+1
WHERE Id=@id
  AND ApprovalStatus='APPROVED'
  AND SyncStatus IN ('READY','FAILED')
"@
            [void](Add-WebApprovalCommandParameter -Command $claim -DbType 1 -Name "@updatedAt" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $claim -DbType 1 -Name "@id" -Value $Id.Trim() -Kind Text -Size 36)
        }
        else {
            $claim.CommandText = @"
UPDATE [BusyCloudWebApproval]
SET
    [SyncStatus]='SYNCING',
    [SyncAttempts]=[SyncAttempts]+1,
    [LastSyncError]=NULL,
    [UpdatedAt]=?,
    [Version]=[Version]+1
WHERE [Id]=?
  AND [ApprovalStatus]='APPROVED'
  AND ([SyncStatus]='READY' OR [SyncStatus]='FAILED')
"@
            [void](Add-WebApprovalCommandParameter -Command $claim -DbType 0 -Name "@p1" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $claim -DbType 0 -Name "@p2" -Value $Id.Trim() -Kind Text -Size 36)
        }

        $affected = [int]$claim.ExecuteNonQuery()

        if ($affected -ne 1) {
            try { $claimTx.Rollback() } catch {}
            $claimTx = $null

            return @{
                success = $false
                httpStatus = 409
                errorCode = "WEB_APPROVAL_SYNC_CLAIM_CONFLICT"
                error = "Another synchronization is already running, or this transaction changed state. Refresh the page."
            }
        }

        Add-WebApprovalActionRow `
            -Context $claimCtx `
            -Transaction $claimTx `
            -WebApprovalId $Id.Trim() `
            -Action "SYNC_STARTED" `
            -ActionBy $ManagerUserName.Trim() `
            -Remarks "BUSY synchronization started." `
            -MetadataJson (@{
                requestedVoucherNo = $requestedNo
                resolvedCandidate = $candidateNo
                automaticNumbering = [bool]$numbering.isAuto
            } | ConvertTo-Json -Compress) `
            -ActionTime $now

        Set-WebApprovalSyncLedgerInternal `
            -Context $claimCtx `
            -Transaction $claimTx `
            -WebApprovalId $Id.Trim() `
            -PayloadHash $storedHash `
            -State "SYNCING" `
            -RequestedVoucherNo $requestedNo `
            -BusyVoucherNo "" `
            -BusyVoucherCode "" `
            -NumberChangeReason "" `
            -NumberResolvedAt $null `
            -Attempts $attemptNumber `
            -LastError ""

        $claimTx.Commit()
        $claimTx = $null
    }
    catch {
        if ($claimTx) { try { $claimTx.Rollback() } catch {} }

        return @{
            success = $false
            httpStatus = 500
            errorCode = "WEB_APPROVAL_SYNC_CLAIM_FAILED"
            error = $_.Exception.Message
        }
    }
    finally {
        Close-WebApprovalDbContext -Context $claimCtx
    }

    # -----------------------------------------------------------------
    # 7. Dangerous boundary: BUSY COM Create-Voucher.
    #    Once this call begins, any uncertain failure becomes REVIEW_REQUIRED.
    # -----------------------------------------------------------------
    $createResult = $null

    try {
        $createResult = Create-Voucher `
            -Data $postingPayload `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }
    catch {
        $message = "BUSY Create-Voucher threw an exception after synchronization started. Candidate number: '$candidateNo'. The outcome is uncertain and automatic retry is blocked. $($_.Exception.Message)"

        # Reload latest attempt count for audit helper.
        $record.SyncAttempts = $attemptNumber

        return Set-WebApprovalPreSyncFailure `
            -Record $record `
            -ManagerUserName $ManagerUserName `
            -SyncStatus "REVIEW_REQUIRED" `
            -ErrorCode "BUSY_CREATE_OUTCOME_AMBIGUOUS" `
            -ErrorMessage $message `
            -NumberReason $message `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }

    $resultInfo = Get-WebApprovalCreateResultInfo -Result $createResult

    if (-not $resultInfo.success) {
        $message = if (-not [string]::IsNullOrWhiteSpace([string]$resultInfo.error)) {
            "BUSY did not confirm a successful create after synchronization started. Candidate number: '$candidateNo'. Automatic retry is blocked until review. $($resultInfo.error)"
        } else {
            "BUSY did not confirm a successful create after synchronization started. Candidate number: '$candidateNo'. Automatic retry is blocked until review."
        }

        $record.SyncAttempts = $attemptNumber

        return Set-WebApprovalPreSyncFailure `
            -Record $record `
            -ManagerUserName $ManagerUserName `
            -SyncStatus "REVIEW_REQUIRED" `
            -ErrorCode "BUSY_CREATE_NOT_CONFIRMED" `
            -ErrorMessage $message `
            -NumberReason $message `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }

    # -----------------------------------------------------------------
    # 8. Confirm the created BUSY header through COM.
    # -----------------------------------------------------------------
    $candidateFromResult = ([string]$resultInfo.vchNo).Trim()
    if ([string]::IsNullOrWhiteSpace($candidateFromResult)) {
        $candidateFromResult = $candidateNo
    }

    $confirmed = Find-WebApprovalBusyVoucherCom `
        -VchType $vchType `
        -VchNo $candidateFromResult `
        -SeriesName $series `
        -VchCode ([int]$resultInfo.vchCode) `
        -InstanceId $InstanceId `
        -CompanyCode $CompanyCode

    if (-not $confirmed.success -or -not $confirmed.found) {
        $message = "Create-Voucher reported success, but the created BUSY voucher could not be confirmed. Candidate number: '$candidateFromResult'. Automatic retry is blocked."

        if (-not $confirmed.success -and $confirmed.error) {
            $message += " $($confirmed.error)"
        }

        $record.SyncAttempts = $attemptNumber

        return Set-WebApprovalPreSyncFailure `
            -Record $record `
            -ManagerUserName $ManagerUserName `
            -SyncStatus "REVIEW_REQUIRED" `
            -ErrorCode "BUSY_POST_CREATE_CONFIRMATION_FAILED" `
            -ErrorMessage $message `
            -NumberReason $message `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }

    $finalNo = ([string]$confirmed.vchNo).Trim()
    $finalCode = 0
    try { $finalCode = [int]$confirmed.vchCode } catch {}

    $numberReason = ""

    if (
        -not $requestedNo.Equals(
            $finalNo,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        if (-not [bool]$numbering.isAuto -and -not [string]::IsNullOrWhiteSpace($ManualVoucherNo)) {
            $numberReason = "Requested manual voucher number '$requestedNo' was changed by the Sales Manager to '$finalNo' before BUSY synchronization."
        }
        else {
            $numberReason = Get-WebApprovalNumberChangeReason `
                -RequestedVoucherNo $requestedNo `
                -BusyVoucherNo $finalNo
        }
    }

    # -----------------------------------------------------------------
    # 9. Persist confirmed final number and SYNCED state.
    # -----------------------------------------------------------------
    $finishCtx = $null
    $finishTx = $null

    try {
        $finishCtx = Get-WebApprovalFiscalDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        $finishTx = $finishCtx.connection.BeginTransaction()
        $now = [datetime]::UtcNow

        $finish = $finishCtx.connection.CreateCommand()
        $finish.Transaction = $finishTx

        if ([int]$finishCtx.dbType -eq 1) {
            $finish.CommandText = @"
UPDATE dbo.BusyCloudWebApproval
SET
    SyncStatus='SYNCED',
    BusyVoucherNo=@busyVoucherNo,
    BusyVoucherCode=@busyVoucherCode,
    NumberChangeReason=@numberChangeReason,
    NumberResolvedAt=@numberResolvedAt,
    LastSyncError=NULL,
    SyncedAt=@syncedAt,
    UpdatedAt=@updatedAt,
    Version=Version+1
WHERE Id=@id
  AND ApprovalStatus='APPROVED'
  AND SyncStatus='SYNCING'
"@

            [void](Add-WebApprovalCommandParameter -Command $finish -DbType 1 -Name "@busyVoucherNo" -Value $finalNo -Kind Text -Size 100)
            [void](Add-WebApprovalCommandParameter -Command $finish -DbType 1 -Name "@busyVoucherCode" -Value ([string]$finalCode) -Kind Text -Size 50)
            [void](Add-WebApprovalCommandParameter -Command $finish -DbType 1 -Name "@numberChangeReason" -Value $numberReason -Kind LongText)
            [void](Add-WebApprovalCommandParameter -Command $finish -DbType 1 -Name "@numberResolvedAt" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $finish -DbType 1 -Name "@syncedAt" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $finish -DbType 1 -Name "@updatedAt" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $finish -DbType 1 -Name "@id" -Value $Id.Trim() -Kind Text -Size 36)
        }
        else {
            $finish.CommandText = @"
UPDATE [BusyCloudWebApproval]
SET
    [SyncStatus]='SYNCED',
    [BusyVoucherNo]=?,
    [BusyVoucherCode]=?,
    [NumberChangeReason]=?,
    [NumberResolvedAt]=?,
    [LastSyncError]=NULL,
    [SyncedAt]=?,
    [UpdatedAt]=?,
    [Version]=[Version]+1
WHERE [Id]=?
  AND [ApprovalStatus]='APPROVED'
  AND [SyncStatus]='SYNCING'
"@

            [void](Add-WebApprovalCommandParameter -Command $finish -DbType 0 -Name "@p1" -Value $finalNo -Kind Text -Size 100)
            [void](Add-WebApprovalCommandParameter -Command $finish -DbType 0 -Name "@p2" -Value ([string]$finalCode) -Kind Text -Size 50)
            [void](Add-WebApprovalCommandParameter -Command $finish -DbType 0 -Name "@p3" -Value $numberReason -Kind LongText)
            [void](Add-WebApprovalCommandParameter -Command $finish -DbType 0 -Name "@p4" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $finish -DbType 0 -Name "@p5" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $finish -DbType 0 -Name "@p6" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $finish -DbType 0 -Name "@p7" -Value $Id.Trim() -Kind Text -Size 36)
        }

        $affected = [int]$finish.ExecuteNonQuery()

        if ($affected -ne 1) {
            throw "The Web Approval row could not be moved from SYNCING to SYNCED."
        }

        Set-WebApprovalSyncLedgerInternal `
            -Context $finishCtx `
            -Transaction $finishTx `
            -WebApprovalId $Id.Trim() `
            -PayloadHash $storedHash `
            -State "SYNCED" `
            -RequestedVoucherNo $requestedNo `
            -BusyVoucherNo $finalNo `
            -BusyVoucherCode ([string]$finalCode) `
            -NumberChangeReason $numberReason `
            -NumberResolvedAt $now `
            -Attempts $attemptNumber `
            -LastError ""

        Add-WebApprovalActionRow `
            -Context $finishCtx `
            -Transaction $finishTx `
            -WebApprovalId $Id.Trim() `
            -Action "SYNC_SUCCEEDED" `
            -ActionBy $ManagerUserName.Trim() `
            -Remarks "Voucher synchronized to BUSY as '$finalNo'." `
            -MetadataJson (@{
                requestedVoucherNo = $requestedNo
                busyVoucherNo = $finalNo
                busyVoucherCode = $finalCode
                numberChangeReason = $numberReason
            } | ConvertTo-Json -Compress) `
            -ActionTime $now

        $reference = "WA-" + $Id.Trim().Substring(0, [Math]::Min(8, $Id.Trim().Length)).ToUpperInvariant()

        Add-WebApprovalNotificationRow `
            -Context $finishCtx `
            -Transaction $finishTx `
            -RecipientUserName ([string]$record.SubmittedBy).Trim() `
            -WebApprovalId $Id.Trim() `
            -NotificationType "WEB_APPROVAL_SYNCED" `
            -Title "Web Approval Synchronized" `
            -Message "$reference was synchronized to BUSY as voucher '$finalNo'." `
            -CreatedAt $now

        $finishTx.Commit()
        $finishTx = $null

        return @{
            success = $true
            message = "Voucher synchronized to BUSY successfully."
            data = @{
                id = $Id.Trim()
                reference = $reference
                requestedVoucherNo = $requestedNo
                busyVoucherNo = $finalNo
                busyVoucherCode = if ($finalCode -gt 0) { $finalCode } else { $null }
                numberChangeReason = $numberReason
                approvalStatus = "APPROVED"
                syncStatus = "SYNCED"
                syncedAt = $now
            }
        }
    }
    catch {
        if ($finishTx) {
            try { $finishTx.Rollback() } catch {}
        }

        # BUSY was already confirmed created. A DB bookkeeping failure after that
        # must never trigger another automatic BUSY create.
        $record.SyncAttempts = $attemptNumber

        return Set-WebApprovalPreSyncFailure `
            -Record $record `
            -ManagerUserName $ManagerUserName `
            -SyncStatus "REVIEW_REQUIRED" `
            -ErrorCode "POST_CREATE_BOOKKEEPING_FAILED" `
            -ErrorMessage "BUSY voucher '$finalNo' was confirmed created, but Web Approval bookkeeping failed. Do not retry Create-Voucher. $($_.Exception.Message)" `
            -NumberReason $numberReason `
            -BusyVoucherNo $finalNo `
            -BusyVoucherCode ([string]$finalCode) `
            -NumberResolvedAt ([datetime]::UtcNow) `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    }
    finally {
        Close-WebApprovalDbContext -Context $finishCtx
    }
}



# ============================================================================
# PHASE 6 - PERSISTENT IN-APP + WEB PUSH NOTIFICATIONS
# ============================================================================

function Get-WebPushConfigPath {
    $backendRoot = Split-Path -Parent $PSScriptRoot
    return Join-Path $backendRoot "data\push\push_config.json"
}

function Get-WebPushConfigInternal {
    $path = Get-WebPushConfigPath

    if (-not (Test-Path -LiteralPath $path)) {
        return @{
            configured = $false
            path = $path
        }
    }

    try {
        $raw = Get-Content -LiteralPath $path -Raw
        $cfg = $raw | ConvertFrom-Json

        $publicKey = ([string]$cfg.publicKey).Trim()
        $privateKey = ([string]$cfg.privateKey).Trim()
        $subject = ([string]$cfg.subject).Trim()
        $workerSecret = ([string]$cfg.workerSecret).Trim()

        return @{
            configured = (
                -not [string]::IsNullOrWhiteSpace($publicKey) -and
                -not [string]::IsNullOrWhiteSpace($privateKey) -and
                -not [string]::IsNullOrWhiteSpace($subject) -and
                -not [string]::IsNullOrWhiteSpace($workerSecret)
            )
            publicKey = $publicKey
            privateKey = $privateKey
            subject = $subject
            workerSecret = $workerSecret
            path = $path
        }
    }
    catch {
        return @{
            configured = $false
            error = $_.Exception.Message
            path = $path
        }
    }
}

function Get-WebPushPublicConfig {
    $cfg = Get-WebPushConfigInternal

    return @{
        success = $true
        data = @{
            configured = [bool]$cfg.configured
            publicKey = if ($cfg.configured) { [string]$cfg.publicKey } else { "" }
        }
    }
}

function Test-WebPushWorkerAuthorized {
    param($Request)

    if ($null -eq $Request) {
        return $false
    }

    $isLoopback = $false

    try {
        $isLoopback = [System.Net.IPAddress]::IsLoopback(
            $Request.RemoteEndPoint.Address
        )
    }
    catch {}

    if (-not $isLoopback) {
        return $false
    }

    $cfg = Get-WebPushConfigInternal

    if (-not $cfg.configured) {
        return $false
    }

    $provided = ([string]$Request.Headers["X-BusyCloud-Push-Worker"]).Trim()

    if ([string]::IsNullOrWhiteSpace($provided)) {
        return $false
    }

    return $provided -eq ([string]$cfg.workerSecret)
}

function Save-WebPushSubscription {
    param(
        [string]$UserName,
        $Data,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ([string]::IsNullOrWhiteSpace($UserName)) {
        return @{
            success = $false
            httpStatus = 401
            error = "Authenticated BUSY user is required."
        }
    }

    $endpoint = ([string](
        Get-WebApprovalPropertyValue `
            -Object $Data `
            -Names @("endpoint") `
            -DefaultValue ""
    )).Trim()

    $keys = Get-WebApprovalPropertyValue `
        -Object $Data `
        -Names @("keys") `
        -DefaultValue $null

    $p256dh = ([string](
        Get-WebApprovalPropertyValue `
            -Object $keys `
            -Names @("p256dh") `
            -DefaultValue ""
    )).Trim()

    $authKey = ([string](
        Get-WebApprovalPropertyValue `
            -Object $keys `
            -Names @("auth") `
            -DefaultValue ""
    )).Trim()

    if (
        [string]::IsNullOrWhiteSpace($endpoint) -or
        [string]::IsNullOrWhiteSpace($p256dh) -or
        [string]::IsNullOrWhiteSpace($authKey)
    ) {
        return @{
            success = $false
            httpStatus = 400
            error = "A valid browser push subscription is required."
        }
    }

    $deviceName = ([string](
        Get-WebApprovalPropertyValue `
            -Object $Data `
            -Names @("deviceName") `
            -DefaultValue ""
    )).Trim()

    if ($deviceName.Length -gt 150) {
        $deviceName = $deviceName.Substring(0, 150)
    }

    $userAgent = ([string](
        Get-WebApprovalPropertyValue `
            -Object $Data `
            -Names @("userAgent") `
            -DefaultValue ""
    )).Trim()

    if ($userAgent.Length -gt 500) {
        $userAgent = $userAgent.Substring(0, 500)
    }

    $hash = Get-WebApprovalSha256 -Text $endpoint
    $ctx = $null

    try {
        $ctx = Get-WebApprovalPermanentDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        [void](Ensure-WebApprovalPermanentTables -Context $ctx)

        $existingId = ""
        $cmd = $ctx.connection.CreateCommand()

        if ([int]$ctx.dbType -eq 1) {
            $cmd.CommandText = @"
SELECT TOP 1 Id
FROM dbo.BusyCloudPushSubscription
WHERE EndpointHash=@hash
"@
            [void](Add-WebApprovalCommandParameter `
                -Command $cmd -DbType 1 -Name "@hash" `
                -Value $hash -Kind Text -Size 64)
        }
        else {
            $cmd.CommandText = @"
SELECT TOP 1 [Id]
FROM [BusyCloudPushSubscription]
WHERE [EndpointHash]=?
"@
            [void](Add-WebApprovalCommandParameter `
                -Command $cmd -DbType 0 -Name "@p1" `
                -Value $hash -Kind Text -Size 64)
        }

        $rawId = $cmd.ExecuteScalar()
        if ($null -ne $rawId -and $rawId -ne [System.DBNull]::Value) {
            $existingId = ([string]$rawId).Trim()
        }

        $now = [datetime]::UtcNow

        if ($existingId) {
            $update = $ctx.connection.CreateCommand()

            if ([int]$ctx.dbType -eq 1) {
                $update.CommandText = @"
UPDATE dbo.BusyCloudPushSubscription
SET
    UserName=@userName,
    Endpoint=@endpoint,
    P256dh=@p256dh,
    AuthKey=@authKey,
    DeviceName=@deviceName,
    UserAgent=@userAgent,
    IsActive=1,
    UpdatedAt=@updatedAt,
    LastSeenAt=@lastSeenAt
WHERE Id=@id
"@
                [void](Add-WebApprovalCommandParameter -Command $update -DbType 1 -Name "@userName" -Value $UserName.Trim() -Kind Text -Size 100)
                [void](Add-WebApprovalCommandParameter -Command $update -DbType 1 -Name "@endpoint" -Value $endpoint -Kind LongText)
                [void](Add-WebApprovalCommandParameter -Command $update -DbType 1 -Name "@p256dh" -Value $p256dh -Kind Text -Size 255)
                [void](Add-WebApprovalCommandParameter -Command $update -DbType 1 -Name "@authKey" -Value $authKey -Kind Text -Size 255)
                [void](Add-WebApprovalCommandParameter -Command $update -DbType 1 -Name "@deviceName" -Value $deviceName -Kind Text -Size 150)
                [void](Add-WebApprovalCommandParameter -Command $update -DbType 1 -Name "@userAgent" -Value $userAgent -Kind Text -Size 500)
                [void](Add-WebApprovalCommandParameter -Command $update -DbType 1 -Name "@updatedAt" -Value $now -Kind Date)
                [void](Add-WebApprovalCommandParameter -Command $update -DbType 1 -Name "@lastSeenAt" -Value $now -Kind Date)
                [void](Add-WebApprovalCommandParameter -Command $update -DbType 1 -Name "@id" -Value $existingId -Kind Text -Size 36)
            }
            else {
                $update.CommandText = @"
UPDATE [BusyCloudPushSubscription]
SET
    [UserName]=?,
    [Endpoint]=?,
    [P256dh]=?,
    [AuthKey]=?,
    [DeviceName]=?,
    [UserAgent]=?,
    [IsActive]=1,
    [UpdatedAt]=?,
    [LastSeenAt]=?
WHERE [Id]=?
"@
                [void](Add-WebApprovalCommandParameter -Command $update -DbType 0 -Name "@p1" -Value $UserName.Trim() -Kind Text -Size 100)
                [void](Add-WebApprovalCommandParameter -Command $update -DbType 0 -Name "@p2" -Value $endpoint -Kind LongText)
                [void](Add-WebApprovalCommandParameter -Command $update -DbType 0 -Name "@p3" -Value $p256dh -Kind Text -Size 255)
                [void](Add-WebApprovalCommandParameter -Command $update -DbType 0 -Name "@p4" -Value $authKey -Kind Text -Size 255)
                [void](Add-WebApprovalCommandParameter -Command $update -DbType 0 -Name "@p5" -Value $deviceName -Kind Text -Size 150)
                [void](Add-WebApprovalCommandParameter -Command $update -DbType 0 -Name "@p6" -Value $userAgent -Kind LongText)
                [void](Add-WebApprovalCommandParameter -Command $update -DbType 0 -Name "@p7" -Value $now -Kind Date)
                [void](Add-WebApprovalCommandParameter -Command $update -DbType 0 -Name "@p8" -Value $now -Kind Date)
                [void](Add-WebApprovalCommandParameter -Command $update -DbType 0 -Name "@p9" -Value $existingId -Kind Text -Size 36)
            }

            [void]$update.ExecuteNonQuery()

            return @{
                success = $true
                data = @{
                    id = $existingId
                    enabled = $true
                }
            }
        }

        $id = [guid]::NewGuid().ToString()
        $insert = $ctx.connection.CreateCommand()

        if ([int]$ctx.dbType -eq 1) {
            $insert.CommandText = @"
INSERT INTO dbo.BusyCloudPushSubscription
(
    Id, UserName, EndpointHash, Endpoint,
    P256dh, AuthKey, DeviceName, UserAgent,
    IsActive, CreatedAt, UpdatedAt, LastSeenAt
)
VALUES
(
    @id, @userName, @endpointHash, @endpoint,
    @p256dh, @authKey, @deviceName, @userAgent,
    1, @createdAt, @updatedAt, @lastSeenAt
)
"@
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@id" -Value $id -Kind Text -Size 36)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@userName" -Value $UserName.Trim() -Kind Text -Size 100)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@endpointHash" -Value $hash -Kind Text -Size 64)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@endpoint" -Value $endpoint -Kind LongText)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@p256dh" -Value $p256dh -Kind Text -Size 255)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@authKey" -Value $authKey -Kind Text -Size 255)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@deviceName" -Value $deviceName -Kind Text -Size 150)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@userAgent" -Value $userAgent -Kind Text -Size 500)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@createdAt" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@updatedAt" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@lastSeenAt" -Value $now -Kind Date)
        }
        else {
            $insert.CommandText = @"
INSERT INTO [BusyCloudPushSubscription]
(
    [Id], [UserName], [EndpointHash], [Endpoint],
    [P256dh], [AuthKey], [DeviceName], [UserAgent],
    [IsActive], [CreatedAt], [UpdatedAt], [LastSeenAt]
)
VALUES
(
    ?,?,?,?,?,?,?,?,1,?,?,?
)
"@
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p1" -Value $id -Kind Text -Size 36)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p2" -Value $UserName.Trim() -Kind Text -Size 100)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p3" -Value $hash -Kind Text -Size 64)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p4" -Value $endpoint -Kind LongText)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p5" -Value $p256dh -Kind Text -Size 255)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p6" -Value $authKey -Kind Text -Size 255)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p7" -Value $deviceName -Kind Text -Size 150)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p8" -Value $userAgent -Kind LongText)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p9" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p10" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p11" -Value $now -Kind Date)
        }

        [void]$insert.ExecuteNonQuery()

        return @{
            success = $true
            data = @{
                id = $id
                enabled = $true
            }
        }
    }
    catch {
        return @{
            success = $false
            httpStatus = 500
            error = $_.Exception.Message
        }
    }
    finally {
        Close-WebApprovalDbContext -Context $ctx
    }
}

function Disable-WebPushSubscription {
    param(
        [string]$UserName,
        [string]$Endpoint,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ([string]::IsNullOrWhiteSpace($UserName)) {
        return @{
            success = $false
            httpStatus = 401
            error = "Authenticated BUSY user is required."
        }
    }

    $endpointText = ([string]$Endpoint).Trim()

    if ([string]::IsNullOrWhiteSpace($endpointText)) {
        return @{
            success = $false
            httpStatus = 400
            error = "Push endpoint is required."
        }
    }

    $hash = Get-WebApprovalSha256 -Text $endpointText
    $ctx = $null

    try {
        $ctx = Get-WebApprovalPermanentDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        [void](Ensure-WebApprovalPermanentTables -Context $ctx)

        $cmd = $ctx.connection.CreateCommand()
        $now = [datetime]::UtcNow

        if ([int]$ctx.dbType -eq 1) {
            $cmd.CommandText = @"
UPDATE dbo.BusyCloudPushSubscription
SET IsActive=0, UpdatedAt=@updatedAt
WHERE EndpointHash=@hash
  AND UserName=@userName
"@
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@updatedAt" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@hash" -Value $hash -Kind Text -Size 64)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@userName" -Value $UserName.Trim() -Kind Text -Size 100)
        }
        else {
            $cmd.CommandText = @"
UPDATE [BusyCloudPushSubscription]
SET [IsActive]=0, [UpdatedAt]=?
WHERE [EndpointHash]=?
  AND [UserName]=?
"@
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p1" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p2" -Value $hash -Kind Text -Size 64)
            [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p3" -Value $UserName.Trim() -Kind Text -Size 100)
        }

        [void]$cmd.ExecuteNonQuery()

        return @{
            success = $true
            data = @{ enabled = $false }
        }
    }
    catch {
        return @{
            success = $false
            httpStatus = 500
            error = $_.Exception.Message
        }
    }
    finally {
        Close-WebApprovalDbContext -Context $ctx
    }
}

function Get-WebPushSubscriptionsInternal {
    param(
        [string]$InstanceId,
        [string]$CompanyCode
    )

    $ctx = $null
    $rdr = $null
    $items = @()

    try {
        $ctx = Get-WebApprovalPermanentDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        [void](Ensure-WebApprovalPermanentTables -Context $ctx)

        $cmd = $ctx.connection.CreateCommand()

        if ([int]$ctx.dbType -eq 1) {
            $cmd.CommandText = @"
SELECT
    Id, UserName, EndpointHash, Endpoint,
    P256dh, AuthKey, DeviceName, UserAgent
FROM dbo.BusyCloudPushSubscription
WHERE IsActive<>0
"@
        }
        else {
            $cmd.CommandText = @"
SELECT
    [Id], [UserName], [EndpointHash], [Endpoint],
    [P256dh], [AuthKey], [DeviceName], [UserAgent]
FROM [BusyCloudPushSubscription]
WHERE [IsActive]<>0
"@
        }

        $rdr = $cmd.ExecuteReader()

        while ($rdr.Read()) {
            $items += @{
                id = [string](Read-WebApprovalReaderValue $rdr "Id" "")
                userName = [string](Read-WebApprovalReaderValue $rdr "UserName" "")
                endpointHash = [string](Read-WebApprovalReaderValue $rdr "EndpointHash" "")
                endpoint = [string](Read-WebApprovalReaderValue $rdr "Endpoint" "")
                p256dh = [string](Read-WebApprovalReaderValue $rdr "P256dh" "")
                authKey = [string](Read-WebApprovalReaderValue $rdr "AuthKey" "")
                deviceName = [string](Read-WebApprovalReaderValue $rdr "DeviceName" "")
                userAgent = [string](Read-WebApprovalReaderValue $rdr "UserAgent" "")
            }
        }
    }
    finally {
        if ($rdr) {
            try { $rdr.Close() } catch {}
            try { $rdr.Dispose() } catch {}
        }
        Close-WebApprovalDbContext -Context $ctx
    }

    return @($items)
}

function Set-WebApprovalNotificationRead {
    param(
        [string]$UserName,
        [string]$NotificationId = "",
        [bool]$ReadAll = $false,
        [string]$InstanceId = "",
        [string]$CompanyCode = ""
    )

    if ([string]::IsNullOrWhiteSpace($UserName)) {
        return @{
            success = $false
            httpStatus = 401
            error = "Authenticated BUSY user is required."
        }
    }

    if (-not $ReadAll -and [string]::IsNullOrWhiteSpace($NotificationId)) {
        return @{
            success = $false
            httpStatus = 400
            error = "notificationId is required unless readAll=true."
        }
    }

    $ctx = $null
    $tx = $null

    try {
        $ctx = Get-WebApprovalFiscalDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        [void](Ensure-WebApprovalFiscalTables -Context $ctx)

        $tx = $ctx.connection.BeginTransaction()
        $now = [datetime]::UtcNow

        $cmd = $ctx.connection.CreateCommand()
        $cmd.Transaction = $tx

        if ([int]$ctx.dbType -eq 1) {
            if ($ReadAll) {
                $cmd.CommandText = @"
UPDATE dbo.BusyCloudNotification
SET IsRead=1, ReadAt=@readAt
WHERE RecipientUserName=@userName
  AND IsRead=0
"@
                [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@readAt" -Value $now -Kind Date)
                [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@userName" -Value $UserName.Trim() -Kind Text -Size 100)
            }
            else {
                $cmd.CommandText = @"
UPDATE dbo.BusyCloudNotification
SET IsRead=1, ReadAt=@readAt
WHERE Id=@id
  AND RecipientUserName=@userName
"@
                [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@readAt" -Value $now -Kind Date)
                [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@id" -Value $NotificationId.Trim() -Kind Text -Size 36)
                [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@userName" -Value $UserName.Trim() -Kind Text -Size 100)
            }
        }
        else {
            if ($ReadAll) {
                $cmd.CommandText = @"
UPDATE [BusyCloudNotification]
SET [IsRead]=1, [ReadAt]=?
WHERE [RecipientUserName]=?
  AND [IsRead]=0
"@
                [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p1" -Value $now -Kind Date)
                [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p2" -Value $UserName.Trim() -Kind Text -Size 100)
            }
            else {
                $cmd.CommandText = @"
UPDATE [BusyCloudNotification]
SET [IsRead]=1, [ReadAt]=?
WHERE [Id]=?
  AND [RecipientUserName]=?
"@
                [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p1" -Value $now -Kind Date)
                [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p2" -Value $NotificationId.Trim() -Kind Text -Size 36)
                [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p3" -Value $UserName.Trim() -Kind Text -Size 100)
            }
        }

        [void]$cmd.ExecuteNonQuery()

        # If the user has already read the notification inside the app,
        # pending/retry push deliveries are no longer useful.
        $cancel = $ctx.connection.CreateCommand()
        $cancel.Transaction = $tx

        if ([int]$ctx.dbType -eq 1) {
            if ($ReadAll) {
                $cancel.CommandText = @"
UPDATE d
SET d.State='CANCELLED', d.UpdatedAt=@updatedAt
FROM dbo.BusyCloudNotificationDelivery d
INNER JOIN dbo.BusyCloudNotification n
    ON n.Id=d.NotificationId
WHERE n.RecipientUserName=@userName
  AND d.State IN ('PENDING','FAILED')
"@
                [void](Add-WebApprovalCommandParameter -Command $cancel -DbType 1 -Name "@updatedAt" -Value $now -Kind Date)
                [void](Add-WebApprovalCommandParameter -Command $cancel -DbType 1 -Name "@userName" -Value $UserName.Trim() -Kind Text -Size 100)
            }
            else {
                $cancel.CommandText = @"
UPDATE dbo.BusyCloudNotificationDelivery
SET State='CANCELLED', UpdatedAt=@updatedAt
WHERE NotificationId=@id
  AND State IN ('PENDING','FAILED')
"@
                [void](Add-WebApprovalCommandParameter -Command $cancel -DbType 1 -Name "@updatedAt" -Value $now -Kind Date)
                [void](Add-WebApprovalCommandParameter -Command $cancel -DbType 1 -Name "@id" -Value $NotificationId.Trim() -Kind Text -Size 36)
            }
        }
        else {
            if ($ReadAll) {
                # Access does not support UPDATE JOIN consistently through all
                # BUSY providers. Read-all cancellation is optional; unread
                # notifications are excluded from future job creation anyway.
                $cancel = $null
            }
            else {
                $cancel.CommandText = @"
UPDATE [BusyCloudNotificationDelivery]
SET [State]='CANCELLED', [UpdatedAt]=?
WHERE [NotificationId]=?
  AND ([State]='PENDING' OR [State]='FAILED')
"@
                [void](Add-WebApprovalCommandParameter -Command $cancel -DbType 0 -Name "@p1" -Value $now -Kind Date)
                [void](Add-WebApprovalCommandParameter -Command $cancel -DbType 0 -Name "@p2" -Value $NotificationId.Trim() -Kind Text -Size 36)
            }
        }

        if ($cancel) {
            [void]$cancel.ExecuteNonQuery()
        }

        $tx.Commit()
        $tx = $null

        return @{ success=$true }
    }
    catch {
        if ($tx) {
            try { $tx.Rollback() } catch {}
        }

        return @{
            success = $false
            httpStatus = 500
            error = $_.Exception.Message
        }
    }
    finally {
        Close-WebApprovalDbContext -Context $ctx
    }
}

function Get-WebPushUnreadNotificationsInternal {
    param(
        $Context,
        [int]$Limit = 100
    )

    $limitSafe = [Math]::Max(1, [Math]::Min(500, $Limit))
    $cmd = $Context.connection.CreateCommand()

    if ([int]$Context.dbType -eq 1) {
        $cmd.CommandText = @"
SELECT TOP $limitSafe
    Id, RecipientUserName, WebApprovalId,
    NotificationType, Title, Message, CreatedAt
FROM dbo.BusyCloudNotification
WHERE IsRead=0
ORDER BY CreatedAt ASC
"@
    }
    else {
        $cmd.CommandText = @"
SELECT TOP $limitSafe
    [Id], [RecipientUserName], [WebApprovalId],
    [NotificationType], [Title], [Message], [CreatedAt]
FROM [BusyCloudNotification]
WHERE [IsRead]=0
ORDER BY [CreatedAt] ASC
"@
    }

    $rdr = $null
    $items = @()

    try {
        $rdr = $cmd.ExecuteReader()

        while ($rdr.Read()) {
            $items += @{
                id = [string](Read-WebApprovalReaderValue $rdr "Id" "")
                recipientUserName = [string](Read-WebApprovalReaderValue $rdr "RecipientUserName" "")
                webApprovalId = [string](Read-WebApprovalReaderValue $rdr "WebApprovalId" "")
                type = [string](Read-WebApprovalReaderValue $rdr "NotificationType" "")
                title = [string](Read-WebApprovalReaderValue $rdr "Title" "")
                message = [string](Read-WebApprovalReaderValue $rdr "Message" "")
                createdAt = Read-WebApprovalReaderValue $rdr "CreatedAt" $null
            }
        }
    }
    finally {
        if ($rdr) {
            try { $rdr.Close() } catch {}
            try { $rdr.Dispose() } catch {}
        }
    }

    return @($items)
}

function Get-OrCreate-WebPushDeliveryInternal {
    param(
        $Context,
        $Notification,
        $Subscription
    )

    $cmd = $Context.connection.CreateCommand()

    if ([int]$Context.dbType -eq 1) {
        $cmd.CommandText = @"
SELECT TOP 1
    Id, State, Attempts, NextAttemptAt
FROM dbo.BusyCloudNotificationDelivery
WHERE NotificationId=@notificationId
  AND SubscriptionId=@subscriptionId
"@
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@notificationId" -Value ([string]$Notification.id) -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 1 -Name "@subscriptionId" -Value ([string]$Subscription.id) -Kind Text -Size 36)
    }
    else {
        $cmd.CommandText = @"
SELECT TOP 1
    [Id], [State], [Attempts], [NextAttemptAt]
FROM [BusyCloudNotificationDelivery]
WHERE [NotificationId]=?
  AND [SubscriptionId]=?
"@
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p1" -Value ([string]$Notification.id) -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $cmd -DbType 0 -Name "@p2" -Value ([string]$Subscription.id) -Kind Text -Size 36)
    }

    $rdr = $null

    try {
        $rdr = $cmd.ExecuteReader()

        if ($rdr.Read()) {
            return @{
                id = [string](Read-WebApprovalReaderValue $rdr "Id" "")
                state = [string](Read-WebApprovalReaderValue $rdr "State" "")
                attempts = [int](Read-WebApprovalReaderValue $rdr "Attempts" 0)
                nextAttemptAt = Read-WebApprovalReaderValue $rdr "NextAttemptAt" $null
            }
        }
    }
    finally {
        if ($rdr) {
            try { $rdr.Close() } catch {}
            try { $rdr.Dispose() } catch {}
        }
    }

    $id = [guid]::NewGuid().ToString()
    $now = [datetime]::UtcNow
    $insert = $Context.connection.CreateCommand()

    if ([int]$Context.dbType -eq 1) {
        $insert.CommandText = @"
INSERT INTO dbo.BusyCloudNotificationDelivery
(
    Id, NotificationId, SubscriptionId,
    RecipientUserName, State, Attempts,
    CreatedAt, UpdatedAt
)
VALUES
(
    @id, @notificationId, @subscriptionId,
    @recipientUserName, 'PENDING', 0,
    @createdAt, @updatedAt
)
"@
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@id" -Value $id -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@notificationId" -Value ([string]$Notification.id) -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@subscriptionId" -Value ([string]$Subscription.id) -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@recipientUserName" -Value ([string]$Notification.recipientUserName) -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@createdAt" -Value $now -Kind Date)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 1 -Name "@updatedAt" -Value $now -Kind Date)
    }
    else {
        $insert.CommandText = @"
INSERT INTO [BusyCloudNotificationDelivery]
(
    [Id], [NotificationId], [SubscriptionId],
    [RecipientUserName], [State], [Attempts],
    [CreatedAt], [UpdatedAt]
)
VALUES
(
    ?,?,?,?,'PENDING',0,?,?
)
"@
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p1" -Value $id -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p2" -Value ([string]$Notification.id) -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p3" -Value ([string]$Subscription.id) -Kind Text -Size 36)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p4" -Value ([string]$Notification.recipientUserName) -Kind Text -Size 100)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p5" -Value $now -Kind Date)
        [void](Add-WebApprovalCommandParameter -Command $insert -DbType 0 -Name "@p6" -Value $now -Kind Date)
    }

    try {
        [void]$insert.ExecuteNonQuery()
    }
    catch {
        # Another worker may have created the same unique pair. Re-read once;
        # if it still cannot be found, surface the original insert error.
        $insertError = $_.Exception
        $retry = $Context.connection.CreateCommand()

        if ([int]$Context.dbType -eq 1) {
            $retry.CommandText = @"
SELECT TOP 1 Id, State, Attempts, NextAttemptAt
FROM dbo.BusyCloudNotificationDelivery
WHERE NotificationId=@notificationId
  AND SubscriptionId=@subscriptionId
"@
            [void](Add-WebApprovalCommandParameter -Command $retry -DbType 1 -Name "@notificationId" -Value ([string]$Notification.id) -Kind Text -Size 36)
            [void](Add-WebApprovalCommandParameter -Command $retry -DbType 1 -Name "@subscriptionId" -Value ([string]$Subscription.id) -Kind Text -Size 36)
        }
        else {
            $retry.CommandText = @"
SELECT TOP 1 [Id], [State], [Attempts], [NextAttemptAt]
FROM [BusyCloudNotificationDelivery]
WHERE [NotificationId]=?
  AND [SubscriptionId]=?
"@
            [void](Add-WebApprovalCommandParameter -Command $retry -DbType 0 -Name "@p1" -Value ([string]$Notification.id) -Kind Text -Size 36)
            [void](Add-WebApprovalCommandParameter -Command $retry -DbType 0 -Name "@p2" -Value ([string]$Subscription.id) -Kind Text -Size 36)
        }

        $retryReader = $null

        try {
            $retryReader = $retry.ExecuteReader()

            if ($retryReader.Read()) {
                return @{
                    id = [string](Read-WebApprovalReaderValue $retryReader "Id" "")
                    state = [string](Read-WebApprovalReaderValue $retryReader "State" "")
                    attempts = [int](Read-WebApprovalReaderValue $retryReader "Attempts" 0)
                    nextAttemptAt = Read-WebApprovalReaderValue $retryReader "NextAttemptAt" $null
                }
            }
        }
        finally {
            if ($retryReader) {
                try { $retryReader.Close() } catch {}
                try { $retryReader.Dispose() } catch {}
            }
        }

        throw $insertError
    }

    return @{
        id = $id
        state = "PENDING"
        attempts = 0
        nextAttemptAt = $null
    }
}

function Get-WebPushJobsForCompany {
    param(
        [string]$InstanceId,
        [string]$CompanyCode,
        [int]$Limit = 50
    )

    $permanentSubscriptions = @(
        Get-WebPushSubscriptionsInternal `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode
    )

    if ($permanentSubscriptions.Count -eq 0) {
        return @()
    }

    $byUser = @{}

    foreach ($sub in $permanentSubscriptions) {
        $key = ([string]$sub.userName).Trim().ToLowerInvariant()
        if (-not $key) { continue }

        if (-not $byUser.ContainsKey($key)) {
            $byUser[$key] = @()
        }

        $byUser[$key] += $sub
    }

    $ctx = $null
    $jobs = @()

    try {
        $ctx = Get-WebApprovalFiscalDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        [void](Ensure-WebApprovalFiscalTables -Context $ctx)

        $notifications = @(
            Get-WebPushUnreadNotificationsInternal `
                -Context $ctx `
                -Limit 200
        )

        foreach ($notification in $notifications) {
            if ($jobs.Count -ge $Limit) { break }

            $userKey = ([string]$notification.recipientUserName).Trim().ToLowerInvariant()

            if (-not $byUser.ContainsKey($userKey)) {
                continue
            }

            foreach ($subscription in @($byUser[$userKey])) {
                if ($jobs.Count -ge $Limit) { break }

                $delivery = Get-OrCreate-WebPushDeliveryInternal `
                    -Context $ctx `
                    -Notification $notification `
                    -Subscription $subscription

                $state = ([string]$delivery.state).Trim().ToUpperInvariant()

                if ($state -notin @("PENDING", "FAILED")) {
                    continue
                }

                if ([int]$delivery.attempts -ge 5) {
                    continue
                }

                $nextAt = $delivery.nextAttemptAt

                if ($null -ne $nextAt) {
                    try {
                        if ([datetime]$nextAt -gt [datetime]::UtcNow) {
                            continue
                        }
                    }
                    catch {}
                }

                $jobs += @{
                    deliveryId = [string]$delivery.id
                    notificationId = [string]$notification.id
                    subscriptionId = [string]$subscription.id
                    instanceId = $InstanceId
                    companyCode = $CompanyCode
                    notificationType = [string]$notification.type
                    webApprovalId = [string]$notification.webApprovalId
                    title = [string]$notification.title
                    message = [string]$notification.message
                    attempts = [int]$delivery.attempts
                    subscription = @{
                        endpoint = [string]$subscription.endpoint
                        keys = @{
                            p256dh = [string]$subscription.p256dh
                            auth = [string]$subscription.authKey
                        }
                    }
                }
            }
        }
    }
    finally {
        Close-WebApprovalDbContext -Context $ctx
    }

    return @($jobs)
}

function Get-WebPushPendingJobsAllCompanies {
    param([int]$Limit = 50)

    $safeLimit = [Math]::Max(1, [Math]::Min(200, $Limit))
    $instancesPath = Join-Path (Split-Path -Parent $PSScriptRoot) "instances.json"

    if (-not (Test-Path -LiteralPath $instancesPath)) {
        return @{
            success = $false
            error = "instances.json was not found."
        }
    }

    $config = Get-Content -LiteralPath $instancesPath -Raw | ConvertFrom-Json
    $jobs = @()

    foreach ($inst in @($config.instances)) {
        foreach ($comp in @($inst.companies)) {
            if ($jobs.Count -ge $safeLimit) { break }

            try {
                $remaining = $safeLimit - $jobs.Count
                $companyJobs = @(
                    Get-WebPushJobsForCompany `
                        -InstanceId ([string]$inst.id) `
                        -CompanyCode ([string]$comp.code) `
                        -Limit $remaining
                )

                $jobs += $companyJobs
            }
            catch {
                Write-Host (
                    "  [PUSH-WORKER WARN] Could not read jobs for {0}/{1}: {2}" -f `
                    $inst.id,
                    $comp.code,
                    $_.Exception.Message
                ) -ForegroundColor DarkYellow
            }
        }

        if ($jobs.Count -ge $safeLimit) { break }
    }

    return @{
        success = $true
        data = @{
            jobs = @($jobs)
            count = @($jobs).Count
        }
    }
}

function Complete-WebPushDelivery {
    param(
        [string]$InstanceId,
        [string]$CompanyCode,
        [string]$DeliveryId,
        [string]$SubscriptionId,
        [bool]$Succeeded,
        [bool]$Gone = $false,
        [int]$HttpStatus = 0,
        [string]$ErrorMessage = ""
    )

    if (
        [string]::IsNullOrWhiteSpace($InstanceId) -or
        [string]::IsNullOrWhiteSpace($CompanyCode) -or
        [string]::IsNullOrWhiteSpace($DeliveryId)
    ) {
        return @{
            success = $false
            httpStatus = 400
            error = "instanceId, companyCode and deliveryId are required."
        }
    }

    $ctx = $null

    try {
        $ctx = Get-WebApprovalFiscalDbContext `
            -InstanceId $InstanceId `
            -CompanyCode $CompanyCode

        [void](Ensure-WebApprovalFiscalTables -Context $ctx)

        $attempts = 0
        $read = $ctx.connection.CreateCommand()

        if ([int]$ctx.dbType -eq 1) {
            $read.CommandText = "SELECT TOP 1 Attempts FROM dbo.BusyCloudNotificationDelivery WHERE Id=@id"
            [void](Add-WebApprovalCommandParameter -Command $read -DbType 1 -Name "@id" -Value $DeliveryId.Trim() -Kind Text -Size 36)
        }
        else {
            $read.CommandText = "SELECT TOP 1 [Attempts] FROM [BusyCloudNotificationDelivery] WHERE [Id]=?"
            [void](Add-WebApprovalCommandParameter -Command $read -DbType 0 -Name "@p1" -Value $DeliveryId.Trim() -Kind Text -Size 36)
        }

        $raw = $read.ExecuteScalar()
        if ($null -ne $raw -and $raw -ne [System.DBNull]::Value) {
            try { $attempts = [int]$raw } catch {}
        }

        $attempts++
        $now = [datetime]::UtcNow
        $state = "FAILED"
        $sentAt = $null
        $nextAttemptAt = $null

        if ($Succeeded) {
            $state = "SENT"
            $sentAt = $now
        }
        elseif ($Gone) {
            $state = "GONE"
        }
        elseif ($attempts -lt 5) {
            $delaySeconds = [Math]::Min(
                900,
                [Math]::Pow(2, [Math]::Max(0, $attempts - 1)) * 30
            )

            $nextAttemptAt = $now.AddSeconds([int]$delaySeconds)
        }

        $update = $ctx.connection.CreateCommand()

        if ([int]$ctx.dbType -eq 1) {
            $update.CommandText = @"
UPDATE dbo.BusyCloudNotificationDelivery
SET
    State=@state,
    Attempts=@attempts,
    LastHttpStatus=@httpStatus,
    LastError=@lastError,
    UpdatedAt=@updatedAt,
    SentAt=@sentAt,
    NextAttemptAt=@nextAttemptAt
WHERE Id=@id
"@
            [void](Add-WebApprovalCommandParameter -Command $update -DbType 1 -Name "@state" -Value $state -Kind Text -Size 20)
            [void](Add-WebApprovalCommandParameter -Command $update -DbType 1 -Name "@attempts" -Value $attempts -Kind Int)
            [void](Add-WebApprovalCommandParameter -Command $update -DbType 1 -Name "@httpStatus" -Value $HttpStatus -Kind Int)
            [void](Add-WebApprovalCommandParameter -Command $update -DbType 1 -Name "@lastError" -Value $ErrorMessage -Kind LongText)
            [void](Add-WebApprovalCommandParameter -Command $update -DbType 1 -Name "@updatedAt" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $update -DbType 1 -Name "@sentAt" -Value $sentAt -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $update -DbType 1 -Name "@nextAttemptAt" -Value $nextAttemptAt -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $update -DbType 1 -Name "@id" -Value $DeliveryId.Trim() -Kind Text -Size 36)
        }
        else {
            $update.CommandText = @"
UPDATE [BusyCloudNotificationDelivery]
SET
    [State]=?,
    [Attempts]=?,
    [LastHttpStatus]=?,
    [LastError]=?,
    [UpdatedAt]=?,
    [SentAt]=?,
    [NextAttemptAt]=?
WHERE [Id]=?
"@
            [void](Add-WebApprovalCommandParameter -Command $update -DbType 0 -Name "@p1" -Value $state -Kind Text -Size 20)
            [void](Add-WebApprovalCommandParameter -Command $update -DbType 0 -Name "@p2" -Value $attempts -Kind Int)
            [void](Add-WebApprovalCommandParameter -Command $update -DbType 0 -Name "@p3" -Value $HttpStatus -Kind Int)
            [void](Add-WebApprovalCommandParameter -Command $update -DbType 0 -Name "@p4" -Value $ErrorMessage -Kind LongText)
            [void](Add-WebApprovalCommandParameter -Command $update -DbType 0 -Name "@p5" -Value $now -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $update -DbType 0 -Name "@p6" -Value $sentAt -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $update -DbType 0 -Name "@p7" -Value $nextAttemptAt -Kind Date)
            [void](Add-WebApprovalCommandParameter -Command $update -DbType 0 -Name "@p8" -Value $DeliveryId.Trim() -Kind Text -Size 36)
        }

        [void]$update.ExecuteNonQuery()
    }
    catch {
        return @{
            success = $false
            httpStatus = 500
            error = $_.Exception.Message
        }
    }
    finally {
        Close-WebApprovalDbContext -Context $ctx
    }

    if ($Gone -and -not [string]::IsNullOrWhiteSpace($SubscriptionId)) {
        $perm = $null

        try {
            $perm = Get-WebApprovalPermanentDbContext `
                -InstanceId $InstanceId `
                -CompanyCode $CompanyCode

            $disable = $perm.connection.CreateCommand()
            $now = [datetime]::UtcNow

            if ([int]$perm.dbType -eq 1) {
                $disable.CommandText = @"
UPDATE dbo.BusyCloudPushSubscription
SET IsActive=0, UpdatedAt=@updatedAt
WHERE Id=@id
"@
                [void](Add-WebApprovalCommandParameter -Command $disable -DbType 1 -Name "@updatedAt" -Value $now -Kind Date)
                [void](Add-WebApprovalCommandParameter -Command $disable -DbType 1 -Name "@id" -Value $SubscriptionId.Trim() -Kind Text -Size 36)
            }
            else {
                $disable.CommandText = @"
UPDATE [BusyCloudPushSubscription]
SET [IsActive]=0, [UpdatedAt]=?
WHERE [Id]=?
"@
                [void](Add-WebApprovalCommandParameter -Command $disable -DbType 0 -Name "@p1" -Value $now -Kind Date)
                [void](Add-WebApprovalCommandParameter -Command $disable -DbType 0 -Name "@p2" -Value $SubscriptionId.Trim() -Kind Text -Size 36)
            }

            [void]$disable.ExecuteNonQuery()
        }
        catch {}
        finally {
            Close-WebApprovalDbContext -Context $perm
        }
    }

    return @{
        success = $true
        data = @{
            state = $state
            attempts = $attempts
        }
    }
}



# ============================================================================
# PHASE 6.1 - ZERO-MANUAL-SERVER WEB PUSH BOOTSTRAP
#
# Goal:
#   A newly deployed VPS should not require an administrator to manually run
#   npm install, generate VAPID keys, or start the Push worker.
#
# On first backend start this code will:
#   1. Reuse system Node.js if available.
#   2. Otherwise download a portable Node.js LTS runtime into:
#          <backend>\runtime\node
#   3. Install the local "web-push" npm dependency when missing.
#   4. Generate push_config.json when missing.
#   5. Start the Push worker.
#
# If Internet access is temporarily unavailable, the BUSY API still starts and
# in-app notifications continue to work. A later backend restart retries setup.
# ============================================================================

function Get-BusyCloudBackendRoot {
    return Split-Path -Parent $PSScriptRoot
}

function Get-BusyCloudLocalNodePath {
    $backendRoot = Get-BusyCloudBackendRoot
    return Join-Path $backendRoot "runtime\node\node.exe"
}

function Get-BusyCloudLocalNpmPath {
    $backendRoot = Get-BusyCloudBackendRoot
    return Join-Path $backendRoot "runtime\node\npm.cmd"
}

function Get-BusyCloudNodeRuntime {
    $systemNode = Get-Command node -ErrorAction SilentlyContinue
    $systemNpm = Get-Command npm -ErrorAction SilentlyContinue

    if ($systemNode -and $systemNpm) {
        return @{
            success = $true
            source = "SYSTEM"
            node = $systemNode.Source
            npm = $systemNpm.Source
        }
    }

    $localNode = Get-BusyCloudLocalNodePath
    $localNpm = Get-BusyCloudLocalNpmPath

    if (
        (Test-Path -LiteralPath $localNode) -and
        (Test-Path -LiteralPath $localNpm)
    ) {
        return @{
            success = $true
            source = "LOCAL"
            node = $localNode
            npm = $localNpm
        }
    }

    return @{
        success = $false
        source = "NONE"
        node = ""
        npm = ""
    }
}

function Install-BusyCloudPortableNode {
    $backendRoot = Get-BusyCloudBackendRoot
    $runtimeRoot = Join-Path $backendRoot "runtime"
    $targetDir = Join-Path $runtimeRoot "node"
    $tempRoot = Join-Path $runtimeRoot "_node_download"

    try {
        New-Item -ItemType Directory -Force -Path $runtimeRoot | Out-Null

        if (Test-Path -LiteralPath $tempRoot) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }

        New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null

        # Windows PowerShell 5.1 may otherwise negotiate an older TLS version.
        try {
            [Net.ServicePointManager]::SecurityProtocol = `
                [Net.ServicePointManager]::SecurityProtocol -bor `
                [Net.SecurityProtocolType]::Tls12
        }
        catch {}

        Write-Host "  [PUSH-BOOTSTRAP] Node.js not found. Resolving current LTS release..." -ForegroundColor Yellow

        $indexUrl = "https://nodejs.org/dist/index.json"
        $releases = Invoke-RestMethod `
            -Uri $indexUrl `
            -UseBasicParsing `
            -TimeoutSec 45

        $release = @($releases | Where-Object {
            $_.lts -and ([string]$_.version).StartsWith("v22.")
        } | Select-Object -First 1)

        if (-not $release -or $release.Count -eq 0) {
            $release = @($releases | Where-Object {
                $_.lts
            } | Select-Object -First 1)
        }

        if (-not $release -or $release.Count -eq 0) {
            throw "Could not resolve a Node.js LTS release."
        }

        $version = ([string]$release[0].version).Trim()

        if ([string]::IsNullOrWhiteSpace($version)) {
            throw "Node.js release metadata did not contain a version."
        }

        $platform = if ([Environment]::Is64BitOperatingSystem) {
            "win-x64"
        }
        else {
            "win-x86"
        }

        $fileName = "node-$version-$platform.zip"
        $downloadUrl = "https://nodejs.org/dist/$version/$fileName"
        $zipPath = Join-Path $tempRoot $fileName
        $extractRoot = Join-Path $tempRoot "extract"

        Write-Host "  [PUSH-BOOTSTRAP] Downloading Node.js $version ($platform)..." -ForegroundColor Yellow

        Invoke-WebRequest `
            -Uri $downloadUrl `
            -OutFile $zipPath `
            -UseBasicParsing `
            -TimeoutSec 180

        New-Item -ItemType Directory -Force -Path $extractRoot | Out-Null

        Expand-Archive `
            -LiteralPath $zipPath `
            -DestinationPath $extractRoot `
            -Force

        $extracted = Get-ChildItem `
            -LiteralPath $extractRoot `
            -Directory |
            Select-Object -First 1

        if (-not $extracted) {
            throw "Downloaded Node.js archive did not contain the expected folder."
        }

        if (Test-Path -LiteralPath $targetDir) {
            Remove-Item -LiteralPath $targetDir -Recurse -Force
        }

        Move-Item `
            -LiteralPath $extracted.FullName `
            -Destination $targetDir `
            -Force

        $node = Join-Path $targetDir "node.exe"
        $npm = Join-Path $targetDir "npm.cmd"

        if (
            -not (Test-Path -LiteralPath $node) -or
            -not (Test-Path -LiteralPath $npm)
        ) {
            throw "Portable Node.js installation did not produce node.exe/npm.cmd."
        }

        Write-Host "  [PUSH-BOOTSTRAP] Portable Node.js installed: $version" -ForegroundColor Green

        return @{
            success = $true
            source = "LOCAL"
            version = $version
            node = $node
            npm = $npm
        }
    }
    catch {
        return @{
            success = $false
            error = $_.Exception.Message
        }
    }
    finally {
        if (Test-Path -LiteralPath $tempRoot) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Resolve-BusyCloudNpmExecutable {
    param(
        [string]$NodePath,
        [string]$NpmPath
    )

    $nodeDir = Split-Path -Parent $NodePath

    # npm.cmd beside node.exe is the most reliable Windows launcher.
    $besideNode = Join-Path $nodeDir "npm.cmd"
    if (Test-Path -LiteralPath $besideNode) {
        return $besideNode
    }

    # Get-Command npm often returns npm.ps1. Prefer the sibling npm.cmd
    # because Start-Process / 32-bit Windows PowerShell can behave
    # inconsistently with the PowerShell shim.
    if (-not [string]::IsNullOrWhiteSpace($NpmPath)) {
        $npmDir = Split-Path -Parent $NpmPath
        $siblingCmd = Join-Path $npmDir "npm.cmd"

        if (Test-Path -LiteralPath $siblingCmd) {
            return $siblingCmd
        }

        if (Test-Path -LiteralPath $NpmPath) {
            return $NpmPath
        }
    }

    return ""
}

function Install-BusyCloudWebPushDependency {
    param(
        [string]$NodePath,
        [string]$NpmPath
    )

    $backendRoot = Get-BusyCloudBackendRoot
    $workerDir = Join-Path $backendRoot "push-worker"
    $webPushDir = Join-Path $workerDir "node_modules\web-push"
    $webPushPackage = Join-Path $webPushDir "package.json"
    $pushDataDir = Join-Path $backendRoot "data\push"
    $npmLog = Join-Path $pushDataDir "bootstrap-npm.log"

    if (Test-Path -LiteralPath $webPushPackage) {
        return @{
            success = $true
            installed = $false
        }
    }

    if (-not (Test-Path -LiteralPath $workerDir)) {
        return @{
            success = $false
            error = "push-worker folder is missing: $workerDir"
        }
    }

    $packageJson = Join-Path $workerDir "package.json"

    if (-not (Test-Path -LiteralPath $packageJson)) {
        return @{
            success = $false
            error = "push-worker\package.json is missing."
        }
    }

    if (
        [string]::IsNullOrWhiteSpace($NodePath) -or
        -not (Test-Path -LiteralPath $NodePath)
    ) {
        return @{
            success = $false
            error = "Node.js executable is unavailable: $NodePath"
        }
    }

    $npmExecutable = Resolve-BusyCloudNpmExecutable `
        -NodePath $NodePath `
        -NpmPath $NpmPath

    if (
        [string]::IsNullOrWhiteSpace($npmExecutable) -or
        -not (Test-Path -LiteralPath $npmExecutable)
    ) {
        return @{
            success = $false
            error = "npm executable could not be resolved."
        }
    }

    try {
        New-Item -ItemType Directory -Force -Path $pushDataDir | Out-Null

        Write-Host "  [PUSH-BOOTSTRAP] Installing Web Push dependency..." -ForegroundColor Yellow
        Write-Host "  [PUSH-BOOTSTRAP] npm launcher: $npmExecutable" -ForegroundColor DarkGray
        Write-Host "  [PUSH-BOOTSTRAP] npm prefix  : $workerDir" -ForegroundColor DarkGray

        $nodeDir = Split-Path -Parent $NodePath
        $oldPath = $env:PATH
        $oldNpmPrefix = $env:npm_config_prefix
        $oldNpmCache = $env:npm_config_cache

        try {
            $env:PATH = "$nodeDir;$oldPath"

            # Force npm to install into THIS application's push-worker folder.
            # This is the important fix: relying only on WorkingDirectory can
            # cause npm shims to install elsewhere on some Windows VPS setups.
            $env:npm_config_prefix = $workerDir

            # Keep npm cache application-local so service accounts / VPS users
            # do not need a writable roaming-profile npm cache.
            $npmCache = Join-Path $backendRoot "runtime\npm-cache"
            New-Item -ItemType Directory -Force -Path $npmCache | Out-Null
            $env:npm_config_cache = $npmCache

            $arguments = @(
                "install",
                "--prefix", $workerDir,
                "--omit=dev",
                "--no-audit",
                "--no-fund",
                "--loglevel=warn"
            )

            $output = @()

            Push-Location $workerDir
            try {
                # Use the Windows npm.cmd directly instead of Start-Process.
                # This preserves the correct prefix and gives us the real
                # npm exit code/output under 32-bit Windows PowerShell.
                $output = @(
                    & $npmExecutable @arguments 2>&1 |
                    ForEach-Object { [string]$_ }
                )
                $exitCode = $LASTEXITCODE
            }
            finally {
                Pop-Location
            }

            $output | Set-Content `
                -LiteralPath $npmLog `
                -Encoding UTF8

            if ($exitCode -ne 0) {
                $tail = @($output | Select-Object -Last 8) -join " | "
                throw "npm install exited with code $exitCode. $tail"
            }
        }
        finally {
            $env:PATH = $oldPath

            if ($null -eq $oldNpmPrefix) {
                Remove-Item Env:\npm_config_prefix -ErrorAction SilentlyContinue
            }
            else {
                $env:npm_config_prefix = $oldNpmPrefix
            }

            if ($null -eq $oldNpmCache) {
                Remove-Item Env:\npm_config_cache -ErrorAction SilentlyContinue
            }
            else {
                $env:npm_config_cache = $oldNpmCache
            }
        }

        # npm may report exit code 0 even if a wrapper/shim ignored the target.
        # Verify the actual dependency file, not only the process exit code.
        if (-not (Test-Path -LiteralPath $webPushPackage)) {
            throw (
                "npm returned success but the expected package is still missing: " +
                "$webPushPackage. See $npmLog"
            )
        }

        # Final runtime resolution test using the exact Node executable.
        $resolveScript = @"
const path = require('path');
const worker = process.argv[1];
try {
  const resolved = require.resolve('web-push', { paths: [worker] });
  process.stdout.write(resolved);
  process.exit(0);
} catch (e) {
  process.stderr.write(String(e && e.stack || e));
  process.exit(2);
}
"@

        $resolveOutput = @(
            & $NodePath `
                "-e" `
                $resolveScript `
                $workerDir 2>&1 |
            ForEach-Object { [string]$_ }
        )

        $resolveExit = $LASTEXITCODE

        if ($resolveExit -ne 0) {
            $detail = @($resolveOutput | Select-Object -Last 5) -join " | "
            throw "web-push was installed but Node.js cannot resolve it. $detail"
        }

        Write-Host "  [PUSH-BOOTSTRAP] Web Push dependency installed and verified." -ForegroundColor Green

        return @{
            success = $true
            installed = $true
            npmLog = $npmLog
        }
    }
    catch {
        return @{
            success = $false
            error = $_.Exception.Message
            npmLog = $npmLog
        }
    }
}

function New-BusyCloudWebPushConfig {
    param([string]$NodePath)

    $backendRoot = Get-BusyCloudBackendRoot
    $workerDir = Join-Path $backendRoot "push-worker"
    $generator = Join-Path $workerDir "generate_config.js"
    $configPath = Get-WebPushConfigPath

    if (Test-Path -LiteralPath $configPath) {
        return @{
            success = $true
            generated = $false
            path = $configPath
        }
    }

    if (-not (Test-Path -LiteralPath $generator)) {
        return @{
            success = $false
            error = "push-worker\generate_config.js is missing."
        }
    }

    try {
        $pushDir = Split-Path -Parent $configPath
        New-Item -ItemType Directory -Force -Path $pushDir | Out-Null

        $subject = ([string]$env:BUSYCLOUD_VAPID_SUBJECT).Trim()

        if ([string]::IsNullOrWhiteSpace($subject)) {
            # A syntactically valid VAPID contact URI. Deployments may override
            # this through BUSYCLOUD_VAPID_SUBJECT without editing code.
            $subject = "mailto:busycloud@example.invalid"
        }

        Write-Host "  [PUSH-BOOTSTRAP] Generating VAPID keys + worker secret..." -ForegroundColor Yellow

        $proc = Start-Process `
            -FilePath $NodePath `
            -ArgumentList @(
                "`"$generator`"",
                "`"$subject`""
            ) `
            -WorkingDirectory $workerDir `
            -WindowStyle Hidden `
            -Wait `
            -PassThru

        if ($proc.ExitCode -ne 0) {
            throw "VAPID config generator exited with code $($proc.ExitCode)."
        }

        $cfg = Get-WebPushConfigInternal

        if (-not $cfg.configured) {
            throw "push_config.json was generated but is not valid/configured."
        }

        Write-Host "  [PUSH-BOOTSTRAP] Web Push keys configured automatically." -ForegroundColor Green

        return @{
            success = $true
            generated = $true
            path = $configPath
        }
    }
    catch {
        return @{
            success = $false
            error = $_.Exception.Message
        }
    }
}

function Initialize-BusyCloudPushRuntime {
    $backendRoot = Get-BusyCloudBackendRoot
    $workerJs = Join-Path $backendRoot "push-worker\worker.js"

    if (-not (Test-Path -LiteralPath $workerJs)) {
        return @{
            success = $false
            ready = $false
            reason = "push-worker\worker.js is not installed."
        }
    }

    $runtime = Get-BusyCloudNodeRuntime

    if (-not $runtime.success) {
        $runtime = Install-BusyCloudPortableNode

        if (-not $runtime.success) {
            return @{
                success = $false
                ready = $false
                reason = "Node.js bootstrap failed: $($runtime.error)"
            }
        }
    }

    $dependency = Install-BusyCloudWebPushDependency `
        -NodePath ([string]$runtime.node) `
        -NpmPath ([string]$runtime.npm)

    if (-not $dependency.success) {
        return @{
            success = $false
            ready = $false
            reason = "Web Push dependency bootstrap failed: $($dependency.error)"
        }
    }

    $config = New-BusyCloudWebPushConfig `
        -NodePath ([string]$runtime.node)

    if (-not $config.success) {
        return @{
            success = $false
            ready = $false
            reason = "VAPID bootstrap failed: $($config.error)"
        }
    }

    return @{
        success = $true
        ready = $true
        node = [string]$runtime.node
        npm = [string]$runtime.npm
        source = [string]$runtime.source
        configPath = [string]$config.path
    }
}

function Start-BusyCloudPushWorker {
    param([int]$ApiPort = 8081)

    $backendRoot = Split-Path -Parent $PSScriptRoot
    $workerDir = Join-Path $backendRoot "push-worker"
    $workerJs = Join-Path $workerDir "worker.js"
    $pushDataDir = Join-Path $backendRoot "data\push"
    $pidFile = Join-Path $pushDataDir "worker.pid"

    $bootstrap = Initialize-BusyCloudPushRuntime

    if (-not $bootstrap.ready) {
        return @{
            success = $false
            started = $false
            configured = $false
            reason = [string]$bootstrap.reason
            error = [string]$bootstrap.reason
        }
    }

    $nodePath = ([string]$bootstrap.node).Trim()

    if (
        [string]::IsNullOrWhiteSpace($nodePath) -or
        -not (Test-Path -LiteralPath $nodePath)
    ) {
        return @{
            success = $false
            started = $false
            configured = $false
            reason = "Node.js runtime is unavailable after bootstrap."
            error = "Node.js runtime is unavailable after bootstrap."
        }
    }

    New-Item -ItemType Directory -Force -Path $pushDataDir | Out-Null

    # Respect an already-running worker.
    if (Test-Path -LiteralPath $pidFile) {
        try {
            $existingPid = [int](Get-Content -LiteralPath $pidFile -Raw).Trim()
            $existing = Get-Process -Id $existingPid -ErrorAction SilentlyContinue

            if ($existing) {
                return @{
                    success = $true
                    started = $false
                    configured = $true
                    alreadyRunning = $true
                    pid = $existingPid
                }
            }

            Remove-Item -LiteralPath $pidFile -Force -ErrorAction SilentlyContinue
        }
        catch {
            Remove-Item -LiteralPath $pidFile -Force -ErrorAction SilentlyContinue
        }
    }

    $stdout = Join-Path $pushDataDir "worker.stdout.log"
    $stderr = Join-Path $pushDataDir "worker.stderr.log"

    try {
        $arguments = @(
            "`"$workerJs`"",
            "--parent-pid",
            [string]$PID,
            "--api",
            "http://127.0.0.1:$ApiPort"
        )

        $proc = Start-Process `
            -FilePath $nodePath `
            -ArgumentList $arguments `
            -WorkingDirectory $workerDir `
            -WindowStyle Hidden `
            -RedirectStandardOutput $stdout `
            -RedirectStandardError $stderr `
            -PassThru

        return @{
            success = $true
            started = $true
            configured = $true
            pid = $proc.Id
            stdout = $stdout
            stderr = $stderr
        }
    }
    catch {
        return @{
            success = $false
            started = $false
            configured = $true
            error = $_.Exception.Message
        }
    }
}
