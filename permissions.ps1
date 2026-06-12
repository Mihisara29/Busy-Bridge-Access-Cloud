# permissions.ps1 — Dump UserControl + all users

$logFile = "$env:USERPROFILE\Desktop\busy_permissions_log.txt"
function Log($msg) {
    $line = "[$((Get-Date).ToString('HH:mm:ss'))] $msg"
    Add-Content -Path $logFile -Value $line
    Write-Host $line
}
"" | Out-File $logFile -Encoding UTF8

$busyPath    = "C:\Enterprise02"
$dataPath    = "C:\Enterprise02\DATA"
$companyCode = "COMP0002"

Log "=== BUSY UserControl Dump ==="

try {
    $fi = New-Object -ComObject "Busy2L21.CFixedInterface"
    $fi.OpenDB($busyPath, $dataPath, $companyCode, 0) | Out-Null
    Log "Connected OK"
} catch {
    Log "Connection FAILED: $_"
    exit
}

# ── 1. All users from Master1 (MasterType=60) ──
Log ""
Log "[1] All users in Master1 (MasterType=60)..."
try {
    $rs = $fi.GetRecordset("SELECT Code, Name FROM Master1 WHERE MasterType=60 ORDER BY Code")
    if ($rs -and -not $rs.EOF) {
        $rs.MoveFirst()
        while (-not $rs.EOF) {
            Log "    Code=$($rs.Fields.Item('Code').Value) | Name=$($rs.Fields.Item('Name').Value)"
            $rs.MoveNext()
        }
        $rs.Close()
    } else {
        Log "    No users found"
    }
} catch { Log "    Error: $_" }

# ── 2. ALL rows from UserControl (common DB) ──
Log ""
Log "[2] All rows in UserControl..."
try {
    $rs = $fi.GetRecordsetFromCompanyDB("SELECT * FROM UserControl ORDER BY Name")
    if ($rs -and -not $rs.EOF) {
        $rs.MoveFirst()
        $rowNum = 0
        while (-not $rs.EOF) {
            $rowNum++
            Log ""
            Log "  --- Row $rowNum ---"
            for ($i = 0; $i -lt $rs.Fields.Count; $i++) {
                $colName = $rs.Fields.Item($i).Name
                $colVal  = $rs.Fields.Item($i).Value
                # Translate boolean-like values
                if ($colVal -eq -1) { $colVal = "YES" }
                elseif ($colVal -eq 0 -and $colName -ne "RecType" -and $colName -ne "I1" -and $colName -ne "I2" -and $colName -ne "I3" -and $colName -ne "I4") { $colVal = "NO" }
                Log "    $($colName.PadRight(12)) = $colVal"
            }
            $rs.MoveNext()
        }
        $rs.Close()
        Log ""
        Log "  Total rows: $rowNum"
    } else {
        Log "    No rows found"
    }
} catch { Log "    Error: $_" }

# ── 3. UserControl for Admin only ──
Log ""
Log "[3] UserControl row for 'Admin' only..."
try {
    $rs = $fi.GetRecordsetFromCompanyDB("SELECT * FROM UserControl WHERE Name='Admin'")
    if ($rs -and -not $rs.EOF) {
        $rs.MoveFirst()
        Log "  Found:"
        for ($i = 0; $i -lt $rs.Fields.Count; $i++) {
            $colName = $rs.Fields.Item($i).Name
            $colVal  = $rs.Fields.Item($i).Value
            if ($colVal -eq -1) { $colVal = "YES" }
            Log "    $($colName.PadRight(12)) = $colVal"
        }
        $rs.Close()
    } else {
        Log "    Admin not found in UserControl"
        Log "    Trying case-insensitive search..."
        $rs2 = $fi.GetRecordsetFromCompanyDB("SELECT DISTINCT Name FROM UserControl")
        if ($rs2 -and -not $rs2.EOF) {
            $rs2.MoveFirst()
            Log "    Names found in UserControl:"
            while (-not $rs2.EOF) {
                Log "      '$($rs2.Fields.Item('Name').Value)'"
                $rs2.MoveNext()
            }
            $rs2.Close()
        }
    }
} catch { Log "    Error: $_" }

# ── 4. RecType values in UserControl ──
Log ""
Log "[4] Distinct RecType values in UserControl..."
try {
    $rs = $fi.GetRecordsetFromCompanyDB("SELECT DISTINCT RecType, Name FROM UserControl ORDER BY RecType, Name")
    if ($rs -and -not $rs.EOF) {
        $rs.MoveFirst()
        while (-not $rs.EOF) {
            Log "    RecType=$($rs.Fields.Item('RecType').Value) | Name=$($rs.Fields.Item('Name').Value)"
            $rs.MoveNext()
        }
        $rs.Close()
    }
} catch { Log "    Error: $_" }

try { $fi.CloseDB() }                                                               catch {}
try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($fi) | Out-Null } catch {}

Log ""
Log "=== DONE === Log: $logFile"