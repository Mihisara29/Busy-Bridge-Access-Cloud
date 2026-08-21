# modules/routes.ps1
# HTTP Route Handlers — Targeted Native Auth & Audit Trail Integration

function Read-RequestBody {
    param($request)
    $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
    return $reader.ReadToEnd()
}

function Send-Response {
    param($response, $data, $statusCode = $null)
    try {
        if ($null -ne $statusCode) { $response.StatusCode = $statusCode }
        $json   = $data | ConvertTo-Json -Depth 10
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
        $response.OutputStream.Write($buffer, 0, $buffer.Length)
        $response.OutputStream.Close()
    } catch { }
}

function Get-QueryStringValue {
    param($queryString, [string]$key, [string]$defaultValue = "")
    $value = $queryString[$key]
    if ($value -and $value -ne "") { return $value }
    return $defaultValue
}

function Get-ServerId {
    return "busy-server-native-unified"
}

function Start-BUSYServer {
    param(
        [int]$Port = 8081
    )

    $listener = New-Object System.Net.HttpListener
    try {
        $listener.Prefixes.Add("http://*:$Port/")
        Write-Host "  Added wildcard binding: http://*:$Port" -ForegroundColor Green
    } catch {
        Write-Host "  Could not add wildcard - run as Administrator" -ForegroundColor Yellow
        $listener = New-Object System.Net.HttpListener
        $listener.Prefixes.Add("http://localhost:$Port/")
        $listener.Prefixes.Add("http://127.0.0.1:$Port/")
    }

    try {
        $listener.Start()
        Write-Host "  Server started successfully!" -ForegroundColor Green
    } catch {
        Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    $config = Get-Config
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  BUSY 21 Enterprise API Bridge v4.5"   -ForegroundColor Green
    Write-Host "  Targeted Login + Audit Trail Engine"   -ForegroundColor Green
    Write-Host "  Server ID : $(Get-ServerId)"           -ForegroundColor Yellow
    Write-Host "  Port      : $Port"                     -ForegroundColor Yellow
    Write-Host "  Press Ctrl+C to stop."                 -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Cyan

    while ($listener.IsListening) {
        $context  = $listener.GetContext()
        $request  = $context.Request
        $response = $context.Response

        $response.Headers.Add("Content-Type",                 "application/json")
        $response.Headers.Add("Access-Control-Allow-Origin",  "*")
        $response.Headers.Add("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
        $response.Headers.Add("Access-Control-Allow-Headers", "Content-Type, Authorization, X-Bridge-Secret, X-Instance-ID, X-Company-Code, Idempotency-Key")

        if ($request.HttpMethod -eq "OPTIONS") {
            $response.StatusCode = 200
            $response.OutputStream.Close()
            continue
        }

        $path   = $request.Url.LocalPath
        $method = $request.HttpMethod
        Write-Host "[$method] $path" -ForegroundColor Yellow

        try {
            $result = $null

            # ════════════════════════════════════════════════
            #  PUBLIC ENDPOINTS (no auth required)
            # ════════════════════════════════════════════════

            if ($path -eq "/health" -and $method -eq "GET") {
                $result = @{ success = $true; status = "BUSY Bridge is running"; version = "4.5"; serverId = Get-ServerId }
                Send-Response $response $result
                Write-Host "  [OK] $path" -ForegroundColor Green
                continue
            }

            # ── LEGACY TARGETED LOGIN (company required) ──
            if ($path -eq "/auth/login" -and $method -eq "POST") {
                $body = Read-RequestBody $request | ConvertFrom-Json

                if (-not $body.company -or -not $body.username -or -not $body.password) {
                    Send-Response $response @{success=$false; error="company, username, and password required"} 400
                    continue
                }

                $loginResult = Invoke-BusyLogin -CompanyIdentifier $body.company -Username $body.username -Password $body.password

                if ($loginResult.success) {
                    Send-Response $response @{ success=$true; token=$loginResult.token; user=$loginResult.user }
                    Write-Host "  [LOGIN SUCCESS] $($body.username) connected to $($body.company)" -ForegroundColor Green
                } else {
                    Send-Response $response @{ success=$false; error=$loginResult.error } 401
                    Write-Host "  [LOGIN FAIL] $($body.username) - $($loginResult.error)" -ForegroundColor Red
                }
                continue
            }

            # ── SCAN LOGIN (auto-detect company from credentials) ──
            if ($path -eq "/auth/scan-login" -and $method -eq "POST") {
                $body = Read-RequestBody $request | ConvertFrom-Json

                if (-not $body.username) {
                    Send-Response $response @{ success=$false; error="username is required" } 400
                    continue
                }

                $scanResult = Invoke-ScanLogin `
                    -Username      $body.username `
                    -Password      ([string]$body.password)

                if ($scanResult.success) {
                    Send-Response $response @{ success=$true; token=$scanResult.token; user=$scanResult.user }
                    Write-Host "  [SCAN LOGIN SUCCESS] $($body.username)" -ForegroundColor Green
                } else {
                    $status = if ($scanResult.error -eq "AMBIGUOUS_CREDENTIALS") { 409 } else { 401 }
                    Send-Response $response @{ success=$false; error=$scanResult.error; detail=$scanResult.detail } $status
                    Write-Host "  [SCAN LOGIN FAIL] $($body.username) - $($scanResult.error)" -ForegroundColor Red
                }
                continue
            }

            if ($path -eq "/auth/me" -and $method -eq "GET") {
                $authResult = Invoke-AuthCheck -Request $request
                if (-not $authResult.allowed) {
                    Send-Response $response @{success=$false; error=$authResult.reason} 401
                    Write-Host "  [AUTH FAIL] $($authResult.reason)" -ForegroundColor Red
                    continue
                }
                $user = $authResult.user

                $instArray = @()
                if ($null -ne $user.instances) { $instArray = @($user.instances) }

                $result = @{
                    success  = $true;
                    user     = @{ uid = $user.uid; name = $user.name; role = $user.role; permissions = $user.permissions };
                    instances = $instArray;
                    serverId = Get-ServerId
                }
                Send-Response $response $result
                Write-Host "  [OK] /auth/me ($($user.name))" -ForegroundColor Green
                continue
            }

            # ════════════════════════════════════════════════
            #  PROTECTED ENDPOINTS (Auth & Headers required)
            # ════════════════════════════════════════════════

            $instanceId   = $request.Headers["X-Instance-ID"]
            $companyCode  = $request.Headers["X-Company-Code"]
            $bridgeSecret = $request.Headers["X-Bridge-Secret"]
            $cfg          = Get-Config
            $requireAuth  = $true

            if ($bridgeSecret -eq $cfg.BRIDGE_SECRET -and -not $instanceId) {
                $requireAuth = $false
                $instanceId  = $cfg.INSTANCE_ID
                $companyCode = $cfg.COMP_CODE
            }

            if ($requireAuth) {
                if (-not $instanceId -or -not $companyCode) {
                    Send-Response $response @{success=$false; error="X-Instance-ID and X-Company-Code headers are required"} 400
                    continue
                }
                $authResult = Invoke-AuthCheck -Request $request -InstanceId $instanceId -CompanyCode $companyCode
                if (-not $authResult.allowed) {
                    Send-Response $response @{success=$false; error=$authResult.reason} 403
                    Write-Host "  [DENIED] $($authResult.reason)" -ForegroundColor Red
                    continue
                }
            }

            # --- COMPANY SETTINGS ---
            if ($path -eq "/busy/company" -and $method -eq "GET") {
                $result = Get-CompanyDetails -InstanceId $instanceId -CompanyCode $companyCode

            # --- VOUCHER MANAGEMENT ---
            } elseif ($path -eq "/busy/voucher" -and $method -eq "POST") {
                $bodyObj = Read-RequestBody $request | ConvertFrom-Json

                if ($authResult.user.name) {
                    $bodyObj | Add-Member -MemberType NoteProperty -Name "bridgeUserName" -Value $authResult.user.name -Force
                }

                $result = Create-Voucher -Data $bodyObj -InstanceId $instanceId -CompanyCode $companyCode

            # --- OFFLINE VOUCHER SYNCHRONIZATION ---
            } elseif ($path -eq "/busy/offline-vouchers/sync" -and $method -eq "POST") {
                if (-not (Get-Command Sync-OfflineVoucher -ErrorAction SilentlyContinue)) {
                    $result = @{
                        success   = $false
                        errorCode = "OFFLINE_SYNC_NOT_LOADED"
                        error     = "Offline synchronization functions are not loaded."
                    }
                    $response.StatusCode = 503
                }
                elseif (-not $requireAuth -or $null -eq $authResult -or $null -eq $authResult.user) {
                    $result = @{
                        success   = $false
                        errorCode = "AUTH_REQUIRED"
                        error     = "An authenticated user is required to synchronize offline vouchers."
                    }
                    $response.StatusCode = 401
                }
                else {
                    $bodyText = Read-RequestBody $request

                    if ([string]::IsNullOrWhiteSpace($bodyText)) {
                        $result = @{
                            success   = $false
                            errorCode = "EMPTY_REQUEST"
                            error     = "Offline voucher payload is required."
                        }
                        $response.StatusCode = 400
                    }
                    else {
                        $bodyObj = $null

                        try {
                            $bodyObj = $bodyText | ConvertFrom-Json
                        }
                        catch {
                            $result = @{
                                success   = $false
                                errorCode = "INVALID_JSON"
                                error     = "The offline voucher request body is not valid JSON."
                            }
                            $response.StatusCode = 400
                        }

                        if ($null -ne $bodyObj) {
                            $bodyObj | Add-Member -MemberType NoteProperty -Name "instanceId" -Value ([string]$instanceId) -Force
                            $bodyObj | Add-Member -MemberType NoteProperty -Name "companyCode" -Value ([string]$companyCode) -Force
                            $bodyObj | Add-Member -MemberType NoteProperty -Name "userName" -Value ([string]$authResult.user.name) -Force

                            $idempotencyKey = [string]$request.Headers["Idempotency-Key"]
                            if ([string]::IsNullOrWhiteSpace($idempotencyKey)) {
                                $idempotencyKey = [string]$bodyObj.localId
                            }

                            if ([string]::IsNullOrWhiteSpace($idempotencyKey)) {
                                $result = @{
                                    success   = $false
                                    errorCode = "IDEMPOTENCY_KEY_REQUIRED"
                                    error     = "Idempotency-Key header or localId is required."
                                }
                                $response.StatusCode = 400
                            }
                            elseif (-not [string]::IsNullOrWhiteSpace([string]$bodyObj.localId) -and ([string]$bodyObj.localId).Trim() -ne $idempotencyKey.Trim()) {
                                $result = @{
                                    success   = $false
                                    errorCode = "IDEMPOTENCY_KEY_MISMATCH"
                                    error     = "Idempotency-Key must match localId."
                                }
                                $response.StatusCode = 400
                            }
                            else {
                                $bodyObj | Add-Member -MemberType NoteProperty -Name "localId" -Value $idempotencyKey.Trim() -Force

                                try {
                                    $result = Sync-OfflineVoucher -Data $bodyObj -CurrentUser $authResult.user

                                    if ($null -eq $result) {
                                        $result = @{
                                            success   = $false
                                            localId   = $idempotencyKey.Trim()
                                            errorCode = "EMPTY_SYNC_RESULT"
                                            error     = "Offline synchronization returned no result."
                                        }
                                        $response.StatusCode = 500
                                    }
                                    elseif ($result.success -eq $false) {
                                        if ($result.conflict -eq $true) { $response.StatusCode = 409 }
                                        else { $response.StatusCode = 400 }
                                    }
                                    else {
                                        $response.StatusCode = 200
                                    }
                                }
                                catch {
                                    $message = $_.Exception.Message
                                    $statusCode = 500
                                    $errorCode = "OFFLINE_SYNC_FAILED"
                                    $isConflict = $false

                                    if ($message -match "already being processed") {
                                        $statusCode = 409
                                        $errorCode = "OFFLINE_SYNC_IN_PROGRESS"
                                        $isConflict = $true
                                    }
                                    elseif ($message -match "required" -or $message -match "invalid" -or $message -match "not found") {
                                        $statusCode = 400
                                        $errorCode = "OFFLINE_SYNC_VALIDATION_FAILED"
                                    }

                                    $result = @{
                                        success   = $false
                                        localId   = $idempotencyKey.Trim()
                                        conflict  = $isConflict
                                        errorCode = $errorCode
                                        error     = $message
                                    }
                                    $response.StatusCode = $statusCode
                                }
                            }
                        }
                    }
                }

            } elseif ($path -eq "/busy/voucher/modify" -and $method -eq "POST") {
                $bodyObj = Read-RequestBody $request | ConvertFrom-Json

                if ($authResult.user.name) {
                    $bodyObj | Add-Member -MemberType NoteProperty -Name "bridgeUserName" -Value $authResult.user.name -Force
                }

                $result = Modify-Voucher -Data $bodyObj -InstanceId $instanceId -CompanyCode $companyCode

            } elseif ($path -eq "/busy/voucher" -and $method -eq "DELETE") {
                $body   = Read-RequestBody $request
                $result = Delete-Voucher -Data ($body | ConvertFrom-Json) -InstanceId $instanceId -CompanyCode $companyCode

            } elseif ($path -eq "/busy/vouchers" -and $method -eq "GET") {
                $vchTypeStr = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                if (-not $vchTypeStr) {
                    $result = @{ success=$true; count=0; data=@() }
                } else {
                    $partyVal = Get-QueryStringValue $request.QueryString "party" ""
                    if ($partyVal -eq "") { $partyVal = Get-QueryStringValue $request.QueryString "params[party]" "" }

                    $fromVal = Get-QueryStringValue $request.QueryString "from" ""
                    if ($fromVal -eq "") { $fromVal = Get-QueryStringValue $request.QueryString "params[from]" "" }

                    $toVal = Get-QueryStringValue $request.QueryString "to" ""
                    if ($toVal -eq "") { $toVal = Get-QueryStringValue $request.QueryString "params[to]" "" }

                    $searchVal = Get-QueryStringValue $request.QueryString "search" ""
                    if ($searchVal -eq "") { $searchVal = Get-QueryStringValue $request.QueryString "params[search]" "" }

                    $result = Get-Vouchers `
                        -VchType     ([int]$vchTypeStr) `
                        -From        $fromVal `
                        -To          $toVal `
                        -Party       $partyVal `
                        -Search      $searchVal `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode
                }

            } elseif ($path -eq "/busy/voucher/detail" -and $method -eq "GET") {
                $vchTypeStr = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                $vchNo = Get-QueryStringValue $request.QueryString "vchNo" ""
                if ($vchNo -eq "") { $vchNo = Get-QueryStringValue $request.QueryString "params[vchNo]" "" }

                $vchSeries = Get-QueryStringValue $request.QueryString "vchSeries" "Main"
                if ($vchSeries -eq "Main" -or $vchSeries -eq "") {
                    $pSeries = Get-QueryStringValue $request.QueryString "params[vchSeries]" ""
                    if ($pSeries -ne "") { $vchSeries = $pSeries }
                }

                $vchDate = Get-QueryStringValue $request.QueryString "vchDate" ""
                if ($vchDate -eq "") { $vchDate = Get-QueryStringValue $request.QueryString "params[vchDate]" "" }

                if (-not $vchTypeStr -or $vchNo -eq "" -or $vchDate -eq "") {
                    $result = @{success=$false; error="vchType, vchNo, and vchDate required"}; $response.StatusCode=400
                } else {
                    $result = Get-VoucherDetail `
                        -VchType     ([int]$vchTypeStr) `
                        -VchNo       $vchNo `
                        -VchSeries   $vchSeries `
                        -VchDate     $vchDate `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode
                }

            } elseif ($path -eq "/busy/account-voucher/detail" -and $method -eq "GET") {
                $vchTypeStr = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                $vchNo = Get-QueryStringValue $request.QueryString "vchNo" ""
                if ($vchNo -eq "") { $vchNo = Get-QueryStringValue $request.QueryString "params[vchNo]" "" }

                $vchSeries = Get-QueryStringValue $request.QueryString "vchSeries" "Main"
                if ($vchSeries -eq "Main" -or $vchSeries -eq "") {
                    $pSeries = Get-QueryStringValue $request.QueryString "params[vchSeries]" ""
                    if ($pSeries -ne "") { $vchSeries = $pSeries }
                }

                $vchDate = Get-QueryStringValue $request.QueryString "vchDate" ""
                if ($vchDate -eq "") { $vchDate = Get-QueryStringValue $request.QueryString "params[vchDate]" "" }

                if (-not $vchTypeStr -or $vchNo -eq "" -or $vchDate -eq "") {
                    $result = @{success=$false; error="vchType, vchNo, and vchDate required"}; $response.StatusCode=400
                } else {
                    $result = Get-AccountVoucherDetail `
                        -VchType     ([int]$vchTypeStr) `
                        -VchNo       $vchNo `
                        -VchSeries   $vchSeries `
                        -VchDate     $vchDate `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode
                }

            # --- VOUCHER WORKFLOW (ORDERS & RETURNS) ---
            } elseif ($path -eq "/busy/pending-orders" -and $method -eq "GET") {
                $vchTypeStr = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                $party      = Get-QueryStringValue $request.QueryString "party" ""
                if ($party -eq "") { $party = Get-QueryStringValue $request.QueryString "params[party]" "" }

                if (-not $vchTypeStr -or $party -eq "") {
                    $result = @{ success=$true; count=0; data=@() }
                } else {
                    $result = Get-PendingOrders `
                        -VchType     ([int]$vchTypeStr) `
                        -Party       $party `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode
                }

            } elseif ($path -eq "/busy/pending-challans" -and $method -eq "GET") {
                $vchTypeStr = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                $party      = Get-QueryStringValue $request.QueryString "party" ""
                if ($party -eq "") { $party = Get-QueryStringValue $request.QueryString "params[party]" "" }

                if (-not $vchTypeStr -or $party -eq "") {
                    $result = @{ success=$true; count=0; data=@() }
                } else {
                    $result = Get-PendingChallans `
                        -VchType     ([int]$vchTypeStr) `
                        -Party       $party `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode
                }

            } elseif ($path -eq "/busy/search-for-return" -and $method -eq "GET") {
                # Used by:
                #   - Sale/Purchase return original-voucher search
                #   - Receipt F11 pending bills
                #   - Payment F11 pending bills
                #
                # Supported parent voucher types:
                #   2  = Purchase
                #   3  = Sale Return
                #   9  = Sale
                #   10 = Purchase Return

                $vchTypeStr = Get-QueryStringValue $request.QueryString "vchType" ""
                if ($vchTypeStr -eq "") {
                    $vchTypeStr = Get-QueryStringValue $request.QueryString "params[vchType]" ""
                }

                $parsedVchType = 0
                $validVchType = $false

                if (-not [string]::IsNullOrWhiteSpace($vchTypeStr)) {
                    $validVchType = [int]::TryParse(
                        $vchTypeStr.ToString(),
                        [ref]$parsedVchType
                    )
                }

                if (-not $validVchType -or $parsedVchType -notin @(2, 3, 9, 10)) {
                    $result = @{
                        success = $false
                        error   = "A valid vchType is required. Supported values are 2, 3, 9, and 10."
                        count   = 0
                        data    = @()
                    }
                    $response.StatusCode = 400
                } else {
                    $vchNoVal = Get-QueryStringValue $request.QueryString "vchNo" ""
                    if ($vchNoVal -eq "") {
                        $vchNoVal = Get-QueryStringValue $request.QueryString "params[vchNo]" ""
                    }

                    $partyVal = Get-QueryStringValue $request.QueryString "party" ""
                    if ($partyVal -eq "") {
                        $partyVal = Get-QueryStringValue $request.QueryString "params[party]" ""
                    }

                    $fromVal = Get-QueryStringValue $request.QueryString "from" ""
                    if ($fromVal -eq "") {
                        $fromVal = Get-QueryStringValue $request.QueryString "params[from]" ""
                    }

                    $toVal = Get-QueryStringValue $request.QueryString "to" ""
                    if ($toVal -eq "") {
                        $toVal = Get-QueryStringValue $request.QueryString "params[to]" ""
                    }

                    # Return the complete enriched result from Search-OriginalVouchers.
                    # The response includes:
                    #   transactionType
                    #   parentAmount
                    #   saleReturnedAmount
                    #   purchaseReturnedAmount
                    #   alreadyReceivedAmount
                    #   alreadyPaidAmount
                    #   pendingAmount
                    #   totalAmt / returnedAmt / netAmt compatibility fields
                    $result = Search-OriginalVouchers `
                        -VchType     $parsedVchType `
                        -VchNo       $vchNoVal `
                        -Party       $partyVal `
                        -FromDate    $fromVal `
                        -ToDate      $toVal `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode

                    if ($null -eq $result) {
                        $result = @{
                            success = $false
                            error   = "Search-OriginalVouchers returned no response."
                            count   = 0
                            data    = @()
                        }
                        $response.StatusCode = 500
                    }
                    elseif ($result.success -eq $false) {
                        $response.StatusCode = 500
                    }
                    else {
                        # Ensure predictable response structure without removing
                        # any of the enriched fields returned by vouchers.ps1.
                        if ($null -eq $result.data) {
                            $result.data = @()
                        }

                        if ($null -eq $result.count) {
                            $result.count = @($result.data).Count
                        }
                    }
                }

            } elseif ($path -eq "/busy/return-history" -and $method -eq "GET") {
                $vchTypeStr  = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                $vchCodeStr  = $request.QueryString["vchCode"]
                if (-not $vchCodeStr) { $vchCodeStr = $request.QueryString["params[vchCode]"] }

                if (-not $vchTypeStr -or -not $vchCodeStr) {
                    $result = @{ success=$true; count=0; data=@() }
                } else {
                    $result = Get-ReturnHistory `
                        -OrigVchType  ([int]$vchTypeStr) `
                        -OrigVchCode  ([int]$vchCodeStr) `
                        -InstanceId   $instanceId `
                        -CompanyCode  $companyCode
                }

            # ── DATABASE PERSISTENCE ENDPOINTS ──
            } elseif ($path -eq "/busy/column-config" -and $method -eq "GET") {
                $vchTypeStr = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                $deviceTypeStr = $request.QueryString["deviceType"]
                if (-not $deviceTypeStr) { $deviceTypeStr = $request.QueryString["params[deviceType]"] }

                if (-not $vchTypeStr -or $null -eq $deviceTypeStr) {
                    $result = @{ success = $false; error = "vchType and deviceType required" }
                    $response.StatusCode = 400
                } else {
                    $result = Get-ColumnConfig -VchType ([int]$vchTypeStr) -DeviceType ([int]$deviceTypeStr) -InstanceId $instanceId -CompanyCode $companyCode
                }

            } elseif ($path -eq "/busy/column-config" -and $method -eq "POST") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                $result = Save-ColumnConfig -Data $data -InstanceId $instanceId -CompanyCode $companyCode

            # --- VOUCHER NUMBERING ADMIN CONFIGURATION ---
            } elseif ($path -eq "/busy/voucher-numbering-admin" -and $method -eq "GET") {
                $vchTypeStr = Get-QueryStringValue $request.QueryString "vchType" ""
                if ([string]::IsNullOrWhiteSpace($vchTypeStr)) {
                    $vchTypeStr = Get-QueryStringValue $request.QueryString "params[vchType]" ""
                }

                $seriesName = Get-QueryStringValue $request.QueryString "seriesName" ""
                if ([string]::IsNullOrWhiteSpace($seriesName)) {
                    $seriesName = Get-QueryStringValue $request.QueryString "params[seriesName]" ""
                }

                $voucherDate = Get-QueryStringValue $request.QueryString "voucherDate" ""
                if ([string]::IsNullOrWhiteSpace($voucherDate)) {
                    $voucherDate = Get-QueryStringValue $request.QueryString "params[voucherDate]" ""
                }

                if (
                    [string]::IsNullOrWhiteSpace($vchTypeStr) -or
                    [string]::IsNullOrWhiteSpace($seriesName)
                ) {
                    $result = @{
                        success = $false
                        error   = "vchType and seriesName are required"
                    }
                    $response.StatusCode = 400
                }
                elseif (
                    -not [string]::IsNullOrWhiteSpace($voucherDate) -and
                    -not (
                        $voucherDate -match '^\d{4}-\d{2}-\d{2}$' -or
                        $voucherDate -match '^\d{2}-\d{2}-\d{4}$'
                    )
                ) {
                    $result = @{
                        success = $false
                        error   = "voucherDate must use yyyy-MM-dd or dd-MM-yyyy format"
                    }
                    $response.StatusCode = 400
                }
                else {
                    $result = Get-WebNumberingConfig `
                        -VchType ([int]$vchTypeStr) `
                        -SeriesName $seriesName.Trim() `
                        -VoucherDate $voucherDate `
                        -InstanceId $instanceId `
                        -CompanyCode $companyCode

                    if ($result.success -eq $false) {
                        $response.StatusCode = 400
                    }
                }

            } elseif ($path -eq "/busy/voucher-numbering-admin" -and $method -eq "POST") {
                $normalRole = ""

                if (
                    $requireAuth -and
                    $null -ne $authResult -and
                    $null -ne $authResult.user
                ) {
                    $normalRole = ([string]$authResult.user.role).Trim().ToLower()
                    $normalRole = $normalRole.Replace(" ", "").Replace("_", "").Replace("-", "")
                }

                if ($normalRole -ne "superadmin") {
                    $result = @{
                        success = $false
                        error   = "Only Super Admin can change voucher numbering configuration."
                    }
                    $response.StatusCode = 403
                }
                else {
                    $rawBody = Read-RequestBody $request

                    if ([string]::IsNullOrWhiteSpace($rawBody)) {
                        $result = @{
                            success = $false
                            error   = "Voucher numbering configuration payload is required."
                        }
                        $response.StatusCode = 400
                    }
                    else {
                        $data = $null

                        try {
                            $data = $rawBody | ConvertFrom-Json
                        }
                        catch {
                            $result = @{
                                success = $false
                                error   = "The request body is not valid JSON."
                            }
                            $response.StatusCode = 400
                        }

                        if ($null -ne $data) {
                            $effectiveInstanceId = $instanceId
                            $effectiveCompanyCode = $companyCode

                            if (
                                [string]::IsNullOrWhiteSpace($effectiveInstanceId) -and
                                $data.instance_id
                            ) {
                                $effectiveInstanceId = [string]$data.instance_id
                            }

                            if (
                                [string]::IsNullOrWhiteSpace($effectiveCompanyCode) -and
                                $data.company_code
                            ) {
                                $effectiveCompanyCode = [string]$data.company_code
                            }

                            $dateBasis = "VOUCHER_DATE"
                            $dateBasisProperty = $data.PSObject.Properties["date_basis"]

                            if ($null -ne $dateBasisProperty) {
                                $candidateDateBasis = ([string]$dateBasisProperty.Value).Trim().ToUpperInvariant()

                                if (
                                    $candidateDateBasis -notin @(
                                        "VOUCHER_DATE",
                                        "REAL_TIME"
                                    )
                                ) {
                                    $result = @{
                                        success = $false
                                        error   = "date_basis must be VOUCHER_DATE or REAL_TIME."
                                    }
                                    $response.StatusCode = 400
                                }
                                else {
                                    $dateBasis = $candidateDateBasis
                                }
                            }

                            if ($null -eq $result) {
                                $data |
                                    Add-Member `
                                        -MemberType NoteProperty `
                                        -Name "date_basis" `
                                        -Value $dateBasis `
                                        -Force

                                $updatedBy = ""

                                if (
                                    $null -ne $authResult -and
                                    $null -ne $authResult.user
                                ) {
                                    $updatedBy = [string]$authResult.user.name
                                }

                                $result = Save-WebNumberingConfig `
                                    -Data $data `
                                    -UpdatedBy $updatedBy `
                                    -InstanceId $effectiveInstanceId `
                                    -CompanyCode $effectiveCompanyCode

                                if ($result.success -eq $false) {
                                    $response.StatusCode = 400
                                }
                            }
                        }
                    }
                }

            # --- USER PERMISSIONS (db.bds OLEDB Integration) ---
            } elseif ($path -eq "/busy/permissions" -and $method -eq "GET") {
                $result = Get-UserPermissions -InstanceId $instanceId -CompanyCode $companyCode

            } elseif ($path -eq "/busy/permissions/my" -and $method -eq "GET") {
                $userRes = Get-UserPermissions -InstanceId $instanceId -CompanyCode $companyCode
                if ($userRes.success) {
                    $activeName = $authResult.user.name
                    $myPerm = $userRes.data | Where-Object { $_.name.Trim().ToLower() -eq $activeName.Trim().ToLower() }
                    if ($myPerm) {
                        $result = @{ success = $true; data = $myPerm }
                    } else {
                        $result = @{ success = $true; data = @{ name = $activeName; C1=1;C2=1;C3=1;C4=1;C5=1;C6=1;C7=1;C8=1;C9=1;C10=1;I1=1;I2=1;I3=1;I4=1;I5=1;I6=1;I7=1;I8=1;I9=1;I10=1;I11=1;I12=1;I13=1;I14=1;M1="{}";M2="{}" } }
                    }
                } else {
                    $result = $userRes
                }

            } elseif ($path -eq "/busy/permissions/save" -and $method -eq "POST") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                $result = Save-UserPermissions -Data $data -InstanceId $instanceId -CompanyCode $companyCode

                if ($result.success -eq $false) {
                    Send-Response $response $result 500
                    continue
                }

            } elseif ($path -eq "/busy/users" -and $method -eq "GET") {
                $result = Get-CompanyUsers -InstanceId $instanceId -CompanyCode $companyCode

            # --- MASTER DATA ---

            # --- PRODUCTION BOM MASTER DATA ---
            } elseif ($path -eq "/busy/boms" -and $method -eq "GET") {
                $searchVal = Get-QueryStringValue $request.QueryString "search" ""
                if ($searchVal -eq "") {
                    $searchVal = Get-QueryStringValue $request.QueryString "params[search]" ""
                }

                $result = Get-BomList `
                    -Search       $searchVal `
                    -InstanceId   $instanceId `
                    -CompanyCode  $companyCode

            } elseif ($path -eq "/busy/bom/detail" -and $method -eq "GET") {
                $bomCodeStr = Get-QueryStringValue $request.QueryString "code" ""
                if ($bomCodeStr -eq "") {
                    $bomCodeStr = Get-QueryStringValue $request.QueryString "params[code]" ""
                }

                $bomName = Get-QueryStringValue $request.QueryString "name" ""
                if ($bomName -eq "") {
                    $bomName = Get-QueryStringValue $request.QueryString "params[name]" ""
                }

                $bomCode = 0
                $validBomCode = $false

                if (-not [string]::IsNullOrWhiteSpace($bomCodeStr)) {
                    $validBomCode = [int]::TryParse(
                        $bomCodeStr.ToString(),
                        [ref]$bomCode
                    )
                }

                if (
                    (-not $validBomCode -or $bomCode -le 0) -and
                    [string]::IsNullOrWhiteSpace($bomName)
                ) {
                    $result = @{
                        success = $false
                        error   = "BOM code or BOM name is required"
                    }
                    $response.StatusCode = 400
                } else {
                    $result = Get-BomDetail `
                        -BomCode     $bomCode `
                        -BomName     $bomName `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode
                }

            } elseif ($path -eq "/busy/bom" -and $method -eq "POST") {
                $data = Read-RequestBody $request | ConvertFrom-Json

                if (
                    -not $data.name -or
                    -not ($data.item -or $data.mainItemName) -or
                    -not ($data.unit -or $data.mainUnit)
                ) {
                    $result = @{
                        success = $false
                        error   = "BOM name, Item to Produce and Unit are required"
                    }
                    $response.StatusCode = 400
                }
                else {
                    $result = Create-Bom `
                        -Data        $data `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode

                    if ($result.success -eq $false) {
                        $response.StatusCode = 400
                    }
                }

            } elseif ($path -eq "/busy/bom" -and $method -eq "PUT") {
                $data = Read-RequestBody $request | ConvertFrom-Json

                if (
                    -not $data.name -or
                    -not ($data.item -or $data.mainItemName) -or
                    -not ($data.unit -or $data.mainUnit)
                ) {
                    $result = @{
                        success = $false
                        error   = "BOM name, Item to Produce and Unit are required"
                    }
                    $response.StatusCode = 400
                }
                else {
                    $result = Update-Bom `
                        -Data        $data `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode

                    if ($result.success -eq $false) {
                        $response.StatusCode = 400
                    }
                }

            } elseif ($path -eq "/busy/account-groups" -and $method -eq "GET") {
                $result = Get-AccountGroups -InstanceId $instanceId -CompanyCode $companyCode
            } elseif ($path -eq "/busy/account-group" -and $method -eq "POST") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                if (-not $data.name) { $result = @{success=$false;error="name required"}; $response.StatusCode=400 }
                else {
                    $pGrp = if ($data.parentGroup) { $data.parentGroup } else { "Primary" }
                    $result = Create-AccountGroup -Name $data.name -ParentGroup $pGrp -InstanceId $instanceId -CompanyCode $companyCode
                }
            } elseif ($path -eq "/busy/account-group" -and $method -eq "PUT") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                if (-not $data.name) { $result = @{success=$false;error="name required"}; $response.StatusCode=400 }
                else {
                    $pGrp = if ($data.parentGroup) { $data.parentGroup } else { "Primary" }
                    $result = Update-AccountGroup -Name $data.name -ParentGroup $pGrp -InstanceId $instanceId -CompanyCode $companyCode
                }
            } elseif ($path -eq "/busy/accounts" -and $method -eq "GET") {
                $pageVal = Get-QueryStringValue $request.QueryString "page" "1"
                if ($pageVal -eq "") { $pageVal = Get-QueryStringValue $request.QueryString "params[page]" "1" }

                $pageSizeVal = Get-QueryStringValue $request.QueryString "pageSize" "30"
                if ($pageSizeVal -eq "") { $pageSizeVal = Get-QueryStringValue $request.QueryString "params[pageSize]" "30" }

                $groupVal = $request.QueryString["group"]
                if ($groupVal -eq "") { $groupVal = Get-QueryStringValue $request.QueryString "params[group]" "" }

                $searchVal = Get-QueryStringValue $request.QueryString "search" ""
                if ($searchVal -eq "") { $searchVal = Get-QueryStringValue $request.QueryString "params[search]" "" }

                $result = Get-Accounts `
                    -GroupName   $groupVal `
                    -Search      $searchVal `
                    -Page        ([int]$pageVal) `
                    -PageSize    ([int]$pageSizeVal) `
                    -InstanceId  $instanceId `
                    -CompanyCode $companyCode

            } elseif ($path -eq "/busy/account" -and $method -eq "GET") {
                $name = $request.QueryString["name"]
                if (-not $name) { $result = @{success=$false;error="name required"}; $response.StatusCode=400 } else { $result = Get-AccountDetail -Name $name -InstanceId $instanceId -CompanyCode $companyCode }
            } elseif ($path -eq "/busy/account" -and $method -eq "POST") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                if (-not $data.name -or -not $data.group) { $result = @{success=$false;error="name and group required"}; $response.StatusCode=400 } else { $result = Create-Account -Data $data -InstanceId $instanceId -CompanyCode $companyCode }
            } elseif ($path -eq "/busy/account" -and $method -eq "PUT") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                if (-not $data.name -or -not $data.group) { $result = @{success=$false;error="name and group required"}; $response.StatusCode=400 } else { $result = Update-Account -Data $data -InstanceId $instanceId -CompanyCode $companyCode }

} elseif ($path -eq "/busy/reports/outstanding" -and $method -eq "GET") {
                $fromVal = Get-QueryStringValue $request.QueryString "from" ""
                $toVal = Get-QueryStringValue $request.QueryString "to" ""
                $asOfVal = Get-QueryStringValue $request.QueryString "asOf" ""

                $typeVal = Get-QueryStringValue $request.QueryString "type" "all"
                $accountVal = Get-QueryStringValue $request.QueryString "account" ""
                $searchVal = Get-QueryStringValue $request.QueryString "search" ""
                $groupVal = Get-QueryStringValue $request.QueryString "group" ""
                $statusVal = Get-QueryStringValue $request.QueryString "status" "all"
                $agingVal = Get-QueryStringValue $request.QueryString "aging" "all"

                $voucherTypeVal = Get-QueryStringValue $request.QueryString "voucherType" "0"
                $minAmountVal = Get-QueryStringValue $request.QueryString "minAmount" "0"
                $maxAmountVal = Get-QueryStringValue $request.QueryString "maxAmount" "0"
                $includeZeroVal = Get-QueryStringValue $request.QueryString "includeZero" "false"

                $pageVal = Get-QueryStringValue $request.QueryString "page" "1"
                $pageSizeVal = Get-QueryStringValue $request.QueryString "pageSize" "50"
                $sortByVal = Get-QueryStringValue $request.QueryString "sortBy" "dueDate"
                $sortDirectionVal = Get-QueryStringValue $request.QueryString "sortDirection" "asc"

                $safePage = 1
                $safePageSize = 50
                $safeVoucherType = 0
                $safeMinAmount = 0.0
                $safeMaxAmount = 0.0

                try {
                    $safePage = [Math]::Max(1, [int]$pageVal)
                }
                catch {
                    $safePage = 1
                }

                try {
                    $requestedPageSize = [int]$pageSizeVal

                    if ($requestedPageSize -eq 0) {
                        # pageSize=0 means "All matching rows".
                        $safePageSize = 0
                    }
                    else {
                        $safePageSize = [Math]::Min(
                            2500,
                            [Math]::Max(1, $requestedPageSize)
                        )
                    }
                }
                catch {
                    $safePageSize = 50
                }

                try {
                    $safeVoucherType = [Math]::Max(
                        0,
                        [int]$voucherTypeVal
                    )
                }
                catch {
                    $safeVoucherType = 0
                }

                # The report supports only:
                # 0  = All Sale/Purchase transaction types
                # 2  = Purchase
                # 3  = Sale Return
                # 9  = Sale
                # 10 = Purchase Return
                if ($safeVoucherType -notin @(0, 2, 3, 9, 10)) {
                    $safeVoucherType = 0
                }

                try {
                    $safeMinAmount = [Math]::Max(
                        0,
                        [double]$minAmountVal
                    )
                }
                catch {
                    $safeMinAmount = 0.0
                }

                try {
                    $safeMaxAmount = [Math]::Max(
                        0,
                        [double]$maxAmountVal
                    )
                }
                catch {
                    $safeMaxAmount = 0.0
                }

                if (
                    $safeMaxAmount -gt 0 -and
                    $safeMinAmount -gt $safeMaxAmount
                ) {
                    $temporaryAmount = $safeMinAmount
                    $safeMinAmount = $safeMaxAmount
                    $safeMaxAmount = $temporaryAmount
                }

                $validTypes = @(
                    "all",
                    "receivable",
                    "payable"
                )

                if ($typeVal -notin $validTypes) {
                    $typeVal = "all"
                }

                $validStatuses = @(
                    "all",
                    "due",
                    "overdue",
                    "not-due",
                    "partially-adjusted",
                    "unadjusted"
                )

                if ($statusVal -notin $validStatuses) {
                    $statusVal = "all"
                }

                $validAgingValues = @(
                    "all",
                    "notDue",
                    "days0To30",
                    "days31To60",
                    "days61To90",
                    "days91To180",
                    "above180"
                )

                if ($agingVal -notin $validAgingValues) {
                    $agingVal = "all"
                }

                $validSortFields = @(
                    "dueDate",
                    "refDate",
                    "account",
                    "pending",
                    "daysOverdue"
                )

                if ($sortByVal -notin $validSortFields) {
                    $sortByVal = "dueDate"
                }

                if ($sortDirectionVal -notin @("asc", "desc")) {
                    $sortDirectionVal = "asc"
                }

                $includeZero = (
                    $includeZeroVal -eq "true" -or
                    $includeZeroVal -eq "1"
                )

                Write-Host (
                    "  [OUTSTANDING-ROUTE] " +
                    "from='$fromVal'; to='$toVal'; asOf='$asOfVal'; " +
                    "type='$typeVal'; status='$statusVal'; aging='$agingVal'; " +
                    "account='$accountVal'; search='$searchVal'; group='$groupVal'; " +
                    "voucherType=$safeVoucherType; min=$safeMinAmount; " +
                    "max=$safeMaxAmount; includeZero=$includeZero; " +
                    "page=$safePage; pageSize=$safePageSize; " +
                    "sortBy='$sortByVal'; direction='$sortDirectionVal'"
                ) -ForegroundColor DarkCyan

                $result = Get-OutstandingReport `
                    -From $fromVal `
                    -To $toVal `
                    -AsOf $asOfVal `
                    -Type $typeVal `
                    -Account $accountVal `
                    -Search $searchVal `
                    -Group $groupVal `
                    -Status $statusVal `
                    -Aging $agingVal `
                    -VoucherType $safeVoucherType `
                    -MinAmount $safeMinAmount `
                    -MaxAmount $safeMaxAmount `
                    -IncludeZero $includeZero `
                    -Page $safePage `
                    -PageSize $safePageSize `
                    -SortBy $sortByVal `
                    -SortDirection $sortDirectionVal `
                    -InstanceId $instanceId `
                    -CompanyCode $companyCode

            } elseif ($path -eq "/busy/reports/stock-status" -and $method -eq "GET") {
                $asOfVal = Get-QueryStringValue $request.QueryString "asOf" ""
                $materialCentreVal = Get-QueryStringValue $request.QueryString "materialCentre" ""
                $itemGroupVal = Get-QueryStringValue $request.QueryString "itemGroup" ""
                $searchVal = Get-QueryStringValue $request.QueryString "search" ""
                $statusVal = Get-QueryStringValue $request.QueryString "status" "all"
                $includeZeroVal = Get-QueryStringValue $request.QueryString "includeZero" "true"
                $lowStockLevelVal = Get-QueryStringValue $request.QueryString "lowStockLevel" "5"
                $valueByVal = Get-QueryStringValue $request.QueryString "valueBy" "purchase"
                $pageVal = Get-QueryStringValue $request.QueryString "page" "1"
                $pageSizeVal = Get-QueryStringValue $request.QueryString "pageSize" "100"

                $result = Get-StockStatusReport `
                    -AsOf           $asOfVal `
                    -MaterialCentre $materialCentreVal `
                    -ItemGroup      $itemGroupVal `
                    -Search         $searchVal `
                    -Status         $statusVal `
                    -IncludeZero    ($includeZeroVal -eq "true" -or $includeZeroVal -eq "1") `
                    -LowStockLevel  ([double]$lowStockLevelVal) `
                    -ValueBy        $valueByVal `
                    -Page           ([int]$pageVal) `
                    -PageSize       ([int]$pageSizeVal) `
                    -InstanceId     $instanceId `
                    -CompanyCode    $companyCode

            } elseif ($path -eq "/busy/item-groups" -and $method -eq "GET") {
                $result = Get-ItemGroups -InstanceId $instanceId -CompanyCode $companyCode
            } elseif ($path -eq "/busy/item-group" -and $method -eq "POST") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                if (-not $data.name) { $result = @{success=$false;error="name required"}; $response.StatusCode=400 } else { $result = Create-ItemGroup -Data $data -InstanceId $instanceId -CompanyCode $companyCode }
            } elseif ($path -eq "/busy/item-group" -and $method -eq "PUT") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                if (-not $data.name) { $result = @{success=$false;error="name required"}; $response.StatusCode=400 } else { $result = Update-ItemGroup -Data $data -InstanceId $instanceId -CompanyCode $companyCode }

            } elseif ($path -eq "/busy/items" -and $method -eq "GET") {
                $pageVal = Get-QueryStringValue $request.QueryString "page" "1"
                if ($pageVal -eq "") { $pageVal = Get-QueryStringValue $request.QueryString "params[page]" "1" }

                $pageSizeVal = Get-QueryStringValue $request.QueryString "pageSize" "30"
                if ($pageSizeVal -eq "") { $pageSizeVal = Get-QueryStringValue $request.QueryString "params[pageSize]" "30" }

                $catVal = Get-QueryStringValue $request.QueryString "category" ""
                if ($catVal -eq "") { $catVal = Get-QueryStringValue $request.QueryString "params[category]" "" }

                $searchVal = Get-QueryStringValue $request.QueryString "search" ""
                if ($searchVal -eq "") { $searchVal = Get-QueryStringValue $request.QueryString "params[search]" "" }

                $result = Get-Items `
                    -Category    $catVal `
                    -Search      $searchVal `
                    -Page        ([int]$pageVal) `
                    -PageSize    ([int]$pageSizeVal) `
                    -InstanceId  $instanceId `
                    -CompanyCode $companyCode

            } elseif ($path -eq "/busy/vouchers/items" -and $method -eq "GET") {
                $searchVal = Get-QueryStringValue $request.QueryString "search" ""
                if ($searchVal -eq "") { $searchVal = Get-QueryStringValue $request.QueryString "params[search]" "" }

                $result = Get-ItemsForVoucher `
                    -Search      $searchVal `
                    -InstanceId  $instanceId `
                    -CompanyCode $companyCode

            } elseif ($path -eq "/busy/vouchers/item-detail" -and $method -eq "GET") {
                $codeStr = Get-QueryStringValue $request.QueryString "code" ""
                if ($codeStr -eq "") {
                    $codeStr = Get-QueryStringValue $request.QueryString "params[code]" ""
                }

                $itemCode = 0
                $isValidCode = $false

                if (-not [string]::IsNullOrWhiteSpace($codeStr)) {
                    $isValidCode = [int]::TryParse(
                        $codeStr.ToString(),
                        [ref]$itemCode
                    )
                }

                if (-not $isValidCode -or $itemCode -le 0) {
                    $result = @{
                        success = $false
                        error   = "Valid item code is required"
                    }
                    $response.StatusCode = 400
                } else {
                    $result = Get-VoucherItemDetail `
                        -Code        $itemCode `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode
                }

            } elseif ($path -eq "/busy/item" -and $method -eq "GET") {
                $codeStr = $request.QueryString["code"]
                if (-not $codeStr) { $result = @{success=$false;error="code required"}; $response.StatusCode=400 } else { $result = Get-ItemByCode -Code ([int]$codeStr) -InstanceId $instanceId -CompanyCode $companyCode }
            } elseif ($path -eq "/busy/item/detail" -and $method -eq "GET") {
                $codeStr = $request.QueryString["code"]
                if (-not $codeStr) { $result = @{success=$false;error="code required"}; $response.StatusCode=400 } else { $result = Get-ItemDetail -Code ([int]$codeStr) -InstanceId $instanceId -CompanyCode $companyCode }
            } elseif ($path -eq "/busy/item" -and $method -eq "POST") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                if (-not $data.name -or -not $data.group) { $result = @{success=$false;error="name and group required"}; $response.StatusCode=400 } else { $result = Create-Item -Data $data -InstanceId $instanceId -CompanyCode $companyCode }
            } elseif ($path -eq "/busy/item" -and $method -eq "PUT") {
                $data = Read-RequestBody $request | ConvertFrom-Json
                if (-not $data.name -or -not $data.group) { $result = @{success=$false;error="name and group required"}; $response.StatusCode=400 } else { $result = Update-Item -Data $data -InstanceId $instanceId -CompanyCode $companyCode }

            } elseif ($path -eq "/busy/parties" -and $method -eq "GET") {
                $searchVal = Get-QueryStringValue $request.QueryString "search" ""
                if ($searchVal -eq "") { $searchVal = Get-QueryStringValue $request.QueryString "params[search]" "" }

                $pageVal = Get-QueryStringValue $request.QueryString "page" "1"
                if ($pageVal -eq "") { $pageVal = Get-QueryStringValue $request.QueryString "params[page]" "1" }

                $pageSizeVal = Get-QueryStringValue $request.QueryString "pageSize" "30"
                if ($pageSizeVal -eq "") { $pageSizeVal = Get-QueryStringValue $request.QueryString "params[pageSize]" "30" }

                $cashBankVal = Get-QueryStringValue $request.QueryString "cashBankOnly" "false"
                if ($cashBankVal -eq "") { $cashBankVal = Get-QueryStringValue $request.QueryString "params[cashBankOnly]" "false" }
                $isCashBankOnly = ($cashBankVal -eq "true" -or $cashBankVal -eq "1")

                $result = Get-Parties `
                    -Search       $searchVal `
                    -CashBankOnly $isCashBankOnly `
                    -Page         ([int]$pageVal) `
                    -PageSize     ([int]$pageSizeVal) `
                    -InstanceId   $instanceId `
                    -CompanyCode  $companyCode

            } elseif ($path -eq "/busy/cash-bank-accounts" -and $method -eq "GET") {
                $searchVal = Get-QueryStringValue $request.QueryString "search" ""
                if ($searchVal -eq "") { $searchVal = Get-QueryStringValue $request.QueryString "params[search]" "" }

                $result = Get-CashBankAccounts `
                    -Search      $searchVal `
                    -InstanceId  $instanceId `
                    -CompanyCode $companyCode
             }elseif ($path -eq "/busy/units" -and $method -eq "GET") {
                $result = Get-Units -InstanceId $instanceId -CompanyCode $companyCode
            } elseif ($path -eq "/busy/tax-categories" -and $method -eq "GET") {
                $result = Get-TaxCategories -InstanceId $instanceId -CompanyCode $companyCode
            } elseif ($path -eq "/busy/bill-sundries" -and $method -eq "GET") {
                $result = Get-BillSundries -InstanceId $instanceId -CompanyCode $companyCode
            } elseif ($path -eq "/busy/material-centers" -and $method -eq "GET") {
                $result = Get-MaterialCenters -InstanceId $instanceId -CompanyCode $companyCode

            # --- MISC UTILITIES & CONFIGURATIONS ---
            } elseif ($path -eq "/busy/numbering-config" -and $method -eq "GET") {
                $vchTypeStr = Get-QueryStringValue $request.QueryString "vchType" ""
                if ([string]::IsNullOrWhiteSpace($vchTypeStr)) {
                    $vchTypeStr = Get-QueryStringValue $request.QueryString "params[vchType]" ""
                }

                $seriesName = Get-QueryStringValue $request.QueryString "seriesName" ""
                if ([string]::IsNullOrWhiteSpace($seriesName)) {
                    $seriesName = Get-QueryStringValue $request.QueryString "params[seriesName]" ""
                }

                $voucherDate = Get-QueryStringValue $request.QueryString "voucherDate" ""
                if ([string]::IsNullOrWhiteSpace($voucherDate)) {
                    $voucherDate = Get-QueryStringValue $request.QueryString "params[voucherDate]" ""
                }

                if (
                    [string]::IsNullOrWhiteSpace($vchTypeStr) -or
                    [string]::IsNullOrWhiteSpace($seriesName)
                ) {
                    $result = @{
                        success = $false
                        error   = "vchType and seriesName required"
                    }
                    $response.StatusCode = 400
                } else {
                    $result = Get-EffectiveNumberingConfig `
                        -VchType ([int]$vchTypeStr) `
                        -SeriesName $seriesName.Trim() `
                        -VoucherDate $voucherDate `
                        -InstanceId $instanceId `
                        -CompanyCode $companyCode

                    if ($result.success -eq $false) {
                        $response.StatusCode = 400
                    }
                }

            } elseif ($path -eq "/busy/voucher-series" -and $method -eq "GET") {
                $vchTypeStr = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                $vchType = 0
                if ($vchTypeStr) { try { $vchType = [int]$vchTypeStr } catch {} }
                $result = Get-VoucherSeries -VchType $vchType -InstanceId $instanceId -CompanyCode $companyCode
            } elseif ($path -eq "/busy/voucher-types" -and $method -eq "GET") {
                $typeParam = Get-QueryStringValue $request.QueryString "type" "All"
                $result = Get-VoucherTypes -Type $typeParam -InstanceId $instanceId -CompanyCode $companyCode
            } elseif ($path -eq "/busy/input-types" -and $method -eq "GET") {
                $result = Get-VoucherInputTypes
            } elseif ($path -eq "/busy/cache/clear" -and $method -eq "POST") {
                Clear-Cache
                $result = @{success=$true; message="All caches cleared"}

            # --- OPTIONAL FIELDS ROUTING ---
            } elseif ($path -eq "/busy/voucher/optional-fields-config" -and $method -eq "GET") {
                $vchTypeStr = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                $seriesName = Get-QueryStringValue $request.QueryString "seriesName" ""
                if ($seriesName -eq "") { $seriesName = Get-QueryStringValue $request.QueryString "params[seriesName]" "" }

                if (-not $vchTypeStr -or $seriesName -eq "") {
                    $result = @{ success = $false; error = "vchType and seriesName required" }
                    $response.StatusCode = 400
                } else {
                    $result = Get-VoucherOptionalFields `
                        -VchType     ([int]$vchTypeStr) `
                        -SeriesName  $seriesName `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode
                }

            } elseif ($path -eq "/busy/voucher/optional-fields-values" -and $method -eq "GET") {
                $vchTypeStr = $request.QueryString["vchType"]
                if (-not $vchTypeStr) { $vchTypeStr = $request.QueryString["params[vchType]"] }

                $seriesName = Get-QueryStringValue $request.QueryString "seriesName" ""
                if ($seriesName -eq "") { $seriesName = Get-QueryStringValue $request.QueryString "params[seriesName]" "" }

                $fieldKeyStr = Get-QueryStringValue $request.QueryString "fieldNo" ""
                if ($fieldKeyStr -eq "") { $fieldKeyStr = Get-QueryStringValue $request.QueryString "params[fieldNo]" "" }
                
                $fieldNo = 1
                if ($fieldKeyStr -match "OptionField(\d+)") {
                    $fieldNo = [int]$Matches[1]
                } elseif ($fieldKeyStr -ne "") {
                    $fieldNo = [int]$fieldKeyStr
                }

                if (-not $vchTypeStr -or $seriesName -eq "" -or $fieldNo -eq 0) {
                    $result = @{ success = $false; error = "vchType, seriesName, and fieldNo required" }
                    $response.StatusCode = 400
                } else {
                    $result = Get-OptionalFieldMasterValues `
                        -VchType     ([int]$vchTypeStr) `
                        -SeriesName  $seriesName `
                        -FieldNo     $fieldNo `
                        -InstanceId  $instanceId `
                        -CompanyCode $companyCode
                }
            } else {
                $result = @{success=$false; error="Endpoint not found: $method $path"}
                $response.StatusCode = 404
            }

            Send-Response $response $result

            if ($result -and $result.success -eq $true) { Write-Host "[OK] $path" -ForegroundColor Green }
            else { Write-Host "  [FAIL] $($result.error)" -ForegroundColor Red }

        } catch {
            Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
            Send-Response $response @{success=$false; error=$_.Exception.Message} 500
        }
    }
}