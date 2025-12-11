# param(
#     [string]$TenantId,
#     [string]$ClientId,
#     [string]$ClientSecret,
#     [string]$UserPrincipalName,
#     [string]$LicenseSkuId,
#     [string]$ActionType   # add OR remove
# )



# Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

# $body = @{
#     grant_type    = "client_credentials"
#     client_id     = $ClientId
#     client_secret = $ClientSecret
#     scope         = "https://graph.microsoft.com/.default"
# }

# try {
#     $tokenResponse = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $body
#     $token = $tokenResponse.access_token
# }
# catch {
#     Write-Host ("Failed to get access token: " + $_.Exception.Message) -ForegroundColor Red
#     exit 1
# }

# $uri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/assignLicense"
# $headers = @{
#     "Authorization" = "Bearer $token"
#     "Content-Type"  = "application/json"
# }

# if ($ActionType -eq "add") {
#     Write-Host "Adding license $LicenseSkuId to $UserPrincipalName..." -ForegroundColor Cyan

#     $bodyJson = @{
#         addLicenses = @(
#             @{ skuId = $LicenseSkuId }
#         )
#         removeLicenses = @()
#     } | ConvertTo-Json -Depth 3
# }
# elseif ($ActionType -eq "remove") {
#     Write-Host "Removing license $LicenseSkuId from $UserPrincipalName..." -ForegroundColor Cyan

#     $bodyJson = @{
#         addLicenses = @()
#         removeLicenses = @(
#             "$LicenseSkuId"
#         )
#     } | ConvertTo-Json -Depth 3
# }
# else {
#     Write-Host "❌ Invalid ActionType. Use: add OR remove" -ForegroundColor Red
#     exit 1
# }

# try {
#     Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $bodyJson

#     if ($ActionType -eq "add") {
#         Write-Host "✅ SUCCESS: License added!" -ForegroundColor Green
#     }
#     else {
#         Write-Host "✅ SUCCESS: License removed!" -ForegroundColor Green
#     }
# }
# catch {
#     Write-Host "❌ API failed" -ForegroundColor Red
#     Write-Host $_.Exception.Message -ForegroundColor Yellow
#     exit 1
# }


## working one
<#

param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret,
    [string]$UserPrincipalName,
    [string]$LicenseSkuId,
    [string]$ActionType   # add license OR remove license
)

# -----------------------------
# Normalize ActionType
# -----------------------------
$ActionType = $ActionType.ToLower().Trim()
if ($ActionType -eq "add license") { $actionNormalized = "add" }
elseif ($ActionType -eq "remove license") { $actionNormalized = "remove" }
else {
    Write-Host "❌ ERROR: Invalid ActionType. Use 'add license' or 'remove license'." -ForegroundColor Red
    Write-Host "WORKNOTES:: Invalid ActionType"
    Write-Host "ERRORFLAG:: True"
   return
}

# -----------------------------
# Validate UPN Format
# -----------------------------
$validDomain = $env:VALID_DOMAIN
if ([string]::IsNullOrWhiteSpace($validDomain)) {
    Write-Host "❌ ERROR: VALID_DOMAIN environment variable is not set!" -ForegroundColor Red
    Write-Host "WORKNOTES:: VALID_DOMAIN not set"
    Write-Host "ERRORFLAG:: True"
   return
}

$escapedDomain = [Regex]::Escape($validDomain)

if ($UserPrincipalName -notmatch "^[a-zA-Z0-9._-]+@$escapedDomain$") {
    Write-Host "❌ ERROR: Invalid UPN format. Only allowed chars: a-z, A-Z, 0-9, ., _, -" -ForegroundColor Red
    Write-Host "WORKNOTES:: Invalid UPN format"
    Write-Host "ERRORFLAG:: True"
    return
}

Write-Host "✔ UPN Format Validated" -ForegroundColor Green

# -----------------------------
# Get Access Token
# -----------------------------
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
$body = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "https://graph.microsoft.com/.default"
}

try {
    $tokenResponse = Invoke-RestMethod -Method POST `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -Body $body
    $token = $tokenResponse.access_token
}
catch {
    Write-Host "❌ Failed to get access token." -ForegroundColor Red
    Write-Host "WORKNOTES:: Failed to get access token"
    Write-Host "ERRORFLAG:: True"
    return
}

$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
}

# -----------------------------
# Get User
# -----------------------------
try {
    $user = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/users/$UserPrincipalName" `
        -Headers $headers `
        -Method GET
    Write-Host "✔ User Found: $($user.userPrincipalName)" -ForegroundColor Green
}
catch {
    Write-Host "❌ ERROR: User not found: $UserPrincipalName" -ForegroundColor Red
    Write-Host "WORKNOTES:: User not found"
    Write-Host "ERRORFLAG:: True"
    return
}

# -----------------------------
# Get Subscribed SKUs
# -----------------------------
try {
    $tenantLicenses = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" `
        -Headers $headers `
        -Method GET
}
catch {
    Write-Host "❌ ERROR: Unable to fetch subscribed SKUs" -ForegroundColor Red
    Write-Host "WORKNOTES:: Failed to fetch tenant SKUs"
    Write-Host "ERRORFLAG:: True"
    return
}

$sku = $tenantLicenses.value | Where-Object { $_.skuId -eq $LicenseSkuId }

if (-not $sku) {
    Write-Host "❌ ERROR: License SKU '$LicenseSkuId' does NOT exist in tenant." -ForegroundColor Red
    Write-Host "WORKNOTES:: License SKU not found"
    Write-Host "ERRORFLAG:: True"
    return
}

Write-Host "✔ License SKU exists: $LicenseSkuId" -ForegroundColor Green

# -----------------------------
# Get License Details
# -----------------------------
try {
    $licenseDetails = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/licenseDetails" `
        -Headers $headers `
        -Method GET
}
catch {
    Write-Host "❌ ERROR: Unable to fetch license details." -ForegroundColor Red
    Write-Host "WORKNOTES:: Failed to fetch license details"
    Write-Host "ERRORFLAG:: True"
    return
}

$assignedSkuIds = $licenseDetails.value.skuId

# Show user licenses
Write-Host ""
Write-Host "========== LICENSES ASSIGNED TO USER ==========" -ForegroundColor Cyan
foreach ($l in $licenseDetails.value) {
    Write-Host "SKU: $($l.skuId)  |  Name: $($l.skuPartNumber)" -ForegroundColor Yellow
}
Write-Host "================================================"
Write-Host ""

# -----------------------------
# ADD LICENSE
# -----------------------------
if ($actionNormalized -eq "add") {

    if ($assignedSkuIds -contains $LicenseSkuId) {
        Write-Host "⚠️ SKIP: License is already assigned to this user." -ForegroundColor Yellow
        Write-Host "WORKNOTES:: License already assigned"
        Write-Host "ERRORFLAG:: True"
        exit 0
    }

    $availableUnits = $sku.prepaidUnits.enabled - $sku.consumedUnits
    if ($availableUnits -le 0) {
        Write-Host "❌ ERROR: No available units left for this license." -ForegroundColor Red
        Write-Host "WORKNOTES:: No available license units"
        Write-Host "ERRORFLAG:: True"
        return
    }

    Write-Host "Adding license $LicenseSkuId to $UserPrincipalName..." -ForegroundColor Cyan
    $bodyJson = @{
        addLicenses    = @(@{ skuId = $LicenseSkuId })
        removeLicenses = @()
    } | ConvertTo-Json -Depth 3
}

# -----------------------------
# REMOVE LICENSE
# -----------------------------
if ($actionNormalized -eq "remove") {

    if ($assignedSkuIds -notcontains $LicenseSkuId) {
        Write-Host "⚠️ SKIP: License is NOT assigned. Nothing to remove." -ForegroundColor Yellow
        Write-Host "WORKNOTES:: License not found for removal"
        Write-Host "ERRORFLAG:: True"
        return
    }

    Write-Host "Removing license $LicenseSkuId from $UserPrincipalName..." -ForegroundColor Cyan
    $bodyJson = @{
        addLicenses    = @()
        removeLicenses = @($LicenseSkuId)
    } | ConvertTo-Json -Depth 3
}

# -----------------------------
# AssignLicense API
# -----------------------------
try {
    Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/assignLicense" `
        -Headers $headers `
        -Method POST `
        -Body $bodyJson

    Write-Host "WORKNOTES:: License $actionNormalized operation completed successfully."
    Write-Host "ERRORFLAG:: False"
    Write-Host "✅ SUCCESS: License $actionNormalized successfully!" -ForegroundColor Green
}
catch {
    Write-Host "❌ API failed" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    Write-Host "WORKNOTES:: API failed: $($_.Exception.Message)"
    Write-Host "ERRORFLAG:: True"
    return
}
#>


param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret,
    [string]$UserPrincipalName,
    [string]$LicenseSkuId,
    [string]$ActionType
)

$global:WorkNotesMessage = ""
$global:ErrorFlag = $false

function Set-Result {
    param(
        [string]$Message,
        [bool]$IsError
    )

    $global:WorkNotesMessage = $Message
    $global:ErrorFlag = $IsError

    Write-Output "WORKNOTES=$global:WorkNotesMessage"
    Write-Output "ERRORFLAG=$($global:ErrorFlag)"
    exit
}


# Normalize Action
$ActionType = $ActionType.ToLower().Trim()
if ($ActionType -eq "add license") { $actionNormalized = "add" }
elseif ($ActionType -eq "remove license") { $actionNormalized = "remove" }
else {
    Set-Result -Message "Invalid ActionType. Use 'add license' or 'remove license'." -IsError $true
}

# Validate UPN Format
$validDomain = $env:VALID_DOMAIN
if ([string]::IsNullOrWhiteSpace($validDomain)) {
    Set-Result -Message "VALID_DOMAIN environment variable is not set." -IsError $true
}

$escapedDomain = [Regex]::Escape($validDomain)
if ($UserPrincipalName -notmatch "^[a-zA-Z0-9._-]+@$escapedDomain$") {
    Set-Result -Message "Invalid UPN format for user." -IsError $true
}

# Get Access Token
try {
    $tokenResponse = Invoke-RestMethod -Method POST `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -Body @{
            grant_type    = "client_credentials"
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = "https://graph.microsoft.com/.default"
        }

    $token = $tokenResponse.access_token
}
catch {
    Set-Result -Message "Failed to get access token." -IsError $true
}

$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
}

# Check User Exists
try {
    $user = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/users/$UserPrincipalName" `
        -Headers $headers `
        -Method GET
}
catch {
    Set-Result -Message "User not found: $UserPrincipalName" -IsError $true
}

# Get Tenant SKUs
try {
    $tenantLicenses = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" `
        -Headers $headers `
        -Method GET
}
catch {
    Set-Result -Message "Failed to fetch subscribed SKUs." -IsError $true
}

$sku = $tenantLicenses.value | Where-Object { $_.skuId -eq $LicenseSkuId }

if (-not $sku) {
    Set-Result -Message "License SKU not found in tenant." -IsError $true
}

# Get License Details of User
try {
    $licenseDetails = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/licenseDetails" `
        -Headers $headers `
        -Method GET
}
catch {
    Set-Result -Message "Failed to fetch user license details." -IsError $true
}

$assignedSkuIds = $licenseDetails.value.skuId

# ADD LICENSE
if ($actionNormalized -eq "add") {

    if ($assignedSkuIds -contains $LicenseSkuId) {
        Set-Result -Message "License already assigned to user." -IsError $false
    }

    $availableUnits = $sku.prepaidUnits.enabled - $sku.consumedUnits
    if ($availableUnits -le 0) {
        Set-Result -Message "No available license units for this SKU." -IsError $true
    }

    $bodyJson = @{
        addLicenses    = @(@{ skuId = $LicenseSkuId })
        removeLicenses = @()
    } | ConvertTo-Json -Depth 3
}

# REMOVE LICENSE
if ($actionNormalized -eq "remove") {

    if ($assignedSkuIds -notcontains $LicenseSkuId) {
        Set-Result -Message "License not assigned to user. Nothing to remove." -IsError $false
    }

    $bodyJson = @{
        addLicenses    = @()
        removeLicenses = @($LicenseSkuId)
    } | ConvertTo-Json -Depth 3
}

# Call AssignLicense API
try {
    Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/assignLicense" `
        -Headers $headers `
        -Method POST `
        -Body $bodyJson

    Set-Result -Message "License $actionNormalized completed successfully." -IsError $false
}
catch {
    Set-Result -Message "Failed to process license operation: $($_.Exception.Message)" -IsError $true
}

