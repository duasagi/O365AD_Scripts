param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret,
    [string]$UserPrincipalName,
    [string]$LicenseSkuId,
    [string]$ActionType
)

# -----------------------------
# GLOBAL VARIABLES
# -----------------------------
$global:WorkNotesMessage = ""
$global:ErrorFlag = $false

function Set-Result {
    param(
        [string]$Message,
        [bool]$IsError
    )

    $global:WorkNotesMessage = $Message
    $global:ErrorFlag = $IsError

    Write-Output "WORKNOTES::$global:WorkNotesMessage"
    Write-Output "ERRORFLAG::$($global:ErrorFlag)"

    if ($IsError) { exit 1 } else { exit 0 }
}

# -----------------------------
# VALIDATE ACTION TYPE
# -----------------------------
if (-not $ActionType) {
    Set-Result -Message "ActionType not provided." -IsError $true
}

$ActionType = $ActionType.ToLower().Trim()

switch ($ActionType) {
    "add"            { $action = "add" }
    "add license"    { $action = "add" }
    "remove"         { $action = "remove" }
    "remove license" { $action = "remove" }
    default {
        Set-Result -Message "Invalid ActionType. Use add/remove." -IsError $true
    }
}

# -----------------------------
# VALIDATE UPN
# -----------------------------
$validDomain = $env:VALID_DOMAIN
if ([string]::IsNullOrWhiteSpace($validDomain)) {
    Set-Result -Message "VALID_DOMAIN environment variable is not set." -IsError $true
}

$escapedDomain = [Regex]::Escape($validDomain)

if ($UserPrincipalName -notmatch "^[a-zA-Z0-9._-]+@$escapedDomain$") {
    Set-Result -Message "Invalid UPN format. Must end with @$validDomain" -IsError $true
}

# -----------------------------
# GET ACCESS TOKEN
# -----------------------------
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
    Set-Result -Message "Failed to get access token: $($_.Exception.Message)" -IsError $true
}

$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
}

# -----------------------------
# CHECK USER EXISTS
# -----------------------------
try {
    $user = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$UserPrincipalName" -Headers $headers
}
catch {
    Set-Result -Message "User not found: $UserPrincipalName" -IsError $true
}

# -----------------------------
# GET TENANT SKU INFO
# -----------------------------
try {
    $tenantSkus = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" -Headers $headers
}
catch {
    Set-Result -Message "Failed to fetch tenant SKUs." -IsError $true
}

$sku = $tenantSkus.value | Where-Object { $_.skuId -eq $LicenseSkuId }

if (-not $sku) {
    Set-Result -Message "License SKU not found in tenant." -IsError $true
}

# -----------------------------
# GET USER LICENSE DETAILS
# -----------------------------
try {
    $licenseInfo = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/licenseDetails" -Headers $headers
}
catch {
    Set-Result -Message "Failed to fetch user license details." -IsError $true
}

$assignedSkus = @()
if ($licenseInfo.value) {
    $assignedSkus = $licenseInfo.value.skuId
}

# -----------------------------
# ADD LICENSE
# -----------------------------
if ($action -eq "add") {

    if ($assignedSkus -contains $LicenseSkuId) {
        Set-Result -Message "License already assigned to user." -IsError $false
    }

    $available = $sku.prepaidUnits.enabled - $sku.consumedUnits

    if ($available -le 0) {
        Set-Result -Message "No available licenses left for this SKU." -IsError $true
    }

    $bodyJson = @"
{
  "addLicenses": [
    { "skuId": "$LicenseSkuId" }
  ],
  "removeLicenses": []
}
"@
}

# -----------------------------
# REMOVE LICENSE
# -----------------------------
if ($action -eq "remove") {

    if ($assignedSkus -notcontains $LicenseSkuId) {
        Set-Result -Message "License not assigned to user. Nothing to remove." -IsError $false
    }

    $bodyJson = @"
{
  "addLicenses": [],
  "removeLicenses": [ "$LicenseSkuId" ]
}
"@
}

# -----------------------------
# PERFORM OPERATION
# -----------------------------
try {
    Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/assignLicense" `
        -Headers $headers `
        -Method POST `
        -Body $bodyJson

    Set-Result -Message "License $action operation completed successfully." -IsError $false
}
catch {
    Set-Result -Message "Failed to process license operation: $($_.Exception.Message)" -IsError $true
}
