
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
        [bool]$OperationSuccess  # TRUE if add/remove actually succeeded
    )

    $global:WorkNotesMessage = $Message

    # Apply your rule for ErrorFlag
    if ($OperationSuccess) {
        $global:ErrorFlag = $true
    } else {
        $global:ErrorFlag = $false
    }

    Write-Output "WORKNOTES::$global:WorkNotesMessage"
    Write-Output "ERRORFLAG::$($global:ErrorFlag)"

    exit 0
}

# -----------------------------
# VALIDATE ACTION TYPE
# -----------------------------
if (-not $ActionType) {
    Set-Result -Message "ActionType not provided." -OperationSuccess $false
}

$ActionType = $ActionType.ToLower().Trim()
switch ($ActionType) {
    "add"            { $action = "add" }
    "add license"    { $action = "add" }
    "remove"         { $action = "remove" }
    "remove license" { $action = "remove" }
    default {
        Set-Result -Message "Invalid ActionType. Use add/remove." -OperationSuccess $false
    }
}

# -----------------------------
# VALIDATE UPN
# -----------------------------
$validDomain = $env:VALID_DOMAIN
if ([string]::IsNullOrWhiteSpace($validDomain)) {
    Set-Result -Message "VALID_DOMAIN environment variable is not set." -OperationSuccess $false
}

$escapedDomain = [Regex]::Escape($validDomain)
if ($UserPrincipalName -notmatch "^[a-zA-Z0-9._-]+@$escapedDomain$") {
    Set-Result -Message "Invalid UPN format. Must end with @$validDomain" -OperationSuccess $false
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
    Set-Result -Message "Failed to get access token: $($_.Exception.Message)" -OperationSuccess $false
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
    Set-Result -Message "User not found: $UserPrincipalName" -OperationSuccess $false
}

# -----------------------------
# GET TENANT SKU INFO
# -----------------------------
try {
    $tenantSkus = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" -Headers $headers
}
catch {
    Set-Result -Message "Failed to fetch tenant SKUs." -OperationSuccess $false
}

$sku = $tenantSkus.value | Where-Object { $_.skuId -eq $LicenseSkuId }
if (-not $sku) {
    Set-Result -Message "License SKU not found in tenant." -OperationSuccess $false
}

# -----------------------------
# GET USER LICENSE DETAILS
# -----------------------------
try {
    $licenseInfo = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/licenseDetails" -Headers $headers
}
catch {
    Set-Result -Message "Failed to fetch user license details." -OperationSuccess $false
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
        Set-Result -Message "License already assigned to user." -OperationSuccess $false
    }

    $available = $sku.prepaidUnits.enabled - $sku.consumedUnits
    if ($available -le 0) {
        Set-Result -Message "No available licenses left for this SKU." -OperationSuccess $false
    }

    $bodyJson = @"
{
  "addLicenses": [
    { "skuId": "$LicenseSkuId" }
  ],
  "removeLicenses": []
}
"@

    try {
        Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/assignLicense" `
                          -Headers $headers -Method POST -Body $bodyJson
        Set-Result -Message "License add operation completed successfully." -OperationSuccess $true
    }
    catch {
        Set-Result -Message "Failed to add license: $($_.Exception.Message)" -OperationSuccess $false
    }
}

# -----------------------------
# REMOVE LICENSE
# -----------------------------
if ($action -eq "remove") {

    if ($assignedSkus -notcontains $LicenseSkuId) {
        Set-Result -Message "License not assigned to user. Nothing to remove." -OperationSuccess $false
    }

    $bodyJson = @"
{
  "addLicenses": [],
  "removeLicenses": [ "$LicenseSkuId" ]
}
"@

    try {
        Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/assignLicense" `
                          -Headers $headers -Method POST -Body $bodyJson
        Set-Result -Message "License remove operation completed successfully." -OperationSuccess $true
    }
    catch {
        Set-Result -Message "Failed to remove license: $($_.Exception.Message)" -OperationSuccess $false
    }
}
