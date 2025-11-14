param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret,
    [string]$UserPrincipalName,
    [string]$LicenseSkuId
)

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

$body = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "https://graph.microsoft.com/.default"
}

try {
    $tokenResponse = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $body
    $token = $tokenResponse.access_token
}
catch {
    Write-Host ("Failed to get access token: " + $_.Exception.Message) -ForegroundColor Red
    exit 1
}

Write-Host "Removing license $LicenseSkuId from $UserPrincipalName ..." -ForegroundColor Cyan

$uri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/assignLicense"

$bodyJson = @{
    addLicenses    = @()                   # nothing to add
    removeLicenses = @(
        "$LicenseSkuId"
    )
} | ConvertTo-Json -Depth 3

$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type"  = "application/json"
}

try {
    Invoke-RestMethod -Uri $uri -Headers $headers -Method POST -Body $bodyJson
    Write-Host "SUCCESS: License removed!" -ForegroundColor Green
}
catch {
    Write-Host "Failed to remove license" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    exit 1
}
