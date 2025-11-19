param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret,
    [string]$UserPrincipalName,
    [string]$LicenseSkuId,
    [string]$ActionType   # add OR remove
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

$uri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/assignLicense"
$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type"  = "application/json"
}

if ($ActionType -eq "add") {
    Write-Host "Adding license $LicenseSkuId to $UserPrincipalName..." -ForegroundColor Cyan

    $bodyJson = @{
        addLicenses = @(
            @{ skuId = $LicenseSkuId }
        )
        removeLicenses = @()
    } | ConvertTo-Json -Depth 3
}
elseif ($ActionType -eq "remove") {
    Write-Host "Removing license $LicenseSkuId from $UserPrincipalName..." -ForegroundColor Cyan

    $bodyJson = @{
        addLicenses = @()
        removeLicenses = @(
            "$LicenseSkuId"
        )
    } | ConvertTo-Json -Depth 3
}
else {
    Write-Host "❌ Invalid ActionType. Use: add OR remove" -ForegroundColor Red
    exit 1
}

try {
    Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $bodyJson

    if ($ActionType -eq "add") {
        Write-Host "✅ SUCCESS: License added!" -ForegroundColor Green
    }
    else {
        Write-Host "✅ SUCCESS: License removed!" -ForegroundColor Green
    }
}
catch {
    Write-Host "❌ API failed" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    exit 1
}
