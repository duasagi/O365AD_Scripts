param(
    [Parameter(Mandatory=$true)][string]$AppId,
    [Parameter(Mandatory=$true)][string]$TenantId,
    [Parameter(Mandatory=$true)][string]$Thumbprint,
    [Parameter(Mandatory=$true)][string]$CertBase64,
    [Parameter(Mandatory=$true)][string]$CertPassword
)

# --- Decode and write PFX to temp ---
$bytes = [Convert]::FromBase64String($CertBase64)
$pfxPath = Join-Path $env:TEMP "github-exch-cert.pfx"
[IO.File]::WriteAllBytes($pfxPath, $bytes)
Write-Host "Wrote PFX to $pfxPath"

# --- Import PFX to CurrentUser\My so MSAL can find it by thumbprint ---
$securePwd = ConvertTo-SecureString -String $CertPassword -AsPlainText -Force
# Remove existing cert with same thumbprint if exists (avoid duplicates)
Try {
    $existing = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object Thumbprint -eq $Thumbprint
    if ($existing) {
        Write-Host "Cert with same thumbprint already exists in store. Removing it first..."
        $existing | Remove-Item -Force
    }
} Catch { Write-Warning $_ }

Write-Host "Importing PFX into Cert:\\CurrentUser\\My ..."
$imported = Import-PfxCertificate -FilePath $pfxPath -CertStoreLocation Cert:\CurrentUser\My -Password $securePwd
if (-not $imported) { throw "Failed to import certificate" }
Write-Host "Certificate imported. Thumbprint: $($imported.Thumbprint)"

# --- Install ExchangeOnlineManagement module if needed ---
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Installing ExchangeOnlineManagement module..."
    Install-Module ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
}

# --- Connect using certificate thumbprint as app-only authentication ---
Write-Host "Connecting to Exchange Online using application certificate..."
Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $Thumbprint -Organization $TenantId -ShowBanner:$false

# Example actions — list first 10 distribution groups
Write-Host "Listing top 10 distribution groups..."
Get-DistributionGroup -ResultSize 10 | Select DisplayName, PrimarySmtpAddress | Format-Table -AutoSize

# Example: add member to a distribution group (uncomment and edit to use)
# Add-DistributionGroupMember -Identity "DL-Example@yourdomain.com" -Member "user@yourdomain.com"

# Disconnect
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected from Exchange Online."

# Clean up: optionally remove the imported cert from store
# Remove-Item -Path "Cert:\CurrentUser\My\$Thumbprint" -Force
