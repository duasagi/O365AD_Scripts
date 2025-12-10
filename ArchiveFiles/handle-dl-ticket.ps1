param(
    [Parameter(Mandatory = $true)]
    [string]$DistributionGroupName,

    [Parameter(Mandatory = $true)]
    [string]$NewOwner_EmailAddress
)

# ======== Read Configuration from GitHub Secrets ========
$certBase64    = $env:CERT_BASE64
$pfxPassword   = $env:CERT_PASSWORD
$appId         = $env:APP_ID
$tenantId      = $env:TENANT_ID
$organization  = $env:ORGANIZATION

# ======== Decode and Import Certificate to Store ========
try {
    Write-Host "🔐 Loading certificate from GitHub secret..."
    $tempPfxPath = Join-Path $env:TEMP "exchange_cert.pfx"

    [System.IO.File]::WriteAllBytes($tempPfxPath, [Convert]::FromBase64String($certBase64))
    $securePwd = ConvertTo-SecureString -String $pfxPassword -AsPlainText -Force

    # Remove any existing cert with same thumbprint (avoid duplicates)
    $tempCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($tempPfxPath, $securePwd)
    $thumb = $tempCert.Thumbprint
    $existing = Get-ChildItem Cert:\CurrentUser\My | Where-Object Thumbprint -eq $thumb
    if ($existing) {
        Write-Host "⚠️ Removing existing certificate with same thumbprint..."
        $existing | Remove-Item -Force
    }

    Write-Host "📥 Importing certificate into Cert:\CurrentUser\My ..."
    Import-PfxCertificate -FilePath $tempPfxPath -CertStoreLocation Cert:\CurrentUser\My -Password $securePwd | Out-Null
    Write-Host "✅ Imported certificate with thumbprint: $thumb"
}
catch {
    Write-Host "❌ Failed to load/import certificate: $_"
    exit 1
}

# ======== Ensure ExchangeOnlineManagement Module ========
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "📦 Installing ExchangeOnlineManagement..."
    Install-Module ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
}

# ======== Connect to Exchange Online ========
$connectionStatus = "Failed"
try {
    Write-Host "🔗 Connecting to Exchange Online..."
    Connect-ExchangeOnline -AppId $appId -CertificateThumbprint $thumb -Organization $organization -ShowBanner:$false -ErrorAction Stop
    Write-Host "✅ Successfully connected to Exchange Online"
    $connectionStatus = "Success"
}
catch {
    Write-Host "❌ Connection to Exchange Online failed: $_"
    exit 1
}

# ======== Distribution Group Logic ========
$validDomain = "@stefaninisandbox.onmicrosoft.com"
$DistributionGroupNamePSTN = $DistributionGroupName -replace "\s", ""
$OwnersArray = $NewOwner_EmailAddress -split ',' | ForEach-Object { $_.Trim() }

$distributionGroupStatusMessage = ""
$operationStatus = "Success"
$anyOwnerAdded = $false
$anyOwnerFailed = $false

# ======== Check if Group Exists ========
try {
    $existingGroup = Get-DistributionGroup -Identity $DistributionGroupName -ErrorAction Stop
    Write-Host "ℹ️ Group already exists: $DistributionGroupName"
    $operationStatus = "Failed"
    $distributionGroupStatusMessage = "Group already exists."
}
catch {
    Write-Host "🆕 Creating new distribution group: $DistributionGroupName"
    try {
        New-DistributionGroup `
            -Name $DistributionGroupName `
            -PrimarySmtpAddress "$DistributionGroupNamePSTN$validDomain" `
            -Alias $DistributionGroupNamePSTN `
            -Type Distribution `
            -ErrorAction Stop | Out-Null

        Write-Host "✅ Created new distribution group."
        $distributionGroupStatusMessage = "Successfully created distribution group $DistributionGroupName."
    }
    catch {
        Write-Host "❌ Failed to create distribution group: $_"
        exit 1
    }
}

# ======== Add Owners ========
foreach ($owner in $OwnersArray) {
    if ($owner -like "*$validDomain") {
        try {
            $recipient = Get-Recipient -Identity $owner -ErrorAction Stop
            Set-DistributionGroup -Identity $DistributionGroupName -ManagedBy @{Add = $owner} -ErrorAction Stop
            Write-Host "✅ Added $owner as owner."
            $anyOwnerAdded = $true
        }
        catch {
            Write-Host "❌ Failed to add $owner as owner: $_"
            $anyOwnerFailed = $true
        }
    }
    else {
        Write-Host "⚠️ $owner is not part of the valid domain ($validDomain)."
        $anyOwnerFailed = $true
    }
}

# ======== Determine Final Status ========
if (($anyOwnerAdded) -and ($anyOwnerFailed)) {
    $operationStatus = "PartialSuccess"
}
elseif (-not $anyOwnerAdded) {
    $operationStatus = "Failed"
}

# ======== Output Summary ========
$output = [ordered]@{
    ConnectionStatus  = $connectionStatus
    OperationStatus   = $operationStatus
    DistributionGroup = $DistributionGroupName
    OwnersAdded       = if ($anyOwnerAdded) { "Yes" } else { "No" }
    Message           = $distributionGroupStatusMessage
}

Write-Host "`n===== Summary ====="
$output | Format-List
$output | ConvertTo-Json -Depth 3

# ======== Cleanup ========
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "🔒 Disconnected from Exchange Online."
