param(
    [Parameter(Mandatory=$true)]
    [string]$DistributionGroupName,

    [Parameter(Mandatory=$true)]
    [string]$NewOwner_EmailAddress
)

# ======== Read Configuration from GitHub Secrets ========
$certBase64       = $env:CERT_BASE64
$pfxPasswordPlain = $env:CERT_PASSWORD
$appId            = $env:APP_ID
$tenantId         = $env:TENANT_ID
$organization     = $env:ORGANIZATION

# ======== Decode and Load Certificate ========
try {
    Write-Host "🔐 Loading certificate from GitHub secret..."
    $tempPfxPath = Join-Path $env:GITHUB_WORKSPACE "temp_cert.pfx"
    [System.IO.File]::WriteAllBytes($tempPfxPath, [Convert]::FromBase64String($certBase64))
    $pfxPassword  = ConvertTo-SecureString -String $pfxPasswordPlain -AsPlainText -Force

    # ✅ Use constructor instead of Import() (immutable fix)
    $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $tempPfxPath,
        $pfxPassword,
        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
    )

    $certThumbprint = $cert.Thumbprint
    Write-Host "✅ Loaded certificate with thumbprint: $certThumbprint"
}
catch {
    Write-Host "❌ Failed to load certificate: $_"
    exit 1
}

# ======== Connect to Exchange Online ========
$connectionStatus = "Failed"
try {
    Write-Host "🔗 Connecting to Exchange Online..."
    Connect-ExchangeOnline `
        -AppId $appId `
        -CertificateThumbprint $certThumbprint `
        -Organization $organization `
        -ShowBanner:$false -ErrorAction Stop

    $connectionStatus = "Success"
    Write-Host "✅ Successfully connected to Exchange Online"
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
        $createdl = New-DistributionGroup -Name $DistributionGroupName `
            -PrimarySmtpAddress "$DistributionGroupNamePSTN$validDomain" `
            -Alias $DistributionGroupNamePSTN -Type Distribution -ErrorAction Stop

write-Host "New-DistributionGroup -Name $DistributionGroupName `
            -PrimarySmtpAddress "$DistributionGroupNamePSTN$validDomain" `
            -Alias $DistributionGroupNamePSTN -Type Distribution -ErrorAction Stop"


        $distributionGroupStatusMessage = "Successfully created distribution group $DistributionGroupName."
        Write-Host "✅ Created new distribution group."
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
            Set-DistributionGroup -Identity $DistributionGroupName -ManagedBy @{Add=$owner} -ErrorAction Stop
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
    ConnectionStatus    = $connectionStatus
    OperationStatus     = $operationStatus
    DistributionGroup   = $DistributionGroupName
    OwnersAdded         = if ($anyOwnerAdded) { "Yes" } else { "No" }
    Message             = $distributionGroupStatusMessage
}

Write-Host "`n===== Summary ====="
$output | Format-List
$output | ConvertTo-Json -Depth 3
