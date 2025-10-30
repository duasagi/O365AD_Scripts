param(
    [Parameter(Mandatory = $true)]
    [string]$DistributionGroupName,

    [Parameter(Mandatory = $true)]
    [string]$NewOwner_EmailAddress
)

# ======== Global Preferences ========
$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"
$InformationPreference = "Continue"

Write-Output "🚀 Script started at $(Get-Date)"

# ======== Read Configuration from GitHub Secrets ========
$certBase64       = $env:CERT_BASE64
$pfxPasswordPlain = $env:CERT_PASSWORD
$appId            = $env:APP_ID
$tenantId         = $env:TENANT_ID
$organization     = $env:ORGANIZATION

# ======== Decode and Load Certificate ========
try {
    Write-Output "🔐 Loading certificate from GitHub secret..."
    $tempPfxPath = Join-Path $env:GITHUB_WORKSPACE "temp_cert.pfx"
    [System.IO.File]::WriteAllBytes($tempPfxPath, [Convert]::FromBase64String($certBase64))
    $pfxPassword = ConvertTo-SecureString -String $pfxPasswordPlain -AsPlainText -Force

    $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $tempPfxPath,
        $pfxPassword,
        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
    )

    $certThumbprint = $cert.Thumbprint
    Write-Output "✅ Loaded certificate with thumbprint: $certThumbprint"
}
catch {
    Write-Output "❌ Failed to load certificate: $_"
    exit 1
}

# ======== Connect to Exchange Online ========
$connectionStatus = "Failed"
try {
    Write-Output "🔗 Connecting to Exchange Online..."
    Connect-ExchangeOnline `
        -AppId $appId `
        -CertificateThumbprint $certThumbprint `
        -Organization $organization `
        -ShowBanner:$false -ErrorAction Stop

    $connectionStatus = "Success"
    Write-Output "✅ Connected to Exchange Online"
}
catch {
    Write-Output "❌ Connection failed: $_"
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
    Write-Output "ℹ️ Group already exists: $DistributionGroupName"
    $distributionGroupStatusMessage = "Group already exists. Owners will be updated if valid."
}
catch {
    Write-Output "🆕 Creating new distribution group: $DistributionGroupName"
    try {
        New-DistributionGroup -Name $DistributionGroupName `
            -PrimarySmtpAddress "$DistributionGroupNamePSTN$validDomain" `
            -Alias $DistributionGroupNamePSTN `
            -Type Distribution -ErrorAction Stop | Out-Null

        $distributionGroupStatusMessage = "Successfully created distribution group."
        Write-Output "✅ Created distribution group."
    }
    catch {
        Write-Output "❌ Failed to create distribution group: $_"
        exit 1
    }
}

# ======== Add Owners ========
foreach ($owner in $OwnersArray) {
    if ($owner -like "*$validDomain") {
        try {
            $recipient = Get-Recipient -Identity $owner -ErrorAction Stop
            Set-DistributionGroup -Identity $DistributionGroupName -ManagedBy @{Add = $owner} -ErrorAction Stop
            Write-Output "✅ Added $owner as owner."
            $anyOwnerAdded = $true
        }
        catch {
            Write-Output "❌ Failed to add $owner as owner: $_"
            $anyOwnerFailed = $true
        }
    }
    else {
        Write-Output "⚠️ $owner does not belong to valid domain ($validDomain)."
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

Write-Output "`n===== Summary ====="
$output | Format-List | Out-String | Write-Output

# Write JSON to console + GitHub summary
$outputJson = $output | ConvertTo-Json -Depth 3
Write-Output "`n📦 JSON Output:"
Write-Output $outputJson
$outputJson | Out-File -FilePath $env:GITHUB_STEP_SUMMARY -Append -Encoding utf8

Write-Output "🏁 Script completed successfully."
