# Exchange Online Setup Script for Distribution List Manager
# Run this script as Administrator in PowerShell

param(
    [string]$AppId = "",
    [string]$TenantId = "",
    [string]$Organization = ""
)

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Exchange Online Setup for DL Manager" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Check if running as admin
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "WARNING: Not running as Administrator. Some steps may fail." -ForegroundColor Yellow
    Write-Host ""
}

# Step 1: Install ExchangeOnlineManagement module
Write-Host "Step 1: Checking ExchangeOnlineManagement module..." -ForegroundColor Yellow
$module = Get-Module -ListAvailable ExchangeOnlineManagement
if (-not $module) {
    Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor White
    Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser -AllowClobber
    Write-Host "Module installed successfully!" -ForegroundColor Green
} else {
    Write-Host "Module already installed (version $($module.Version))" -ForegroundColor Green
}

# Step 2: Read .env file for App ID
Write-Host ""
Write-Host "Step 2: Reading configuration..." -ForegroundColor Yellow
$envPath = Join-Path $PSScriptRoot ".env"
if (Test-Path $envPath) {
    $envContent = Get-Content $envPath
    foreach ($line in $envContent) {
        if ($line -match "^AZURE_CLIENT_ID=(.+)$") {
            $AppId = $matches[1].Trim()
        }
        if ($line -match "^AZURE_TENANT_ID=(.+)$") {
            $TenantId = $matches[1].Trim()
        }
        if ($line -match "^EXCHANGE_ORGANIZATION=(.+)$") {
            $Organization = $matches[1].Trim()
        }
    }
}

if (-not $AppId) {
    $AppId = Read-Host "Enter your Azure App ID (Client ID)"
}
if (-not $Organization) {
    $Organization = Read-Host "Enter your organization (e.g., contoso.onmicrosoft.com)"
}

Write-Host "App ID: $AppId" -ForegroundColor Cyan
Write-Host "Organization: $Organization" -ForegroundColor Cyan

# Step 3: Generate self-signed certificate
Write-Host ""
Write-Host "Step 3: Generating self-signed certificate..." -ForegroundColor Yellow

$certName = "DLManagerExchangeCert"
$certPath = Join-Path $PSScriptRoot "exchange_cert.cer"

# Check if cert already exists
$existingCert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Subject -like "*$certName*" }

if ($existingCert) {
    Write-Host "Certificate already exists. Thumbprint: $($existingCert.Thumbprint)" -ForegroundColor Green
    $cert = $existingCert
} else {
    # Create new certificate
    $cert = New-SelfSignedCertificate `
        -Subject "CN=$certName" `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -KeyExportPolicy Exportable `
        -KeySpec Signature `
        -KeyLength 2048 `
        -KeyAlgorithm RSA `
        -HashAlgorithm SHA256 `
        -NotAfter (Get-Date).AddYears(2)

    Write-Host "Certificate created successfully!" -ForegroundColor Green
}

$thumbprint = $cert.Thumbprint
Write-Host "Thumbprint: $thumbprint" -ForegroundColor Cyan

# Export public key for Azure
Export-Certificate -Cert $cert -FilePath $certPath -Force | Out-Null
Write-Host "Certificate exported to: $certPath" -ForegroundColor Green

# Step 4: Update .env file
Write-Host ""
Write-Host "Step 4: Updating .env file..." -ForegroundColor Yellow

$envContent = Get-Content $envPath -Raw
$envContent = $envContent -replace "EXCHANGE_CERT_THUMBPRINT=.*", "EXCHANGE_CERT_THUMBPRINT=$thumbprint"
$envContent = $envContent -replace "EXCHANGE_ORGANIZATION=.*", "EXCHANGE_ORGANIZATION=$Organization"
Set-Content -Path $envPath -Value $envContent.TrimEnd()

Write-Host ".env file updated!" -ForegroundColor Green

# Step 5: Instructions for Azure
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  IMPORTANT: Complete Setup in Azure Portal" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "1. Go to: https://portal.azure.com" -ForegroundColor White
Write-Host ""
Write-Host "2. Navigate to: Entra ID > App registrations > Your App" -ForegroundColor White
Write-Host ""
Write-Host "3. Upload the certificate:" -ForegroundColor White
Write-Host "   - Click 'Certificates & secrets'" -ForegroundColor Gray
Write-Host "   - Click 'Certificates' tab" -ForegroundColor Gray
Write-Host "   - Click 'Upload certificate'" -ForegroundColor Gray
Write-Host "   - Select: $certPath" -ForegroundColor Yellow
Write-Host ""
Write-Host "4. Add Exchange API permission:" -ForegroundColor White
Write-Host "   - Click 'API permissions'" -ForegroundColor Gray
Write-Host "   - Click 'Add a permission'" -ForegroundColor Gray
Write-Host "   - Select 'APIs my organization uses'" -ForegroundColor Gray
Write-Host "   - Search for 'Office 365 Exchange Online'" -ForegroundColor Gray
Write-Host "   - Select 'Application permissions'" -ForegroundColor Gray
Write-Host "   - Check 'Exchange.ManageAsApp'" -ForegroundColor Gray
Write-Host "   - Click 'Add permissions'" -ForegroundColor Gray
Write-Host "   - Click 'Grant admin consent'" -ForegroundColor Gray
Write-Host ""
Write-Host "5. Assign Exchange Administrator role to the app:" -ForegroundColor White
Write-Host "   - Go to: Entra ID > Roles and administrators" -ForegroundColor Gray
Write-Host "   - Search for 'Exchange Administrator'" -ForegroundColor Gray
Write-Host "   - Click on it > Add assignments" -ForegroundColor Gray
Write-Host "   - Click 'No member selected'" -ForegroundColor Gray
Write-Host "   - Switch to 'Enterprise applications' tab" -ForegroundColor Gray
Write-Host "   - Find and select your app: Distribution List Manager" -ForegroundColor Gray
Write-Host "   - Click 'Add'" -ForegroundColor Gray
Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  Certificate Thumbprint: $thumbprint" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "After completing the Azure setup, run the app with run.bat" -ForegroundColor White
Write-Host ""

# Test connection option
$test = Read-Host "Do you want to test the connection now? (y/n)"
if ($test -eq "y") {
    Write-Host ""
    Write-Host "Testing Exchange Online connection..." -ForegroundColor Yellow

    try {
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $thumbprint -Organization $Organization -ShowBanner:$false

        Write-Host "Connection successful!" -ForegroundColor Green

        # Test getting distribution groups
        $groups = Get-DistributionGroup -ResultSize 1
        Write-Host "Successfully retrieved distribution groups!" -ForegroundColor Green

        Disconnect-ExchangeOnline -Confirm:$false
    } catch {
        Write-Host "Connection failed: $_" -ForegroundColor Red
        Write-Host ""
        Write-Host "Make sure you completed all the Azure setup steps above." -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Setup complete! Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
