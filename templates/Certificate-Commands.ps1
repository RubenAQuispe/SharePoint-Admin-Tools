<#
.SYNOPSIS
    Certificate creation commands for SharePoint Admin Tools authentication

.DESCRIPTION
    This file contains the PowerShell commands needed to create self-signed certificates
    for authenticating with Microsoft Graph API in SharePoint Admin Tools.

.NOTES
    Author: SharePoint Admin Tools Community
    Created: 2025-01-06
    Purpose: Template for certificate creation and management
#>

# =============================================================================
# CERTIFICATE CREATION COMMANDS
# =============================================================================

# 1. Create a self-signed certificate for SharePoint authentication
# Note: Replace "YourSecurePassword" with a strong password
$cert = New-SelfSignedCertificate -Subject "CN=SharePointPermissionsAudit" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256 -NotAfter (Get-Date).AddDays(365)

# 2. Set a secure password for the certificate export
$certPassword = ConvertTo-SecureString -String "YourSecurePassword" -Force -AsPlainText

# 3. Export certificate with private key (.pfx file) - for PowerShell script use
Export-PfxCertificate -Cert $cert -FilePath C:\Temp\SharePointAudit.pfx -Password $certPassword

# 4. Export certificate public key (.cer file) - for Azure app registration upload
Export-Certificate -Cert $cert -FilePath C:\Temp\SharePointAudit.cer

# =============================================================================
# CERTIFICATE VERIFICATION COMMANDS
# =============================================================================

# Verify the certificate was created successfully
Write-Host "Certificate Thumbprint: $($cert.Thumbprint)" -ForegroundColor Green
Write-Host "Certificate Subject: $($cert.Subject)" -ForegroundColor Green
Write-Host "Certificate Expiry: $($cert.NotAfter)" -ForegroundColor Green

# Test certificate file accessibility
if (Test-Path "C:\Temp\SharePointAudit.pfx") {
    Write-Host "PFX file created successfully: C:\Temp\SharePointAudit.pfx" -ForegroundColor Green
} else {
    Write-Host "ERROR: PFX file not found!" -ForegroundColor Red
}

if (Test-Path "C:\Temp\SharePointAudit.cer") {
    Write-Host "CER file created successfully: C:\Temp\SharePointAudit.cer" -ForegroundColor Green
} else {
    Write-Host "ERROR: CER file not found!" -ForegroundColor Red
}

# Test loading the certificate
try {
    $testCert = Get-PfxCertificate -FilePath "C:\Temp\SharePointAudit.pfx"
    Write-Host "Certificate loads successfully!" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Cannot load certificate - $($_.Exception.Message)" -ForegroundColor Red
}

# =============================================================================
# ALTERNATIVE CERTIFICATE CREATION (with different validity periods)
# =============================================================================

# Short-term certificate (30 days) - for testing
# $cert = New-SelfSignedCertificate -Subject "CN=SharePointPermissionsAudit-Test" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256 -NotAfter (Get-Date).AddDays(30)

# Long-term certificate (2 years) - for production
# $cert = New-SelfSignedCertificate -Subject "CN=SharePointPermissionsAudit-Prod" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256 -NotAfter (Get-Date).AddDays(730)

# =============================================================================
# CERTIFICATE MANAGEMENT COMMANDS
# =============================================================================

# List all certificates in current user store
# Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object { $_.Subject -like "*SharePoint*" } | Format-Table Subject, Thumbprint, NotAfter

# Remove a certificate by thumbprint (replace with actual thumbprint)
# Remove-Item -Path "Cert:\CurrentUser\My\CERTIFICATE_THUMBPRINT_HERE"

# =============================================================================
# SECURE STORAGE RECOMMENDATIONS
# =============================================================================

<#
IMPORTANT SECURITY NOTES:

1. NEVER store certificates in public repositories
2. Use Azure Key Vault for production environments:
   - Store the certificate in Key Vault
   - Grant your application access to the Key Vault
   - Retrieve certificate programmatically

3. For local development:
   - Store certificates in a secure location (not Desktop/Downloads)
   - Use strong passwords for certificate files
   - Set appropriate file permissions (readable only by authorized users)

4. Certificate rotation:
   - Set calendar reminders for certificate expiration
   - Plan for certificate renewal before expiry
   - Update Azure app registration when rotating certificates

5. Backup strategy:
   - Keep secure backups of certificate files
   - Document certificate thumbprints and expiry dates
   - Maintain an inventory of all certificates in use

EXAMPLE SECURE PATHS:
Windows: C:\Certificates\SharePoint\
Linux/Mac: /opt/certificates/sharepoint/
Azure Key Vault: https://your-keyvault.vault.azure.net/
#>

# =============================================================================
# CERTIFICATE USAGE EXAMPLES
# =============================================================================

<#
After creating the certificate, use it in scripts like this:

# Load certificate for authentication
$certPath = "C:\Temp\SharePointAudit.pfx"
$certPassword = ConvertTo-SecureString "YourSecurePassword" -AsPlainText -Force
$certificate = Get-PfxCertificate -FilePath $certPath -Password $certPassword

# Use with SharePoint Admin Tools
.\Check-OneDrivePermissions.ps1 `
    -AuthMethod Certificate `
    -CertificatePath $certPath `
    -CertificatePassword $certPassword `
    -ClientId "your-client-id" `
    -TenantId "your-tenant-id" `
    -SourceUserEmail "source@company.com" `
    -TargetUserEmail "target@company.com"
#>
