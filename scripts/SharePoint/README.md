# SharePoint Management Tools

A collection of PowerShell scripts for managing and securing SharePoint sites in Microsoft 365.

## 🛠️ Available Tools

### 1. Scan-SharePointForRansomware.ps1
**Purpose:** Scan SharePoint sites for potential ransomware indicators
- **Use Case:** Security auditing and threat detection
- **Output:** CSV report of suspicious files and patterns
- **Safety:** Read-only operation

## 📋 Detailed Tool Documentation

### Scan-SharePointForRansomware.ps1

**Purpose:** Scan SharePoint sites for files that match known ransomware patterns or indicators of compromise.

**Parameters:**
- `SiteUrl` - URL of the SharePoint site to scan
- `OutputPath` - Custom output file path (optional)
- `TenantId` - Azure AD Tenant ID (optional)
- `ClientId` - Application Client ID (optional)
- `CertificatePath` - Certificate file path for auth (optional)
- `AuthMethod` - Authentication method: Interactive, Certificate, or ClientSecret

**Example:**
```powershell
.\Scan-SharePointForRansomware.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Finance" -OutputPath "C:\Reports\RansomwareScan.csv"
```

**Output Columns:**
- `SiteUrl` - URL of the SharePoint site
- `ItemPath` - Path to the scanned item
- `ItemType` - File or Folder
- `FileName` - Name of the file
- `FileExtension` - Extension of the file
- `SuspiciousPattern` - Detected suspicious pattern (if any)
- `RiskLevel` - High, Medium, or Low
- `LastModifiedBy` - User who last modified the file
- `LastModifiedDate` - Date when the file was last modified

## 🔐 Authentication Methods

### Interactive Authentication (Recommended for first-time use)
```powershell
.\Scan-SharePointForRansomware.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Finance"
# Will prompt for browser login
```

### Certificate Authentication (Recommended for automation)
```powershell
.\Scan-SharePointForRansomware.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Finance" -CertificatePath "C:\cert.pfx" -ClientId "your-app-id" -TenantId "your-tenant-id"
```

### Client Secret Authentication
```powershell
.\Scan-SharePointForRansomware.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Finance" -ClientId "your-app-id" -TenantId "your-tenant-id" -AuthMethod "ClientSecret"
# Will prompt for client secret
```

## 🔧 Required Azure App Permissions

For certificate or client secret authentication, your Azure app registration needs:

### Microsoft Graph API Permissions:
- `Sites.Read.All` - Read SharePoint sites
- `Files.Read.All` - Read all files

### Admin Consent Required:
- All permissions require admin consent
- Grant consent in Azure Portal → App registrations → API permissions

---

**Author:** Ruben Quispe  
**Created:** 2025-06-16  
**Version:** 1.0  
**License:** MIT License - See LICENSE file for details
