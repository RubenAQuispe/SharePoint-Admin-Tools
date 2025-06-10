# OneDrive Permissions Checker

This PowerShell script analyzes a user's OneDrive to find all files and folders that have been shared with another specific user. It uses Microsoft Graph API to retrieve detailed permission information and generates a comprehensive CSV report.

## 🎯 Use Cases

- **Security Audits**: Identify what files a specific user has access to in someone's OneDrive
- **Compliance Reporting**: Generate detailed sharing reports for regulatory requirements
- **Access Reviews**: Periodic review of file sharing permissions
- **Incident Response**: Quickly identify potential data exposure during security incidents
- **Offboarding**: Check what files departing employees have shared or have access to

## 🚀 Quick Start

### Basic Usage (Interactive Authentication)
```powershell
.\Check-OneDrivePermissions.ps1 -SourceUserEmail "john.doe@company.com" -TargetUserEmail "jane.smith@company.com"
```

### Advanced Usage (Certificate Authentication)
```powershell
.\Check-OneDrivePermissions.ps1 `
    -SourceUserEmail "john.doe@company.com" `
    -TargetUserEmail "jane.smith@company.com" `
    -AuthMethod Certificate `
    -CertificatePath "C:\Certs\SharePointAudit.pfx" `
    -ClientId "12345678-1234-1234-1234-123456789012" `
    -TenantId "87654321-4321-4321-4321-210987654321" `
    -OutputPath "C:\Reports\OneDriveSharing.csv"
```

## 📋 Parameters

| Parameter | Required | Type | Description |
|-----------|----------|------|-------------|
| `SourceUserEmail` | No* | String | Email of the OneDrive owner whose files will be analyzed |
| `TargetUserEmail` | No* | String | Email of the user to check permissions for |
| `OutputPath` | No | String | Full path for the CSV report (auto-generated if not specified) |
| `TenantId` | No | String | Azure AD Tenant ID (required for certificate/secret auth) |
| `ClientId` | No | String | Application Client ID (required for certificate/secret auth) |
| `CertificatePath` | No | String | Path to .pfx certificate file |
| `CertificatePassword` | No | SecureString | Password for the certificate file |
| `ClientSecret` | No | SecureString | Client secret for app authentication |
| `AuthMethod` | No | String | Authentication method: Interactive, Certificate, or ClientSecret |
| `Recursive` | No | Boolean | Include subfolders in analysis (default: true) |
| `IncludeInherited` | No | Boolean | Include inherited permissions (default: true) |
| `DebugMode` | No | Switch | Enable detailed debug output for troubleshooting |

*The script will prompt for these if not provided

## 🔐 Authentication Methods

### 1. Interactive Authentication (Default)
- **Best for**: Testing, one-time runs, development
- **Requirements**: SharePoint Administrator or Global Administrator role
- **Setup**: None required - browser login prompt will appear

```powershell
# Simple interactive mode
.\Check-OneDrivePermissions.ps1
```

### 2. Certificate Authentication (Recommended for Production)
- **Best for**: Automated scripts, scheduled tasks, production environments
- **Requirements**: Azure App Registration with certificate
- **Security**: Most secure option, no passwords stored

```powershell
# Certificate authentication
.\Check-OneDrivePermissions.ps1 `
    -AuthMethod Certificate `
    -CertificatePath "C:\Certs\SharePointAudit.pfx" `
    -ClientId "your-app-id" `
    -TenantId "your-tenant-id"
```

### 3. Client Secret Authentication
- **Best for**: Development, testing with service accounts
- **Requirements**: Azure App Registration with client secret
- **Security**: Moderate - secrets can expire and need rotation

```powershell
# Client secret authentication
$secret = ConvertTo-SecureString "your-secret" -AsPlainText -Force
.\Check-OneDrivePermissions.ps1 `
    -AuthMethod ClientSecret `
    -ClientId "your-app-id" `
    -TenantId "your-tenant-id" `
    -ClientSecret $secret
```

## 📊 Report Output

The script generates a detailed CSV report with the following columns:

| Column | Description |
|--------|-------------|
| `ItemPath` | Full path to the file or folder |
| `ItemType` | File or Folder |
| `PermissionLevel` | Read, Write, Owner, etc. |
| `ShareType` | Direct User Permission, Sharing Link, etc. |
| `IsInherited` | Whether permission is inherited from parent |
| `InheritedFrom` | Source of inherited permission |
| `PermissionId` | Unique identifier for the permission |
| `UserDisplayName` | Display name of the user with access |
| `UserEmail` | Email address of the user |
| `UserId` | Azure AD User ID |
| `LinkWebUrl` | URL for sharing links (if applicable) |
| `CreatedDateTime` | When the permission was created |
| `LastModifiedDateTime` | When the permission was last modified |

### Sample Report Preview
```csv
ItemPath,ItemType,PermissionLevel,ShareType,IsInherited,UserDisplayName,UserEmail
"/Financial Reports/Q4 Budget.xlsx",File,"read","Direct User Permission (Email Match)",False,"Jane Smith","jane.smith@company.com"
"/Projects/Marketing",Folder,"write","Direct User Permission (ID Match)",False,"Jane Smith","jane.smith@company.com"
"/Projects/Marketing/Campaign Files",Folder,"write","Direct User Permission (Email Match)",True,"Jane Smith","jane.smith@company.com"
```

## 🔧 Troubleshooting

### Common Issues

#### "Access Denied" Errors
- **Cause**: Insufficient permissions
- **Solution**: Ensure you have SharePoint Administrator or Global Administrator role
- **Alternative**: Use an account with appropriate permissions

#### "Certificate not found" or "Certificate password required"
- **Cause**: Certificate file issues
- **Solution**: 
  ```powershell
  # Verify certificate exists and is accessible
  Test-Path "C:\path\to\certificate.pfx"
  
  # Test certificate loading
  Get-PfxCertificate -FilePath "C:\path\to\certificate.pfx"
  ```

#### "App registration not found"
- **Cause**: Incorrect Client ID or Tenant ID
- **Solution**: Verify the GUIDs in Azure Portal → App registrations

#### "Insufficient privileges to complete the operation"
- **Cause**: App registration lacks required permissions
- **Solution**: Add Files.Read.All, User.Read.All, Sites.Read.All permissions and grant admin consent

### Debug Mode
Enable debug mode for detailed troubleshooting:

```powershell
.\Check-OneDrivePermissions.ps1 -DebugMode -SourceUserEmail "user@company.com" -TargetUserEmail "target@company.com"
```

Debug mode provides:
- Detailed permission detection logic
- User matching information
- File/folder type detection details
- API call results
- Progress tracking information

## 🏗️ Advanced Configuration

### Filtering Options

#### Skip Inherited Permissions
```powershell
.\Check-OneDrivePermissions.ps1 -IncludeInherited $false
```

#### Non-Recursive (Root Level Only)
```powershell
.\Check-OneDrivePermissions.ps1 -Recursive $false
```

### Custom Output Locations

#### Specify Directory (Auto-generate filename)
```powershell
.\Check-OneDrivePermissions.ps1 -OutputPath "C:\Reports\"
```

#### Specify Full Filename
```powershell
.\Check-OneDrivePermissions.ps1 -OutputPath "C:\Reports\CustomReport.csv"
```

## 📈 Performance Considerations

### Large OneDrive Accounts
- The script processes items recursively, which can take time for large OneDrives
- Progress is displayed every 50 items processed
- Rate limiting is handled automatically by the Microsoft Graph SDK

### Optimization Tips
- Use `-Recursive $false` for root-level only analysis (faster)
- Use `-IncludeInherited $false` to reduce noise in reports
- Run during off-peak hours for large analysis jobs
- Consider certificate authentication for better performance in automated scenarios

## 🔒 Security Considerations

### Permissions Required
The script requires these Microsoft Graph permissions:
- `Files.Read.All` - Read all files that the user can access
- `User.Read.All` - Read all users' full profiles
- `Sites.Read.All` - Read items in all site collections

### Data Protection
- Reports contain sensitive information about file sharing
- Store reports securely and limit access appropriately
- Consider encrypting report files for highly sensitive environments
- Clean up old reports according to your retention policies

### Certificate Security
- Store certificates in secure locations (preferably Azure Key Vault)
- Use strong passwords for certificate files
- Rotate certificates regularly (recommended: annually)
- Monitor certificate expiration dates

## 📚 Related Documentation

- [Main Repository README](../../README.md) - Authentication setup and general information
- [Certificate Setup Guide](../../docs/Certificate-Setup.md) - Detailed certificate creation instructions
- [Azure App Registration Guide](../../docs/Azure-App-Registration.md) - Step-by-step app setup
- [Troubleshooting Guide](../../docs/Troubleshooting.md) - Common issues and solutions

## 🤝 Contributing

Found a bug or have a feature request? Please [open an issue](https://github.com/YOUR-USERNAME/SharePoint-Admin-Tools/issues) or submit a pull request!

## 📄 License

This script is part of the SharePoint Admin Tools collection and is licensed under the MIT License.
