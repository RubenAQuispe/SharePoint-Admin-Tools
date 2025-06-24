# OneDrive Permission Management Tools

A comprehensive toolkit for scanning, auditing, and managing OneDrive permissions across your Microsoft 365 environment. These tools work together seamlessly to provide a complete workflow for permission management.

## 🛠️ Available Tools

### 1. Check-OneDrivePermissions.ps1
**Purpose:** Scan a single user's OneDrive for all shared permissions
- **Use Case:** Detailed analysis of one specific OneDrive
- **Output:** CSV report of all permissions in the OneDrive
- **Safety:** Read-only operation

### 2. Scan-AllOneDrivesForUser.ps1  
**Purpose:** Scan all OneDrives in the organization for permissions granted to a specific user
- **Use Case:** Find everywhere a user has been granted access
- **Output:** CSV report showing which OneDrives the user can access
- **Safety:** Read-only operation

### 3. Remove-UserPermissions.ps1
**Purpose:** Remove user permissions from OneDrives based on CSV reports
- **Use Case:** Clean up permissions after user departure or role change
- **Input:** CSV files from the scanning tools above
- **Safety:** 🛡️ **CRITICAL SAFETY FEATURE** - Never removes permissions from user's own OneDrive
- **Default:** Dry-run mode (preview only)
- **Enhanced:** Robust handling of drive ID formats and improved folder navigation

## 🚀 Quick Start Guide

### Complete Workflow Example
This is the recommended workflow for safely removing a user's permissions:

1. **Find all OneDrives where the user has access:**
   ```powershell
   # Scan the entire organization to find where jane.smith@company.com has access
   .\Scan-AllOneDrivesForUser.ps1 -TargetUserEmail "jane.smith@company.com" -AuthMethod "Certificate" -CertificatePath "C:\Certs\SharePointAudit.pfx" -ClientId "4be377c6-ed0f-4457-abe2-2cde63c9f897" -TenantId "7182fc16-7e98-4fd7-ab66-75893f065754"
   ```

2. **Preview the permission removal (safe dry run):**
   ```powershell
   # Preview what would be removed without making any changes
   .\Remove-UserPermissions.ps1 -InputCsvPath "AllOneDrivesReport_jane_smith_company_com_20250624_083000.csv" -TargetUserEmail "jane.smith@company.com" -AuthMethod "Certificate" -CertificatePath "C:\Certs\SharePointAudit.pfx" -ClientId "4be377c6-ed0f-4457-abe2-2cde63c9f897" -TenantId "7182fc16-7e98-4fd7-ab66-75893f065754" -DryRun $true
   ```

3. **Execute the actual permission removal:**
   ```powershell
   # After confirming the dry run results, perform the actual removal
   .\Remove-UserPermissions.ps1 -InputCsvPath "AllOneDrivesReport_jane_smith_company_com_20250624_083000.csv" -TargetUserEmail "jane.smith@company.com" -AuthMethod "Certificate" -CertificatePath "C:\Certs\SharePointAudit.pfx" -ClientId "4be377c6-ed0f-4457-abe2-2cde63c9f897" -TenantId "7182fc16-7e98-4fd7-ab66-75893f065754" -DryRun $false
   ```

### Prerequisites
- PowerShell 5.1 or later
- Microsoft.Graph PowerShell modules (auto-installed by scripts)
- SharePoint Administrator or Global Administrator role
- Azure App Registration (for certificate/secret auth) or Interactive authentication

### Basic Workflow

1. **Scan for permissions:**
   ```powershell
   # Option A: Check all permissions in John's OneDrive
   .\Check-OneDrivePermissions.ps1 -OneDriveOwnerEmail "john.doe@company.com"
   
   # Option B: Find all OneDrives where Jane has access
   .\Scan-AllOneDrivesForUser.ps1 -TargetUserEmail "jane.smith@company.com"
   ```

2. **Preview permission removal:**
   ```powershell
   # Dry run (safe preview mode)
   .\Remove-UserPermissions.ps1 -InputCsvPath "ScanResults.csv" -TargetUserEmail "jane.smith@company.com" -DryRun $true
   ```

3. **Execute permission removal:**
   ```powershell
   # Live removal (after confirming dry run results)
   .\Remove-UserPermissions.ps1 -InputCsvPath "ScanResults.csv" -TargetUserEmail "jane.smith@company.com" -DryRun $false
   ```

## 📋 Detailed Tool Documentation

### Check-OneDrivePermissions.ps1

**Purpose:** Comprehensive scan of a single user's OneDrive to identify all shared files and folders.

**Parameters:**
- `OneDriveOwnerEmail` - Email of the OneDrive owner to scan
- `OutputPath` - Custom output file path (optional)
- `TenantId` - Azure AD Tenant ID (optional)
- `ClientId` - Application Client ID (optional)
- `CertificatePath` - Certificate file path for auth (optional)
- `AuthMethod` - Authentication method: Interactive, Certificate, or ClientSecret

**Example:**
```powershell
.\Check-OneDrivePermissions.ps1 -OneDriveOwnerEmail "user@company.com" -OutputPath "C:\Reports\UserPermissions.csv"
```

**Output Columns:**
- `OneDriveOwner` - Owner of the OneDrive
- `ItemPath` - Path to the shared item
- `ItemType` - File or Folder
- `PermissionId` - Unique permission identifier
- `UserEmail` - Email of user with access
- `UserDisplayName` - Display name of user
- `PermissionType` - Read, Write, etc.
- `ShareType` - Direct, Inherited, etc.

### Scan-AllOneDrivesForUser.ps1

**Purpose:** Organization-wide scan to find all OneDrives where a specific user has been granted permissions.

**Parameters:**
- `TargetUserEmail` - Email of user to search for
- `OutputPath` - Custom output file path (optional)
- `TenantId` - Azure AD Tenant ID (optional)
- `ClientId` - Application Client ID (optional)
- `CertificatePath` - Certificate file path for auth (optional)
- `AuthMethod` - Authentication method: Interactive, Certificate, or ClientSecret
- `MaxConcurrentJobs` - Number of parallel scans (default: 5)

**Example:**
```powershell
.\Scan-AllOneDrivesForUser.ps1 -TargetUserEmail "contractor@company.com" -OutputPath "C:\Reports\ContractorAccess.csv"
```

**Output Columns:**
- `OneDriveOwnerEmail` - Owner of the OneDrive being accessed
- `OneDriveOwnerDisplayName` - Display name of OneDrive owner
- `ItemPath` - Path to the shared item
- `ItemType` - File or Folder
- `PermissionId` - Unique permission identifier
- `TargetUserEmail` - Email of target user (the one being searched for)
- `TargetUserDisplayName` - Display name of target user
- `PermissionType` - Read, Write, etc.
- `ShareType` - Direct, Inherited, etc.

### Remove-UserPermissions.ps1

**Purpose:** Safely remove user permissions from OneDrives based on CSV reports from scanning tools.

**🛡️ CRITICAL SAFETY FEATURES:**
- **Never removes permissions from user's own OneDrive**
- **Dry-run mode by default** (no changes until explicitly enabled)
- **Confirmation prompts** before making changes
- **Comprehensive audit logging** of all actions
- **Batch processing** with rate limiting protection

**Parameters:**
- `InputCsvPath` - Path to CSV from scanning tools
- `TargetUserEmail` - Email of user whose permissions to remove
- `DryRun` - Preview mode (default: $true)
- `BatchSize` - Permissions per batch (default: 10)
- `Force` - Skip confirmation prompts
- `LogPath` - Custom audit log path
- `TenantId` - Azure AD Tenant ID (optional)
- `ClientId` - Application Client ID (optional)
- `CertificatePath` - Certificate file path for auth (optional)
- `AuthMethod` - Authentication method: Interactive, Certificate, or ClientSecret

**Example:**
```powershell
# Safe preview (recommended first step)
.\Remove-UserPermissions.ps1 -InputCsvPath "C:\Reports\UserAccess.csv" -TargetUserEmail "user@company.com" -DryRun $true

# Execute actual removal (after confirming preview)
.\Remove-UserPermissions.ps1 -InputCsvPath "C:\Reports\UserAccess.csv" -TargetUserEmail "user@company.com" -DryRun $false
```

**Safety Example:**
```
❌ SKIP: jane.smith@company.com owns this OneDrive - will not remove her own permissions
✅ REMOVE: jane.smith@company.com access from john.doe@company.com's OneDrive
```

## 🔐 Authentication Methods

### Interactive Authentication (Recommended for first-time use)
```powershell
.\Check-OneDrivePermissions.ps1 -OneDriveOwnerEmail "user@company.com"
# Will prompt for browser login
```

### Certificate Authentication (Recommended for automation)
```powershell
.\Scan-AllOneDrivesForUser.ps1 -TargetUserEmail "user@company.com" -CertificatePath "C:\cert.pfx" -ClientId "your-app-id" -TenantId "your-tenant-id"
```

### Client Secret Authentication
```powershell
.\Remove-UserPermissions.ps1 -InputCsvPath "report.csv" -TargetUserEmail "user@company.com" -ClientId "your-app-id" -TenantId "your-tenant-id" -AuthMethod "ClientSecret"
# Will prompt for client secret
```

## 📊 Common Use Cases

### Employee Departure Cleanup
1. **Scan for all access:**
   ```powershell
   .\Scan-AllOneDrivesForUser.ps1 -TargetUserEmail "departing.employee@company.com"
   ```

2. **Preview removal:**
   ```powershell
   .\Remove-UserPermissions.ps1 -InputCsvPath "ScanResults.csv" -TargetUserEmail "departing.employee@company.com" -DryRun $true
   ```

3. **Execute removal:**
   ```powershell
   .\Remove-UserPermissions.ps1 -InputCsvPath "ScanResults.csv" -TargetUserEmail "departing.employee@company.com" -DryRun $false
   ```

### Contractor Access Audit
1. **Check what contractor can access:**
   ```powershell
   .\Scan-AllOneDrivesForUser.ps1 -TargetUserEmail "contractor@external.com"
   ```

2. **Audit specific user's sharing:**
   ```powershell
   .\Check-OneDrivePermissions.ps1 -OneDriveOwnerEmail "project.manager@company.com"
   ```

### Department Reorganization
1. **Scan affected users:**
   ```powershell
   .\Scan-AllOneDrivesForUser.ps1 -TargetUserEmail "user1@company.com"
   .\Scan-AllOneDrivesForUser.ps1 -TargetUserEmail "user2@company.com"
   ```

2. **Bulk permission cleanup:**
   ```powershell
   .\Remove-UserPermissions.ps1 -InputCsvPath "User1Access.csv" -TargetUserEmail "user1@company.com"
   .\Remove-UserPermissions.ps1 -InputCsvPath "User2Access.csv" -TargetUserEmail "user2@company.com"
   ```

## 🔧 Required Azure App Permissions

For certificate or client secret authentication, your Azure app registration needs:

### Microsoft Graph API Permissions:
- `Files.Read.All` - Read all files (for scanning tools)
- `Files.ReadWrite.All` - Read and modify files (for removal tool)
- `User.Read.All` - Read user profiles
- `Sites.Read.All` - Read SharePoint sites
- `Sites.ReadWrite.All` - Modify SharePoint sites (for removal tool)

### Admin Consent Required:
- All permissions require admin consent
- Grant consent in Azure Portal → App registrations → API permissions

## 📝 Output Files and Logs

### CSV Reports
- **Location:** Script directory (unless custom path specified)
- **Format:** UTF-8 CSV with headers
- **Naming:** Timestamped for easy identification

### Audit Logs (Remove-UserPermissions.ps1)
- **Contains:** All actions taken, including skipped items
- **Columns:** Timestamp, Type, Message, OneDriveOwner, ItemPath, PermissionId, Action, TargetUser
- **Use Case:** Compliance reporting, troubleshooting, rollback planning

## ⚠️ Important Safety Notes

### Before Running Removal Tool:
1. **Always run with `-DryRun $true` first**
2. **Review the preview output carefully**
3. **Verify the target user email is correct**
4. **Ensure you have proper authorization**
5. **Keep audit logs for compliance**

### Safety Mechanisms:
- 🛡️ **User's own OneDrive is always protected**
- 📋 **Dry-run mode prevents accidental changes**
- ⏸️ **Batch processing with delays prevents API overload**
- 📊 **Comprehensive logging for audit trails**
- ❓ **Confirmation prompts before destructive actions**

## 🐛 Troubleshooting

### Drive ID Format Issues:
The `Remove-UserPermissions.ps1` script has been enhanced to handle multiple drive ID formats. In some Microsoft 365 environments, the Graph API may return multiple drive IDs or malformed drive IDs. The script now:

- Automatically detects and cleans up drive IDs that contain spaces or multiple values
- Selects the primary business OneDrive when multiple drives are returned
- Uses improved folder navigation with case-insensitive and fuzzy matching
- Provides detailed logging about drive ID handling

If you encounter the error `The provided drive id appears to be malformed, or does not represent a valid drive`, the script will automatically handle this by cleaning up the drive ID format.

### Common Issues:

**Authentication Failed:**
- Verify app permissions and admin consent
- Check certificate is uploaded to Azure app
- Ensure user has required admin roles

**Script Runs Slowly:**
- Normal for large tenants (API rate limiting)
- Consider running during off-hours
- Monitor progress output

**Permission Not Found:**
- Permission may have been removed manually
- User may have lost access since scan
- Check audit logs for details

**CSV Format Issues:**
- Ensure using UTF-8 encoding
- Don't modify CSV structure
- Check for special characters in file paths

### Getting Help:
1. Check PowerShell execution policy: `Get-ExecutionPolicy`
2. Verify module installation: `Get-Module Microsoft.Graph.* -ListAvailable`
3. Test basic Graph connection: `Connect-MgGraph -Scopes "User.Read"`
4. Review audit logs for detailed error information

## 📚 Additional Resources

- [Microsoft Graph PowerShell SDK Documentation](https://docs.microsoft.com/en-us/powershell/microsoftgraph/)
- [Azure App Registration Guide](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [OneDrive API Documentation](https://docs.microsoft.com/en-us/onedrive/developer/)

---

**Author:** Ruben Quispe  
**Created:** 2025-01-06  
**Version:** 1.1  
**Last Updated:** 2025-06-24  
**License:** MIT License - See LICENSE file for details

## 🔄 Recent Updates

### Version 1.1 (2025-06-24)
- Enhanced `Remove-UserPermissions.ps1` to handle multiple drive ID formats
- Improved folder navigation with case-insensitive and fuzzy matching
- Added better error handling and debugging information
- Fixed issues with malformed drive IDs returned by Microsoft Graph API
- Updated documentation with more detailed examples
