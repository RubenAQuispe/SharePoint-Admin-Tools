# SharePoint Administration Tools

A comprehensive collection of PowerShell scripts for SharePoint Online and OneDrive administration tasks. These tools help administrators efficiently manage permissions, analyze sharing patterns, and generate detailed reports.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-API-green)
![License](https://img.shields.io/badge/license-MIT-brightgreen)

## 🚀 Quick Start

### Prerequisites
- Windows PowerShell 5.1+ or PowerShell Core 7+
- Microsoft Graph PowerShell SDK
- SharePoint Administrator or Global Administrator role
- Azure App Registration (for certificate/secret authentication)

### Installation
1. Clone this repository:
   ```powershell
   git clone https://github.com/RubenAQuispe/SharePoint-Admin-Tools.git
   cd SharePoint-Admin-Tools
   ```

2. Install required modules (scripts will auto-install if missing):
   ```powershell
   Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
   Install-Module Microsoft.Graph.Users -Scope CurrentUser
   Install-Module Microsoft.Graph.Files -Scope CurrentUser
   ```

## 📁 Available Scripts

### 🔍 OneDrive Permissions Checker
**Location:** `scripts/OneDrive-Permissions/`

Analyzes OneDrive sharing between users and generates detailed permission reports.

**Features:**
- ✅ Deep recursive folder analysis
- ✅ Multiple authentication methods (Interactive, Certificate, Client Secret)
- ✅ Comprehensive permission detection
- ✅ CSV export with detailed metadata
- ✅ Real-time progress tracking
- ✅ Debug mode for troubleshooting

**Quick Usage:**
```powershell
cd scripts/OneDrive-Permissions
.\Check-OneDrivePermissions.ps1 -SourceUserEmail "john.doe@company.com" -TargetUserEmail "jane.smith@company.com"
```

**[📖 Full Documentation](scripts/OneDrive-Permissions/README.md)**

---

### 🔜 Upcoming Scripts
- **SharePoint Site Permission Analyzer** - Comprehensive site-level permission auditing
- **Teams Channel File Scanner** - Analyze file sharing in Teams channels
- **External Sharing Reporter** - Identify files shared with external users
- **Permission Cleanup Utility** - Bulk permission management tools

## 🔐 Authentication Setup

### Option 1: Interactive Authentication (Recommended for testing)
No setup required - you'll be prompted to sign in via browser.

### Option 2: Certificate Authentication (Recommended for automation)

1. **Create a self-signed certificate:**
   ```powershell
   $cert = New-SelfSignedCertificate -Subject "CN=SharePointPermissionsAudit" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256 -NotAfter (Get-Date).AddDays(365)
   
   $certPassword = ConvertTo-SecureString -String "YourSecurePassword" -Force -AsPlainText
   Export-PfxCertificate -Cert $cert -FilePath C:\Temp\SharePointAudit.pfx -Password $certPassword
   Export-Certificate -Cert $cert -FilePath C:\Temp\SharePointAudit.cer
   ```

2. **Register Azure Application:**
   - Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
   - Click "New registration"
   - Name: "SharePoint Admin Tools"
   - Supported account types: "Accounts in this organizational directory only"
   - Click "Register"

3. **Configure API Permissions:**
   - Go to "API permissions" → "Add a permission" → "Microsoft Graph" → "Application permissions"
   - Add these permissions:
     - `Files.Read.All`
     - `User.Read.All`
     - `Sites.Read.All`
   - Click "Grant admin consent"

4. **Upload Certificate:**
   - Go to "Certificates & secrets" → "Certificates" → "Upload certificate"
   - Upload the `.cer` file created in step 1

5. **Usage:**
   ```powershell
   .\Check-OneDrivePermissions.ps1 -AuthMethod Certificate -CertificatePath "C:\Temp\SharePointAudit.pfx" -ClientId "your-app-id" -TenantId "your-tenant-id"
   ```

### Option 3: Client Secret Authentication

Follow steps 2-3 from Certificate Authentication, then:

1. **Create Client Secret:**
   - Go to "Certificates & secrets" → "Client secrets" → "New client secret"
   - Copy the secret value (save it securely!)

2. **Usage:**
   ```powershell
   $clientSecret = ConvertTo-SecureString "your-client-secret" -AsPlainText -Force
   .\Check-OneDrivePermissions.ps1 -AuthMethod ClientSecret -ClientId "your-app-id" -TenantId "your-tenant-id" -ClientSecret $clientSecret
   ```

## 📊 Sample Reports

The scripts generate comprehensive CSV reports with details like:
- File/folder paths and types
- Permission levels (Read, Write, Owner, etc.)
- Share types (Direct, Inherited, Link-based)
- User information (Name, Email, ID)
- Timestamps (Created, Modified)
- Debug information for troubleshooting

## 🛡️ Security Best Practices

- ✅ **Never commit certificates or secrets** to version control
- ✅ **Use certificate authentication** for production automation
- ✅ **Rotate certificates and secrets** regularly
- ✅ **Follow principle of least privilege** for app permissions
- ✅ **Monitor app usage** through Azure AD logs
- ✅ **Store certificates securely** (Azure Key Vault recommended)

## 🤝 Contributing

We welcome contributions! Please see our [Contributing Guidelines](CONTRIBUTING.md) for details.

### How to Contribute
1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support

- 📖 **Documentation**: Check the `docs/` folder for detailed guides
- 🐛 **Issues**: Report bugs via [GitHub Issues](https://github.com/RubenAQuispe/SharePoint-Admin-Tools/issues)
- 💬 **Discussions**: Join the conversation in [GitHub Discussions](https://github.com/RubenAQuispe/SharePoint-Admin-Tools/discussions)

## 🔗 Useful Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [SharePoint REST API Reference](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/rest-api-reference)
- [PowerShell Gallery - Microsoft.Graph](https://www.powershellgallery.com/packages/Microsoft.Graph)
- [Azure App Registration Guide](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)

---

**⭐ If this project helps you, please consider giving it a star!**

Made with ❤️ by the SharePoint Admin Tools Community
