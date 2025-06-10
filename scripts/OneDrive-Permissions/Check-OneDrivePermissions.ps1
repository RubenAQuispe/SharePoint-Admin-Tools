<#
.SYNOPSIS
    Check OneDrive file and folder permissions between two users using Microsoft Graph API.

.DESCRIPTION
    This script analyzes a user's OneDrive to find all files and folders that have been shared 
    with another specific user. It uses Microsoft Graph API to retrieve detailed permission 
    information and generates a comprehensive CSV report.

.PARAMETER SourceUserEmail
    Email address of the OneDrive owner (user whose OneDrive will be analyzed)

.PARAMETER TargetUserEmail
    Email address of the user to check permissions for (recipient of shared files)

.PARAMETER OutputPath
    Full path where the CSV report will be saved (optional, defaults to script directory)

.PARAMETER TenantId
    Azure AD Tenant ID (optional, will prompt if not provided)

.PARAMETER ClientId
    Application Client ID for authentication (optional, uses interactive auth if not provided)

.PARAMETER Recursive
    Include subfolders in the analysis (default: $true)

.PARAMETER IncludeInherited
    Include inherited permissions in the report (default: $true)

.EXAMPLE
    .\Check-OneDrivePermissions.ps1 -SourceUserEmail "john.doe@company.com" -TargetUserEmail "jane.smith@company.com"

.EXAMPLE
    .\Check-OneDrivePermissions.ps1 -SourceUserEmail "john.doe@company.com" -TargetUserEmail "jane.smith@company.com" -OutputPath "C:\Reports\ShareReport.csv"

.NOTES
    Author: Ruben Quispe
    Created: 2025-01-06
    Requires: Microsoft.Graph PowerShell Module
    Permissions Required: Files.Read.All, User.Read.All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Email of OneDrive owner")]
    [string]$SourceUserEmail,
    
    [Parameter(Mandatory = $false, HelpMessage = "Email of user to check permissions for")]
    [string]$TargetUserEmail,
    
    [Parameter(Mandatory = $false, HelpMessage = "Output file path")]
    [string]$OutputPath,
    
    [Parameter(Mandatory = $false, HelpMessage = "Azure AD Tenant ID")]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false, HelpMessage = "Application Client ID")]
    [string]$ClientId,
    
    [Parameter(Mandatory = $false, HelpMessage = "Path to certificate file (.pfx)")]
    [string]$CertificatePath,
    
    [Parameter(Mandatory = $false, HelpMessage = "Certificate password")]
    [SecureString]$CertificatePassword,
    
    [Parameter(Mandatory = $false, HelpMessage = "Client Secret for app authentication")]
    [SecureString]$ClientSecret,
    
    [Parameter(Mandatory = $false, HelpMessage = "Authentication method")]
    [string]$AuthMethod,
    
    [Parameter(Mandatory = $false, HelpMessage = "Include subfolders")]
    [bool]$Recursive = $true,
    
    [Parameter(Mandatory = $false, HelpMessage = "Include inherited permissions")]
    [bool]$IncludeInherited = $true,
    
    [Parameter(Mandatory = $false, HelpMessage = "Enable detailed debug output")]
    [switch]$DebugMode
)

# Global variables
$Script:Results = @()
$Script:ProcessedItems = 0
$Script:CachedCertificate = $null
$Script:StartTime = $null

#region Helper Functions

function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

function Install-RequiredModules {
    Write-ColorOutput "Checking required PowerShell modules..." "Yellow"
    
    $RequiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Users", 
        "Microsoft.Graph.Files"
    )
    
    foreach ($Module in $RequiredModules) {
        if (!(Get-Module -ListAvailable -Name $Module)) {
            Write-ColorOutput "Installing module: $Module" "Yellow"
            try {
                Install-Module -Name $Module -Force -AllowClobber -Scope CurrentUser
                Write-ColorOutput "Successfully installed $Module" "Green"
            }
            catch {
                Write-ColorOutput "Failed to install $Module. Error: $($_.Exception.Message)" "Red"
                return $false
            }
        }
        else {
            Write-ColorOutput "Module $Module is already installed" "Green"
        }
    }
    
    # Import modules
    foreach ($Module in $RequiredModules) {
        try {
            Import-Module $Module -Force
        }
        catch {
            Write-ColorOutput "Failed to import $Module. Error: $($_.Exception.Message)" "Red"
            return $false
        }
    }
    
    return $true
}

function Connect-ToGraph {
    Write-ColorOutput "Connecting to Microsoft Graph using $($Script:AuthMethod) authentication..." "Yellow"
    
    $RequiredScopes = @("Files.Read.All", "User.Read.All", "Sites.Read.All")
    
    try {
        switch ($Script:AuthMethod) {
            "Certificate" {
                # Certificate-based authentication - use cached certificate
                if ($Script:CachedCertificate) {
                    $Certificate = $Script:CachedCertificate
                } else {
                    if ($Script:CertificatePassword) {
                        $Certificate = Get-PfxCertificate -FilePath $Script:CertificatePath -Password $Script:CertificatePassword
                    } else {
                        $Certificate = Get-PfxCertificate -FilePath $Script:CertificatePath
                    }
                    $Script:CachedCertificate = $Certificate
                }
                
                Connect-MgGraph -ClientId $Script:ClientId -TenantId $Script:TenantId -Certificate $Certificate
                Write-ColorOutput "Connected using certificate authentication." "Green"
            }
            
            "ClientSecret" {
                # Client Secret authentication
                $ClientSecretCredential = New-Object System.Management.Automation.PSCredential($Script:ClientId, $Script:ClientSecret)
                Connect-MgGraph -TenantId $Script:TenantId -ClientSecretCredential $ClientSecretCredential
                Write-ColorOutput "Connected using client secret authentication." "Green"
            }
            
            "Interactive" {
                # Interactive authentication
                if ($Script:TenantId) {
                    Connect-MgGraph -TenantId $Script:TenantId -Scopes $RequiredScopes
                } else {
                    Connect-MgGraph -Scopes $RequiredScopes
                }
                Write-ColorOutput "Connected using interactive authentication." "Green"
            }
            
            default {
                # Fallback to interactive
                Connect-MgGraph -Scopes $RequiredScopes
                Write-ColorOutput "Connected using interactive authentication (fallback)." "Green"
            }
        }
        
        # Verify connection and display tenant info
        $Context = Get-MgContext
        if ($Context) {
            Write-ColorOutput "Successfully connected to tenant: $($Context.TenantId)" "Green"
            
            # Try to get tenant details for confirmation
            try {
                $TenantInfo = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($TenantInfo) {
                    Write-ColorOutput "Organization: $($TenantInfo.DisplayName)" "Green"
                }
            }
            catch {
                # Ignore if can't get org info
            }
            
            return $true
        } else {
            Write-ColorOutput "Failed to establish Graph context." "Red"
            return $false
        }
    }
    catch {
        Write-ColorOutput "Failed to connect to Microsoft Graph. Error: $($_.Exception.Message)" "Red"
        
        # Provide specific guidance based on auth method
        switch ($Script:AuthMethod) {
            "Certificate" {
                Write-ColorOutput "Certificate authentication troubleshooting:" "Yellow"
                Write-ColorOutput "- Verify certificate is uploaded to your Azure app registration" "Yellow"
                Write-ColorOutput "- Ensure app has required permissions: Files.Read.All, User.Read.All, Sites.Read.All" "Yellow"
                Write-ColorOutput "- Check that admin consent has been granted" "Yellow"
            }
            "ClientSecret" {
                Write-ColorOutput "Client secret authentication troubleshooting:" "Yellow"
                Write-ColorOutput "- Verify client secret is valid and not expired" "Yellow"
                Write-ColorOutput "- Ensure app has required permissions: Files.Read.All, User.Read.All, Sites.Read.All" "Yellow"
                Write-ColorOutput "- Check that admin consent has been granted" "Yellow"
            }
            "Interactive" {
                Write-ColorOutput "Interactive authentication troubleshooting:" "Yellow"
                Write-ColorOutput "- Ensure you have SharePoint Administrator or Global Administrator role" "Yellow"
                Write-ColorOutput "- Try clearing browser cache or using incognito mode" "Yellow"
            }
        }
        
        return $false
    }
}

function Get-UserDriveId {
    param(
        [string]$UserEmail
    )
    
    try {
        Write-ColorOutput "Getting drive information for user: $UserEmail" "Yellow"
        
        # Get user information
        $User = Get-MgUser -UserId $UserEmail -ErrorAction Stop
        
        # Get user's drives
        $Drives = Get-MgUserDrive -UserId $User.Id -ErrorAction Stop
        
        if ($Drives) {
            # If multiple drives, get the default/primary drive (OneDrive for Business)
            if ($Drives -is [array]) {
                $PrimaryDrive = $Drives | Where-Object { $_.DriveType -eq "business" } | Select-Object -First 1
                if (-not $PrimaryDrive) {
                    $PrimaryDrive = $Drives[0]  # Fallback to first drive
                }
            } else {
                $PrimaryDrive = $Drives
            }
            
            Write-ColorOutput "Found primary drive for user: $($PrimaryDrive.Id)" "Green"
            return $PrimaryDrive.Id
        }
        else {
            Write-ColorOutput "No drive found for user: $UserEmail" "Red"
            return $null
        }
    }
    catch {
        Write-ColorOutput "Error getting drive for user $UserEmail : $($_.Exception.Message)" "Red"
        return $null
    }
}

function Get-UserByEmail {
    param(
        [string]$Email
    )
    
    try {
        $User = Get-MgUser -UserId $Email -ErrorAction Stop
        return $User
    }
    catch {
        Write-ColorOutput "Error finding user $Email : $($_.Exception.Message)" "Red"
        return $null
    }
}

function Get-DriveItemPermissions {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [string]$ItemPath,
        [string]$ItemType,
        [string]$TargetUserId,
        [string]$TargetUserEmail
    )
    
    try {
        $Permissions = Get-MgDriveItemPermission -DriveId $DriveId -DriveItemId $ItemId -ErrorAction Stop
        
        if ($DebugMode) {
            Write-ColorOutput "    🔍 Found $($Permissions.Count) permissions for: $ItemPath" "DarkCyan"
        }
        
        foreach ($Permission in $Permissions) {
            $MatchFound = $false
            $PermissionLevel = ""
            $ShareType = ""
            $IsInherited = $false
            $InheritedFrom = ""
            $DebugInfo = @()
            
            # Enhanced debugging - show ALL permission data
            if ($DebugMode) {
                Write-ColorOutput "    📋 Permission ID: $($Permission.Id)" "DarkGray"
            }
            
            # Determine permission level with more sources
            if ($Permission.Roles) {
                $PermissionLevel = $Permission.Roles -join ", "
                $DebugInfo += "Roles: $PermissionLevel"
            }
            
            # Check if permission is inherited
            if ($Permission.InheritedFrom) {
                $IsInherited = $true
                $InheritedFrom = $Permission.InheritedFrom.Path
                $DebugInfo += "Inherited from: $InheritedFrom"
            }
            
            # Skip inherited permissions if not requested
            if ($IsInherited -and -not $IncludeInherited) {
                if ($DebugMode) {
                    Write-ColorOutput "    ⏭️  Skipping inherited permission" "DarkYellow"
                }
                continue
            }
            
            # Enhanced user matching - check multiple properties
            $UserInfo = $null
            $UserName = ""
            $UserEmail = ""
            $UserId = ""
            
            # Check GrantedToV2 (modern property)
            if ($Permission.GrantedToV2) {
                if ($Permission.GrantedToV2.User) {
                    $UserInfo = $Permission.GrantedToV2.User
                    $UserName = $UserInfo.DisplayName
                    $UserEmail = $UserInfo.Email
                    $UserId = $UserInfo.Id
                    $DebugInfo += "GrantedToV2 User: $UserName ($UserEmail) ID: $UserId"
                }
                elseif ($Permission.GrantedToV2.SiteUser) {
                    $UserInfo = $Permission.GrantedToV2.SiteUser
                    $UserName = $UserInfo.DisplayName
                    $UserEmail = $UserInfo.Email
                    $UserId = $UserInfo.Id
                    $DebugInfo += "GrantedToV2 SiteUser: $UserName ($UserEmail) ID: $UserId"
                }
            }
            
            # Check GrantedTo (legacy property)
            if (-not $UserInfo -and $Permission.GrantedTo) {
                if ($Permission.GrantedTo.User) {
                    $UserInfo = $Permission.GrantedTo.User
                    $UserName = $UserInfo.DisplayName
                    $UserEmail = $UserInfo.Email
                    $UserId = $UserInfo.Id
                    $DebugInfo += "GrantedTo User: $UserName ($UserEmail) ID: $UserId"
                }
            }
            
            # Check grantedToIdentitiesV2 (newer property)
            if (-not $UserInfo -and $Permission.GrantedToIdentitiesV2) {
                foreach ($Identity in $Permission.GrantedToIdentitiesV2) {
                    if ($Identity.User) {
                        $UserInfo = $Identity.User
                        $UserName = $UserInfo.DisplayName
                        $UserEmail = $UserInfo.Email
                        $UserId = $UserInfo.Id
                        $DebugInfo += "GrantedToIdentitiesV2 User: $UserName ($UserEmail) ID: $UserId"
                        break
                    }
                }
            }
            
            # Log debug information
            if ($DebugMode) {
                foreach ($Debug in $DebugInfo) {
                    Write-ColorOutput "    📄 $Debug" "DarkGray"
                }
            }
            
            # Enhanced matching logic
            if ($UserInfo) {
                # Match by User ID (most reliable)
                if ($UserId -and $UserId -eq $TargetUserId) {
                    $MatchFound = $true
                    $ShareType = "Direct User Permission (ID Match)"
                    if ($DebugMode) {
                        Write-ColorOutput "    ✅ MATCH FOUND: User ID matches!" "Green"
                    }
                }
                # Match by Email (case insensitive)
                elseif ($UserEmail -and $UserEmail.ToLower() -eq $TargetUserEmail.ToLower()) {
                    $MatchFound = $true
                    $ShareType = "Direct User Permission (Email Match)"
                    if ($DebugMode) {
                        Write-ColorOutput "    ✅ MATCH FOUND: Email matches!" "Green"
                    }
                }
                # Match by Display Name (contains target user's name)
                elseif ($UserName -and $UserName -like "*$($TargetUserEmail.Split('@')[0])*") {
                    $MatchFound = $true
                    $ShareType = "Direct User Permission (Name Match)"
                    if ($DebugMode) {
                        Write-ColorOutput "    ✅ MATCH FOUND: Display name contains user!" "Green"
                    }
                }
                else {
                    if ($DebugMode) {
                        Write-ColorOutput "    ❌ No match: Target=$TargetUserEmail/$TargetUserId vs Found=$UserEmail/$UserId" "DarkRed"
                    }
                }
            }
            
            # Check sharing links
            if (-not $MatchFound -and $Permission.Link) {
                if ($Permission.Link.Type) {
                    $ShareType = "Sharing Link - " + $Permission.Link.Type
                    $DebugInfo += "Link Type: $($Permission.Link.Type)"
                    # Include sharing links as potential access
                    $MatchFound = $true
                    if ($DebugMode) {
                        Write-ColorOutput "    🔗 SHARING LINK FOUND: $($Permission.Link.Type)" "Cyan"
                    }
                }
            }
            
            if ($MatchFound) {
                $Result = [PSCustomObject]@{
                    ItemPath = $ItemPath
                    ItemType = $ItemType
                    PermissionLevel = $PermissionLevel
                    ShareType = $ShareType
                    IsInherited = $IsInherited
                    InheritedFrom = $InheritedFrom
                    PermissionId = $Permission.Id
                    UserDisplayName = $UserName
                    UserEmail = $UserEmail
                    UserId = $UserId
                    LinkWebUrl = if ($Permission.Link) { $Permission.Link.WebUrl } else { "" }
                    CreatedDateTime = $Permission.AdditionalProperties.createdDateTime
                    LastModifiedDateTime = $Permission.AdditionalProperties.lastModifiedDateTime
                }
                
                $Script:Results += $Result
                if ($DebugMode) {
                    Write-ColorOutput "    📝 Added permission to results!" "Green"
                }
            }
        }
    }
    catch {
        Write-ColorOutput "Error getting permissions for item $ItemPath : $($_.Exception.Message)" "Red"
    }
}

function Process-DriveItem {
    param(
        [string]$DriveId,
        [object]$Item,
        [string]$TargetUserId,
        [string]$TargetUserEmail,
        [string]$BasePath = ""
    )
    
    $Script:ProcessedItems++
    $ItemPath = if ($BasePath) { "$BasePath/$($Item.Name)" } else { $Item.Name }
    
    # Update progress with new system
    Show-ProgressUpdate -CurrentItem $ItemPath -Status "Analyzing permissions"
    
    # Enhanced folder/file detection logic with multiple fallbacks
    $ItemType = "Unknown"
    
    # Force file detection for obvious file extensions FIRST
    if ($Item.Name -match '\.(docx|pptx|xlsx|pdf|txt|jpg|png|gif|zip|mp4|avi|csv)$') {
        $ItemType = "File"
        Write-ColorOutput "  🔧 FORCED FILE detection for extension: $ItemPath" "Yellow"
    }
    # Force folder detection for obvious folder names
    elseif ($Item.Name -match '^(Documents|Pictures|Desktop|Downloads|Music|Videos|Apps|Attachments|Personal Folders|Microsoft Teams Chat Files|Whiteboards|Zoom Recordings|New folder|Files\(\d+\)|AFSP Logos|AFSP Templates|Templates)') {
        $ItemType = "Folder"
        Write-ColorOutput "  📁 FORCED FOLDER detection for name: $ItemPath" "Yellow"
    }
    # Check if item has a size property but is 0 (often indicates folder)
    elseif ($Item.Size -eq 0 -and $Item.Name -notmatch '\.[a-zA-Z0-9]+$') {
        $ItemType = "Folder"
        Write-ColorOutput "  📁 FOLDER detected (zero size, no extension): $ItemPath" "Yellow"
    }
    # Primary detection - check File property first (most reliable)
    elseif ($Item.File -ne $null) {
        $ItemType = "File"
    }
    # Secondary detection - check Folder property
    elseif ($Item.Folder -ne $null) {
        $ItemType = "Folder"
    }
    # Fallback detection - use file extension pattern for files
    elseif ($Item.Name -match '\.[a-zA-Z0-9]+$') {
        $ItemType = "File"
        Write-ColorOutput "  ⚠️  Using extension fallback for: $ItemPath" "Yellow"
    }
    # Last resort - assume folder if no extension
    else {
        $ItemType = "Folder"
        Write-ColorOutput "  ⚠️  Using no-extension fallback (assuming folder): $ItemPath" "Yellow"
    }
    
    # Debug logging for detection issues
    if ($DebugMode) {
        Write-ColorOutput "    🔍 Item: $($Item.Name)" "DarkGray"
        Write-ColorOutput "    📝 Has File property: $(if ($Item.File) { 'YES' } else { 'NO' })" "DarkGray"
        Write-ColorOutput "    📁 Has Folder property: $(if ($Item.Folder) { 'YES' } else { 'NO' })" "DarkGray"
        Write-ColorOutput "    📋 Detected as: $ItemType" "DarkGray"
    }
    
    if ($ItemType -eq "Folder") {
        Write-ColorOutput "  � Checking FOLDER permissions: $ItemPath" "Cyan"
    } else {
        Write-ColorOutput "  � Checking FILE permissions: $ItemPath" "Blue"
    }
    
    # Always check permissions first - this is crucial for folders!
    Get-DriveItemPermissions -DriveId $DriveId -ItemId $Item.Id -ItemPath $ItemPath -ItemType $ItemType -TargetUserId $TargetUserId -TargetUserEmail $TargetUserEmail
    
    # Enhanced recursive processing logic - use our determined ItemType, not raw properties
    if ($Recursive -and $ItemType -eq "Folder") {
        try {
            Write-ColorOutput "  └─ Processing folder contents of: $ItemPath" "DarkGray"
            
            $Children = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $Item.Id -ErrorAction Stop
            
            if ($Children -and $Children.Count -gt 0) {
                Write-ColorOutput "  └─ Found $($Children.Count) items in folder" "DarkGray"
                
                foreach ($Child in $Children) {
                    # Validate child object before processing
                    if ($Child -and $Child.Id -and $Child.Name) {
                        Process-DriveItem -DriveId $DriveId -Item $Child -TargetUserId $TargetUserId -TargetUserEmail $TargetUserEmail -BasePath $ItemPath
                    } else {
                        Write-ColorOutput "  └─ Skipping invalid child object in $ItemPath" "DarkYellow"
                    }
                }
            } else {
                Write-ColorOutput "  └─ Empty folder: $ItemPath" "DarkGray"
            }
        }
        catch {
            Write-ColorOutput "Error processing folder $ItemPath : $($_.Exception.Message)" "Red"
        }
    } elseif ($ItemType -eq "File") {
        # Explicitly log files - no children to process
        Write-ColorOutput "  └─ File (no children to process): $ItemPath" "DarkGray"
    } else {
        # Unknown type - log for debugging
        Write-ColorOutput "  └─ Unknown item type ($ItemType) - skipping children: $ItemPath" "DarkYellow"
    }
}

function Show-ProgressUpdate {
    param(
        [string]$CurrentItem,
        [string]$Status = "Processing"
    )
    
    $Elapsed = (Get-Date) - $Script:StartTime
    $ElapsedStr = "{0:mm\:ss}" -f $Elapsed
    
    # Calculate rate
    $Rate = if ($Elapsed.TotalSeconds -gt 0) { 
        [math]::Round($Script:ProcessedItems / $Elapsed.TotalSeconds, 1) 
    } else { 0 }
    
    # Update progress bar
    $ProgressParams = @{
        Activity = "OneDrive Permission Analysis"
        Status = "$Status - Processed: $($Script:ProcessedItems) | Permissions Found: $($Script:Results.Count) | Rate: $Rate items/sec | Elapsed: $ElapsedStr"
        CurrentOperation = "Current: $CurrentItem"
    }
    
    Write-Progress @ProgressParams
    
    # Show periodic updates in console
    if ($Script:ProcessedItems % 50 -eq 0 -or $Script:ProcessedItems -eq 1) {
        Write-ColorOutput "Progress Update: $($Script:ProcessedItems) items processed, $($Script:Results.Count) permissions found" "Cyan"
    }
}

function Export-Results {
    param(
        [string]$OutputPath
    )
    
    if ($Script:Results.Count -eq 0) {
        Write-ColorOutput "No shared items found between the specified users." "Yellow"
        return
    }
    
    try {
        $Script:Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-ColorOutput "Report exported successfully to: $OutputPath" "Green"
        Write-ColorOutput "Total shared items found: $($Script:Results.Count)" "Green"
    }
    catch {
        Write-ColorOutput "Error exporting results: $($_.Exception.Message)" "Red"
    }
}

function Get-UserInput {
    param(
        [string]$Prompt,
        [bool]$Required = $true,
        [switch]$IsFilePath
    )
    
    do {
        $Input = Read-Host $Prompt
        if (-not $Required -or $Input.Trim()) {
            if ($IsFilePath -and $Input.Trim()) {
                if (-not (Test-Path $Input.Trim())) {
                    Write-ColorOutput "File not found: $($Input.Trim()). Please enter a valid file path." "Red"
                    continue
                }
            }
            return $Input.Trim()
        }
        Write-ColorOutput "This field is required. Please enter a value." "Red"
    } while ($Required)
}

function Show-AuthMethodMenu {
    Write-Host ""
    Write-ColorOutput "Select Authentication Method:" "Yellow"
    Write-Host "1. Interactive Login (Browser-based - Recommended for first time)"
    Write-Host "2. Certificate Authentication (.pfx file)"
    Write-Host "3. App Registration with Client Secret"
    Write-Host ""
    
    do {
        $Choice = Read-Host "Choose authentication method [1-3]"
        switch ($Choice) {
            "1" { return "Interactive" }
            "2" { return "Certificate" }
            "3" { return "ClientSecret" }
            default { Write-ColorOutput "Invalid choice. Please enter 1, 2, or 3." "Red" }
        }
    } while ($true)
}

function Get-SecureInput {
    param(
        [string]$Prompt
    )
    
    do {
        $SecureInput = Read-Host $Prompt -AsSecureString
        if ($SecureInput.Length -gt 0) {
            return $SecureInput
        }
        Write-ColorOutput "This field is required. Please enter a value." "Red"
    } while ($true)
}

function Test-GuidFormat {
    param(
        [string]$InputString
    )
    
    try {
        $null = [System.Guid]::Parse($InputString)
        return $true
    }
    catch {
        return $false
    }
}

function Get-AuthenticationDetails {
    # Auto-detect authentication method if not specified
    if (-not $Script:AuthMethod) {
        if ($Script:CertificatePath) {
            $Script:AuthMethod = "Certificate"
            Write-ColorOutput "Auto-detected Certificate authentication method." "Green"
        } elseif ($Script:ClientSecret) {
            $Script:AuthMethod = "ClientSecret"
            Write-ColorOutput "Auto-detected Client Secret authentication method." "Green"
        } else {
            $Script:AuthMethod = Show-AuthMethodMenu
        }
    }
    
    Write-Host ""
    Write-ColorOutput "Setting up $($Script:AuthMethod) authentication..." "Yellow"
    
    switch ($Script:AuthMethod) {
        "Certificate" {
            if (-not $Script:CertificatePath) {
                $Script:CertificatePath = Get-UserInput "Enter certificate file path (.pfx)" -IsFilePath
            }
            
            # Test if certificate can be loaded and cache it
            try {
                if ($Script:CertificatePassword) {
                    $Script:CachedCertificate = Get-PfxCertificate -FilePath $Script:CertificatePath -Password $Script:CertificatePassword -ErrorAction Stop
                } else {
                    try {
                        $Script:CachedCertificate = Get-PfxCertificate -FilePath $Script:CertificatePath -ErrorAction Stop
                    }
                    catch {
                        Write-ColorOutput "Certificate appears to be password protected." "Yellow"
                        $Script:CertificatePassword = Get-SecureInput "Enter certificate password"
                        $Script:CachedCertificate = Get-PfxCertificate -FilePath $Script:CertificatePath -Password $Script:CertificatePassword -ErrorAction Stop
                    }
                }
                Write-ColorOutput "Certificate loaded successfully." "Green"
            }
            catch {
                Write-ColorOutput "Error loading certificate: $($_.Exception.Message)" "Red"
                return $false
            }
            
            if (-not $Script:ClientId) {
                do {
                    $Script:ClientId = Get-UserInput "Enter Application Client ID (GUID format)"
                    if (-not (Test-GuidFormat $Script:ClientId)) {
                        Write-ColorOutput "Invalid GUID format. Please enter a valid Client ID." "Red"
                        $Script:ClientId = ""
                    }
                } while (-not $Script:ClientId)
            }
            
            if (-not $Script:TenantId) {
                do {
                    $Script:TenantId = Get-UserInput "Enter Tenant ID (GUID format)"
                    if (-not (Test-GuidFormat $Script:TenantId)) {
                        Write-ColorOutput "Invalid GUID format. Please enter a valid Tenant ID." "Red"
                        $Script:TenantId = ""
                    }
                } while (-not $Script:TenantId)
            }
        }
        
        "ClientSecret" {
            if (-not $Script:ClientId) {
                do {
                    $Script:ClientId = Get-UserInput "Enter Application Client ID (GUID format)"
                    if (-not (Test-GuidFormat $Script:ClientId)) {
                        Write-ColorOutput "Invalid GUID format. Please enter a valid Client ID." "Red"
                        $Script:ClientId = ""
                    }
                } while (-not $Script:ClientId)
            }
            
            if (-not $Script:TenantId) {
                do {
                    $Script:TenantId = Get-UserInput "Enter Tenant ID (GUID format)"
                    if (-not (Test-GuidFormat $Script:TenantId)) {
                        Write-ColorOutput "Invalid GUID format. Please enter a valid Tenant ID." "Red"
                        $Script:TenantId = ""
                    }
                } while (-not $Script:TenantId)
            }
            
            if (-not $Script:ClientSecret) {
                $Script:ClientSecret = Get-SecureInput "Enter Client Secret"
            }
        }
        
        "Interactive" {
            Write-ColorOutput "Interactive authentication selected. You will be prompted to sign in via browser." "Green"
            # No additional parameters needed for interactive auth
        }
    }
    
    return $true
}

#endregion

#region Main Script

function Main {
    # Display banner
    Write-ColorOutput "==================================" "Cyan"
    Write-ColorOutput "OneDrive Permission Checker v1.0" "Cyan"
    Write-ColorOutput "==================================" "Cyan"
    Write-Host ""
    
    # Install required modules
    if (-not (Install-RequiredModules)) {
        Write-ColorOutput "Failed to install required modules. Exiting." "Red"
        return
    }
    
    # Copy parameters to script scope for global access
    $Script:AuthMethod = if ($AuthMethod) { $AuthMethod } else { $null }
    $Script:TenantId = $TenantId
    $Script:ClientId = $ClientId
    $Script:CertificatePath = $CertificatePath
    $Script:CertificatePassword = $CertificatePassword
    $Script:ClientSecret = $ClientSecret
    
    # Setup authentication details
    if (-not (Get-AuthenticationDetails)) {
        Write-ColorOutput "Failed to setup authentication. Exiting." "Red"
        return
    }
    
    # Get user input if parameters not provided
    if (-not $SourceUserEmail) {
        $SourceUserEmail = Get-UserInput "Enter the source user's email (OneDrive owner)"
    }
    
    if (-not $TargetUserEmail) {
        $TargetUserEmail = Get-UserInput "Enter the target user's email (recipient to check)"
    }
    
    if (-not $OutputPath) {
        $DefaultPath = Join-Path (Get-Location) "OneDrivePermissionsReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $OutputPath = Get-UserInput "Enter output file path or directory (or press Enter for default: $DefaultPath)" $false
        if (-not $OutputPath) {
            $OutputPath = $DefaultPath
        }
    }
    
    # Handle directory path vs full file path
    if (Test-Path $OutputPath -PathType Container) {
        # It's a directory, generate filename
        $Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $Filename = "OneDrivePermissionsReport_${SourceUserEmail}_to_${TargetUserEmail}_$Timestamp.csv"
        $Filename = $Filename -replace '[<>:"/\\|?*]', '_'  # Replace invalid chars
        $OutputPath = Join-Path $OutputPath $Filename
        Write-ColorOutput "Auto-generated filename: $OutputPath" "Green"
    }
    
    # Connect to Microsoft Graph
    if (-not (Connect-ToGraph)) {
        Write-ColorOutput "Failed to connect to Microsoft Graph. Exiting." "Red"
        return
    }
    
    # Get source user's drive
    $SourceDriveId = Get-UserDriveId -UserEmail $SourceUserEmail
    if (-not $SourceDriveId) {
        Write-ColorOutput "Could not find OneDrive for source user: $SourceUserEmail" "Red"
        return
    }
    
    # Get target user information
    $TargetUser = Get-UserByEmail -Email $TargetUserEmail
    if (-not $TargetUser) {
        Write-ColorOutput "Could not find target user: $TargetUserEmail" "Red"
        return
    }
    
    Write-ColorOutput "Starting permission analysis..." "Yellow"
    Write-ColorOutput "Source User: $SourceUserEmail" "White"
    Write-ColorOutput "Target User: $TargetUserEmail" "White"
    Write-ColorOutput "Output Path: $OutputPath" "White"
    Write-ColorOutput "Recursive: $Recursive" "White"
    Write-ColorOutput "Include Inherited: $IncludeInherited" "White"
    Write-Host ""
    
    # Initialize start time for progress tracking
    $Script:StartTime = Get-Date
    Write-ColorOutput "Processing started at: $($Script:StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" "Green"
    Write-Host ""
    
    # First, check permissions on the root OneDrive folder itself
    try {
        Write-ColorOutput "🏠 Checking ROOT OneDrive permissions..." "Magenta"
        $Script:ProcessedItems++
        Show-ProgressUpdate -CurrentItem "OneDrive Root" -Status "Checking root permissions"
        
        # Check permissions on the root drive item
        Get-DriveItemPermissions -DriveId $SourceDriveId -ItemId "root" -ItemPath "/OneDrive Root" -ItemType "Folder" -TargetUserId $TargetUser.Id -TargetUserEmail $TargetUserEmail
        
        Write-ColorOutput "✅ Root OneDrive permission check completed" "Green"
        Write-Host ""
    }
    catch {
        Write-ColorOutput "Warning: Could not check root OneDrive permissions: $($_.Exception.Message)" "Yellow"
    }
    
    # Get root items and process them
    try {
        Write-ColorOutput "📂 Getting root folder contents..." "Yellow"
        $RootItems = Get-MgDriveItemChild -DriveId $SourceDriveId -DriveItemId "root" -ErrorAction Stop
        
        Write-ColorOutput "Found $($RootItems.Count) items in root folder" "Green"
        
        foreach ($Item in $RootItems) {
            Process-DriveItem -DriveId $SourceDriveId -Item $Item -TargetUserId $TargetUser.Id -TargetUserEmail $TargetUserEmail
        }
        
        Write-Progress -Activity "Processing OneDrive Items" -Completed
        
        # Export results
        Write-Host ""
        Write-ColorOutput "Processing complete. Exporting results..." "Yellow"
        Export-Results -OutputPath $OutputPath
        
        # Display summary
        Write-Host ""
        Write-ColorOutput "==================================" "Cyan"
        Write-ColorOutput "SUMMARY" "Cyan"
        Write-ColorOutput "==================================" "Cyan"
        Write-ColorOutput "Items processed: $Script:ProcessedItems" "White"
        Write-ColorOutput "Shared items found: $($Script:Results.Count)" "White"
        Write-ColorOutput "Report location: $OutputPath" "White"
        
        if ($Script:Results.Count -gt 0) {
            Write-Host ""
            Write-ColorOutput "Preview of found permissions:" "Yellow"
            $Script:Results | Select-Object -First 5 | Format-Table -AutoSize
        }
    }
    catch {
        Write-ColorOutput "Error processing OneDrive items: $($_.Exception.Message)" "Red"
    }
    finally {
        # Disconnect from Graph
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Write-ColorOutput "Disconnected from Microsoft Graph." "Green"
        }
        catch {
            # Ignore disconnect errors
        }
    }
}

#endregion

# Execute main function
try {
    Main
}
catch {
    Write-ColorOutput "Unexpected error: $($_.Exception.Message)" "Red"
    Write-ColorOutput "Stack trace: $($_.ScriptStackTrace)" "Red"
}
finally {
    Write-Host ""
    Write-ColorOutput "Script execution completed." "Cyan"
}
