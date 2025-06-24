<#
.SYNOPSIS
    Scan SharePoint sites to find folders containing ransomware-encrypted files.

.DESCRIPTION
    This script scans all SharePoint sites (or a specific site) in the organization to identify 
    folders that contain files with ransomware extensions (e.g., .NITROGEN). It provides a 
    comprehensive report of infected locations to help with restoration planning.

.PARAMETER TargetSiteUrl
    (Optional) Specific SharePoint site URL to scan. If not provided, scans all sites.

.PARAMETER RansomwareExtension
    File extension to search for (e.g., ".NITROGEN", ".LOCKED"). Default: ".NITROGEN"

.PARAMETER OutputPath
    Full path where the CSV report will be saved (optional, defaults to script directory)

.PARAMETER TenantId
    Azure AD Tenant ID

.PARAMETER ClientId
    Application Client ID for authentication

.PARAMETER CertificatePath
    Path to certificate file (.pfx) for certificate authentication

.PARAMETER CertificatePassword
    Certificate password for certificate authentication

.PARAMETER ClientSecret
    Client Secret for app authentication

.PARAMETER AuthMethod
    Authentication method: Interactive, Certificate, or ClientSecret

.PARAMETER ResumeFromFile
    Resume scanning from a previous checkpoint file

.PARAMETER DebugMode
    Enable detailed debug output for troubleshooting

.EXAMPLE
    .\Scan-SharePointForRansomware.ps1

.EXAMPLE
    .\Scan-SharePointForRansomware.ps1 -RansomwareExtension ".LOCKED" -OutputPath "C:\Reports\RansomwareReport.csv"

.EXAMPLE
    .\Scan-SharePointForRansomware.ps1 -ResumeFromFile "C:\Reports\checkpoint.json"

.EXAMPLE
    .\Scan-SharePointForRansomware.ps1 -TargetSiteUrl "https://yourtenant.sharepoint.com/sites/TestSite" -DebugMode

.EXAMPLE
    .\Scan-SharePointForRansomware.ps1 -TargetSiteUrl "https://yourtenant.sharepoint.com/sites/ProjectAlpha" -RansomwareExtension ".NITROGEN" -OutputPath "C:\Reports\SingleSite_Test.csv"

.NOTES
    Author: Ruben Quispe
    Created: 2025-01-06
    Based on: Scan-AllOneDrivesForUser.ps1
    Requires: Microsoft.Graph PowerShell Module
    Permissions Required: Files.Read.All, Sites.Read.All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Specific SharePoint site URL to scan")]
    [string]$TargetSiteUrl,
    
    [Parameter(Mandatory = $false, HelpMessage = "Ransomware file extension to search for")]
    [string]$RansomwareExtension = ".NITROGEN",
    
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
    
    [Parameter(Mandatory = $false, HelpMessage = "Resume from checkpoint file")]
    [string]$ResumeFromFile,
    
    [Parameter(Mandatory = $false, HelpMessage = "Enable detailed debug output")]
    [switch]$DebugMode
)

# Global variables
$Script:Results = @()
$Script:ProcessedSites = 0
$Script:ProcessedLibraries = 0
$Script:ProcessedFolders = 0
$Script:ProcessedFiles = 0
$Script:InfectedFolders = 0
$Script:SkippedSites = 0
$Script:ErrorSites = @()
$Script:CachedCertificate = $null
$Script:StartTime = $null
$Script:CheckpointFile = ""
$Script:LastCheckpointSave = $null
$Script:CheckpointInterval = 120 # Save checkpoint every 2 minutes
$Script:Checkpoint = @{
    ProcessedSites = @()
    ProcessedLibraries = @()
    CurrentSiteId = ""
    CurrentLibraryId = ""
    StartTime = $null
    RansomwareExtension = ""
    LastSavedAt = $null
    Results = @()
    Statistics = @{
        ProcessedSites = 0
        ProcessedLibraries = 0
        ProcessedFolders = 0
        ProcessedFiles = 0
        InfectedFolders = 0
        SkippedSites = 0
        ErrorSites = @()
    }
}

# Comprehensive SharePoint supported file extensions
$Script:KnownFileExtensions = @(
    # Microsoft Office Documents
    '.doc', '.docx', '.docm', '.dot', '.dotx', '.dotm',
    '.xls', '.xlsx', '.xlsm', '.xlsb', '.xlt', '.xltx', '.xltm', '.csv',
    '.ppt', '.pptx', '.pptm', '.pot', '.potx', '.potm', '.pps', '.ppsx', '.ppsm',
    '.pub', '.one', '.onenote', '.mpp', '.vsd', '.vsdx', '.visio',
    
    # PDF and Text Documents
    '.pdf', '.txt', '.rtf', '.odt', '.ods', '.odp', '.pages', '.numbers', '.key',
    
    # Images
    '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif', '.svg', '.webp', 
    '.ico', '.heic', '.heif', '.raw', '.psd', '.ai', '.eps', '.indd',
    
    # Audio Files
    '.mp3', '.wav', '.wma', '.aac', '.flac', '.ogg', '.m4a', '.aiff', '.au',
    
    # Video Files
    '.mp4', '.avi', '.mov', '.wmv', '.flv', '.mkv', '.webm', '.m4v', '.3gp',
    '.mpg', '.mpeg', '.m2v', '.asf', '.vob', '.ts', '.mts',
    
    # Archives and Compressed Files
    '.zip', '.rar', '.7z', '.tar', '.gz', '.bz2', '.xz', '.cab', '.ace', '.arj',
    
    # Web and Code Files
    '.html', '.htm', '.css', '.js', '.json', '.xml', '.xsl', '.xslt',
    '.aspx', '.ascx', '.asmx', '.ashx', '.php', '.jsp', '.asp',
    '.py', '.java', '.cs', '.cpp', '.c', '.h', '.hpp', '.rb', '.pl', '.sh',
    '.sql', '.bat', '.cmd', '.ps1', '.psm1', '.psd1',
    
    # Email and Calendar
    '.msg', '.eml', '.ics', '.vcs', '.vcf',
    
    # Ebook and Document Formats
    '.epub', '.mobi', '.azw', '.fb2', '.lit',
    
    # CAD and Design
    '.dwg', '.dxf', '.step', '.iges', '.stl', '.obj',
    
    # Data and Database
    '.mdb', '.accdb', '.dbf', '.sqlite', '.db',
    
    # Font Files
    '.ttf', '.otf', '.woff', '.woff2', '.eot',
    
    # Common Ransomware Extensions
    '.locked', '.encrypted', '.crypto', '.crypt', '.vault', '.cerber',
    '.locky', '.zepto', '.thor', '.aesir', '.odin', '.shit', '.fuck',
    '.xxx', '.ttt', '.micro', '.mp3', '.kimcilware', '.encrypted',
    '.R5A', '.R4A', '.R3A', '.biz', '.XRNT', '.XTBL', '.crypt',
    '.vault', '.EXX', '.ezz', '.ecc', '.ABC', '.aaa', '.zzz',
    '.xyz', '.CCC', '.DOC', '.russian', '.putin', '.weapologize',
    '.nuclear55', '.comrade', '.coverton', '.EnCiPhErEd',
    '.LeChiffre', '.keypass', '.SecureCrypted', '.AlWaysDT',
    '.carote', '.deadbolt', '.lol', '.OMG', '.RRK', '.encryptedRSA',
    '.crjoker', '.EnCrYpTeD', '.coded', '.bugware', '.frtrss',
    '.CCCRRRPPP', '.pzdc', '.good', '.LOL', '.PoAr2w', '.paym',
    '.NITROGEN' # The specific extension we're scanning for
)

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
        "Microsoft.Graph.Sites",
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
    
    $RequiredScopes = @("Files.Read.All", "Sites.Read.All")
    
    try {
        switch ($Script:AuthMethod) {
            "Certificate" {
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
                $ClientSecretCredential = New-Object System.Management.Automation.PSCredential($Script:ClientId, $Script:ClientSecret)
                Connect-MgGraph -TenantId $Script:TenantId -ClientSecretCredential $ClientSecretCredential
                Write-ColorOutput "Connected using client secret authentication." "Green"
            }
            
            "Interactive" {
                if ($Script:TenantId) {
                    Connect-MgGraph -TenantId $Script:TenantId -Scopes $RequiredScopes
                } else {
                    Connect-MgGraph -Scopes $RequiredScopes
                }
                Write-ColorOutput "Connected using interactive authentication." "Green"
            }
            
            default {
                Connect-MgGraph -Scopes $RequiredScopes
                Write-ColorOutput "Connected using interactive authentication (fallback)." "Green"
            }
        }
        
        $Context = Get-MgContext
        if ($Context) {
            Write-ColorOutput "Successfully connected to tenant: $($Context.TenantId)" "Green"
            return $true
        } else {
            Write-ColorOutput "Failed to establish Graph context." "Red"
            return $false
        }
    }
    catch {
        Write-ColorOutput "Failed to connect to Microsoft Graph. Error: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Get-AllSharePointSites {
    if ($TargetSiteUrl) {
        Write-ColorOutput "Targeting specific site: $TargetSiteUrl" "Yellow"
        
        try {
            $TargetSite = $null
            
            # Method 1: Try by original URL
            try {
                $TargetSite = Get-MgSite -SiteId $TargetSiteUrl -ErrorAction Stop
                Write-ColorOutput "✅ Successfully retrieved site by original URL" "Green"
            }
            catch {
                Write-ColorOutput "⚠️  Method 1 failed: $($_.Exception.Message)" "Yellow"
                
                # Method 2: Try by site name search
                if ($TargetSiteUrl -match "/sites/([^/]+)") {
                    $SiteName = $Matches[1]
                    try {
                        $SearchResults = Get-MgSite -Search $SiteName -ErrorAction Stop
                        $TargetSite = $SearchResults | Where-Object { 
                            $_.DisplayName -eq $SiteName -or 
                            $_.WebUrl -like "*$SiteName*"
                        } | Select-Object -First 1
                        
                        if ($TargetSite) {
                            Write-ColorOutput "✅ Found site via name search: $($TargetSite.DisplayName)" "Green"
                        }
                    }
                    catch {
                        Write-ColorOutput "⚠️  Method 2 failed: $($_.Exception.Message)" "Yellow"
                    }
                }
            }
            
            if ($TargetSite) {
                Write-ColorOutput "📍 Target Site Found:" "Green"
                Write-ColorOutput "  - Name: $($TargetSite.DisplayName)" "Green"
                Write-ColorOutput "  - URL: $($TargetSite.WebUrl)" "Green"
                return @($TargetSite)
            } else {
                Write-ColorOutput "❌ ERROR: Could not find site with URL: $TargetSiteUrl" "Red"
                return @()
            }
        }
        catch {
            Write-ColorOutput "❌ Error retrieving target site: $($_.Exception.Message)" "Red"
            return @()
        }
    }
    
    # Enhanced site discovery with multiple fallback methods
    Write-ColorOutput "🔍 Getting all SharePoint sites in the organization..." "Yellow"
    Write-ColorOutput "   Using enhanced discovery methods for better compatibility..." "Gray"
    
    $AllSites = @()
    $MethodUsed = ""
    
    # Method 1: Get-MgSite -All (Most comprehensive, works in most tenants)
    try {
        Write-ColorOutput "   📋 Method 1: Attempting Get-MgSite -All..." "Gray"
        $AllSites = Get-MgSite -All -ErrorAction Stop
        $MethodUsed = "Get-MgSite -All"
        Write-ColorOutput "   ✅ Method 1 SUCCESS: Found $($AllSites.Count) sites" "Green"
    }
    catch {
        Write-ColorOutput "   ⚠️  Method 1 failed: $($_.Exception.Message)" "Yellow"
        
        # Method 2: Get-MgSite with high Top value and pagination
        try {
            Write-ColorOutput "   📋 Method 2: Attempting Get-MgSite -Top 999..." "Gray"
            $AllSites = Get-MgSite -Top 999 -ErrorAction Stop
            $MethodUsed = "Get-MgSite -Top 999"
            Write-ColorOutput "   ✅ Method 2 SUCCESS: Found $($AllSites.Count) sites" "Green"
        }
        catch {
            Write-ColorOutput "   ⚠️  Method 2 failed: $($_.Exception.Message)" "Yellow"
            
            # Method 3: Basic Get-MgSite call
            try {
                Write-ColorOutput "   📋 Method 3: Attempting basic Get-MgSite..." "Gray"
                $AllSites = Get-MgSite -ErrorAction Stop
                $MethodUsed = "Get-MgSite (basic)"
                Write-ColorOutput "   ✅ Method 3 SUCCESS: Found $($AllSites.Count) sites" "Green"
            }
            catch {
                Write-ColorOutput "   ⚠️  Method 3 failed: $($_.Exception.Message)" "Yellow"
                
                # Method 4: Search with wildcard (original method)
                try {
                    Write-ColorOutput "   📋 Method 4: Attempting Get-MgSite -Search '*'..." "Gray"
                    $AllSites = Get-MgSite -Search "*" -Top 200 -ErrorAction Stop
                    $MethodUsed = "Get-MgSite -Search '*'"
                    Write-ColorOutput "   ✅ Method 4 SUCCESS: Found $($AllSites.Count) sites" "Green"
                }
                catch {
                    Write-ColorOutput "   ⚠️  Method 4 failed: $($_.Exception.Message)" "Yellow"
                    
                    # Method 5: Alternative search patterns
                    try {
                        Write-ColorOutput "   📋 Method 5: Attempting alternative search patterns..." "Gray"
                        $SearchTerms = @("", "site", "team", "project", "department")
                        foreach ($Term in $SearchTerms) {
                            try {
                                $SearchResults = Get-MgSite -Search $Term -ErrorAction Stop
                                if ($SearchResults -and $SearchResults.Count -gt 0) {
                                    $AllSites += $SearchResults
                                }
                            }
                            catch {
                                # Continue to next search term
                            }
                        }
                        
                        # Remove duplicates
                        $AllSites = $AllSites | Sort-Object Id -Unique
                        $MethodUsed = "Alternative search patterns"
                        Write-ColorOutput "   ✅ Method 5 SUCCESS: Found $($AllSites.Count) unique sites" "Green"
                    }
                    catch {
                        Write-ColorOutput "   ❌ Method 5 failed: $($_.Exception.Message)" "Red"
                        
                        # Method 6: Last resort - try to get root site and enumerate
                        try {
                            Write-ColorOutput "   📋 Method 6: Attempting root site enumeration..." "Gray"
                            $Context = Get-MgContext
                            if ($Context -and $Context.TenantId) {
                                # Try to get tenant root site and enumerate from there
                                $TenantInfo = Get-MgOrganization -OrganizationId $Context.TenantId -ErrorAction SilentlyContinue
                                if ($TenantInfo) {
                                    # Try common SharePoint URL patterns
                                    $TenantName = $TenantInfo.VerifiedDomains | Where-Object {$_.IsDefault -eq $true} | Select-Object -ExpandProperty Name
                                    if ($TenantName) {
                                        $SharePointDomain = $TenantName -replace '\.onmicrosoft\.com$', ''
                                        $RootSiteUrl = "https://$SharePointDomain.sharepoint.com"
                                        
                                        try {
                                            $RootSite = Get-MgSite -SiteId $RootSiteUrl -ErrorAction Stop
                                            $AllSites = @($RootSite)
                                            $MethodUsed = "Root site only (limited)"
                                            Write-ColorOutput "   ⚠️  Method 6 LIMITED SUCCESS: Found root site only" "Yellow"
                                        }
                                        catch {
                                            Write-ColorOutput "   ❌ Method 6 failed: Cannot access root site" "Red"
                                        }
                                    }
                                }
                            }
                        }
                        catch {
                            Write-ColorOutput "   ❌ Method 6 failed: $($_.Exception.Message)" "Red"
                        }
                    }
                }
            }
        }
    }
    
    # Validate and filter results
    if ($AllSites -and $AllSites.Count -gt 0) {
        Write-ColorOutput "✅ Site discovery successful using: $MethodUsed" "Green"
        
        # Filter out OneDrive personal sites and invalid entries
        $FilteredSites = $AllSites | Where-Object {
            $_ -and 
            $_.WebUrl -and 
            $_.DisplayName -and
            $_.WebUrl -notlike "*/personal/*" -and
            $_.DisplayName -ne "OneDrive" -and
            $_.WebUrl -notlike "*-my.sharepoint.com*"
        }
        
        # Additional filtering for common SharePoint site patterns
        $SharePointSites = $FilteredSites | Where-Object {
            $_.WebUrl -like "*/sites/*" -or 
            $_.WebUrl -like "*/teams/*" -or
            $_.WebUrl -match "^https://[^/]+\.sharepoint\.com/?$"  # Root site
        }
        
        Write-ColorOutput "📊 Site Discovery Results:" "Cyan"
        Write-ColorOutput "   - Total sites found: $($AllSites.Count)" "White"
        Write-ColorOutput "   - After filtering: $($FilteredSites.Count)" "White"
        Write-ColorOutput "   - SharePoint sites to scan: $($SharePointSites.Count)" "Green"
        Write-ColorOutput "   - Method used: $MethodUsed" "Gray"
        
        if ($DebugMode -and $SharePointSites.Count -gt 0) {
            Write-ColorOutput "🔍 Sample sites found:" "Gray"
            $SharePointSites | Select-Object -First 5 | ForEach-Object {
                Write-ColorOutput "   - $($_.DisplayName): $($_.WebUrl)" "DarkGray"
            }
        }
        
        return $SharePointSites
    }
    else {
        Write-ColorOutput "❌ ERROR: All site discovery methods failed!" "Red"
        Write-ColorOutput "   This may indicate:" "Yellow"
        Write-ColorOutput "   - Insufficient permissions (need Sites.Read.All)" "Yellow"
        Write-ColorOutput "   - Tenant policy restrictions" "Yellow"
        Write-ColorOutput "   - Network connectivity issues" "Yellow"
        Write-ColorOutput "   - Application registration issues" "Yellow"
        
        # Provide troubleshooting guidance
        Write-ColorOutput "" "White"
        Write-ColorOutput "🔧 Troubleshooting suggestions:" "Cyan"
        Write-ColorOutput "   1. Verify app permissions include Sites.Read.All" "White"
        Write-ColorOutput "   2. Check if admin consent has been granted" "White"
        Write-ColorOutput "   3. Try running: Get-MgSite manually to test" "White"
        Write-ColorOutput "   4. Verify certificate has SharePoint permissions" "White"
        
        return @()
    }
}

function Check-FolderForRansomware {
    param(
        [string]$SiteId,
        [string]$DriveId,
        [object]$Folder,
        [string]$SiteName,
        [string]$SiteUrl,
        [string]$LibraryName,
        [string]$BasePath = "",
        [int]$CurrentDepth = 0
    )
    
    $FolderPath = if ($BasePath) { "$BasePath/$($Folder.Name)" } else { $Folder.Name }
    $FullPath = "$LibraryName/$FolderPath"
    
    try {
        $Items = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $Folder.Id -ErrorAction Stop
        
        if ($DebugMode) {
            Write-ColorOutput "      📁 Checking folder: $FullPath ($($Items.Count) items)" "DarkGray"
        }
        
        $Script:ProcessedFolders++
        $EncryptedFiles = @()
        $EncryptedFileFound = $false
        
        # Check each item for ransomware files
        foreach ($Item in $Items) {
            if ($Item -and $Item.Name) {
                $Script:ProcessedFiles++
                
                if ($Item.Name -like "*$RansomwareExtension") {
                    $EncryptedFiles += $Item.Name
                    $EncryptedFileFound = $true
                    
                    if ($DebugMode) {
                        Write-ColorOutput "        🦠 INFECTED FILE FOUND: $($Item.Name)" "Red"
                    }
                }
            }
        }
        
        # If ransomware found, add to results
        if ($EncryptedFileFound) {
            $Script:InfectedFolders++
            
            $Result = [PSCustomObject]@{
                SiteName = $SiteName
                SiteUrl = $SiteUrl
                Library = $LibraryName
                FolderPath = $FullPath
                InfectedFiles = $EncryptedFiles -join "; "
                TotalItemsInFolder = $Items.Count
                NeedsRestoration = "YES"
                ScanDateTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                FolderDepth = $CurrentDepth
            }
            
            $Script:Results += $Result
            Write-ColorOutput "    🚨 INFECTED FOLDER: $FullPath" "Red"
        }
        
        # Process subfolders with comprehensive file extension detection
        foreach ($Item in $Items) {
            if ($Item -and $Item.Id -and $Item.Name) {
                # Get file extension
                $Extension = ""
                if ($Item.Name -match '\.([^.]+)$') {
                    $Extension = $Matches[1].ToLower()
                    $Extension = ".$Extension"
                }
                
                # Check if it's a known file extension
                if ($Script:KnownFileExtensions -contains $Extension) {
                    # Known file extension = File (skip processing)
                    if ($DebugMode) {
                        Write-ColorOutput "        [FILE] $($Item.Name) (extension: $Extension)" "DarkGray"
                    }
                } else {
                    # Unknown extension or no extension = Folder (process it)
                    if ($DebugMode) {
                        Write-ColorOutput "        [FOLDER] $($Item.Name)" "Green"
                    }
                    Check-FolderForRansomware -SiteId $SiteId -DriveId $DriveId -Folder $Item -SiteName $SiteName -SiteUrl $SiteUrl -LibraryName $LibraryName -BasePath $FolderPath -CurrentDepth ($CurrentDepth + 1)
                }
            }
        }
    }
    catch {
        if ($_.Exception.Message -notlike "*Null reference*") {
            Write-ColorOutput "[ERROR] Error processing folder $FullPath : $($_.Exception.Message)" "Red"
        }
    }
}

function Process-SharePointSite {
    param(
        [object]$Site
    )
    
    try {
        Write-ColorOutput "  🔍 Scanning site: $($Site.DisplayName)" "Yellow"
        $Script:ProcessedSites++
        
        # Get drives (document libraries)
        $Drives = Get-MgSiteDrive -SiteId $Site.Id -ErrorAction Stop
        
        Write-ColorOutput "  📚 Found $($Drives.Count) document libraries" "Green"
        
        foreach ($Drive in $Drives) {
            try {
                Write-ColorOutput "    📖 Scanning library: $($Drive.Name)" "Yellow"
                $Script:ProcessedLibraries++
                
                # Get root items
                $RootItems = Get-MgDriveItemChild -DriveId $Drive.Id -DriveItemId "root" -ErrorAction Stop
                
                foreach ($Item in $RootItems) {
                    if ($Item -and $Item.Id -and $Item.Name) {
                        # Get file extension
                        $Extension = ""
                        if ($Item.Name -match '\.([^.]+)$') {
                            $Extension = $Matches[1].ToLower()
                            $Extension = ".$Extension"
                        }
                        
                        # Check if it's a known file extension
                        if ($Script:KnownFileExtensions -contains $Extension) {
                            # Known file extension = File (skip processing)
                            if ($DebugMode) {
                                Write-ColorOutput "      [FILE] $($Item.Name) (extension: $Extension)" "DarkGray"
                            }
                        } else {
                            # Unknown extension or no extension = Folder (process it)
                            if ($DebugMode) {
                                Write-ColorOutput "      [FOLDER] $($Item.Name)" "Green"
                            }
                            Check-FolderForRansomware -SiteId $Site.Id -DriveId $Drive.Id -Folder $Item -SiteName $Site.DisplayName -SiteUrl $Site.WebUrl -LibraryName $Drive.Name
                        }
                    }
                }
            }
            catch {
                Write-ColorOutput "    [ERROR] Error processing library $($Drive.Name): $($_.Exception.Message)" "Red"
            }
        }
        
        Write-ColorOutput "  ✅ Completed scanning site: $($Site.DisplayName)" "Green"
        return $true
    }
    catch {
        Write-ColorOutput "  ❌ Error processing site $($Site.DisplayName): $($_.Exception.Message)" "Red"
        $Script:ErrorSites += $Site.DisplayName
        return $false
    }
}

function Export-Results {
    param(
        [string]$OutputPath
    )
    
    try {
        # Export main results
        if ($Script:Results.Count -gt 0) {
            $Script:Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
            Write-ColorOutput "🚨 RANSOMWARE REPORT exported successfully to: $OutputPath" "Green"
        }
        
        # Create summary report
        $SummaryPath = $OutputPath -replace '\.csv$', '_Summary.txt'
        $Summary = @"
SharePoint Ransomware Scan Summary
Ransomware Extension: $RansomwareExtension
Scan Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Sites Processed: $Script:ProcessedSites
Sites Skipped: $Script:SkippedSites
Folders Analyzed: $Script:ProcessedFolders
Files Scanned: $Script:ProcessedFiles
Infected Folders Found: $Script:InfectedFolders
Errors: $($Script:ErrorSites.Count)

$(if ($Script:InfectedFolders -eq 0) {
"Result: No ransomware-infected folders found. Your SharePoint environment appears clean."
} else {
"Result: $Script:InfectedFolders infected folder(s) found. Immediate restoration required.

Infected Locations:
$($Script:Results | ForEach-Object { "- $($_.SiteName): $($_.FolderPath)" } | Out-String)
"})
"@
        $Summary | Out-File -FilePath $SummaryPath -Encoding UTF8
        Write-ColorOutput "Total infected folders found: $Script:InfectedFolders" "Green"
        Write-ColorOutput "Summary report saved to: $SummaryPath" "Green"
        
    }
    catch {
        Write-ColorOutput "Error exporting results: $($_.Exception.Message)" "Red"
    }
}

function Get-AuthenticationDetails {
    if (-not $Script:AuthMethod) {
        if ($Script:CertificatePath) {
            $Script:AuthMethod = "Certificate"
        } elseif ($Script:ClientSecret) {
            $Script:AuthMethod = "ClientSecret"
        } else {
            $Script:AuthMethod = "Interactive"
        }
    }
    
    return $true
}

function Save-Checkpoint {
    param(
        [string]$CheckpointPath
    )
    
    try {
        $Script:Checkpoint.Results = $Script:Results
        $Script:Checkpoint.LastSavedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $Script:Checkpoint.Statistics.ProcessedSites = $Script:ProcessedSites
        $Script:Checkpoint.Statistics.ProcessedLibraries = $Script:ProcessedLibraries
        $Script:Checkpoint.Statistics.ProcessedFolders = $Script:ProcessedFolders
        $Script:Checkpoint.Statistics.ProcessedFiles = $Script:ProcessedFiles
        $Script:Checkpoint.Statistics.InfectedFolders = $Script:InfectedFolders
        $Script:Checkpoint.Statistics.SkippedSites = $Script:SkippedSites
        $Script:Checkpoint.Statistics.ErrorSites = $Script:ErrorSites
        
        $Script:Checkpoint | ConvertTo-Json -Depth 10 | Out-File -FilePath $CheckpointPath -Encoding UTF8
        
        if ($DebugMode) {
            Write-ColorOutput "💾 Checkpoint saved: $CheckpointPath" "Gray"
        }
        
        $Script:LastCheckpointSave = Get-Date
    }
    catch {
        Write-ColorOutput "⚠️  Warning: Failed to save checkpoint: $($_.Exception.Message)" "Yellow"
    }
}

function Load-Checkpoint {
    param(
        [string]$CheckpointPath
    )
    
    try {
        if (Test-Path $CheckpointPath) {
            $LoadedCheckpoint = Get-Content -Path $CheckpointPath -Raw | ConvertFrom-Json
            
            # Restore checkpoint data
            $Script:Checkpoint.ProcessedSites = $LoadedCheckpoint.ProcessedSites
            $Script:Checkpoint.ProcessedLibraries = $LoadedCheckpoint.ProcessedLibraries
            $Script:Checkpoint.CurrentSiteId = $LoadedCheckpoint.CurrentSiteId
            $Script:Checkpoint.CurrentLibraryId = $LoadedCheckpoint.CurrentLibraryId
            $Script:Checkpoint.StartTime = $LoadedCheckpoint.StartTime
            $Script:Checkpoint.RansomwareExtension = $LoadedCheckpoint.RansomwareExtension
            $Script:Checkpoint.Results = $LoadedCheckpoint.Results
            
            # Restore script variables
            if ($LoadedCheckpoint.Results) {
                $Script:Results = $LoadedCheckpoint.Results
            }
            
            if ($LoadedCheckpoint.Statistics) {
                $Script:ProcessedSites = $LoadedCheckpoint.Statistics.ProcessedSites
                $Script:ProcessedLibraries = $LoadedCheckpoint.Statistics.ProcessedLibraries
                $Script:ProcessedFolders = $LoadedCheckpoint.Statistics.ProcessedFolders
                $Script:ProcessedFiles = $LoadedCheckpoint.Statistics.ProcessedFiles
                $Script:InfectedFolders = $LoadedCheckpoint.Statistics.InfectedFolders
                $Script:SkippedSites = $LoadedCheckpoint.Statistics.SkippedSites
                if ($LoadedCheckpoint.Statistics.ErrorSites) {
                    $Script:ErrorSites = $LoadedCheckpoint.Statistics.ErrorSites
                }
            }
            
            Write-ColorOutput "📂 Resuming from checkpoint: $CheckpointPath" "Green"
            Write-ColorOutput "   - Previously processed sites: $($Script:ProcessedSites)" "Green"
            Write-ColorOutput "   - Previously found infected folders: $($Script:InfectedFolders)" "Green"
            Write-ColorOutput "   - Last saved: $($LoadedCheckpoint.LastSavedAt)" "Green"
            
            return $true
        }
        else {
            Write-ColorOutput "⚠️  Checkpoint file not found: $CheckpointPath" "Yellow"
            return $false
        }
    }
    catch {
        Write-ColorOutput "❌ Failed to load checkpoint: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Should-SaveCheckpoint {
    if ($Script:LastCheckpointSave -eq $null) {
        return $true
    }
    
    $TimeSinceLastSave = (Get-Date) - $Script:LastCheckpointSave
    return $TimeSinceLastSave.TotalSeconds -ge $Script:CheckpointInterval
}

function Initialize-CheckpointSystem {
    param(
        [string]$OutputPath
    )
    
    # Setup checkpoint file path
    if ($Script:CheckpointFile -eq "") {
        $CheckpointDir = Split-Path $OutputPath -Parent
        $CheckpointName = "RansomwareCheckpoint_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
        $Script:CheckpointFile = Join-Path $CheckpointDir $CheckpointName
    }
    
    # Setup Ctrl+C handler for graceful interruption
    $null = Register-ObjectEvent -InputObject ([Console]) -EventName "CancelKeyPress" -Action {
        Write-Host ""
        Write-ColorOutput "🛑 Scan interrupted by user. Saving checkpoint..." "Yellow"
        Save-Checkpoint -CheckpointPath $Script:CheckpointFile
        Write-ColorOutput "💾 Checkpoint saved. Use -ResumeFromFile `"$($Script:CheckpointFile)`" to resume." "Green"
        [Environment]::Exit(0)
    }
    
    # Initialize checkpoint structure
    $Script:Checkpoint.StartTime = $Script:StartTime
    $Script:Checkpoint.RansomwareExtension = $RansomwareExtension
    
    Write-ColorOutput "✅ Checkpoint system initialized: $($Script:CheckpointFile)" "Green"
}

function Test-SiteAlreadyProcessed {
    param(
        [string]$SiteId
    )
    
    return $Script:Checkpoint.ProcessedSites -contains $SiteId
}

function Add-ProcessedSite {
    param(
        [string]$SiteId
    )
    
    if ($Script:Checkpoint.ProcessedSites -notcontains $SiteId) {
        $Script:Checkpoint.ProcessedSites += $SiteId
    }
}

#endregion

#region Main Script

function Main {
    Write-ColorOutput "=============================================" "Cyan"
    Write-ColorOutput "SharePoint Ransomware Scanner" "Cyan"
    Write-ColorOutput "=============================================" "Cyan"
    Write-Host ""
    
    # Install required modules
    if (-not (Install-RequiredModules)) {
        Write-ColorOutput "Failed to install required modules. Exiting." "Red"
        return
    }
    
    # Copy parameters to script scope
    $Script:AuthMethod = $AuthMethod
    $Script:TenantId = $TenantId
    $Script:ClientId = $ClientId
    $Script:CertificatePath = $CertificatePath
    $Script:CertificatePassword = $CertificatePassword
    $Script:ClientSecret = $ClientSecret
    
    # Setup authentication
    if (-not (Get-AuthenticationDetails)) {
        Write-ColorOutput "Failed to setup authentication. Exiting." "Red"
        return
    }
    
    # Setup output path
    if (-not $OutputPath) {
        $DefaultPath = Join-Path (Get-Location) "RansomwareReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $OutputPath = $DefaultPath
    }
    
    # Handle checkpoint resume
    $ResumeMode = $false
    if ($ResumeFromFile -and (Test-Path $ResumeFromFile)) {
        Write-ColorOutput "🔄 Resume mode detected" "Yellow"
        $ResumeMode = $true
        $Script:CheckpointFile = $ResumeFromFile
    }
    
    # Connect to Microsoft Graph
    if (-not (Connect-ToGraph)) {
        Write-ColorOutput "Failed to connect to Microsoft Graph. Exiting." "Red"
        return
    }
    
    # Load checkpoint if resuming
    if ($ResumeMode) {
        if (Load-Checkpoint -CheckpointPath $ResumeFromFile) {
            Write-ColorOutput "✅ Successfully loaded checkpoint" "Green"
        } else {
            Write-ColorOutput "❌ Failed to load checkpoint, starting fresh scan" "Yellow"
            $ResumeMode = $false
        }
    }
    
    # Get SharePoint sites
    $Sites = Get-AllSharePointSites
    
    if ($Sites.Count -eq 0) {
        Write-ColorOutput "No sites found to process. Exiting." "Red"
        return
    }
    
    Write-ColorOutput "Starting ransomware scan..." "Yellow"
    Write-ColorOutput "Extension to scan for: $RansomwareExtension" "White"
    Write-ColorOutput "Sites to process: $($Sites.Count)" "White"
    Write-ColorOutput "Output path: $OutputPath" "White"
    Write-Host ""
    
    # Initialize checkpoint system if not resuming
    if (-not $ResumeMode) {
        $Script:StartTime = Get-Date
        Initialize-CheckpointSystem -OutputPath $OutputPath
    }
    
    # Process each site
    $SiteIndex = 0
    $SitesToProcess = if ($ResumeMode) {
        # Filter out already processed sites when resuming
        $Sites | Where-Object { -not (Test-SiteAlreadyProcessed -SiteId $_.Id) }
    } else {
        $Sites
    }
    
    Write-ColorOutput "Sites to process: $($SitesToProcess.Count) (Total: $($Sites.Count))" "White"
    
    foreach ($Site in $SitesToProcess) {
        $SiteIndex++
        
        # Skip if already processed (additional safety check)
        if (Test-SiteAlreadyProcessed -SiteId $Site.Id) {
            Write-ColorOutput "⏭️  Skipping already processed site: $($Site.DisplayName)" "Gray"
            continue
        }
        
        # Update progress
        $TotalSites = if ($ResumeMode) { $Sites.Count } else { $SitesToProcess.Count }
        $CurrentSiteNumber = if ($ResumeMode) { $Script:ProcessedSites + $SiteIndex } else { $SiteIndex }
        
        Write-Progress -Activity "Scanning SharePoint Sites" -Status "Processing site $CurrentSiteNumber/$TotalSites - $($Site.DisplayName)" -PercentComplete (($CurrentSiteNumber / $TotalSites) * 100)
        
        # Process the site
        $ProcessSuccess = Process-SharePointSite -Site $Site
        
        # Mark site as processed and save checkpoint periodically
        if ($ProcessSuccess) {
            Add-ProcessedSite -SiteId $Site.Id
            
            # Save checkpoint every few sites or based on time interval
            if ((Should-SaveCheckpoint) -or ($SiteIndex % 3 -eq 0)) {
                Save-Checkpoint -CheckpointPath $Script:CheckpointFile
            }
        }
        
        # Progress reporting
        if ($SiteIndex % 5 -eq 0) {
            Write-ColorOutput "Progress: $CurrentSiteNumber/$TotalSites sites processed ($($Script:InfectedFolders) infected folders found)" "Cyan"
        }
    }
    
    # Final checkpoint save
    if ($Script:CheckpointFile -ne "") {
        Save-Checkpoint -CheckpointPath $Script:CheckpointFile
        Write-ColorOutput "🔄 Final checkpoint saved" "Green"
    }
    
    Write-Progress -Activity "Scanning SharePoint Sites" -Completed
    
    # Export results
    Write-Host ""
    Write-ColorOutput "Scan complete. Exporting results..." "Yellow"
    Export-Results -OutputPath $OutputPath
    
    # Display final summary
    Write-Host ""
    Write-ColorOutput "=============================================" "Cyan"
    Write-ColorOutput "SCAN SUMMARY" "Cyan"
    Write-ColorOutput "=============================================" "Cyan"
    Write-ColorOutput "Sites Processed: $Script:ProcessedSites" "White"
    Write-ColorOutput "Libraries Scanned: $Script:ProcessedLibraries" "White"
    Write-ColorOutput "Folders Analyzed: $Script:ProcessedFolders" "White"
    Write-ColorOutput "Files Scanned: $Script:ProcessedFiles" "White"
    Write-ColorOutput "Infected Folders: $Script:InfectedFolders" "White"
    Write-ColorOutput "Runtime: $("{0:hh\:mm\:ss}" -f ((Get-Date) - $Script:StartTime))" "White"
    
    if ($Script:InfectedFolders -gt 0) {
        Write-Host ""
        Write-ColorOutput "⚠️  RANSOMWARE DETECTED! Immediate action required." "Red"
        Write-ColorOutput "Review the report at: $OutputPath" "Yellow"
    } else {
        Write-Host ""
        Write-ColorOutput "✅ No ransomware detected. Environment appears clean." "Green"
    }
}

#endregion

# Execute main function
try {
    Main
}
catch {
    Write-ColorOutput "Unexpected error: $($_.Exception.Message)" "Red"
}
finally {
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-ColorOutput "Disconnected from Microsoft Graph." "Green"
    }
    catch {
        # Ignore disconnect errors
    }
    
    Write-Host ""
    Write-ColorOutput "Script execution completed." "Cyan"
}
