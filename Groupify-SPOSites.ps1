<#
.SYNOPSIS
    Comprehensive SharePoint site validation and Microsoft 365 Group connection script.
    Validates sites for group-connection eligibility and optionally performs bulk groupification.

.DESCRIPTION
    This script analyzes SharePoint Online sites to determine their eligibility for Microsoft 365 Group connection.
    It can either auto-discover all sites in the tenant or process sites from a CSV file exported from SharePoint Admin Center.
    
    The script performs comprehensive validation including:
    - Template compatibility checking
    - Publishing feature validation
    - Modern blocking feature detection
    - Alias availability verification
    - Group connection status assessment
    
    Optionally, the script can automatically groupify eligible sites by creating Microsoft 365 Groups and connecting them.

.PARAMETER appID
    The Azure Active Directory Application ID (Client ID) used for authentication.
    This should be configured in the script's USER CONFIGURATION section.

.PARAMETER tenant
    The Azure Active Directory Tenant ID.
    This should be configured in the script's USER CONFIGURATION section.

.PARAMETER csvInputFilePath
    Optional path to a CSV file containing sites to process.
    If empty, the script will auto-discover all sites in the tenant.
    Expected CSV format should include 'URL' or 'Site URL' column.

.PARAMETER tenantAdminUrl
    The SharePoint tenant admin URL (e.g., https://contoso-admin.sharepoint.com).

.PARAMETER enableGroupify
    Boolean flag to enable automatic groupification of eligible sites.
    When set to $true, the script will create Microsoft 365 Groups for eligible sites.
    Default is $false for safety.

.PARAMETER groupifyMaxSites
    Maximum number of sites to groupify in a single run as a safety measure.
    Default is 5 sites.

.EXAMPLE
    .\Groupify-SPOSites.ps1
    
    Runs the script using auto-discovery mode to find all eligible sites in the tenant.
    Only validates sites without performing groupification (safe mode).

.EXAMPLE
    # Configure CSV input and enable groupification
    $csvInputFilePath = "C:\temp\sites.csv"
    $enableGroupify = $true
    .\Groupify-SPOSites.ps1
    
    Processes sites from the specified CSV file and automatically groupifies eligible sites.

.INPUTS
    CSV file with SharePoint site information (optional)
    Expected columns:
    - 'URL' or 'Site URL' (required)
    - 'Site name' or 'Site Title' (optional)
    - 'Sensitivity' or 'Sensitivity Label' (optional)
    - 'Microsoft 365 group' (optional)

.OUTPUTS
    1. Console output with detailed processing information
    2. Log file: SPO-Site-Groupify-Sitefinder_log_[timestamp].txt
    3. Error log: SPO-Site-Groupify-Sitefinder_error_[timestamp].txt
    4. CSV export: Sites-Ready-For-Groups_[timestamp].csv

.NOTES
    File Name      : Groupify-SPOSites.ps1
    Author         : Mike Lee / Tania Menice
    Date           : 9/29/2025
    Prerequisite   : PnP.PowerShell module (minimum version 2.0.0)
    
    Required Permissions:
    - SharePoint Administrator role in Microsoft 365
    - Application must have appropriate permissions for SharePoint and Microsoft Graph
    
    Authentication:
    - Uses interactive authentication with Azure AD application
    - Requires user sign-in during execution
    
    Safety Features:
    - Groupification is disabled by default ($enableGroupify = $false)
    - Maximum sites limit for groupification ($groupifyMaxSites)
    - Comprehensive validation before groupification
    - Detailed logging and error handling

.FUNCTIONALITY
    Core Functions:
    - LogWrite: Handles console and file logging with color coding
    - LogError: Writes errors to separate error log
    - IsGuid: Validates GUID format
    - IsGroupConnected: Checks if site is already group-connected
    - IsTemplateCompatibleWithGroups: Validates site template compatibility
    - Test-PublishingFeatures: Checks for blocking publishing features
    - Test-ModernBlockingFeatures: Identifies modern UI blocking features
    - Invoke-SiteGroupify: Performs comprehensive validation and groupification

.LINK
    https://docs.microsoft.com/en-us/sharepoint/dev/transform/modernize-connect-to-office365-group
    https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets
    https://gist.github.com/joerodgers/4b19e7eef9935a96c1af3d0e9138bcc8
#>


# =================================================================================================
# USER CONFIGURATION - Update the variables in this section
# =================================================================================================
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"                 # This is your Entra App ID
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"                # This is your Tenant ID
$csvInputFilePath = ""         # Path to CSV file (leave empty to be prompted)
$tenantAdminUrl = "https://m365cpi13246019-admin.sharepoint.com"

# Log file configuration
$logDirectory = "$env:TEMP"                                      # Directory for log files (use "." for current directory)
$logFilePrefix = "SPO-Site-Groupify-Sitefinder"                 # Prefix for log file names

# CSV export configuration
$exportFilePath = "$env:TEMP\Sites-Ready-For-Groups_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"  # Path for export file

# Groupify configuration (OPTIONAL - Set to $true to actually create groups for eligible sites)
$enableGroupify = $false                                         # Set to $true to automatically groupify eligible sites (USE WITH CAUTION!)
$groupifyDisplayNameSuffix = " Group"                            # Suffix to add to site title for group display name
$groupifyIsPublic = $false                                       # Set to $true to create public groups, $false for private groups
$groupifyKeepOldHomePage = $false                                # Set to $true to keep the current site home page
$groupifyClassification = ""                                     # Optional classification for the groups (leave empty if not needed)
$groupifyMaxSites = 5                                            # Maximum number of sites to groupify in one run (safety limit)
# =================================================================================================
# END OF USER CONFIGURATION
# =================================================================================================



#region Logging and generic functions
function LogWrite {
    param([string] $log , [string] $ForegroundColor)

    $global:strmWrtLog.writeLine($log)
    if ([string]::IsNullOrEmpty($ForegroundColor)) {
        Write-Host $log
    }
    else {    
        Write-Host $log -ForegroundColor $ForegroundColor
    }
}

# Function to write error messages to a separate error log
# This function maintains a dedicated error log for troubleshooting and audit purposes
function LogError {
    param([string] $log)               # The error message to log
    
    # Write error to dedicated error stream for centralized error tracking
    $global:strmWrtError.writeLine($log)
}

# Function to validate if a string represents a properly formatted GUID
# GUIDs are extensively used in SharePoint for identifying groups, sites, lists, and other objects
function IsGuid {
    param([string] $owner)             # The string to validate as a GUID

    try {
        # Attempt to parse the string as a GUID - this will throw an exception if invalid
        [GUID]$g = $owner
        $t = $g.GetType()
        # Verify that the parsed object is actually a GUID type
        return ($t.Name -eq "Guid")
    }
    catch {
        # If parsing fails, the string is not a valid GUID
        return $false
    }
}

# Function to determine if a SharePoint site is already connected to a Microsoft 365 Group
# Group-connected sites have a distinctive Owner property format that this function analyzes
function IsGroupConnected {
    param([string] $owner)             # The site's Owner property to analyze

    # Check if owner property exists and is not empty
    if (-not [string]::IsNullOrEmpty($owner)) {
        # Group-connected sites have Owner in specific format: GUID_o (total length 38 characters)
        # Example: "12345678-1234-1234-1234-123456789012_o"
        if ($owner.Length -eq 38) {
            
            # Validate that first 36 characters form a valid GUID and last 2 characters are "_o"
            if ((IsGuid $owner.Substring(0, 36)) -and ($owner.Substring(36, 2) -eq "_o")) {
                return $true
            }
        }        
    }

    # If any validation fails, the site is not group-connected
    return $false
}

# Function to determine if a SharePoint site template is compatible with Microsoft 365 Groups
# This implements Microsoft's official guidance on template compatibility for group connection
function IsTemplateCompatibleWithGroups {
    param([string] $template)          # The site template to validate (e.g., "STS#0", "BLOG#0")

    # Return false for null, empty, or whitespace-only templates
    if ([string]::IsNullOrWhiteSpace($template)) {
        return $false
    }

    # Normalize template name to uppercase for consistent comparison
    $normalizedTemplate = $template.ToUpperInvariant()

    # Microsoft's official list of templates that CANNOT be connected to groups
    # These templates have architectural limitations or conflicts that prevent group connection
    $incompatibleTemplates = @(
        "BICENTERSITE#0",        # Business Intelligence Center - complex BI features conflict with groups
        "BLANKINTERNET#0",       # Blank Internet site - designed for public-facing sites
        "ENTERWIKI#0",           # Enterprise Wiki - complex wiki functionality conflicts
        "SRCHCEN#0",             # Search Center - search-specific features conflict with groups
        "SRCHCENTERLITE#0",      # Search Center Lite - search functionality conflicts
        "POINTPUBLISHINGHUB#0",  # Publishing Hub - publishing infrastructure conflicts
        "POINTPUBLISHINGTOPIC#0",
        "CMSPUBLISHING#0",
        "SPSMSITEHOST#0",
        "TEAMCHANNEL#1",
        "APPCATALOG#0",
        "REDIRECTSITE#0"
    )

    return ($incompatibleTemplates -notcontains $normalizedTemplate)
}

function Test-PublishingFeatures {
    param([string] $SiteUrl, $Connection)
    
    try {
        # SharePoint Publishing features block group connection because they fundamentally
        # change how sites work and are incompatible with the modern group-connected experience
        
        # Check for Site Collection level Publishing Infrastructure feature
        # GUID: F6924D36-2FA8-4F0B-B16D-06B7250180FA
        $publishingSiteFeature = Get-PnPFeature -Identity "F6924D36-2FA8-4F0B-B16D-06B7250180FA" -Scope Site -Connection $Connection -ErrorAction SilentlyContinue
        
        # Check for Web level SharePoint Server Publishing feature  
        # GUID: 94C94CA6-B32F-4DA9-A9E3-1F3D343D7ECB
        $publishingWebFeature = Get-PnPFeature -Identity "94C94CA6-B32F-4DA9-A9E3-1F3D343D7ECB" -Scope Web -Connection $Connection -ErrorAction SilentlyContinue
        
        # If either publishing feature is enabled, the site cannot be group-connected
        if ($publishingSiteFeature -or $publishingWebFeature) {
            return $false, "Publishing features are enabled and block group connection"
        }
        return $true, "No blocking publishing features found"
    }
    catch {
        # If we can't check features, proceed with caution but don't block the process
        return $true, "Could not check publishing features - proceeding with caution"
    }
}

function Test-ModernBlockingFeatures {
    param([string] $SiteUrl, $Connection)
    
    try {
        # These features specifically block the modern list and library experience
        # While they don't prevent group connection, they should be disabled for optimal modern experience
        
        # Check for Site Collection level Modern List and Library blocking feature
        # GUID: E3540C7D-6BEA-403C-A224-1A12EAFEE4C4
        $modernListBlockSite = Get-PnPFeature -Identity "E3540C7D-6BEA-403C-A224-1A12EAFEE4C4" -Scope Site -Connection $Connection -ErrorAction SilentlyContinue
        
        # Check for Web level Modern List and Library blocking feature
        # GUID: 52E14B6F-B1BB-4969-B89B-C4FAA56745EF  
        $modernListBlockWeb = Get-PnPFeature -Identity "52E14B6F-B1BB-4969-B89B-C4FAA56745EF" -Scope Web -Connection $Connection -ErrorAction SilentlyContinue
        
        # Collect any blocking features found
        $blockingFeatures = @()
        if ($modernListBlockSite) { $blockingFeatures += "Modern List UI blocking (Site)" }
        if ($modernListBlockWeb) { $blockingFeatures += "Modern List UI blocking (Web)" }
        
        # Return whether blocking features were found and details
        if ($blockingFeatures.Count -gt 0) {
            return $true, "Blocking features found: $($blockingFeatures -join ', ')"
        }
        return $false, "No modern blocking features found"
    }
    catch {
        # If we can't check features, assume no blocking features are present
        return $false, "Could not check modern blocking features - proceeding with caution"
    }
}

function Invoke-SiteGroupify {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $SiteUrl,
        [Parameter(Mandatory = $true)][string] $SiteTitle,
        [Parameter(Mandatory = $true)][string] $Template,
        [Parameter(Mandatory = $true)] $Connection
    )

    LogWrite "  ── Starting comprehensive validation and groupify for $SiteUrl" Cyan

    try {
        # Step 1: Check template compatibility
        if (-not (IsTemplateCompatibleWithGroups $Template)) {
            LogWrite "    ❌ Site template '$Template' is not compatible with Microsoft 365 Groups. Skipping." Red
            return $false
        }
        LogWrite "    ✅ Template '$Template' is compatible with groups" Green

        # Step 2: Connect to the specific site for detailed validation
        $siteConnection = $null
        try {
            $siteConnection = Connect-PnPOnline -Url $SiteUrl -ClientId $appID -Tenant $tenant -ReturnConnection -Interactive
            LogWrite "    ✅ Successfully connected to site for validation" Green
        }
        catch {
            LogWrite "    ❌ Could not connect to site for validation: $($_.Exception.Message)" Red
            return $false
        }

        # Step 3: Check for publishing features that prevent group connection
        $publishingCheck, $publishingMessage = Test-PublishingFeatures $SiteUrl $siteConnection
        if ($publishingCheck) {
            # Publishing features are present - this blocks group connection completely
            LogWrite "    ❌ Publishing feature validation failed: $publishingMessage" Red
            return $false
        }
        LogWrite "    ✅ Publishing feature check: $publishingMessage" Green

        # Step 4: Check for modern blocking features (these can be automatically fixed)
        $modernCheck, $modernMessage = Test-ModernBlockingFeatures $SiteUrl $siteConnection
        if ($modernCheck) {
            # Modern blocking features detected - warn but these can be disabled during groupification
            LogWrite "    ⚠ Modern blocking features detected: $modernMessage" Yellow
            LogWrite "    💡 These will be automatically disabled during groupification" Cyan
        }
        else {
            LogWrite "    ✅ Modern features check: $modernMessage" Green
        }

        # Step 5: Generate group properties
        $siteUri = [System.Uri]$SiteUrl
        $displayNameBase = if ([string]::IsNullOrWhiteSpace($SiteTitle)) {
            $siteUri.Segments[-1].TrimEnd('/')
        }
        else {
            $SiteTitle
        }

        $displayName = if ([string]::IsNullOrWhiteSpace($groupifyDisplayNameSuffix)) {
            $displayNameBase
        }
        else {
            ($displayNameBase + $groupifyDisplayNameSuffix).Trim()
        }

        if ([string]::IsNullOrWhiteSpace($displayName)) {
            $displayName = "Microsoft 365 Group for $SiteUrl"
        }

        # Step 6: Generate and validate alias
        $aliasSource = $displayName
        $alias = ($aliasSource -replace '[^A-Za-z0-9]', '')
        if ([string]::IsNullOrWhiteSpace($alias)) {
            $alias = ($siteUri.Segments[-1] -replace '[^A-Za-z0-9]', '')
        }
        if ([string]::IsNullOrWhiteSpace($alias)) {
            $alias = "grp$([Guid]::NewGuid().ToString('N').Substring(0, 8))"
        }
        if ($alias.Length -gt 60) {
            $alias = $alias.Substring(0, 60)
        }
        $alias = $alias.ToLowerInvariant()

        # Check if alias contains spaces (not allowed)
        if ($alias.Contains(" ")) {
            LogWrite "    ❌ Generated alias '$alias' contains spaces - this is not allowed" Red
            return $false
        }

        LogWrite "    Display Name: '$displayName'" Gray
        LogWrite "    Alias: '$alias'" Gray

        # Step 7: Check if alias is already in use
        try {
            $aliasInUse = Test-PnPMicrosoft365GroupAliasIsUsed -Alias $alias -Connection $Connection
            if ($aliasInUse) {
                LogWrite "    ❌ Alias '$alias' is already in use by another Microsoft 365 Group" Red
                return $false
            }
            LogWrite "    ✅ Alias '$alias' is available" Green
        }
        catch {
            LogWrite "    ⚠ Could not verify alias availability: $($_.Exception.Message)" Yellow
            LogWrite "    💡 Proceeding with groupification - will fail if alias exists" Cyan
        }

        LogWrite "    → Attempting groupification with Add-PnPMicrosoft365GroupToSite..." Cyan

        # Step 8: Perform the actual groupification
        try {
            Add-PnPMicrosoft365GroupToSite `
                -Url $SiteUrl `
                -Alias $alias `
                -DisplayName $displayName `
                -Description "Microsoft 365 Group for $SiteTitle" `
                -IsPublic:$groupifyIsPublic `
                -KeepOldHomePage:$groupifyKeepOldHomePage `
                -Connection $siteConnection `
                -ErrorAction Stop

            LogWrite "    ✅ Successfully groupified site! Group alias: '$alias'" Green
            return $true
        }
        catch {
            LogWrite "    ❌ Groupification failed: $($_.Exception.Message)" Red

            # Provide specific error guidance
            if ($_.Exception.Message -like "*alias*already*exists*") {
                LogWrite "    💡 The alias '$alias' already exists. Try a different site or check existing groups." Yellow
            }
            elseif ($_.Exception.Message -like "*permission*" -or $_.Exception.Message -like "*unauthorized*" -or $_.Exception.Message -like "*forbidden*") {
                LogWrite "    💡 Permission issue. Verify you have SharePoint Administrator role in your tenant." Yellow
            }
            elseif ($_.Exception.Message -like "*owner*") {
                LogWrite "    💡 Owner-related issue. The site may need at least one owner." Yellow
            }
            else {
                LogWrite "    💡 Generic troubleshooting:" Yellow
                LogWrite "      1. Verify you have SharePoint Administrator role" Yellow
                LogWrite "      2. Check if alias '$alias' already exists" Yellow
                LogWrite "      3. Ensure site isn't already group-connected" Yellow
                LogWrite "      4. Verify tenant allows group creation" Yellow
            }

            # Log error details
            LogError "GROUPIFICATION FAILED for $SiteUrl"
            LogError "  Alias: '$alias'"
            LogError "  DisplayName: '$displayName'"
            LogError "  Error: $($_.Exception.Message)"

            return $false
        }
        finally {
            # Clean up site connection
            if ($siteConnection) {
                try { Disconnect-PnPOnline -Connection $siteConnection } catch { }
            }
        }
    }
    catch {
        LogWrite "    ❌ Unexpected error in Invoke-SiteGroupify: $($_.Exception.Message)" Red
        LogError "Unexpected error in Invoke-SiteGroupify for $SiteUrl : $($_.Exception.Message)"
        return $false
    }
}


#endregion

#######################################################
# MAIN SCRIPT EXECUTION BEGINS HERE                  #
#######################################################
# This section contains the main logic that:
# 1. Sets up logging and loads required PowerShell modules
# 2. Determines whether to use auto-discovery or CSV input
# 3. Connects to SharePoint using interactive authentication  
# 4. Processes sites for group connection eligibility
# 5. Optionally performs groupification on eligible sites
# 6. Exports results to CSV and displays summary
#######################################################
#region Setup Logging
$date = Get-Date
$logfile = ((Get-Item -Path $logDirectory -Verbose).FullName + "\" + $logFilePrefix + "_log_" + $date.ToFileTime() + ".txt")
$global:strmWrtLog = [System.IO.StreamWriter]$logfile
$global:Errorfile = ((Get-Item -Path $logDirectory -Verbose).FullName + "\" + $logFilePrefix + "_error_" + $date.ToFileTime() + ".txt")
$global:strmWrtError = [System.IO.StreamWriter]$Errorfile
#endregion

#region Load needed PowerShell modules
#Ensure PnP PowerShell is loaded
$minimumVersion = New-Object System.Version("2.0.0")
if (-not (Get-InstalledModule -Name PnP.PowerShell -MinimumVersion $minimumVersion -ErrorAction Ignore)) {
    Install-Module PnP.PowerShell -MinimumVersion $minimumVersion -Scope CurrentUser -Force
}
Import-Module PnP.PowerShell -DisableNameChecking -MinimumVersion $minimumVersion
#endregion

# =================================================================================================
# DETERMINE INPUT METHOD: CSV FILE vs AUTO-DISCOVERY
# =================================================================================================
# The script can operate in two modes:
# 1. CSV Mode: Process specific sites from a CSV file (typical for targeted groupification)
# 2. Auto-Discovery: Scan all sites in the tenant (useful for tenant-wide analysis)

if ([String]::IsNullOrEmpty($csvInputFilePath)) {
    # AUTO-DISCOVERY MODE: No CSV file specified, so discover all sites automatically
    LogWrite "No CSV file provided - will discover all sites automatically using Get-PnPTenantSite" Yellow
    $useAutoDiscovery = $true
    $csvContent = @()     # Empty array since no CSV data
    $csvHeaders = @()     # Empty array since no CSV headers
}
else {
    # CSV MODE: Process sites from the specified CSV file
    # Validate the CSV file exists before proceeding
    if (-not (Test-Path $csvInputFilePath)) {
        Write-Host "Error: CSV file not found at path: $csvInputFilePath" -ForegroundColor Red
        exit 1
    }
    
    LogWrite "Using CSV file: $csvInputFilePath"
    $useAutoDiscovery = $false
    
    # Import CSV and analyze its structure to understand the column format
    $csvContent = Import-Csv $csvInputFilePath
    $csvHeaders = ($csvContent[0] | Get-Member -MemberType NoteProperty).Name
    
    LogWrite "Detected CSV columns: $($csvHeaders -join ', ')"
}

# =================================================================================================
# CSV COLUMN VALIDATION AND FORMAT DETECTION
# =================================================================================================
# The script supports multiple CSV formats to accommodate different export sources:
# - SharePoint Admin Center exports: "URL", "Site name", "Sensitivity", "Microsoft 365 group"
# - Custom exports: "Site URL", "Site Title", "Sensitivity Label", etc.
# This flexibility allows the script to work with various data sources without modification.

if ($useAutoDiscovery) {
    # AUTO-DISCOVERY MODE: No CSV validation needed
    LogWrite "Auto-discovery mode: Will process all sites found in tenant" Green
    $hasUrlColumn = $false           # Not applicable in auto-discovery
    $hasSiteNameColumn = $false      # Not applicable in auto-discovery  
    $hasSensitivityColumn = $false   # Not applicable in auto-discovery
    $hasGroupConnectedColumn = $false # Not applicable in auto-discovery
}
else {
    # CSV MODE: Validate that required columns exist and detect format variations
    
    # Check for URL column (required) - support both naming conventions
    $hasUrlColumn = ($csvHeaders -contains "URL") -or ($csvHeaders -contains "Site URL")
    
    # Check for optional columns that provide additional context
    $hasSiteNameColumn = ($csvHeaders -contains "Site name") -or ($csvHeaders -contains "Site Title")
    $hasSensitivityColumn = ($csvHeaders -contains "Sensitivity") -or ($csvHeaders -contains "Sensitivity Label")
    $hasGroupConnectedColumn = $csvHeaders -contains "Microsoft 365 group"
    
    if ($hasUrlColumn) {
        LogWrite "Processing SharePoint sites from CSV file." Green
        
        # Log which optional columns are available for enhanced processing
        if ($hasSiteNameColumn) {
            LogWrite "Found site name column for site identification." Green
        }
        if ($hasSensitivityColumn) {
            LogWrite "Found sensitivity column for sensitivity label information." Green
        }
        if ($hasGroupConnectedColumn) {
            LogWrite "Found 'Microsoft 365 group' column for group connection status." Green
        }
    }
    else {
        # Critical error: No URL column means we can't process any sites
        LogWrite "Error: No 'URL' or 'Site URL' column found. Please ensure this is a valid CSV file with site URLs." Red
        exit 1
    }
}

# Validate app authentication configuration
# =================================================================================================
# AUTHENTICATION VALIDATION AND SHAREPOINT CONNECTION
# =================================================================================================
# This section validates that required authentication parameters are provided and establishes
# a connection to SharePoint Online using interactive authentication.

# Validate that required authentication parameters are configured
if ([String]::IsNullOrEmpty($appID) -or [String]::IsNullOrEmpty($tenant)) {
    Write-Host "Error: App authentication configuration is incomplete. Please ensure appID and tenant are configured." -ForegroundColor Red
    exit 1
}

LogWrite "Using interactive authentication with App ID: $appID"

#region Connect to SharePoint using Interactive Authentication
# =================================================================================================
# SHAREPOINT ONLINE CONNECTION ESTABLISHMENT
# =================================================================================================
# Interactive authentication is used because it:
# - Works with all tenant configurations (MFA, conditional access, etc.)
# - Doesn't require certificate management or app secret rotation  
# - Provides reliable authentication for both validation and groupification operations
# - Automatically handles token refresh during long-running operations

LogWrite "Connect to tenant admin site $tenantAdminUrl using Interactive Authentication"
try {
    # Connect using the Microsoft Graph PowerShell app ID for reliable interactive authentication
    $tenantContext = Connect-PnPOnline -Url $tenantAdminUrl -ClientId $appID -Tenant $tenant -ReturnConnection -Interactive
    LogWrite "Successfully connected to SharePoint" Green
}
catch {
    LogWrite "Failed to connect to SharePoint: $($_.Exception.Message)" Red
    LogError "Failed to connect to SharePoint: $($_.Exception.Message)"
    exit 1
}
#endregion

#region Site Validation Setup
# =================================================================================================
# SITE VALIDATION AND PROCESSING INITIALIZATION
# =================================================================================================
# This section sets up the validation pipeline that will process SharePoint sites.
# The validation pipeline performs comprehensive checks to ensure only eligible sites
# are connected to Microsoft 365 Groups.

LogWrite "Starting site validation..."

# Initialize counters to track processing statistics
$totalSites = 0              # Total number of sites processed
$activeSites = 0             # Sites that are active and accessible
$groupConnectedSites = 0     # Sites already connected to Microsoft 365 Groups
$inactiveSites = 0           # Sites that are inactive or inaccessible
$groupifiedSites = 0         # Sites successfully connected to groups during this run

# Array to collect detailed information about sites ready for group connection
$sitesReadyForGroups = @()

# =================================================================================================
# GROUPIFICATION MODE SAFETY WARNINGS AND SETTINGS
# =================================================================================================
# When groupification is enabled, the script will make live changes to SharePoint sites.
# This section provides clear warnings and displays current settings to administrators.

if ($enableGroupify) {
    LogWrite "⚠️  WARNING: Groupify mode is ENABLED!" Yellow
    LogWrite "   This will automatically create Microsoft 365 Groups and connect them to eligible sites." Yellow
    LogWrite "   Maximum sites to groupify in this run: $groupifyMaxSites" Yellow
    LogWrite "   Group settings: Public=$groupifyIsPublic, KeepOldHomePage=$groupifyKeepOldHomePage" Yellow
    LogWrite "   Authentication method for groupification: Delegated (Interactive)" Yellow
    if (-not [string]::IsNullOrEmpty($groupifyClassification)) {
        LogWrite "   Classification: $groupifyClassification" Yellow
    }
    LogWrite "   📝 NOTE: You will be prompted to sign in interactively for groupification operations" Cyan
    LogWrite "" 
}

LogWrite "Processing sites from $(if ($useAutoDiscovery) { 'tenant auto-discovery' } else { $csvInputFilePath })..."

if ($useAutoDiscovery) {
    # Auto-discovery mode: Get all sites from tenant, filtering out incompatible templates
    try {
        LogWrite "Discovering all sites in tenant (filtering out incompatible templates)..." Cyan
        
        # Define incompatible templates for filtering
        $incompatibleTemplates = @(
            "BICENTERSITE#0",
            "BLANKINTERNET#0", 
            "ENTERWIKI#0",
            "SRCHCEN#0",
            "SRCHCENTERLITE#0",
            "POINTPUBLISHINGHUB#0",
            "POINTPUBLISHINGTOPIC#0",
            "CMSPUBLISHING#0",
            "SPSMSITEHOST#0",
            "TEAMCHANNEL#1",
            "APPCATALOG#0",
            "RedirectSite#0",
            "SITEPAGEPUBLISHING#0",
            "GROUP#0"        # Already group-connected sites
        )
        
        # Try to get all sites first, then filter in PowerShell
        # (SharePoint Online filter syntax is very limited and error-prone)
        LogWrite "Getting all tenant sites..." Gray
        $allSites = Get-PnPTenantSite -Connection $tenantContext
        LogWrite "Retrieved $($allSites.Count) total sites from tenant" Gray
        
        # Filter out incompatible templates and already group-connected sites in PowerShell
        LogWrite "Filtering out incompatible templates and group-connected sites..." Gray
        $originalCount = $allSites.Count
        $compatibleSites = $allSites | Where-Object {
            $template = $_.Template
            $owner = $_.Owner
            
            # Exclude incompatible templates
            $isIncompatible = $incompatibleTemplates -contains $template
            # Exclude personal sites (SPSPERS*)
            $isPersonalSite = $template -like "*SPSPERS*"
            # Exclude GROUP#0 template sites (these are already group-connected)
            $isGroupTemplate = $template -eq "GROUP#0"
            # Check for group-connected sites by examining the Owner property format
            # Group-connected sites have Owner in format: GUID_o (38 characters total)
            $isAlreadyGroupConnected = $false
            if (-not [string]::IsNullOrEmpty($owner) -and $owner.Length -eq 38 -and $owner.EndsWith("_o")) {
                try {
                    # Verify the first 36 characters form a valid GUID
                    $null = [GUID]$owner.Substring(0, 36)
                    $isAlreadyGroupConnected = $true
                }
                catch {
                    # Not a valid GUID, so not group-connected
                    $isAlreadyGroupConnected = $false
                }
            }
            
            # Return sites that are NOT incompatible, NOT personal sites, NOT group template, and NOT already group-connected
            -not $isIncompatible -and -not $isPersonalSite -and -not $isGroupTemplate -and -not $isAlreadyGroupConnected
        }
        
        $filteredOutCount = $originalCount - $compatibleSites.Count
        $allSites = $compatibleSites
        LogWrite "Found $($allSites.Count) sites ready for group connection after filtering (excluded $filteredOutCount incompatible/group-connected sites)" Green
        
        # Process each discovered site
        foreach ($siteFromList in $allSites) {
            $siteUrl = $siteFromList.Url
            $totalSites++
            
            # Set default values for auto-discovery mode (no CSV data available)
            $sensitivity = "None"  # Sensitivity label information not available in auto-discovery
            
            LogWrite "[PROCESSING] $siteUrl"
            
            try {
                # Get detailed site information to check group connection status
                $site = Get-PnPTenantSite -Url $siteUrl -Connection $tenantContext -ErrorAction SilentlyContinue
                
                if ($null -ne $site -and $site.Status -eq "Active") {
                    $activeSites++
                    
                    # Check if site template is compatible with groups
                    $isTemplateCompatible = IsTemplateCompatibleWithGroups $site.Template
                    
                    # Check if site is group-connected using the Owner property
                    $isActuallyGroupConnected = IsGroupConnected $site.Owner
                    
                    if ($isActuallyGroupConnected) {
                        $groupConnectedSites++
                        LogWrite "  ✓ Active Site | Group Connected: Yes | Sensitivity: $sensitivity" Green
                    }
                    elseif ($isTemplateCompatible) {
                        LogWrite "  ✓ Active Site | Group Connected: No | Sensitivity: $sensitivity | Can be connected to group" Yellow
                        
                        # Add to sites ready for group connection
                        $siteInfo = [PSCustomObject]@{
                            'Site URL'           = $siteUrl
                            'Site Title'         = if ($site.Title) { $site.Title } else { "N/A" }
                            'Template'           = $site.Template
                            'Owner'              = if ($site.Owner) { $site.Owner } else { "N/A" }
                            'Storage Used (MB)'  = $site.StorageUsageCurrent
                            'Last Modified'      = $site.LastContentModifiedDate
                            'Sensitivity Label'  = $sensitivity
                            'Current Status'     = "Ready for Group Connection"
                            'Site Status'        = $site.Status
                            'Sharing Capability' = if ($site.SharingCapability) { $site.SharingCapability } else { "N/A" }
                        }
                        $sitesReadyForGroups += $siteInfo
                        
                        # Attempt groupification if enabled and under limit
                        if ($enableGroupify -and $groupifiedSites -lt $groupifyMaxSites) {
                            $groupifySuccess = Invoke-SiteGroupify -SiteUrl $siteUrl -SiteTitle $site.Title -Template $site.Template -Connection $tenantContext
                            if ($groupifySuccess) {
                                $groupifiedSites++
                                $groupConnectedSites++  # Update count since this site is now group-connected
                            }
                        }
                        elseif ($enableGroupify -and $groupifiedSites -ge $groupifyMaxSites) {
                            LogWrite "    ⚠️ Maximum groupify limit ($groupifyMaxSites) reached - skipping remaining sites" Yellow
                        }
                    }
                    else {
                        LogWrite "  ⚠ Active Site | Group Connected: No | Template: $($site.Template) | Not compatible with groups" Magenta
                    }
                    
                    # Additional site information
                    LogWrite "    Template: $($site.Template) | Storage Used: $($site.StorageUsageCurrent) MB | Last Modified: $($site.LastContentModifiedDate)"
                }
                else {
                    $inactiveSites++
                    LogWrite "  ✗ Site Status: $(if ($site) { $site.Status } else { 'Not Found' }) | Cannot be processed" Red
                    LogError "Site $siteUrl has status $(if ($site) { $site.Status } else { 'Not Found' })"
                }
            }
            catch [Exception] {
                $inactiveSites++
                $ErrorMessage = $_.Exception.Message
                LogWrite "  ✗ Error processing site: $ErrorMessage" Red
                LogError "Error processing $siteUrl : $ErrorMessage"
            }
        }
    }
    catch {
        LogWrite "Error discovering sites: $($_.Exception.Message)" Red
        LogError "Error discovering sites: $($_.Exception.Message)"
        exit 1
    }
}
else {
    # CSV mode: Process sites from CSV file
    $csvRows = $csvContent
    
    foreach ($row in $csvRows) {
        # Handle different CSV formats for URL - support both "URL" and "Site URL" columns
        $siteUrl = ""
        if ($null -ne $row."Site URL") {
            $siteUrl = $row."Site URL".Trim()
        }
        elseif ($null -ne $row.URL) {
            $siteUrl = $row.URL.Trim()
        }
        
        if ([string]::IsNullOrEmpty($siteUrl)) {
            continue # Skip rows without URL
        }
        
        $totalSites++
        
        # Get site information from CSV - support both old and new column formats
        $sensitivity = "None"
        if ($hasSensitivityColumn) {
            if ($null -ne $row."Sensitivity Label") {
                $sensitivity = $row."Sensitivity Label".Trim()
            }
            elseif ($null -ne $row.Sensitivity) {
                $sensitivity = $row.Sensitivity.Trim()
            }
        }
        
        LogWrite "[PROCESSING] $siteUrl"
        
        try {
            # Get site details from SharePoint
            $site = Get-PnPTenantSite -Url $siteUrl -Connection $tenantContext -ErrorAction SilentlyContinue
            
            if ($null -ne $site -and $site.Status -eq "Active") {
                $activeSites++
                
                # Check if site template is compatible with groups
                $isTemplateCompatible = IsTemplateCompatibleWithGroups $site.Template
                
                # Check if site is group-connected using the Owner property
                $isActuallyGroupConnected = IsGroupConnected $site.Owner
                
                if ($isActuallyGroupConnected) {
                    $groupConnectedSites++
                    LogWrite "  ✓ Active Site | Group Connected: Yes | Sensitivity: $sensitivity" Green
                }
                elseif ($isTemplateCompatible) {
                    LogWrite "  ✓ Active Site | Group Connected: No | Sensitivity: $sensitivity | Can be connected to group" Yellow
                    
                    # Add to sites ready for group connection
                    $siteInfo = [PSCustomObject]@{
                        'Site URL'           = $siteUrl
                        'Site Title'         = if ($site.Title) { $site.Title } else { "N/A" }
                        'Template'           = $site.Template
                        'Owner'              = if ($site.Owner) { $site.Owner } else { "N/A" }
                        'Storage Used (MB)'  = $site.StorageUsageCurrent
                        'Last Modified'      = $site.LastContentModifiedDate
                        'Sensitivity Label'  = $sensitivity
                        'Current Status'     = "Ready for Group Connection"
                        'Site Status'        = $site.Status
                        'Sharing Capability' = if ($site.SharingCapability) { $site.SharingCapability } else { "N/A" }
                    }
                    $sitesReadyForGroups += $siteInfo
                    
                    # Attempt groupification if enabled and under limit
                    if ($enableGroupify -and $groupifiedSites -lt $groupifyMaxSites) {
                        $groupifySuccess = Invoke-SiteGroupify -SiteUrl $siteUrl -SiteTitle $site.Title -Template $site.Template -Connection $tenantContext
                        if ($groupifySuccess) {
                            $groupifiedSites++
                            $groupConnectedSites++  # Update count since this site is now group-connected
                        }
                    }
                    elseif ($enableGroupify -and $groupifiedSites -ge $groupifyMaxSites) {
                        LogWrite "    ⚠️ Maximum groupify limit ($groupifyMaxSites) reached - skipping remaining sites" Yellow
                    }
                }
                else {
                    LogWrite "  ⚠ Active Site | Group Connected: No | Template: $($site.Template) | Not compatible with groups" Magenta
                }
                
                # Additional site information
                LogWrite "    Template: $($site.Template) | Storage Used: $($site.StorageUsageCurrent) MB | Last Modified: $($site.LastContentModifiedDate)"
            }
            elseif ($null -ne $site) {
                $inactiveSites++
                LogWrite "  ✗ Site Status: $($site.Status) | Cannot be processed" Red
                LogError "Site $siteUrl has status $($site.Status)"
            }
            else {
                $inactiveSites++
                LogWrite "  ✗ Site not found or inaccessible" Red
                LogError "Site $siteUrl not found or inaccessible"
            }
        }
        catch [Exception] {
            $inactiveSites++
            $ErrorMessage = $_.Exception.Message
            LogWrite "  ✗ Error processing site: $ErrorMessage" Red
            LogError "Error processing $siteUrl : $ErrorMessage"
        }
    }
}

# Export sites ready for group connection to CSV
if ($sitesReadyForGroups.Count -gt 0) {
    try {
        $sitesReadyForGroups | Export-Csv -Path $exportFilePath -NoTypeInformation -Encoding UTF8
        LogWrite ""
        LogWrite "✓ Exported $($sitesReadyForGroups.Count) sites ready for group connection to: $exportFilePath" Green
        LogWrite "  Use this CSV file as input for the bulk Office 365 Group connection process." Cyan
    }
    catch {
        LogWrite "✗ Failed to export CSV file: $($_.Exception.Message)" Red
        LogError "Failed to export CSV file: $($_.Exception.Message)"
    }
}
else {
    LogWrite ""
    LogWrite "ℹ No sites found that are ready for group connection - no CSV export created." Yellow
}

# Display summary
LogWrite ""
LogWrite "=== SUMMARY ===" Cyan
LogWrite "Total sites processed: $totalSites" White
LogWrite "Active sites: $activeSites" Green
LogWrite "Group-connected sites: $groupConnectedSites" Green
LogWrite "Sites available for group connection: $($activeSites - $groupConnectedSites)" Yellow
LogWrite "Inactive/Error sites: $inactiveSites" Red
LogWrite "Sites exported to CSV: $($sitesReadyForGroups.Count)" Cyan
if ($enableGroupify) {
    LogWrite "Sites successfully groupified: $groupifiedSites" Green
    if ($groupifiedSites -eq $groupifyMaxSites) {
        LogWrite "Maximum groupify limit reached - remaining eligible sites not processed" Yellow
    }
}
LogWrite "===============" Cyan
#endregion

#region Close log files
if ($global:strmWrtLog -ne $NULL) {
    $global:strmWrtLog.Close()
    $global:strmWrtLog.Dispose()
}

if ($global:strmWrtError -ne $NULL) {
    $global:strmWrtError.Close()
    $global:strmWrtError.Dispose()
}
#endregion

#######################################################################################################################
# SCRIPT COMPLETION - ADMINISTRATOR REFERENCE GUIDE                                                                 #
#######################################################################################################################
#
# CONGRATULATIONS! The SharePoint Site Groupification script has completed execution.
#
# WHAT THIS SCRIPT ACCOMPLISHED:
# ==============================
# ✅ Comprehensive Site Validation: Analyzed SharePoint sites for Microsoft 365 Group compatibility
# ✅ Template Compatibility Check: Verified site templates support group connection
# ✅ Publishing Feature Detection: Identified blocking SharePoint Publishing features
# ✅ Modern UI Validation: Checked for features that impact modern experience
# ✅ Group Connection Analysis: Determined current group connection status
# ✅ Detailed Logging: Created comprehensive logs for audit and troubleshooting
# ✅ Results Export: Generated CSV reports for further analysis
#
# KEY VALIDATION CHECKS PERFORMED:
# ================================
# 1. SITE ACCESSIBILITY - Verified sites exist and are accessible
# 2. TEMPLATE COMPATIBILITY - Ensured site templates support group connection
# 3. PUBLISHING FEATURES - Detected blocking SharePoint Publishing infrastructure
# 4. MODERN UI FEATURES - Identified features that impact modern experience
# 5. GROUP STATUS - Determined if sites are already group-connected
# 6. PERMISSIONS - Validated sufficient permissions for operations
#
# UNDERSTANDING THE RESULTS:
# =========================
# 📊 TOTAL SITES PROCESSED: Check the summary statistics for complete counts
# ✅ SUCCESSFUL VALIDATIONS: Sites that passed all compatibility checks
# ❌ FAILED VALIDATIONS: Sites with compatibility issues (see logs for details)
# ⚠️ ALREADY GROUP-CONNECTED: Sites that already have Microsoft 365 Groups
# 📁 CSV EXPORTS: Detailed results saved for administrative review
#
# NEXT STEPS FOR ADMINISTRATORS:
# ==============================
# 
# FOR VALIDATION-ONLY RUNS (PerformGroupify = $false):
# → Review the exported CSV file for detailed site analysis
# → Address any compatibility issues identified in failed sites
# → Plan groupification strategy based on validation results
# → Test with a small subset before processing all sites
# → Re-run with -PerformGroupify $true when ready for live changes
#
# FOR GROUPIFICATION RUNS (PerformGroupify = $true):
# → Monitor newly group-connected sites for proper integration
# → Verify groups appear in Outlook, Teams, and other Microsoft 365 apps
# → Check that users can access site content through group interfaces
# → Address any post-groupification issues identified in logs
# → Communicate changes to affected users and site owners
#
# TROUBLESHOOTING GUIDANCE:
# ========================
# 
# COMMON ISSUES AND SOLUTIONS:
# → Publishing Features: Sites with publishing features cannot be group-connected
#   Solution: Disable publishing features or exclude these sites
# 
# → Template Incompatibility: Some site templates don't support groups
#   Solution: Migrate content to compatible templates or exclude these sites
# 
# → Permission Errors: Insufficient permissions for site access or modification
#   Solution: Ensure you have SharePoint Administrator or site owner permissions
# 
# → Authentication Issues: Problems with interactive authentication
#   Solution: Verify tenant name, check network connectivity, confirm user permissions
#
# LOG FILE LOCATIONS:
# ==================
# 📄 Main Log: Contains detailed execution information and validation results
# 🚨 Error Log: Contains specific error messages for troubleshooting
# 📊 Results CSV: Contains detailed site information and validation outcomes
#
# MONITORING RECOMMENDATIONS:
# ===========================
# → Monitor group-connected sites for 24-48 hours after groupification
# → Check Microsoft 365 admin center for any group-related issues
# → Verify group permissions align with organizational policies
# → Ensure proper group naming conventions are followed
# → Review group classification and sensitivity labels as needed
#
# SUPPORT RESOURCES:
# =================
# 📖 Microsoft Documentation: https://docs.microsoft.com/sharepoint/dev/transform/modernize-connect-to-office365-group
# 🔧 PnP PowerShell: https://pnp.github.io/powershell/
# 🏢 SharePoint Admin Center: https://admin.microsoft.com/sharepoint
# 👥 Microsoft 365 Groups Admin: https://admin.microsoft.com/groups
#
# Remember: Always test changes in a non-production environment first!
#
#######################################################################################################################
