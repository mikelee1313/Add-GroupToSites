<#
.SYNOPSIS
    Retrieves all site collection administrators from SharePoint Online sites and exports them to a CSV file.

.DESCRIPTION
    This script connects to a SharePoint Online tenant and retrieves all site collection administrators from 
    all site collections. It includes direct user admins, members of the site's owners group, and members of 
    Entra ID (formerly Azure AD) groups that have site collection admin rights. The results are exported to a CSV file.

    The script includes throttling protection with retry logic to handle SharePoint Online throttling.
    
    You can add Entra ID group IDs to the $ignoreGroupIds array to exclude specific groups from being processed,
    which can be useful for very large groups or service accounts that don't need to be included in the report.
    
    The $Debug variable controls logging verbosity. When set to $false (default), only essential information and
    errors/warnings are logged. When set to $true, verbose logging is enabled for troubleshooting.

.PARAMETER SiteURL
    The SharePoint admin center URL.

.PARAMETER appID
    The application (client) ID for the app registration in Azure AD.

.PARAMETER thumbprint
    The certificate thumbprint for authentication.

.PARAMETER tenant
    The tenant ID for the Microsoft 365 tenant.

.NOTES
    File Name     : Get-SCA-AllSites.ps1
    
    Prerequisite  : PnP PowerShell module installed (Tested with  3.1.0)
                  : API Perms:
                      Application: Sharepoint: Sites.FullControl.All
                      Application: Graph: Directory.Read.All
    
    Author        : Mike Lee | Vijay Kumar
    Date          : 5/14/2025

.EXAMPLE
    .\Get-SCA-AllSites.ps1

.OUTPUTS
    A CSV file with all site collection administrators is created in the %TEMP% folder.
    A log file is also created in the %TEMP% folder for troubleshooting purposes.
#>

# Set Variables
$tenantname = "m365x61250205" #This is your tenant name
$appID = "5baa1427-1e90-4501-831d-a8e67465f0d9"  #This is your Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9" #This is certificate thumbprint
$tenant = "85612ccb-4c28-4a34-88df-a538cc139a51" #This is your Tenant ID
$Debug = $false #Set to $true for verbose logging, $false for essential logs only

# Groups to ignore - Add group IDs to this array to skip processing them
$ignoreGroupIds = @(
    "cd733102-d898-4be0-a80b-4f8a833a8795",
    "85706948-d975-413c-9be0-1f3e9c2dedbc"
    # Example: "12345678-1234-5678-abcd-1234567890ab",
    # Example: "87654321-4321-8765-dcba-0987654321fe"
)

# Function to check if a group ID should be ignored
function Test-IgnoreGroup {
    param (
        [string]$groupId
    )
    return $ignoreGroupIds -contains $groupId
}

#Define Log path
$startime = Get-Date -Format "yyyyMMdd_HHmmss"
$ouputpath = "$env:TEMP\" + 'SiteCollectionAdmins_' + $startime + ".csv"
$logFilePath = "$env:TEMP\" + 'SiteCollectionAdmins_' + $startime + ".log"

# Function to write site data to CSV incrementally
function Write-SiteDataToCSV {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$SiteData,
        [Parameter(Mandatory = $true)]
        [string]$CsvPath,
        [switch]$CreateHeader
    )
    
    # Prepare the data for CSV export
    $csvData = [PSCustomObject]@{
        SiteUrl          = $SiteData.SiteUrl
        DirectAdmins     = ($SiteData.DirectAdmins -join "; ")
        SPGroupAdmins    = ($SiteData.SPGroupAdmins -join "; ")
        EntraGroupAdmins = ($SiteData.EntraGroupAdmins -join "; ")
    }
    
    # Determine if we should include headers
    $exportParams = @{
        Path              = $CsvPath
        Append            = (-not $CreateHeader)
        NoTypeInformation = $true
        Encoding          = 'UTF8'
    }
    
    # Export data to CSV
    $csvData | Export-Csv @exportParams
    
    Write-Log "Data for site $($SiteData.SiteUrl) written to CSV" -level "INFO" -ForceLog:$CreateHeader
}

#This is the logging function with debug support
function Write-Log {
    param (
        [string]$message,
        [string]$level = "INFO",
        [switch]$ForceLog = $false
    )
    
    # Only log if Debug is true, or if it's an ERROR/WARNING, or if ForceLog is specified
    if ($Debug -or $level -eq "ERROR" -or $level -eq "WARNING" -or $ForceLog) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "$timestamp - $level - $message"
        Add-Content -Path $logFilePath -Value $logMessage
        
        # Also echo important messages to console with color coding
        if ($level -eq "ERROR") {
            Write-Host "$timestamp - $message" -ForegroundColor Red
        }
        elseif ($level -eq "WARNING") {
            Write-Host "$timestamp - $message" -ForegroundColor Yellow
        }
        elseif ($ForceLog) {
            # Important INFO messages that are forced to log
            Write-Host "$timestamp - $message" -ForegroundColor White
        }
    }
}

Write-Host "Starting script to get Site Collection Admins at $startime" -ForegroundColor Yellow
Write-Log "Starting script to get Site Collection Admins" -level "INFO" -ForceLog
if ($Debug) {
    Write-Host "Debug mode is ON. Verbose logging enabled." -ForegroundColor Cyan
    Write-Log "Debug mode is ON. Verbose logging enabled." -level "INFO"
}
$SiteURL = "https://$tenantname-admin.sharepoint.com"
Connect-PnPOnline -Url $SiteURL -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant

# Function to handle throttling
function Invoke-WithRetry {
    param (
        [scriptblock]$ScriptBlock
    )
    $retryCount = 0
    $maxRetries = 5
    while ($retryCount -lt $maxRetries) {
        try {
            $result = & $ScriptBlock
            return $result
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq 429) {
                $retryAfter = $_.Exception.Response.Headers["Retry-After"]
                if (-not $retryAfter) {
                    $retryAfter = 30 # Default retry interval in seconds
                    Write-Warning "Throttled. 'Retry-After' header missing. Using default retry interval of $retryAfter seconds."
                }
                else {
                    Write-Warning "Throttled. Retrying after $retryAfter seconds."
                }
                Start-Sleep -Seconds $retryAfter
                $retryCount++ 
            }
            else {
                throw $_
            }
        }
    }
    throw "Max retries reached. Exiting."
}

# Get all site collections
$sites = Invoke-WithRetry { Get-PnPTenantSite | Where-Object { $_.Url -notlike "*-my.sharepoint.com*" } }
Write-Log "Retrieved $($sites.Count) site collections" -level "INFO" -ForceLog

# Track whether this is the first site (to create CSV header)
$firstSite = $true
$processedSites = 0

foreach ($site in $sites) {
    # Connect to each site collection
    Write-Host "Getting Site Collection Admins from: $($site.Url)" -ForegroundColor Green
    Write-Log "Getting Site Collection Admins from: $($site.Url)" -level "INFO"
    
    try {
        # Connect to the site collection
        Connect-PnPOnline -Url $site.Url -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant

        Invoke-WithRetry { Connect-PnPOnline -Url $site.Url -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant }

        # Initialize data structure for this site
        $siteData = [PSCustomObject]@{
            SiteUrl          = $site.Url
            DirectAdmins     = @()
            SPGroupAdmins    = @()
            EntraGroupAdmins = @()
        }

        # Get site collection administrators
        $admins = Invoke-WithRetry { Get-PnPSiteCollectionAdmin }
        Write-Log "Found $($admins.Count) site collection admins for $($site.Url)" -level "INFO" -ForceLog

        foreach ($admin in $admins) {
            if ($admin.PrincipalType -eq "User") {
                # Add user admin details - combine name and email
                $siteData.DirectAdmins += "$($admin.Title) <$($admin.Email)>"
            }
            elseif ($admin.PrincipalType -eq "SecurityGroup" -and $admin.Title.ToLower().Contains("owners")) {
                try {
                    $groupMembers = Invoke-WithRetry { Get-PnPGroupMember -Identity $admin.Title }
                
                    foreach ($member in $groupMembers) {
                        $siteData.SPGroupAdmins += "$($admin.Title): $($member.Title) <$($member.Email)>"
                    }
                }
                catch {
                    Write-Log "Group '$($admin.Title)' in site '$($site.Url)' is deleted or inaccessible. Trying Fallback: $_" -level "WARNING"
                    try {
                        $spgroup = Invoke-WithRetry { Get-PnPGroup -Identity $site.Url }
                        if ($spgroup.Title.ToLower().Contains("owners")) {
                            $groupMembers = Invoke-WithRetry { Get-PnPGroupMember -Identity $spgroup.Title }
                        
                            #Check if there are members in the group
                            if ($groupMembers.Count -ge 0) {
                                foreach ($member in $groupMembers) {
                                    $siteData.SPGroupAdmins += "$($spgroup.Title): $($member.Title) <$($member.Email)>"
                                }
                            }
                        }
                    }
                    catch {
                        Write-Log "Failed to retrieve members for group '$($admin.Title)' in site '$($site.Url)': $_" -level "ERROR"
                    }
                }
            }
            elseif ($admin.PrincipalType -eq "SecurityGroup" -and $admin.Title.ToLower() -notlike '*owners*') {
                # Check if this is an Entra ID (Azure AD) group
                if ($admin.LoginName -like "c:0t.c|tenant|*") {
                    try {
                        # Extract the group ID from the login name
                        $entraGroupId = $admin.LoginName.Replace("c:0t.c|tenant|", "")
                        
                        # Check if this group should be ignored
                        if (Test-IgnoreGroup -groupId $entraGroupId) {
                            Write-Log "Skipping ignored Entra Group '$($admin.Title)' with ID: $entraGroupId" -level "DEBUG" #-ForceLog
                            $siteData.EntraGroupAdmins += "$($admin.Title): [Group excluded from processing]"
                            continue
                        }
                    
                        # Get group members using Microsoft Graph
                        $entraGroupMembers = Invoke-WithRetry { Get-PnPMicrosoft365GroupMembers -Identity $entraGroupId }
                    
                        # Get group owners as well
                        $entraGroupOwners = Invoke-WithRetry { Get-PnPMicrosoft365GroupOwners -Identity $entraGroupId }
                        
                        # Add group owners with a special designation
                        if ($entraGroupOwners -and $entraGroupOwners.Count -gt 0) {
                            foreach ($owner in $entraGroupOwners) {
                                $siteData.EntraGroupAdmins += "$($admin.Title): [OWNER] $($owner.DisplayName) <$($owner.Email)>"
                            }
                        }
                        
                        # Add regular members
                        if ($entraGroupMembers.Count -ge 0) {
                            foreach ($member in $entraGroupMembers) {
                                $siteData.EntraGroupAdmins += "$($admin.Title): $($member.DisplayName) <$($member.Email)>"
                            }
                        }
                    }
                    catch {
                        Write-Log "Failed to retrieve members for Entra ID group '$($admin.Title)' in site '$($site.Url)': $_" -level "ERROR"
                    }
                }
                else {
                    # Regular SharePoint group
                    $siteData.SPGroupAdmins += "$($admin.Title) <$($admin.Email)>"
                }
            }
        }

        # Write the data for this site to the CSV file
        Write-SiteDataToCSV -SiteData $siteData -CsvPath $ouputpath -CreateHeader:$firstSite
        $firstSite = $false
        $processedSites++
        
        # Clean up to free memory
        Remove-Variable siteData

    }
    catch {
        Write-Log "Failed to connect to site '$($site.Url)': $_" -level "ERROR"
        continue
    }
}

Write-Log "Processed $processedSites site collections successfully" -level "INFO" -ForceLog
Write-Host "Operations completed successfully and results exported to $ouputpath" -ForegroundColor Yellow
Write-Host "Check Log file any issues: $logFilePath" -ForegroundColor Cyan
Write-Log "Operations completed successfully and results exported to $ouputpath" -level "INFO" -ForceLog
