<#
.SYNOPSIS
    Retrieves all site collection administrators from SharePoint Online sites and exports them to a CSV file.

.DESCRIPTION
    This script connects to a SharePoint Online tenant and retrieves all site collection administrators from 
    all site collections. It includes direct user admins, members of the site's owners group, and members of 
    Entra ID (formerly Azure AD) groups that have site collection admin rights. The results are exported to a CSV file.

    The script includes throttling protection with retry logic to handle SharePoint Online throttling.

.PARAMETER SiteURL
    The SharePoint admin center URL.

.PARAMETER appID
    The application (client) ID for the app registration in Azure AD.

.PARAMETER thumbprint
    The certificate thumbprint for authentication.

.PARAMETER tenant
    The tenant ID for the Microsoft 365 tenant.

.NOTES
    File Name      : Get-SCA-AllSites.ps1
    Prerequisite   : PnP PowerShell module installed
    Author         : Mike Lee | Vijay Kumar
    Date           : 4/11/2025

.EXAMPLE
    .\Get-SCA-AllSites.ps1

.OUTPUTS
    A CSV file with all site collection administrators is created in the %TEMP% folder.
    A log file is also created in the %TEMP% folder for troubleshooting purposes.
#>

# Set Variables
$tenantname = "m365cpi13246019" #This is your tenant name
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"  #This is your Entra App ID
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082" #This is certificate thumbprint
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3" #This is your Tenant ID


#Define Log path
$startime = Get-Date -Format "yyyyMMdd_HHmmss"
$ouputpath = "$env:TEMP\" + 'SiteCollectionAdmins_' + $startime + ".csv"
$logFilePath = "$env:TEMP\" + 'SiteCollectionAdmins_' + $startime + ".log"

#This is the logging function
function Write-Log {
    param (
        [string]$message,
        [string]$level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $level - $message"
    Add-Content -Path $logFilePath -Value $logMessage
}

Write-Host "Starting script to get Site Collection Admins at $startime" -ForegroundColor Yellow
Write-Log "Starting script to get Site Collection Admins" -level "INFO"
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

# Create a hashtable to store unique site/group combinations
$resultsHash = @{}

foreach ($site in $sites) {
    # Connect to each site collection
    Write-Host "Getting Site Collection Admins from: $($site.Url)" -ForegroundColor Green
    Write-Log "Getting Site Collection Admins from: $($site.Url)"
    Invoke-WithRetry { Connect-PnPOnline -Url $site.Url -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant }

    # Get site collection administrators
    $admins = Invoke-WithRetry { Get-PnPSiteCollectionAdmin }

    foreach ($admin in $admins) {
        if ($admin.PrincipalType -eq "User") {
            # Add user admin details directly - these are still individual entries
            $key = "$($site.Url)_User_$($admin.Email)"
            $resultsHash[$key] = [PSCustomObject]@{
                SiteUrl       = $site.Url
                Title         = $admin.Title
                Email         = $admin.Email
                Members       = $admin.Title
                MemberEmails  = $admin.Email
                PrincipalType = "User"
            }
        }
        elseif ($admin.PrincipalType -eq "SecurityGroup" -and $admin.Title.ToLower().Contains("owners")) {
            try {
                # Create a key for this site and group combination
                $key = "$($site.Url)_SPGroup_$($admin.Title)"
                
                # Initialize with empty members if this is a new entry
                if (-not $resultsHash.ContainsKey($key)) {
                    $resultsHash[$key] = [PSCustomObject]@{
                        SiteUrl       = $site.Url
                        Title         = $admin.Title
                        Email         = ""
                        Members       = ""
                        MemberEmails  = ""
                        PrincipalType = "SharePoint Owners Group"
                    }
                }
                
                try {
                    $groupMembers = Invoke-WithRetry { Get-PnPGroupMember -Identity $admin.Title }
                    
                    $memberNames = @()
                    $memberEmails = @()
                    
                    foreach ($member in $groupMembers) {
                        $memberNames += $member.Title
                        $memberEmails += $member.Email
                    }
                    
                    # Update the entry with consolidated members
                    $resultsHash[$key].Members = $memberNames -join "; "
                    $resultsHash[$key].MemberEmails = $memberEmails -join "; "
                }
                catch {
                    Write-Log "Group '$($admin.Title)' in site '$($site.Url)' is deleted or inaccessible: $_" -level "WARNING"
                }
            }
            catch {
                Write-Log "Failed to retrieve members for group '$($admin.Title)' in site '$($site.Url)': $_" -level "WARNING"
            }
        }
        elseif ($admin.PrincipalType -eq "SecurityGroup" -and $admin.Title.ToLower() -notlike '*owners*') {
            # Check if this is an Entra ID (Azure AD) group or a SharePoint group
            if ($admin.LoginName -like "c:0t.c|tenant|*") {
                # Create a key for this site and group combination
                $key = "$($site.Url)_EntraGroup_$($admin.Title)"
                
                # Initialize with empty members if this is a new entry
                if (-not $resultsHash.ContainsKey($key)) {
                    $resultsHash[$key] = [PSCustomObject]@{
                        SiteUrl       = $site.Url
                        Title         = $admin.Title
                        Email         = ""
                        Members       = ""
                        MemberEmails  = ""
                        PrincipalType = "Entra ID Group"
                    }
                }
                
                try {
                    # Extract the group ID from the login name
                    $entraGroupId = $admin.LoginName.Replace("c:0t.c|tenant|", "")
                    
                    # Get group members using Microsoft Graph
                    $entraGroupMembers = Invoke-WithRetry { Get-PnPMicrosoft365GroupMembers -Identity $entraGroupId }
                    
                    $memberNames = @()
                    $memberEmails = @()
                    
                    foreach ($member in $entraGroupMembers) {
                        $memberNames += $member.DisplayName
                        $memberEmails += $member.Email
                    }
                    
                    # Update the entry with consolidated members
                    $resultsHash[$key].Members = $memberNames -join "; "
                    $resultsHash[$key].MemberEmails = $memberEmails -join "; "
                }
                catch {
                    Write-Log "Failed to retrieve members for Entra ID group '$($admin.Title)' in site '$($site.Url)': $_" -level "WARNING"
                }
            }
        }
        else {
            # Add SharePoint group details
            $key = "$($site.Url)_SPGroup_$($admin.Title)"
            $resultsHash[$key] = [PSCustomObject]@{
                SiteUrl       = $site.Url
                Title         = $admin.Title
                Email         = $admin.Email
                Members       = $admin.Title
                MemberEmails  = $admin.Email
                LoginName     = $admin.LoginName
                PrincipalType = "SharePoint Group"
            }
        }
    }
}

# Convert hashtable to array for CSV export
$consolidatedResults = $resultsHash.Values

# Export results to CSV
$consolidatedResults | Export-Csv -Path $ouputpath -NoTypeInformation -Encoding UTF8
Write-Host "Operations completed successfully and results exported to $ouputpath" -ForegroundColor Yellow
Write-Host "Check Log file any issues: $logFilePath" -ForegroundColor Cyan
Write-Log "Operations completed successfully and results exported to $ouputpath" -level "INFO"
