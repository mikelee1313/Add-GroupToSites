<#
.SYNOPSIS
    Retrieves all Site Collection Administrators across SharePoint Online sites in a tenant.

.DESCRIPTION
    This script connects to SharePoint Online using a certificate-based app registration and identifies
    all site collection administrators across regular SharePoint sites (excluding OneDrive sites).
    It exports the results to a CSV file with details about each admin, including site URL, admin name,
    user/email, and principal type. For SharePoint groups containing "Owners" in the name, the script
    attempts to resolve group membership.

.PARAMETER None
    This script uses predefined variables in the script body that need to be customized:
    - $tenantname: The SharePoint tenant name
    - $appID: The Entra App ID for authentication
    - $thumbprint: The certificate thumbprint for authentication
    - $tenant: The Tenant ID

.NOTES
    File Name      : GGet-SCA-AllSites-noEntraMembers_SimpleOutput.ps1
    Prerequisite   : PnP PowerShell module installed
    Author         : Mike Lee | Vijay Kumar
    Date           : 4/18/2025

    The script implements throttling retry logic to handle SharePoint Online throttling.
    Results are exported to the user's temp directory with a timestamp in the filename.
    A log file is also created in the temp directory to track script execution.

.EXAMPLE
    .\Get-SCA-AllSites-noEntraMembers_SimpleOutput.ps1
    
    Runs the script with the hardcoded tenant information and exports results to the temp directory.

.OUTPUTS
    - CSV file with columns: SiteUrl, Title, User, PrincipalType
    - Log file tracking script execution

.FUNCTIONALITY
    - Connects to SharePoint Admin Center
    - Retrieves all site collections (excluding OneDrive sites)
    - Connects to each site collection
    - Gets site collection administrators
    - For SharePoint groups with "owners" in the name, resolves group membership
    - Exports all data to CSV
#>
# Set Variables
$tenantname = "CONTOSO" #This is your tenant name
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

# Replace the results processing and export section with this:
# Prepare a hashtable to store results by site URL
$siteResults = @{}

foreach ($site in $sites) {
    # Connect to each site collection
    Write-Host "Getting Site Collection Admins from: $($site.Url)" -ForegroundColor Green
    Write-Log "Getting Site Collection Admins from: $($site.Url)"
    Invoke-WithRetry { Connect-PnPOnline -Url $site.Url -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant }

    # Get site collection administrators
    $admins = Invoke-WithRetry { Get-PnPSiteCollectionAdmin }
    
    # Create arrays to store different admin types for this site
    $userAdmins = @()
    $groupAdmins = @()
    $otherAdmins = @()

    foreach ($admin in $admins) {
        if ($admin.PrincipalType -eq "User") {
            # Add user admin details
            $userAdmins += "$($admin.Title) ($($admin.Email))"
        }
        elseif ($admin.PrincipalType -eq "SecurityGroup" -and $admin.Title.ToLower().Contains("owners")) {
            try {
                # Retrieve group members explicitly
                try {
                    $groupMembers = Invoke-WithRetry { Get-PnPGroupMember -Identity $admin.Title }
                    $membersList = ($groupMembers | ForEach-Object { "$($_.Title) ($($_.Email))" }) -join "; "
                    $groupAdmins += "Group: $($admin.Title) - Members: [$membersList]"
                }
                catch {
                    try {
                        $spgroup = Invoke-WithRetry { Get-PnPGroup -Identity $site.Url }
                        if ($spgroup.Title.ToLower().Contains("owners")) {
                            $groupMembers = Invoke-WithRetry { Get-PnPGroupMember -Identity $spgroup.Title }
                            $membersList = ($groupMembers | ForEach-Object { "$($_.Title) ($($_.Email))" }) -join "; "
                            $groupAdmins += "Group: $($admin.Title) - Members: [$membersList]"
                        }
                    }
                    catch {
                        Write-Log "Failed to retrieve members for group '$($admin.Title)' in site '$($site.Url)': $_" -level "WARNING"
                        $groupAdmins += "Group: $($admin.Title) - Members: [Unable to retrieve]"
                    }
                }
            }
            catch {
                Write-Log "Error processing admin '$($admin.Title)' in site '$($site.Url)': $_" -level "ERROR"
                $groupAdmins += "Group: $($admin.Title) - Members: [Error retrieving]"
            }
        }
        else {
            # Add all other types of admins
            $otherAdmins += "$($admin.Title) ($($admin.PrincipalType))"
        }
    }
    
    # Store the consolidated information for this site
    $siteResults[$site.Url] = [PSCustomObject]@{
        SiteUrl = $site.Url
        UserAdmins = ($userAdmins -join " | ")
        GroupAdmins = ($groupAdmins -join " | ")
        OtherAdmins = ($otherAdmins -join " | ")
    }
}

# Convert hashtable to array and export results to CSV
$consolidatedResults = $siteResults.Values
$consolidatedResults | Export-Csv -Path $ouputpath -NoTypeInformation -Encoding UTF8
Write-Host "Operations completed successfully and results exported to $ouputpath" -ForegroundColor Yellow
Write-Host "Check Log file for any issues: $logFilePath" -ForegroundColor Cyan
Write-Log "Operations completed successfully and results exported to $ouputpath" -level "INFO"
