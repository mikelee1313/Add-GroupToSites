<#
.SYNOPSIS
    Removes a specified Microsoft 365 group as a site collection administrator from all SharePoint sites in the tenant.

.DESCRIPTION
    This script connects to the SharePoint Admin center and iterates through all OneDrive sites 
    and removes a specified Microsoft 365 group as a site collection administrator if 
    it exists in that role. The script includes throttling management and logging functionality.

.PARAMETER AdminURL
    The URL of the SharePoint Admin center.

.PARAMETER $loginname
    The identity of the Microsoft 365 group to be removed as a site collection administrator.

.PARAMETER appID
    The application ID (client ID) for the app registration in Azure AD used for authentication.

.PARAMETER thumbprint
    The certificate thumbprint used for authentication.

.PARAMETER tenant
    The tenant ID (GUID) of the Microsoft 365 tenant.

.NOTES
    File Name      : Remove-ODBUsers.ps1
    Author         : Mike Lee | Mike Thames
    Created On     : 7/21/25

.EXAMPLE
    .\Remove-ODBUsers.ps1

.OUTPUTS
    Log file at $env:TEMP\Removing_Group_from_Sites_[timestamp].txt with information about the operations performed.
#>

#########################################################################
# USER CONFIGURATION - MODIFY THESE SETTINGS BEFORE RUNNING THE SCRIPT
#########################################################################

$AdminURL = "https://M365CPI13246019-admin.sharepoint.com" 
$loginname = "i:0#.f|membership|admin@m365cpi13246019.onmicrosoft.com"
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082" 
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"

#########################################################################
# END OF USER CONFIGURATION
#########################################################################

$startime = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = "$env:TEMP\Removing_Group_from_Sites_$startime.txt"

# Function to handle throttling
function Wait-Throttling {
    param (
        [int]$RetryAfter
    )
    Write-Host "Throttled. Retrying after $RetryAfter seconds..."
    Start-Sleep -Seconds $RetryAfter
}

# Connect to SharePoint Online
Connect-PnPOnline -ApplicationId $appID -Tenant $tenant -Url $AdminURL -Thumbprint $thumbprint

# Function to check if the $loginname already exists in the site collection
function Test-SiteCollectionAdminExists {
    param (
        [string]$loginname
    )
    try {
        $admins = Get-PnPSiteCollectionAdmin 
        foreach ($admin in $admins) {
            if ($admin.LoginName -eq $loginname) {
                return $true
            }
        }
        return $false
    }
    catch {
        Write-Host "Error checking site collection admins: $_"
        return $false
    }
}

# Function to log messages
function Write-LogMessage {
    param (
        [string]$message,
        [string]$level
    )
    $logFile = $logFilePath
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $logEntry = "$timestamp - $level - $message"
    Add-Content -Path $logFile -Value $logEntry
    Write-Host $message
}

# Function to process sites and remove site collection admin if exists
function Invoke-Sites {
    param (
        [string]$loginname
    )
    $Sites = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '*-my.sharepoint.com/personal*'"
    foreach ($Site in $Sites) {
        try {
            Connect-PnPOnline -Url $Site.Url -ApplicationId $appID -Tenant $tenant -Thumbprint $thumbprint
            Write-LogMessage "Processing Site Collection: $($Site.URL) for user: $loginname" "INFO"
            if (Test-SiteCollectionAdminExists -loginname $loginname) {
                Write-LogMessage "Removing Site Collection Admin '$loginname' for: $($Site.URL)" "REMOVED"
                Remove-PnPSiteCollectionAdmin -Owners $loginname
                Write-LogMessage "Successfully removed Site Collection Admin '$loginname' for: $($Site.URL)" "INFO"
            }
            else {
                Write-LogMessage "User '$loginname' does not exist as Site Collection Admin for: $($Site.URL)" "NOT_PRESENT"
            }
        }
        catch {
            Write-LogMessage "Error accessing site $($Site.URL): $_" "ERROR"
            Write-LogMessage "Continuing with next site..." "INFO"
            continue
        }
    }
}

# Main function to handle throttling and process sites
function remove-groups {
    while ($true) {
        try {
            Invoke-Sites -loginname $loginname
            break
        }
        catch {
            if ($_.Exception.Response.StatusCode.Value__ -eq 429) {
                $RetryAfter = [int]$_.Exception.Response.Headers["Retry-After"]
                Wait-Throttling -RetryAfter $RetryAfter
            }
            else {
                Write-LogMessage "Error processing sites: $_" "ERROR"
                throw $_
            }
        }
    }
}

# Run the main function
remove-groups
