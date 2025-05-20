<#
.SYNOPSIS
    Adds a specified Microsoft 365 group as a site collection administrator to all SharePoint sites in the tenant.

.DESCRIPTION
    This script connects to the SharePoint Admin center and iterates through all SharePoint sites in the tenant
    (excluding OneDrive sites), adding a specified Microsoft 365 group as a site collection administrator if 
    it doesn't already exist in that role. The script includes throttling management and logging functionality.

.PARAMETER AdminURL
    The URL of the SharePoint Admin center.

.PARAMETER groupname
    The identity of the Microsoft 365 group to be added as a site collection administrator.

.PARAMETER appID
    The application ID (client ID) for the app registration in Azure AD used for authentication.

.PARAMETER thumbprint
    The certificate thumbprint used for authentication.

.PARAMETER tenant
    The tenant ID (GUID) of the Microsoft 365 tenant.

.NOTES
    File Name      : Add-GroupToSites.ps1
    Author         : Mike Lee / Vijay Kumar / Darin Roulston
    Created On     : 3/11/25

.EXAMPLE
    .\Add-GroupToSites.ps1

.OUTPUTS
    Log file at $env:TEMP\Adding_Group_to_Sites_[timestamp].txt with information about the operations performed.
#>
# Variables for processing
$AdminURL = "" #Example:"https://contoso-admin.sharepoint.com/"
$groupname = "" #Example:"c:0t.c|tenant|ed046cb9-86bc-47e7-95f5-912cfe343fc2"
$appID = "" #Example: "1e488dc4-1977-48ef-8d4d-9856f4e04536"
$thumbprint = "" #Example: "5EAD7303A5C7E27DB4245878AD554642940BA082"
$tenant = "" #Example: "9cfc42cb-51da-4055-87e9-b20a170b6ba3"

$startime = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = "$env:TEMP\Adding_Group_to_Sites_$startime.txt"

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

# Function to check if the groupname already exists in the site collection
function Test-SiteCollectionAdminExists {
    param (
        [string]$groupname
    )
    try {
        $admins = Get-PnPSiteCollectionAdmin 
        foreach ($admin in $admins) {
            if ($admin.LoginName -eq $groupname) {
                return $true
            }
        }
        return $false
    } catch {
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

# Function to process sites and add site collection admin if not exists
function Invoke-Sites {
    param (
        [string]$groupname
    )
    $Sites = Get-PnPTenantSite -Filter "Url -notlike '-my.sharepoint.com'"
    foreach ($Site in $Sites) {
        Connect-PnPOnline -Url $Site.Url -ApplicationId $appID -Tenant $tenant -Thumbprint $thumbprint
        Write-LogMessage "Processing Site Collection: $($Site.URL)" "INFO"
        if (-not (Test-SiteCollectionAdminExists -groupname $groupname)) {
            Write-LogMessage "Adding Site Collection Admin for: $($Site.URL)" "ADDED"
            Set-PnPTenantSite -Url $Site.Url -Owners $groupname
            Write-LogMessage "Successfully added Site Collection Admin for: $($Site.URL)" "INFO"
        } else {
            Write-LogMessage "Group already exists as Site Collection Admin for: $($Site.URL)" "PRESENT"
        }
    }
}

# Main function to handle throttling and process sites
function add-groups {
    while ($true) {
        try {
            Invoke-Sites -groupname $groupname
            break
        } catch {
            if ($_.Exception.Response.StatusCode.Value__ -eq 429) {
                $RetryAfter = [int]$_.Exception.Response.Headers["Retry-After"]
                Wait-Throttling -RetryAfter $RetryAfter
            } else {
                Write-LogMessage "Error processing sites: $_" "ERROR"
                throw $_
            }
        }
    }
}

# Run the main function
add-groups
