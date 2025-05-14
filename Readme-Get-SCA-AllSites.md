# Get-SCA-AllSites.ps1

## Overview

`Get-SCA-AllSites.ps1` is a PowerShell script designed to retrieve all site collection administrators (SCAs) from SharePoint Online sites and export them to a CSV file. The script includes features to handle throttling, exclude specific groups, and log the process for troubleshooting.

## Features

- Retrieves SCAs from all SharePoint Online site collections.
- Identifies direct user admins, members of the site's owners group, and members of Entra ID (formerly Azure AD) groups with SCA rights.
- Handles SharePoint Online throttling using retry logic.
- Excludes specific Entra ID groups using an `ignoreGroupIds` array.
- Exports results to a CSV file and creates a log file for troubleshooting.

## Prerequisites

- **PnP PowerShell Module**: Tested with version 3.1.0.
- **API Permissions**:
  - Application: SharePoint: `Sites.FullControl.All`
  - Application: Microsoft Graph: `Directory.Read.All`
- An app registration in Azure AD with the necessary permissions.

## Parameters

- **`SiteURL`**: The SharePoint admin center URL.
- **`appID`**: The application (client) ID for the app registration in Azure AD.
- **`thumbprint`**: The certificate thumbprint for authentication.
- **`tenant`**: The tenant ID for the Microsoft 365 tenant.

## How to Use

1. **Set Variables**: Update the `$tenantname`, `$appID`, `$thumbprint`, and `$tenant` variables in the script with your tenant details.
2. **Run the Script**: Execute the script in PowerShell:
   ```powershell
   .\Get-SCA-AllSites.ps1
   ```
3. **Output**: 
   - A CSV file containing SCAs is created in the `%TEMP%` folder.
   - A log file is also created in the `%TEMP%` folder for troubleshooting purposes.

## Logging and Debugging

- **Logging**: Essential information, errors, and warnings are logged to a log file in the `%TEMP%` folder.
- **Verbose Logging**: Enable verbose logging by setting the `$Debug` variable to `$true`.

## Customizations

- **Exclude Groups**: Add group IDs to the `$ignoreGroupIds` array to exclude specific groups.
- **CSV Export**: The `Write-SiteDataToCSV` function handles CSV export and can be customized as needed.

## Example

```powershell
.\Get-SCA-AllSites.ps1
```

## Example Output:

![image](https://github.com/user-attachments/assets/252871dd-9bb4-44af-bd83-ec10c0db3ea9)


## Notes

- **Author**: Mike Lee | Vijay Kumar
- **Last Updated**: 5/14/2025
- **File Name**: `Get-SCA-AllSites.ps1`

---

For more information, refer to the [source script](https://github.com/mikelee1313/Add-GroupToSites/blob/main/Get-SCA-AllSites.ps1).
