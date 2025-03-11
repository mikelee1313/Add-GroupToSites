What does this script do?

This script connects to the SharePoint Admin center and iterates through all SharePoint sites in the tenant (excluding OneDrive sites), adding a specified Microsoft 365 group as a site collection administrator if it doesn't already exist in that role. The script includes throttling management and logging functionality.

The log will default the uses temp folder %temp%

Example of output log:

![image](https://github.com/user-attachments/assets/eb2901f0-3bd5-4c16-b69e-96d8d4433e0e)

To execute this script you will need the PNP module installed.

Ensure to modify these settings to match your Tenant

AdminURL = "https://contoso-admin.sharepoint.com/"
$groupname = "c:0t.c|tenant|ed046cb9-86bc-47e7-95f5-912cfe343fc2"
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082"
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"
