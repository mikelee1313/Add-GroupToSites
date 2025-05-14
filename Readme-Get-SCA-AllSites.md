# Get-SCA-AllSites.ps1

`Get-SCA-AllSites.ps1` is a PowerShell script designed to retrieve a list of all site collection administrators (SCAs) across all SharePoint sites within your Microsoft 365 tenant. This script is part of the `Add-GroupToSites` repository, which focuses on managing and automating administrative tasks for SharePoint Online.

## Features

- Enumerates all SharePoint sites in the tenant.
- Retrieves the site collection administrators for each site.
- Outputs the data in a structured format for further analysis or reporting.

## Prerequisites

Before running this script, ensure you have the following:

1. **PowerShell**  
   Make sure PowerShell is installed on your system.
   
2. **SharePoint Online Management Shell**  
   Install the SharePoint Online Management Shell by following the instructions [here](https://learn.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online?view=sharepoint-ps).

3. **Administrative Permissions**  
   You must have tenant administrator privileges to execute this script.

4. **Microsoft 365 Credentials**  
   Ensure you have credentials with sufficient permissions to access all SharePoint sites and retrieve SCAs.

## Setup and Installation

1. Clone the repository or download the script directly:
   ```bash
   git clone https://github.com/mikelee1313/Add-GroupToSites.git
   ```
   Alternatively, download the `Get-SCA-AllSites.ps1` script from the following URL:  
   [Get-SCA-AllSites.ps1](https://github.com/mikelee1313/Add-GroupToSites/blob/main/Get-SCA-AllSites.ps1)

2. Place the script in a directory of your choice.

## Usage

1. Open PowerShell and navigate to the directory containing the script.

2. Run the script using the following command:
   ```powershell
   ./Get-SCA-AllSites.ps1
   ```

3. Provide the required credentials when prompted.

4. The script will output the list of site collection administrators for all sites in the tenant. You can redirect the output to a file if needed:
   ```powershell
   ./Get-SCA-AllSites.ps1 > output.txt
   ```

## Output

The script generates a structured list of site collection administrators for all SharePoint sites. You can customize the script to export the data in a CSV format for easier consumption:
```powershell
./Get-SCA-AllSites.ps1 | Export-Csv -Path "SCAs.csv" -NoTypeInformation
```

## Troubleshooting

- **Authentication Issues**: Ensure your credentials have the necessary permissions to access the tenant and SharePoint sites.
- **Module Missing**: If you encounter errors related to missing modules, ensure the SharePoint Online Management Shell is installed and imported:
  ```powershell
  Import-Module Microsoft.Online.SharePoint.PowerShell
  ```

## Contributing

Contributions, issues, and feature requests are welcome! Feel free to check the [issues page](https://github.com/mikelee1313/Add-GroupToSites/issues) for existing issues or create a new one.

## License

This script is distributed under the [MIT License](https://github.com/mikelee1313/Add-GroupToSites/blob/main/LICENSE). Feel free to use, modify, and distribute it.

## About the Repository

The `Add-GroupToSites` repository is designed to streamline the management of SharePoint Online administration. For more scripts and tools, check out the repository:  
[Add-GroupToSites on GitHub](https://github.com/mikelee1313/Add-GroupToSites)

---

Let me know if you'd like to make any modifications or add further details!
