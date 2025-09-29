# SharePoint Site Groupification Script

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)](https://github.com/PowerShell/PowerShell)
[![PnP PowerShell](https://img.shields.io/badge/PnP%20PowerShell-2.0%2B-orange)](https://pnp.github.io/powershell/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

A comprehensive PowerShell script for validating SharePoint Online sites and connecting them to Microsoft 365 Groups. This script provides enterprise-grade validation, safety features, and detailed reporting for bulk SharePoint site groupification operations.

## 🎯 Overview

The **Groupify-SPOSites.ps1** script automates the process of analyzing SharePoint Online sites for Microsoft 365 Group compatibility and optionally connecting eligible sites to groups. It performs extensive validation to ensure only compatible sites are processed, preventing errors and maintaining site functionality.

### Key Features

- **🔍 Comprehensive Validation Pipeline**: 8-step validation process including template compatibility, publishing features, and modern UI checks
- **📊 Dual Input Methods**: Auto-discovery of all tenant sites or CSV-based processing for specific sites
- **🛡️ Safety-First Design**: Validation-only mode by default with extensive safety checks and limits
- **🔐 Interactive Authentication**: Works with MFA, conditional access, and all tenant configurations
- **📈 Detailed Reporting**: Comprehensive logging and CSV exports for audit and analysis
- **⚡ Enterprise Ready**: Built for large-scale operations with proper error handling and monitoring

## 📋 Prerequisites

### Required Software
- **PowerShell 5.1** or **PowerShell 7+**
- **PnP PowerShell module 2.0+**
  ```powershell
  Install-Module PnP.PowerShell -Force
  ```

### Required Permissions
- **SharePoint Administrator** or **Global Administrator** role in Microsoft 365
- Access to SharePoint Admin Center
- Permissions to create Microsoft 365 Groups (for groupification operations)

### Supported Site Templates
The script validates and supports the following SharePoint site templates:
- `STS#0` - Team Site (classic)
- `STS#3` - Team Site (no Microsoft 365 Group)
- `BLOG#0` - Blog site
- `DEV#0` - Developer Site
- `PROJECTSITE#0` - Project Site

## 🚀 Quick Start

### 1. Download and Configure
```powershell
# Download the script
# Configure the USER CONFIGURATION section with your tenant details
```

### 2. Run Validation (Recommended First)
```powershell
# Safe validation-only run - no changes made
.\Groupify-SPOSites.ps1
```

### 3. Review Results
Check the generated CSV file and logs to understand which sites are eligible for groupification.

### 4. Perform Groupification (Optional)
```powershell
# Enable groupification after reviewing validation results
# Set $enableGroupify = $true in the script configuration
.\Groupify-SPOSites.ps1
```

## ⚙️ Configuration

### User Configuration Section
Edit the following variables in the script's `USER CONFIGURATION` section:

```powershell
# =================================================================================================
# USER CONFIGURATION - Update the variables in this section
# =================================================================================================
$appID = "your-app-id"                                    # Your Entra App ID
$tenant = "your-tenant-id"                                # Your Tenant ID  
$csvInputFilePath = ""                                    # Path to CSV file (empty for auto-discovery)
$tenantAdminUrl = "https://yourtenant-admin.sharepoint.com"

# Groupify configuration
$enableGroupify = $false                                  # Set to $true to perform actual groupification
$groupifyMaxSites = 5                                     # Safety limit for number of sites per run
$groupifyIsPublic = $false                                # Create private groups by default
$groupifyKeepOldHomePage = $false                         # Use new group home page
```

## 🔧 Usage Examples

### Auto-Discovery Mode (Scan All Sites)
```powershell
# Discover and validate all sites in the tenant
.\Groupify-SPOSites.ps1
```

### CSV Input Mode (Process Specific Sites)
```powershell
# Process sites from SharePoint Admin Center export
$csvInputFilePath = "C:\SharePoint\sites-export.csv"
.\Groupify-SPOSites.ps1
```

### Groupification Mode (Live Changes)
```powershell
# Enable groupification for eligible sites
$enableGroupify = $true
$groupifyMaxSites = 10
.\Groupify-SPOSites.ps1
```

## 📊 Input Formats

### CSV File Format
The script supports multiple CSV formats from SharePoint Admin Center:

**Standard Format:**
```csv
URL,Site name,Sensitivity,Microsoft 365 group
https://tenant.sharepoint.com/sites/site1,Team Site 1,None,No
https://tenant.sharepoint.com/sites/site2,Team Site 2,Confidential,Yes
```

**Alternative Format:**
```csv
Site URL,Site Title,Sensitivity Label
https://tenant.sharepoint.com/sites/site1,Team Site 1,None
https://tenant.sharepoint.com/sites/site2,Team Site 2,Confidential
```

## 🔍 Validation Pipeline

The script performs comprehensive validation through these steps:

1. **Site Accessibility** - Verifies site exists and is accessible
2. **Template Compatibility** - Ensures site template supports group connection
3. **Publishing Features** - Detects blocking SharePoint Publishing features
4. **Modern UI Features** - Identifies features that impact modern experience
5. **Group Connection Status** - Determines if site is already group-connected
6. **Permission Validation** - Confirms sufficient permissions for operations
7. **Alias Availability** - Validates proposed group alias is available
8. **Final Eligibility** - Comprehensive eligibility determination

## 📈 Output and Reporting

### Generated Files
- **Main Log**: `SPO-Site-Groupify-Sitefinder_log_[timestamp].txt`
- **Error Log**: `SPO-Site-Groupify-Sitefinder_error_[timestamp].txt`
- **Results CSV**: `Sites-Ready-For-Groups_[timestamp].csv`

### Console Output
The script provides color-coded console output:
- 🟢 **Green**: Successful operations and validations
- 🟡 **Yellow**: Warnings and informational messages  
- 🔴 **Red**: Errors and validation failures
- 🔵 **Cyan**: Important notifications and headers

### Results CSV Fields
```csv
Site URL,Site Title,Template,Owner,Storage Used (MB),Last Modified,
Sensitivity Label,Current Status,Site Status,Sharing Capability
```

## 🛡️ Safety Features

### Built-in Protections
- **Validation-Only Default**: No changes made unless explicitly enabled
- **Maximum Site Limits**: Configurable safety limits for bulk operations
- **Comprehensive Logging**: Detailed audit trail of all operations
- **Error Recovery**: Graceful handling of individual site failures
- **Pre-flight Checks**: Extensive validation before making any changes

### Best Practices
1. **Always run validation first** before enabling groupification
2. **Test with small batches** before processing large numbers of sites
3. **Review logs and CSV exports** to understand validation results
4. **Monitor group-connected sites** for 24-48 hours after groupification
5. **Back up important sites** before bulk operations

## 🔧 Troubleshooting

### Common Issues and Solutions

#### Authentication Problems
```
Issue: "Failed to connect to SharePoint Online"
Solution: Verify tenant name, check network connectivity, confirm user permissions
```

#### Publishing Feature Conflicts
```
Issue: "Publishing features are enabled and block group connection"
Solution: Disable publishing features or exclude these sites from groupification
```

#### Template Incompatibility
```
Issue: "Site template not compatible with groups"  
Solution: Migrate content to compatible templates or exclude incompatible sites
```

#### Permission Errors
```
Issue: "Insufficient permissions for site access"
Solution: Ensure you have SharePoint Administrator or Global Administrator role
```

### Debug Mode
Enable verbose logging by modifying the log level in the configuration section.

## 📚 Documentation References

- [Microsoft 365 Groups and SharePoint](https://docs.microsoft.com/en-us/sharepoint/dev/transform/modernize-connect-to-office365-group)
- [PnP PowerShell Documentation](https://pnp.github.io/powershell/)
- [SharePoint Site Templates](https://docs.microsoft.com/en-us/sharepoint/sites/create-site-collection)
- [SharePoint Admin Center](https://docs.microsoft.com/en-us/sharepoint/get-started-new-admin-center)

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

### Development Guidelines
1. Follow PowerShell best practices and coding standards
2. Add comprehensive comments for new functionality
3. Include error handling for all new operations
4. Update documentation for any new features
5. Test thoroughly in non-production environments

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## ⚠️ Disclaimer

This script modifies SharePoint Online sites and creates Microsoft 365 Groups. Always test in a non-production environment first. The authors are not responsible for any data loss or service disruption resulting from the use of this script.

## 🏷️ Version History

- **v1.0** - Initial release with basic groupification functionality
- **v2.0** - Enhanced validation pipeline and safety features
- **v3.0** - Comprehensive documentation and admin guidance
- **Current** - Production-ready script with extensive validation and reporting

## 🆘 Support

For support and questions:
1. Check the [troubleshooting section](#-troubleshooting) above
2. Review the detailed logs generated by the script
3. Consult Microsoft's official SharePoint documentation
4. Open an issue in this repository for bugs or feature requests

---

**⭐ If this script helps you, please consider giving it a star!**
