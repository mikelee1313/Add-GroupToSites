# Add-GroupToSites

**Adds or removes a specified Microsoft 365 group as a site collection administrator across all SharePoint Online and OneDrive sites in a tenant. Also includes tools to report on site collection administrators.**

---

## Table of Contents

- [Overview](#overview)
- [Scripts](#scripts)
  - [Add-GroupToSPOSites.ps1](#add-grouptospositesps1)
  - [Add-GroupToOneDriveSites.ps1](#add-grouptoonedrivesitesps1)
  - [Remove-SPOUsers.ps1](#remove-spousersps1)
  - [Remove-ODBUsers.ps1](#remove-odbusersps1)
  - [Get-SCA-AllSites.ps1](#get-sca-allsitesps1)
- [Requirements](#requirements)
- [Authentication](#authentication)
- [Logging](#logging)
- [Notes](#notes)

---

## Overview

This repository contains PowerShell scripts to automate the management of site collection administrators in Microsoft 365. The scripts allow you to add or remove a Microsoft 365 group as a site collection admin across all SharePoint Online sites, including OneDrive sites, and report on admin configurations. All scripts support certificate-based authentication and robust logging.

---

## Scripts

### Add-GroupToSPOSites.ps1

**Adds a specified Microsoft 365 group as a site collection administrator to all SharePoint Online sites (excluding OneDrive).**

- Connects to the SharePoint Admin Center using certificate authentication.
- Iterates all SharePoint Online sites (excluding OneDrive and portals).
- Adds the group as site collection admin if not already present.
- Handles throttling and logs all actions.

---

### Add-GroupToOneDriveSites.ps1

**Adds a specified Microsoft 365 group as a site collection administrator to all OneDrive sites in the tenant.**

- Connects to the SharePoint Admin Center using certificate authentication.
- Iterates all OneDrive sites.
- Adds the group as site collection admin if not already present.
- Handles throttling and logs all actions.

---

### Remove-SPOUsers.ps1

**Removes a specified Microsoft 365 group as a site collection administrator from all SharePoint Online sites (excluding OneDrive).**

- Connects to the SharePoint Admin Center using certificate authentication.
- Iterates all SharePoint Online sites (excluding OneDrive and portals).
- Removes the group as site collection admin if present.
- Handles throttling and logs all actions.

---

### Remove-ODBUsers.ps1

**Removes a specified Microsoft 365 group as a site collection administrator from all OneDrive sites.**

- Connects to the SharePoint Admin Center using certificate authentication.
- Iterates all OneDrive sites.
- Removes the group as site collection admin if present.
- Handles throttling and logs all actions.

---

### Get-SCA-AllSites.ps1

**Reports all site collection administrators for every SharePoint Online site (excluding OneDrive).**

- Connects to the SharePoint Admin Center using certificate authentication.
- Iterates all SharePoint Online sites (excluding OneDrive).
- Exports direct admins, site owner group members, and Entra ID group members with admin rights to a CSV file.
- Supports ignoring specific groups and includes verbose logging for troubleshooting.

---

## Requirements

- PowerShell 7+
- [PnP PowerShell Module](https://pnp.github.io/powershell/)
- Certificate-based authentication set up in Azure AD for the required permissions:
  - SharePoint: `Sites.FullControl.All`
  - Microsoft Graph: `Directory.Read.All`

## Authentication

Each script uses certificate-based authentication with parameters for:
- Admin Center URL
- Microsoft 365 group identity
- Application (client) ID
- Certificate thumbprint
- Tenant ID

**Configure these parameters at the top of each script before running.**

---

## Logging

All scripts generate a timestamped log file in the `%TEMP%` directory, recording details about the operations performed and any errors encountered.

---

## Notes

- Scripts include built-in handling for SharePoint Online throttling (HTTP 429).
- Modify the user configuration section at the top of each script before running.
- For large tenants or very large groups, consider using the ignore list in `Get-SCA-AllSites.ps1` to skip specific Entra ID groups.

---

**For more details, see the script headers or the in-line comments within each script.**

---

If you need usage examples or help with configuring any script, please refer to the `.EXAMPLE` section in each script file or open an issue!
