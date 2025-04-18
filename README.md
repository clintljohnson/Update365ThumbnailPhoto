# Update365ThumbnailPhotos

A PowerShell script that synchronizes user thumbnail photos from Active Directory to Microsoft 365 services (SharePoint, Exchange Online, and Teams). The script provides both a dry-run mode for testing and a force mode for actual updates.

## Overview

This script addresses a common limitation in Microsoft 365 synchronization: while the `thumbnailPhoto` attribute is initially synced from on-premises Active Directory to Exchange Online, subsequent changes to this attribute don't automatically sync. This script provides an automated solution to keep user photos synchronized across your Microsoft 365 environment.

## Prerequisites

### Required PowerShell Modules
- **ActiveDirectory**: For accessing on-premises AD
- **Microsoft.Graph**: For Microsoft 365 operations (minimum version 2.26.1)
- **PSSQLite** (Optional): For photo change tracking

### Azure AD App Registration
1. An Azure AD application registration with the following:
   - API Permissions: `ProfilePhoto.ReadWrite.All` (Application permission)
   - A valid client secret
   - The application must be granted admin consent for the permissions

## Installation

1. Clone or download the script to your desired location
2. Ensure all required PowerShell modules are installed:
```powershell
Install-Module -Name Microsoft.Graph -Force
Install-Module -Name ActiveDirectory -Force
Install-Module -Name PSSQLite -Force  # Optional, for change tracking
```

## Client Secret Management

### When Creating/Updating Client Secret

When your client secret expires (maximum duration is 2 years) or when setting up the script for the first time, run the script with the `-SaveCredentials` parameter:

```powershell
.\Update365ThumbnailPhotos.ps1 `
    -OrganizationalUnit "OU=Whoville,DC=CONTOSO,DC=COM" `
    -Force `
    -TenantId '917dd6ff-2b18-43c4-95fa-1bcf78e31ea2' `
    -ClientId '891296c1-5e3f-9ef2-ebf2-5e2f649f12c1' `
    -ClientSecret 'Apo8Q~q3RieMEyWC5KXvfe2ZeQ4SbNmtXiY8jchh' `
    -SaveCredentials
```

This will save the credentials to `Update365ThumbnailPhotos.config` in the script directory.

### Regular Usage

Once credentials are saved, you can run the script without providing the authentication parameters:

```powershell
.\Update365ThumbnailPhotos.ps1 `
    -OrganizationalUnit "OU=Whoville,DC=CONTOSO,DC=COM" `
    -Force
```

## Features

- **Dry Run Mode**: Test changes without applying them (omit `-Force` parameter)
- **Force Mode**: Apply actual changes to Microsoft 365 (use `-Force` parameter)
- **Photo Change Tracking**: Tracks which photos have been updated (requires PSSQLite module)
- **Detailed Logging**: HTML-formatted logs with color-coded status messages
- **Credential Storage**: Secure storage of app registration credentials
- **Selective Updates**: Update specific users or organizational units
- **Email Notifications**: Optional email notifications of successful updates

## Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| OrganizationalUnit | Yes* | Distinguished name of the OU containing users to update |
| SamAccountName | Yes* | Specific user account to update |
| Domain | No | Domain name (default: CONTOSO.COM) |
| LogFile | No | Path to log file (default: .\PhotoUpdate_timestamp.html) |
| PlainTextLog | No | Use plain text logging instead of HTML |
| Force | No | Apply changes instead of dry run |
| TenantId | No** | Azure AD tenant ID |
| ClientId | No** | Azure AD application ID |
| ClientSecret | No** | Azure AD application secret |
| SaveCredentials | No | Save provided credentials for future use |
| EmailResults | No | Send email with update results |
| EmailFrom | Yes*** | Sender email address |
| EmailTo | Yes*** | Recipient email address(es) - multiple addresses can be specified, separated by commas |
| EmailSubject | No | Email subject (default: Microsoft 365 Photo Update Results) |
| EmailRelay | No | SMTP relay server (default: smtp.office365.com) |
| EmailRelayPort | No | SMTP relay port (default: 25) |

\* Either OrganizationalUnit or SamAccountName must be specified, but not both
\** Required only when credentials haven't been saved previously
\*** Required when -EmailResults is used

## How It Works

1. **Authentication**: The script authenticates to Microsoft Graph using client credentials flow:
```powershell
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential
```

2. **User Discovery**: Retrieves users with thumbnail photos from Active Directory:
```powershell
$users = Get-ADUser -SearchBase $OrganizationalUnit -Filter * -Properties thumbnailPhoto, mail
```

3. **Photo Synchronization**: For each user:
   - Checks if the photo has changed (if using PSSQLite)
   - Creates a temporary file with the photo data
   - Updates the photo in Microsoft 365 using Microsoft Graph
   - Saves a copy in the Images folder
   - Updates the tracking database (if using PSSQLite)

4. **Email Notifications**: If enabled, sends an HTML-formatted email with:
   - List of successfully updated users
   - User display names and email addresses
   - Photo sizes
   - Total count of updates

## Logging

The script creates detailed logs in HTML format (by default) with color-coded status messages:
- 🟢 Success: Green
- 🔴 Error: Red
- 🟡 Warning: Yellow
- ⚪ Info: White
- 🔵 Current: Cyan

Logs are saved in the script directory with the naming pattern: `PhotoUpdate_YYYYMMDD_HHMMSS.html`

## Best Practices

1. Always run in dry-run mode first (without `-Force`)
2. Review the generated logs for any potential issues
3. Keep track of when your client secret expires (maximum 2 years)
4. Regularly backup the credentials configuration file
5. Monitor the size of the SQLite database if using photo change tracking
6. When using email notifications, ensure your SMTP relay is properly configured
7. Test email functionality with a small set of users first

## Troubleshooting

### Common Issues

1. **Authentication Errors**
   - Verify app registration permissions
   - Ensure admin consent has been granted
   - Check if client secret has expired

2. **Photo Update Failures**
   - Verify photo size (should be less than 100KB)
   - Check user's email address in Active Directory
   - Ensure photo is in a compatible format (JPEG recommended)

3. **Module-Related Errors**
   - Update Microsoft.Graph module to latest version
   - Ensure all required modules are installed
   - Check PowerShell execution policy

4. **Email Delivery Issues**
   - Verify SMTP relay server is accessible
   - Check if port 25 is open (or specified port)
   - Ensure sender email address is valid
   - Verify recipient email address is correct

## License

This script is provided as-is without any warranty. Use at your own risk.

## Email Notifications

If enabled, sends an HTML-formatted email with:
- List of successfully updated users
- User display names and email addresses
- Photo sizes
- Total count of updates

Multiple recipients can be specified by separating email addresses with commas in the `-EmailTo` parameter:

```powershell
.\Update365ThumbnailPhotos.ps1 `
    -OrganizationalUnit "OU=Whoville,DC=CONTOSO,DC=COM" `
    -Force `
    -EmailResults `
    -EmailFrom "photosync@contoso.com" `
    -EmailTo "admin@contoso.com,helpdesk@contoso.com" `
    -EmailSubject "Daily Office 365 Photo Synchronization" `
    -EmailRelay "smtp.contoso.com"
```
