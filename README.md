# Exchange Mailbox Export Script

## Overview
This PowerShell script automates the process of exporting a user's mailbox to a PST file in Microsoft Exchange. It ensures the necessary directory structure exists, initiates the export request, and provides email notifications at both the start and completion of the process.

## Features
- Automatically retrieves the full display name of the user from Exchange.
- Ensures the target folder exists before exporting.
- Sends email notifications when the export starts and when it completes.
- Cleans up completed export requests automatically.

## Prerequisites
- The script must be run on an **Exchange Management Shell** or a system with Exchange PowerShell modules installed.
- The executing user must have **Mailbox Export** permissions.
- The target folder must be accessible by the Exchange server.

## Installation & Usage
### 1. Assign Mailbox Export Permissions
Before exporting a mailbox, ensure the executing user has permission to export mailboxes:
```powershell
New-ManagementRoleAssignment –Role "Mailbox Import Export" –User <YourAdminUser>
```
> Restart the Exchange Management Shell after assigning the role.

### 2. Run the Script
Execute the script from an Exchange Management Shell:
```powershell
.\Export-Mailbox-To-PST.ps1
```
> Update variables in the script with real mailbox names and network paths before execution.

## Configuration
Modify the following variables in the script as needed:
- **`$mailboxName`** - The username of the mailbox to export.
- **`$emailRecipients`** - Recipients of email notifications.
- **`$exportPath`** - Network path where the PST file will be saved.
- **`$exchangeServer`** - Your Exchange server address.

## Troubleshooting
1. **Error: Path not found**
   - Ensure the target export directory exists or is accessible by Exchange.
   - Verify the script is running with appropriate permissions.
2. **Error: User does not have permission**
   - Assign `Mailbox Import Export` permissions to the user.
3. **No Email Notifications Received**
   - Verify SMTP settings and email recipient addresses.

## Disclaimer
This script is provided **as is** without any warranty. Use it at your own risk. Ensure you have proper backups before making changes to your Exchange environment.

## License
MIT License - Free to use and modify. Credit is appreciated.

## Author
Script Author: [aviado1](https://github.com/aviado1)
