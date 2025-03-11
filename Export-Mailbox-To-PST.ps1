# Exchange Mailbox Export Script
# Author: aviado
# Description:
# This script automates the process of exporting a user's mailbox to a PST file in Microsoft Exchange.
# It ensures that the target directory exists before exporting, initiates the export request,
# and sends email notifications upon request initiation and completion.
#
# Features:
# - Automatically retrieves the full display name of the user from Exchange.
# - Ensures the target folder exists before exporting.
# - Sends email notifications when the export starts and when it completes.
# - Automatically cleans up completed export requests.
#
# Requirements:
# - Exchange Management Shell must be installed and running on the system.
# - The executing user must have the necessary permissions to export mailboxes.
# - The target folder must be accessible by the Exchange server.

# Load Exchange PowerShell module
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

# Variables
$mailboxName = "sampleuser"  # Change this to the username of the mailbox to export
$fullName = (Get-Mailbox -Identity $mailboxName).DisplayName  # Automatically retrieves the full name
$exportPath = "\\192.168.1.100\ExportShare\PST_Exports\$fullName\$mailboxName.pst"
$emailRecipients = "admin@example.com", "support@example.com"  # Modify with actual recipient emails
$exchangeServer = "mail.example.com"  # Set your Exchange server address

# Ensure the export directory exists
if (!(Test-Path -Path "\\192.168.1.100\ExportShare\PST_Exports\$fullName")) {
    New-Item -ItemType Directory -Path "\\192.168.1.100\ExportShare\PST_Exports\$fullName" -Force
}

# Start email notification
$startTime = Get-Date
$startBody = @"
Your Export PST request has been received. You will receive additional updates during its progress.

File:    $exportPath
Mailbox:    $fullName
Started by:    Administrator
Start time:    $($startTime.ToString("g"))
Run time:    00:00:00

Please do not reply to this email. It was sent from an unmonitored account.
"@

Send-MailMessage -From "exchange-alert@example.com" `
                 -To $emailRecipients `
                 -Subject "Your Export PST request has been received" `
                 -Body $startBody `
                 -SmtpServer $exchangeServer

# Initiate export request
New-MailboxExportRequest -Mailbox $mailboxName -FilePath $exportPath

# Check export request status until completed
$status = "Queued"
while ($status -eq "Queued" -or $status -eq "InProgress") {
    Start-Sleep -Seconds 30
    $status = (Get-MailboxExportRequest -Mailbox $mailboxName).Status
}

# Completion time
$endTime = Get-Date
$runTime = New-TimeSpan -Start $startTime -End $endTime

# Completed email notification
$completionBody = @"
Export PST has finished.

File:    $exportPath
Mailbox:    $fullName
Started by:    Administrator
Start time:    $($startTime.ToString("g"))
Run time:    $($runTime.ToString("hh\:mm\:ss"))

Please do not reply to this email. It was sent from an unmonitored account.
"@

Send-MailMessage -From "exchange-alert@example.com" `
                 -To $emailRecipients `
                 -Subject "Export PST has finished" `
                 -Body $completionBody `
                 -SmtpServer $exchangeServer

# Optional: Cleanup completed requests
Get-MailboxExportRequest -Mailbox $mailboxName | Remove-MailboxExportRequest -Confirm:$false
