# 1. Install the Exchange Online Management Module (if you haven't already)
# Install-Module ExchangeOnlineManagement

# 2. Connect to Exchange Online
Connect-ExchangeOnline

# 3. Define the output path
$OutputPath = "C:\scripts\AllInboxRules.csv"

# 4. Get all mailboxes and iterate through them to get rules
Get-Mailbox -ResultSize Unlimited | ForEach-Object {
    Get-InboxRule -Mailbox $_.PrimarySmtpAddress | Select-Object @{Name="Mailbox";Expression={$_.MailboxOwnerId}}, Name, Enabled, Priority, Description, From, RedirectTo, ForwardTo
} | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

# 5. Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false

Write-Host "Report exported to $OutputPath"
