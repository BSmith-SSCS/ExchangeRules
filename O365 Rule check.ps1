
# Check for c:\Temp location.  This is where files are written.
$path = "C:\temp"

if (-Not (Test-Path -Path $path)) {
    New-Item -ItemType Directory -Path $path
    Write-Output "Directory created: $path"
} else {
    Write-Output "Directory already exists: $path"
}

# Connect to O365 Online.

$Credential = Get-Credential
Install-Module -Name ExchangeOnlineManagement
Connect-ExchangeOnline

#Show all mailbox rules
$Mailboxes = Get-Mailbox -ResultSize unlimited  | where {$_.RecipientTypeDetails -eq "UserMailbox"}
foreach ($Mailbox in $Mailboxes){Get-InboxRule -mailbox $Mailbox.Name | fl DateCreated,Name,Description,RuleIdentity > c:/temp/$Mailbox.txt}
