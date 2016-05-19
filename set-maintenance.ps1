# start exchange maintenance on one node

. \\sthdcsrvb152\script$\_lib\connect-exchange.ps1

# Main program
connect

Get-MailboxDatabaseCopyStatus -Server sthdcsrvb152 | ? {$_.Status -eq "Mounted"} | % {Move-ActiveMailboxDatabase $_.DatabaseName -ActivateOnServer sthdcsrvb153 -Confirm:$false}
Set-MailboxServer sthdcsrvb152 -DatabaseCopyAutoActivationPolicy Blocked
Set-ServerComponentState sthdcsrvb152 -Component ServerWideOffline -State Inactive -Requester Maintenance

Get-ServerComponentState sthdcsrvb152 -Component ServerWideOffline
