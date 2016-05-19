##########################################
### start exchange maintenance on one node
### Version History
### ===============
### 1.0 -- * Initial version
### ================
### End History info
####################

[CmdletBinding()]
param(
    [ValidateSet("sthdcsrvb152","sthdcsrvb153","sthdcsrvb152.martinservera.net","sthdcsrvb153.martinservera.net")]
    [ValidateNotNullOrEmpty()]
    [Parameter(Mandatory = $true,Position = 0,valueFromPipeline=$true)] [string] $server = "",
    [switch] $confirm = $false
   )

# Include files and variable setup
Unblock-File \\sthdcsrvb174.martinservera.net\script$\_lib\logFunctions.ps1 -Confirm:$false
Unblock-File \\sthdcsrvb174.martinservera.net\script$\_lib\ad.ps1 -Confirm:$false
. \\sthdcsrvb174.martinservera.net\script$\_lib\connect-exchange.ps1
. \\sthdcsrvb174.martinservera.net\script$\_lib\logfunctions.ps1

Function set-Maintenance([ValidateNotNullOrEmpty()]
                            [string] $node)
{
    If ($node.tolower() -like "sthdcsrvb153*") {$otherServer = "sthdcsrvb152"}
    Else {$otherServer = "sthdcsrvb152"} 
    Get-MailboxDatabaseCopyStatus -Server $node | ? {$_.Status -eq "Mounted"} | % {Move-ActiveMailboxDatabase $_.DatabaseName -ActivateOnServer $otherServer -Confirm:$false}
    Set-MailboxServer $node -DatabaseCopyAutoActivationPolicy Blocked
    Set-ServerComponentState $node -Component ServerWideOffline -State Inactive -Requester Maintenance
}

Function verify-Maintenance([ValidateNotNullOrEmpty()]
                            [string] $node)
{
    ( Get-ServerComponentState $node -Component ServerWideOffline | % {$_.State -eq "Inactive"})
}

# Main program
#Set up logging
$scriptFileName = ($MyInvocation.MyCommand.Name).split(".")[0]
$logFilePath = "\\sthdcsrvb174.martinservera.net\script$\_log\"
openLogFile "$logFilePath$(($MyInvocation.MyCommand.name).split('.')[0])-$(get-date -uformat %D)-$env:USERNAME.log"

#Do work
connect
If ($confirm) {set-Maintenance -node $server} else {LogLine "Would have entered maintenance mode on Exchange node $($server)"}
If (($confirm) -and (verify-Maintenance -node $server)) {LogLine "Maintenance mode entered on Exhchange node $($server)"}
Else { If ($confirm) {LogErrorLine "Maintenance mode note entered for Exchange node $($Server)! Please Investigate" ; LogWarningLine $Error[0]}}

