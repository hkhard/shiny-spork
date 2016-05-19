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
    [Parameter(Mandatory = $true,Position = 0,valueFromPipeline=$true)] [string $server = "sthdcsrvb152",
    [switch] $confirm = $false,
   )

# Include files and variable setup
. \\sthdcsrvb174.martinservera.net\script$\_lib\connect-exchange.ps1
. \\sthdcsrvb174.martinservera.net\script$\_lib\logfunctions.ps1


Function set-Maintenance([ValidateNotNullOrEmpty()]
                            [string $node])
{
    If ($node.tolower() -like "sthdcsrvb153*") {$otherServer = "sthdcsrvb152"}
    Else
    {$otherServer = "sthdcsrvb152"} 
    Get-MailboxDatabaseCopyStatus -Server $node | ? {$_.Status -eq "Mounted"} | % {Move-ActiveMailboxDatabase $_.DatabaseName -ActivateOnServer $otherServer -Confirm:$false}
    Set-MailboxServer $node -DatabaseCopyAutoActivationPolicy Blocked
    Set-ServerComponentState $node -Component ServerWideOffline -State Inactive -Requester Maintenance
}

Function verify-Maintenance([ValidateNotNullOrEmpty()]
                            [string $node])
{
    ( Get-ServerComponentState $node -Component ServerWideOffline | % {$_.State -eq "Inactive"})
}

# Main program
connect
set-Maintenance -node $server
If (verify-Maintenance -node $server) {LogLine "Maintenance mode enterned on Exhchange node $($server)"}
Else {LogErrorLine "Maintenance mode note entered for Exchange node $($Server)! Please Investigate" ; LogWarningLine $Error[0]}
