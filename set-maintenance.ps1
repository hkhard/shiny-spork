##########################################
### start exchange maintenance on one node
### Version History
### ===============
### 1.0 -- * Initial version
### 1.1 -- * With stop-Maintenance
### ================
### End History info
####################

[CmdletBinding()]
param(
    [ValidateSet("sthdcsrvb152","sthdcsrvb153","sthdcsrvb152.martinservera.net","sthdcsrvb153.martinservera.net")]
    [ValidateNotNullOrEmpty()]
    [Parameter(Mandatory = $true,Position = 0,valueFromPipeline=$true)] [string] $server = "",
    [switch] $stopMaintenance = $false,
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
    $return = $true
    If ($node.tolower() -like "sthdcsrvb153*") {$otherServer = "sthdcsrvb152"}
    Else {$otherServer = "sthdcsrvb152"}
    try {
        Get-MailboxDatabaseCopyStatus -Server $node | ? {$_.Status -eq "Mounted"} | % {Move-ActiveMailboxDatabase $_.DatabaseName -ActivateOnServer $otherServer -Confirm:$false}
    }
    catch [System.Exception] {
        LogErrorLine $Error[0]
        $return = $false
    }
    try {
        Set-MailboxServer $node -DatabaseCopyAutoActivationPolicy Blocked
    }
    catch [System.Exception] {
        LogErrorLine $Error[0]
        $return = $false
    }
    try {
        Set-ServerComponentState $node -Component ServerWideOffline -State Inactive -Requester Maintenance
    }
    catch [System.Exception] {
       LogErrorLine $Error[0]
       $return = $false  
    }
    $return
}

Function verify-Maintenance([ValidateNotNullOrEmpty()]
                            [string] $node)
{
    ( Get-ServerComponentState $node -Component ServerWideOffline | % {$_.State -eq "Inactive"})
}

Function get-serverInMaintenanceMode()
{
    $return = $null
    $serverStatus = Get-exchangeServer | Get-ServerComponentState -Component ServerWideOffline 
    foreach ($item in $serverStatus) {
        if ($($item.State) -eq "Inactive") {
            $return = $($item).Identity
        }
    }
    $return
}

Function stop-Maintenance([ValidateNotNullOrEmpty()] 
                          [String] $node)
{
    Try {Set-ServerComponentState $node -Component ServerWideOffline -State Active -Requester Maintenance} Catch {LogErrorLine $Error[0]}
    Try {Set-MailboxServer $node -DatabaseCopyAutoActivationPolicy Unrestricted} Catch {LogErrorLine $Error[0]}
}                          

# Main program
# Set up logging
$scriptFileName = ($MyInvocation.MyCommand.Name).split(".")[0]
$logFilePath = "\\sthdcsrvb174.martinservera.net\script$\_log\"
openLogFile "$logFilePath$(($MyInvocation.MyCommand.name).split('.')[0])-$(get-date -uformat %D)-$env:USERNAME.log"

# Do work
If (!($stopMaintenance))
{
    If ($confirm) {
        connect
        if (!(set-Maintenance -node $server))
        {       
            LogErrorLine "Could not enter maintenance mode on Exchange node $($server)"
        }
        If (verify-Maintenance -node $server) {LogLine "Maintenance mode entered on Exhchange node $($server)"}
        Else { LogErrorLine "Maintenance mode note entered for Exchange node $($Server)! Please Investigate" ; LogWarningLine $Error[0]}
}
else {
    If ($confirm) {
        connect
        $server = get-serverInMaintenanceMode
        If (($confirm) -and ($server)) {stop-Maintenance -node $server}
        If (($confirm) -and (!(verify-Maintenance -node $server))) {LogLine "Maintenance mode exited on Exhchange node $($server)"}
    }
}

