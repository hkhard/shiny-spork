<#
    .SYNOPSIS
    #####################################################################
    # Created by Kontract 2012-2016, v2.1 (c)
    # (Hans.Hard@kontract.se)
    #####################################################################	
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	Version 2.1, 20th May, 2016
	
    .DESCRIPTION
	Move all active Exchange databases off from the provided server name and onto
    the other server node within Martin & Servera's Exchange 2016 system.
      
	.PARAMETER Server
	A server name from the list (it is possible to tab between the alternatives):
        * sthdcsrvb152
        * sthdcsrvb153
        * sthdcsrvb152.martinservera.net
        * sthdcsrvb153.martinservera.net
	
	.PARAMETER StopMaintenance
	Switch, If supplied on command line or set to true, will try and stop maintenance mode on the given server name

    .PARAMETER Confirm
	If supplied on command line or set to true, will actually execute the maintenance operation.

    #>
###########################################
### start exchange maintenance on one node
### (c) 2016, Kontract IS AB // Hans K Hård
### Version History
### ===============
### 1.0 -- * Initial version
### 1.1 -- * With stop-Maintenance
### 2.0 -- * Adjusted start-maintenance with new routines
### 2.1 -- * Added updated stop-maintenance according to new start-routine
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

####################
# Include files
####################
Unblock-File \\sthdcsrvb174.martinservera.net\script$\_lib\logFunctions.ps1 -Confirm:$false
Unblock-File \\sthdcsrvb174.martinservera.net\script$\_lib\ad.ps1 -Confirm:$false
. \\sthdcsrvb174.martinservera.net\script$\_lib\connect-exchange.ps1
. \\sthdcsrvb174.martinservera.net\script$\_lib\logfunctions.ps1

################
# Set up logging
################
$scriptFileName = ($MyInvocation.MyCommand.Name).split(".")[0]
$logFilePath = "\\sthdcsrvb174.martinservera.net\script$\_log\"
openLogFile "$logFilePath$(($MyInvocation.MyCommand.name).split('.')[0])-$(get-date -uformat %D)-$env:USERNAME.log"

############################
# Start function definitions
############################

#####################################################################
# Function set-Maintenance by Kontract (c)
#  (Hans.Hard@kontract.se)
#
# Sets the supplied $node into maintenance mode
#
#####################################################################
Function set-Maintenance([ValidateNotNullOrEmpty()]
                            [string] $node)
{
    $return = $true
    If ($node.tolower() -like "sthdcsrvb153*") {$otherServer = "sthdcsrvb152"}
    Else {$otherServer = "sthdcsrvb152"}
    $nodeFQDN = $node + "martinservera.net"
    $otherServerFQDN = $otherServer + "martinservera.net"
    if ($return) { try
        {
            Set-ServerComponentState -server $node -Component HubTransport -State Draining -Requester Maintenance
            Restart-Service MSExchangeTransport -server $node -Confirm:$false
            Redirect-Message -Server $node -Target $otherServerFQDN -Requester Maintenance -Confirm:$false
        }
    catch [System.Exception] {
        LogErrorLine $Error[0]
        $return = $false
    }
    }
    if ($return) { try 
        {
            Set-ServerComponentState $node -Component UMCallRouter -State Draining -Requester Maintenance
        }
    catch [System.Exception] {
        LogErrorLine $Error[0]
        $return = $false
    }
    }
    if (($return) -and (((Get-DatabaseAvailabilityGroup -Identity DAG01 -Status).primaryActiveManager).tolower -eq $node )) { try 
        {
            LogLine " Moving Cluster-group to $($otherServer).."
             Move-ClusterGroup "Cluster Group" -Node $otherServer -Confirm:$false 
             Suspend-ClusterNode -Node $node -Confirm:$false  
             while (((Get-DatabaseAvailabilityGroup -Identity DAG01 -Status).primaryActiveManager).tolower -eq $node ) {
                 write-host -ForegroundColor Red "." -NoNewline
                 start-sleep 0.5
             }
             Write-Host " "
             LogLineWithColour -sLine "Done!" -sColour "Green" 
    }
    catch [System.Exception] {
        LogErrorLine $Error[0]
        $return = $false
    }
    if ($return) { try 
        {
            Get-MailboxDatabaseCopyStatus -Server $node | ? {$_.Status -eq "Mounted"} | % {Move-ActiveMailboxDatabase $_.DatabaseName -ActivateOnServer $otherServer -Confirm:$false}
            while ( ((Get-MailboxDatabaseCopyStatus -Server $node  | ? {$_.Status -eq "Mounted"}).count) -gt 0 ) {
                 write-host -ForegroundColor Red "." -NoNewline
                 start-sleep 0.5
             }
             Write-Host " "
             LogLineWithColour -sLine "Done!" -sColour "Green" 
        }
        catch [System.Exception] {
            LogErrorLine $Error[0]
            $return = $false
        }
    }
    if ($return) { try
        {
            Set-MailboxServer $node -DatabaseCopyAutoActivationPolicy Blocked
        }
        catch [System.Exception] {
            LogErrorLine $Error[0]
            $return = $false
        }
    }
    if ($return) { try 
        {
            Set-ServerComponentState $node -Component ServerWideOffline -State Inactive -Requester Maintenance
        }
        catch [System.Exception] {
           LogErrorLine $Error[0]
        $return = $false  
        }
    }
    $return }
}


#####################################################################
# Function verify-Maintenance by Kontract (c)
#  (Hans.Hard@kontract.se)
#
# Checks to see if the supplied $node is in maintenance mode,
#  returns $true if it is so.
#
#####################################################################
Function verify-Maintenance([ValidateNotNullOrEmpty()]
                            [string] $node)
{
    ( Get-ServerComponentState $node -Component ServerWideOffline | % {$_.State -eq "Inactive"})
}

#####################################################################
# Function get-serverInMaintenanceMode by Kontract (c)
#  (Hans.Hard@kontract.se)
#
# Returns the last server currently in maintenance mode
#
#####################################################################
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

#####################################################################
# Function stop-Maintenance by Kontract (c)
#  (Hans.Hard@kontract.se)
#
# Stops maintenance mode on the supplide $node
#
#####################################################################
Function stop-Maintenance([ValidateNotNullOrEmpty()] 
                          [String] $node)
{
    $return = $true
    If ($return){Try {Set-ServerComponentState $node -Component ServerWideOffline -State Active -Requester Maintenance} Catch {LogErrorLine $Error[0];$return=$false}}
    If ($return){Try {Set-ServerComponentState $node -Component UMCallRouter -State Active -Requester Maintenance} Catch {LogErrorLine $Error[0];$return=$false}}
    If ($return) {Try {Resume-ClusterNode $node -Confirm:$false} Catch {LogErrorLine $Error[0];$return=$false}}
    If ($return){Try {Set-MailboxServer $node -DatabaseCopyAutoActivationPolicy Unrestricted} Catch {LogErrorLine $Error[0];$return=$false}}
    If ($return){Try {
        Set-ServerComponentState $node -Component HubTransport -State Active -Requester Maintenance -confirm:$false
        Restart-Service MSExchangeTransport -server $node -Confirm:$false
    }
    Catch {LogErrorLine $Error[0];$return=$false}}
    $report = Get-ServerComponentState $node | ft Component,State -AutoSize
    LogLine " Status of exchange components on server $($node):"
    LogLine $report
}                          

############################
# Start main program
############################
StartStopInfo -sAction "start"
#Check that we are in elevated powershell environment
If (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
 LogErrorLine "Script is not run with administrative user. Please, restart Powershell with elevated privileges... Quitting!"
 Remove-PSSession *
 LogLine "Script set-maintenance aborted!"
 StartStopInfo -sAction "stop"
 Break
}

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
}
else {
    If ($confirm) {
        connect
        $server = get-serverInMaintenanceMode
        If ($server) {stop-Maintenance -node $server}
        If (!(verify-Maintenance -node $server)) {
            LogLine "Maintenance mode exited on Exhchange node $($server)"
            LogLine " "
            LogLinewithColour -sLine "Now go to either Exchange server and in PowerShell, run the following:" -sColour Green
            LogLine " "
            LogLine " cd ´$exscripts" 
            LogLine ".\RedistributeActiveDatabases.ps1 -BalanceDbsByActivationPreference -ShowFinalDatabaseDistribution -confirm:´$false"
        }
    }
}
StartStopInfo -sAction "stop"