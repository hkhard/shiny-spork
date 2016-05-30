<#
    .SYNOPSIS
    #####################################################################
    # Created by Kontract (c) 2012-2016, v2.12
    #  (Stefan.Alkman@kontract.se)
    #  (Hans.Hard@kontract.se)
    #####################################################################	
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	Version 2.12, 12th May, 2016
	
    .DESCRIPTION
	Migrate Exchange mailboxes from Martin & Servera's legacy exchange into
	new Exchange 2016 system
      
	.PARAMETER Users
	A comma separated list of user names. These should be on the form:
        xyxy,xzxz,zxzx,zyzy and so on
    Parameter can also be a piped list of user names
	
	.PARAMETER IncrementalSync
	If supplied on command line or set to true will set migration batch to stop in syncing mode. Requires manual intervention at ECP to complete.

	.PARAMETER AutoComplete
	If supplied on command line or set to true will set migration batch to automatically complete when ready.
	
    .PARAMETER Confirm
	If supplied on command line or set to true, will send migration batch command to Exchange.
	
	.PARAMETER inputFile
	A file containing a list of users to be migrated.

	.PARAMETER notifyME
	Information messages will be sent to the logged on administrator's e-mail address.


    #>

### Version History
### ===============
### 1.0 -- * Initial version
### 2.0 -- * First version adapted for M&S AB
### 2.01 - * Cleaned out some unneeded subroutines
### 2.02 - * Fix for Size HT error and changed migration command to local (we do not have remote migrations at this point.)
### 2.03 - * Ugly hack for converting Dict.Item of MDB sizes to HT
### 2.1 -- * Added group membership check and correction (for denySCP to make autodiscover work as it should)
### 2.11 - * Ugly hack to properly compare database sizes in FindBestFitDatabase soubroutine.
### 2.12 - * Added unblock-file to include files.
### ================
### End History info
####################

[CmdletBinding()]
param(
   [Parameter(Mandatory = $true,Position = 0,valueFromPipeline=$true)] [string[]] $users,
   [switch] $confirm = $false,
   [switch] $autoComplete = $false,
   [string] $inputFile,
   [switch] $NotifyME = $false,
   [Parameter(Mandatory = $false)] [int] $BadItemLimit = 0
   )

####################
# Include files
####################
Unblock-File \\sthdcsrvb174.martinservera.net\script$\_lib\logFunctions.ps1 -Confirm:$false
Unblock-File \\sthdcsrvb174.martinservera.net\script$\_lib\ad.ps1 -Confirm:$false
. \\sthdcsrvb174.martinservera.net\script$\_lib\logFunctions.ps1
. \\sthdcsrvb174.martinservera.net\script$\_lib\ad.ps1
Import-Module \\sthdcsrvb174.martinservera.net\script$\_lib\SIDHistory.psm1
Import-Module ActiveDirectory


$scriptFileName = ($MyInvocation.MyCommand.Name).split(".")[0]
$logFilePath = "\\sthdcsrvb174.martinservera.net\script$\_log\"
openLogFile "$logFilePath$(($MyInvocation.MyCommand.name).split('.')[0])-$(get-date -uformat %D)-$env:USERNAME.log"
$User2MDBdataBaseFile = "\\sthdcsrvb174.martinservera.net\script$\migrateExchange\user2mailbox.xml"
$MDBSizeHTdataBaseFile = "\\sthdcsrvb174.martinservera.net\script$\migrateExchange\exchangedbsizes.xml"

####################
# Variables
####################
#$Key = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43,55,234,155,54,8,5,34,7)
$failArray = @()			## Array to hold users not found as migrated in new domain
$users2migrate = @{} 		## HashTable to hold username, oldDomain whose mailboxes should be migrated into new domain
$commandData = @()          ## Array to hold command parameters for user objects to migrate
$UserMBX2DBHash = @{}       ## Hash table to hold user to database mappings
$mArray = @()               ## Array to hold what Migration endpoints are used by the input user accounts
$mdbHT = @{}                ## Hashtable for mailboxdatabases and number of mailboxes
$mdbSizeHT = @{}            ## Hashtable for mailboxdatabases and their sizes
$Rnd = Get-Random -Minimum 10 -Maximum 99 ### Random number 10 to 99
# $BadItemLimit = 1
$NumberofAllowedLargeItems = "Unlimited"
$scpDenyGroupNAme = "secSCPDeny"
$activesyncGroupName = "SEC-EAS-Users"
$OWAGroupName = "SEC-OWA-Users"
$mbx2dbUpdated = $false
$mbxHTupdated = $false
$mdbSizeHTUpdated = $false
$proxyAddressIsUpdated = $false 
$PSSessionOption.MaximumConnectionRedirectionCount = 20   ## To allow more than standard (5) Exchange 2013 PowerShell redirections
$todayTime = Get-date -uformat %y%m%d%H%M%S

####################
# Start function definitions
####################

#####################################################################
# Function initiateExchangePS by Kontract (c)
#  (Hans.Hard@kontract.se)
#
# Starts all Exchange Powershell sessions
#
#####################################################################
Function initiateExchangePS
{
If ((Get-PSSession | Where-Object {$_.ComputerName -like "sthdcsrvb**"}).count -lt 1)
 {
  Write-Host -ForegroundColor Green "... Please wait... will load Exchange 2016 PowerShell..."
  $session = New-PSSession -ConfigurationName Microsoft.Exchange2016.MS -ConnectionURI https://sthdcsrvb153.martinservera.net/Powershell/ -Authentication NegotiateWithImplicitCredential -AllowRedirection
  Write-Host -ForegroundColor Green "... Loading Exchange 2016 PowerShell for Martin & Servera AB ..."
  Import-PSSession $Session -AllowClobber
 }
}

#####################################################################
# Function loadUsers ($filename) by Kontract (c)
#  (Hans.Hard@kontract.se)
#
# Reads $filename as a CSV-file and
#  executes checkUserMigrationStatus function on them
#
#####################################################################
Function loadUsers ( $fileName )
{
 if ("$fileName" -ne "")
 {
  if (test-path -path $fileName)
  {
 # Import CSV file contents to string array
   $importFileContents = Get-Content $fileName  
 # Split them up into an array
   $inFile = $importFileContents -split ',' -replace '^\s*|\s*$'
   $failArray.Clear()
   foreach ($row in $inFile)
   {
    checkUserMigrationStatus($row)
   }
  }
 }
}

#####################################################################
# Function checkUserMigrationStatus ($strUser) by Kontract (c)
#  (Hans.Hard@kontract.se)
#
# Accepts a user name string. Checks in the local domain if it is migrated
#  Returns a hashtable containing sidHistory, samAccountName
#####################################################################
function checkUserMigrationStatus ( $strUser )
{
 Try {$uldp = Get-ADUser -Identity $strUser -properties mail,proxyAddresses -Server sthdcsrvb170.martinservera.net -ErrorAction Stop ; $objFound = $true } Catch {$uldp = "" ; $objFound = $false}
 If ($objFound)
 {
  $uname = $uldp.samaccountname
  Try {$UserMailbox = Get-Mailbox -identity $uName} Catch {$userMailbox = $null}
 }
 If ($UserMailbox)
 {
  LogInfoLine " Processing: $uname"
  $domainTemp = $($uldp.UserPrincipalName).split("@")[1]
  $domainPart = $($domainTemp).split(".")[0]
  try { $script:users2migrate.add($uname, $domainPart) } catch { }
 }
 else { LogWarningLine " $strUser does not have a mailbox" ; LogErrorLine $error[0] ; $script:failArray += $strUser}
}

#####################################################################
# Function getOldMailboxSize ($valPair) by Kontract (c)
#  (Hans.Hard@kontract.se)
#
### $($valPair).value is "DOMAIN" $($valPair).Key is "USERNAME"
####
# Accepts a HashTable value pair with user name as key and domain as value
# and retreives the users total mailbox statistic.
#
# Returns an object with .TotalItemSizeInBytes and TotalItemSizeGB as result
#
#####################################################################
Function getOldMailboxSize ( $valPair )
{
 
 Switch ( $($valPair).value.tolower() )
 {
  "martinservera"
  {
   LogLine " Trying to get mailbox statistics for: $($valPair.Key)"
   $UserMailboxStats = Get-Mailbox -identity $($valPair).Key | Get-MailboxStatistics
  }
 }
 $UserMailboxStats | Add-Member -MemberType ScriptProperty -Name TotalItemSizeInBytes -Value {$this.TotalItemSize -replace "(.*\()|,| [a-z]*\)", ""}
 $returnValue = $UserMailboxStats | Select-Object DisplayName, TotalItemSizeInBytes,@{Name="TotalItemSizeGB"; Expression={[math]::Round( (($_.TotalItemSizeInBytes/1GB)),2)}}
 $returnValue.TotalItemSizeGB
}

###########################################################################
# Function findBestFitTargetMDB ($sizeOption, $mdbArray) by Kontract (c)
#  (Hans.Hard@kontract.se)
#
# Inputs are chosen size option and an array of databases for that size
#  returns a text string IDentity for the chosen database.
###########################################################################
Function findBestFitTargetMDB ( $userName )
{
 ### First check if we have already processed this user...
 If (!( $Script:UserMBX2DBHash.Get_item($userName) -ge 0) )
 { 
  $temp = $Script:mdbsizeHT.Item(0).Value
  $newDBsize = $temp + $Script:cSize
  $indexName = "" ; $indexName = $Script:mdbsizeHT.Item(0).Key
  $Script:mdbSizeHT.Item(0).Value = $newDBsize  ### Update the hashtable with mailbox size

  ### Store the new user name to database mapping on the XML HashTable
  $dbName="" ; $dbName = $script:mdbsizeHT.Get_Item(0).Name
  $Script:UserMBX2DBHash.Add($userName, $dbName )
  $Script:mbx2dbUpdated = $true
  ### Return the database with lowest usage 
  $dbName  ### Returns the database chosen
 }
 Else
 {
  # Return the previously chosen DB name to the calling function
  # (The assignment already exists in the DB to MBX mapping file)
  $Script:UserMBX2DBHash.Get_item($userName)
 }
}

###########################################################################
# Function getMailboxCounts by Kontract (c)
#  (Hans.Hard@kontract.se)
#
# Enumerates all databases in the Exchange environmment and stores current
#  number of mailboxes in $mdbHT ( $mbxID, $mbxCount ) end returns it
###########################################################################
Function getMailboxCounts
{
 ### Get all MBX DB
 $mdbArray = Get-MailboxDatabase
 ForEach ($mdb in $mdbArray)
 {
  $mbxID = $mdb.AdminDisplayName
  $DBExists = $True
  Try { $numberOfMBXinMDB = (get-mailbox -database $mbxID -resultsize Unlimited -ErrorAction SilentlyContinue | Measure-Object).Count } Catch {$DBExists = $False  ; LogErrorLine $Error[0] }
  If ((!($numberOfMBXinMDB)) -and ($DBExists)) {$numberOfMBXinMDB = 0}
  ### If we have more mbx:s on physical DB, store that number. Otherwise choose the HT number
  If ( $numberOfMBXinMDB -gt $mdbHT.get_item( $mbxID) ) {$mdbHT.Set_item($mbxID, $numberOfMBXinMDB)
  }
 }
}

###########################################################################
# Function getMailboxCounts by Kontract (c)
#  (Hans.Hard@kontract.se)
#
# Enumerates all databases in the Exchange environmment and stores current
#  size of mailboxdatabase in $mdbHT ( $mbxID, $mbxCount ) end returns it
###########################################################################
Function getMailboxDatabaseSizes
{
 $a = @{}
 $label = ""
 ### Get all MBX DB and sizes and put them on an array
 $mdb =  @(Get-MailboxDatabase -Status  | select Name,DatabaseSize | Sort-Object -Property Databasesize)
 ### Clean out the DatabaseSize column of the array and put into Size table
 $mdb | Add-Member -MemberType ScriptProperty -Name Size -Value {$this.DatabaseSize -replace "(.*\()|,| [A-Z][a-z]*\)", ""}
 ### Recalculate the Size column into GB and store on a new Array containing @(Name, SizeGB)
 $a = @($mdb | Select-Object Name, @{Name="SizeGB"; Expression={[math]::Round( (($_.Size/1GB)),2)}})
 foreach ($row in $a.GetEnumerator() )
 {
  If ( "$($row.SizeGB)" -gt "$($mdbSizeHT.get_item( $($row.Name)))" ) { $mdbSizeHT.Set_item($row.Name, $row.SizeGB) ;  $mdbSizeHTUpdated = $true  }
 }
 $Script:mdbSizeHT = $Script:mdbSizeHT.GetEnumerator() | Sort-Object -property Value
}


#####################################################################
# Function checkandEditGroupMembership by Kontract (c)
#  (Hans.Hard@kontract.se)
#
#  Checks if user is member of the correct groups, if not adds them
#####################################################################
function checkandEditGroupMembership ( $user )
{
 $isScpDenyGroupMember = $False
 $isActiveSyncMember = $False
 $isOWAMember = $False
 ### Get the user LDAP object
 $userAcct = Get-ADUser -Identity $user
 $userLDAPobject = getLDAPADObject -objPath $($userAcct.DistinguishedName)
 $colUserIsMemberOf = $userLDAPObject.MemberOf
 foreach($group in $colUserIsMemberOf )
 {
  If ($group.tolower().Contains($($scpDenyGroupNAme).ToLower()))     
  {
   $isScpDenyGroupMember = $true
  }
  if ($group.tolower().Contains($($activesyncGroupName).ToLower()))
  {
   $isActiveSyncMember = $true
  }
  if ($group.tolower().Contains($($OWAGroupName).ToLower()))        
  {
   $isOWAMember = $true
  }
  }
 If (!($isScpDenyGroupMember))
 {
  LogLine " Adding $($userLDAPObject.userPrincipalName) to SCP Deny group."
  Try {  If ($Confirm) { Add-ADGroupMember $scpDenyGroupNAme $($userLDAPObject.samAccountName) ; $isScpDenyGroupMember = $true } Else {LogWarningLine " Would have added $($userLDAPObject.samAccountName) to $scpDenyGroupNAme."} ; $isScpDenyGroupMember = $true } Catch { LogErrorLine $Error[0] }
 }
If (!($isActiveSyncMember))
 {
  LogLine " Adding $($userLDAPObject.userPrincipalName) to activesync AD Group."
  Try {  If ($Confirm) {  Add-ADGroupMember $activesyncGroupName $($userLDAPObject.samAccountName) ; $isActiveSyncMember = $true } Else {LogWarningLine " Would have added $($userLDAPObject.samAccountName) to $activesyncGroupName."} ; $isActiveSyncMember = $true } Catch { LogErrorLine $Error[0] }
 }
If (!($isOWAMember))
 {
  LogLine " Adding $($userLDAPObject.userPrincipalName) to Outlook Web Access AD Group."
  Try {  If ($Confirm) {  Add-ADGroupMember $OWAGroupName $($userLDAPObject.samAccountName) ; $isOWAMember = $true } Else {LogWarningLine " Would have added $($userLDAPObject.samAccountName) to $OWAGroupName."} ; $isOWAMember = $true } Catch { LogErrorLine $Error[0] }
 }
 ($isScpDenyGroupMember -and $isOWAMember -and $isActiveSyncMember)
}


#####################################################################
# Function getmigrationendpoint by Kontract (c)
#  (Hans.Hard@kontract.se)
#
# Returns $true is address already exists or is updated without errors
#  otherwise returns $false
#####################################################################
function getMigrationEndPoint ( [string] $strValue )
{
 ### Switch domain to get what Migration endpoint to use
 Switch ( $strValue.tolower() )
   {
    "martinservera"
    {
     $LocalmigrationEndPoint = "Martinservera"
     $RemoteHostname = "xxx.yyy.zzz"
    }
   }
 $LocalmigrationEndPoint
}

#####################################################################
# Function checkMailbox4otherFullAccessPermissions by Kontract (c)
#  (Hans.Hard@kontract.se)
#
# Returns an array of user accounts with full access
#  otherwise returns $null
#####################################################################
function checkMailbox4otherFullAccessPermissions ( [string] $strValue,  [string] $strDomain )
{
 $tmpArr = @()
 Switch ( $strDomain.ToLower() )
 {
  "martinservera"
  {
   $tmpArr = @((Get-Mailbox -identity $strValue | Get-MailboxPermission | where { ($_.AccessRights -eq "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF") }).User)
  }
 }
 If (!($tmpArr.count -gt 0)) { $tmpArr = $null }
 $tmpArr
}

#####################################################################
# Function checkMailbox4SendAsPermissions by Kontract (c)
#  (Hans.Hard@kontract.se)
#
# Returns an array of user accounts with sendAs-permissions
#  otherwise returns $null
#####################################################################
function checkMailbox4SendAsPermissions ( [string] $strValue,  [string] $strDomain )
{
 ### Get-Mailbox -ResultSize unlimited | Get-ADPermission | Where {$_.ExtendedRights -like �Send-As� -and $_.User -notlike �NT AUTHORIT\SELF� -and $_.Deny -eq $false} | ft Identity,User,IsInherited -AutoSize
 $tmpArr = @()
 Switch ( $strDomain.ToLower() )
 {
  "martinservera"
  {
   $tmpArr = @((Get-Mailbox -identity $strValue | Get-ADPermission | where { ($_.ExtendedRights -like "SendAs") -and -not ($_.User -like "NT AUTHORITY\SELF")  -and $_.Deny -eq $false }).User)
  }
 }
 If (!($tmpArr.count -gt 0)) { $tmpArr = $null }
 $tmpArr
}


####################
# End functions
####################


####################
# Main program
####################
StartStopInfo -sAction "start"

######## Initiate all Powershells needed!
initiateExchangePS

######### Execute Pre checks
If ($inputFile)
{
 loadusers $inputFile
}
Else
{
 If ($input.count -gt 1)
 {
  $users.Clear()
  $input | foreach { $users += $_ }
 }
 foreach ($row in $users)
 {
  If ("$row" -ne "" ) { checkUserMigrationStatus($row) }
 }
}
If (-not $failArray)
{
 # Import databases of MBX counts and User to MBX assignments
 #$mdbHT = Import-Clixml -Path $MDBdataBaseFile
 $mdbSizeHT = Import-Clixml -Path $MDBSizeHTdataBaseFile
 $UserMBX2DBHash = Import-Clixml -Path $User2MDBdataBaseFile


 # Find number of mailboxes in all databases
 #getMailboxCounts
 getMailboxDatabaseSizes

 ######## Main program starts here! Iterate all valid users and perform actions on them
 If ($users2migrate)
 {
  foreach ($try in $users2migrate.getenumerator() )
  {
   $mbxSize = getOldMailboxSize ( $try )
   $compareSize = $mbxSize
   $cSize = $compareSize
   LogLine " $($try.Key) has a mailbox size of: $($cSize)GB."

   $MDBSizeOption = "Unlimited"
   $targetMDB = findBestFitTargetMDB -userName $($try.Key)

   $mdbSizeHT = $mdbSizeHT.GetEnumerator() | Sort-Object -property Value  ### Re-Sort the size hash table with new mailbox size added

   
   ### Make sure that the correct target domain exists on the user object
    #$proxyAddressIsUpdated = checkandEditExchangeAttribute -user $($try.Key)

   ### Make sure that the correct group memberships exists on the user object
   $groupMembershipIsUpdated = checkandEditGroupMembership -user $($try.Key)

   LogLine " Choosing $($targetMDB) as database for user $($try.Key)"
   
   ### Switch domain to get what Migration endpoint to use
    #$migEndPoint = getMigrationEndPoint( $($try.Value).tolower() )

   ### Check if we have other users with access to the particular mailbox
   $mbxGrants = @()
   $mbxSendAs = @()
   $mbxGrants = checkMailbox4otherFullAccessPermissions -strValue $($try.Key) -strDomain $($try.Value)
   $mbxSendAs = checkMailbox4SendAsPermissions -strValue $($try.Key) -strDomain $($try.Value)
   If ($mbxGrants)
   {
    LogWarningLine "User has $($mbxGrants.Count) other objects with FullAccess permissions!, they are as follows:"
    If ($mbxGrants.count -eq 1)
    { LogWarningLine " $($mbxGrants)" }
    Else
    { ForEach ( $a in $mbxGrants ) { LogWarningLine " $($a)" } }
   }
   If ($mbxSendAs)
   {
    LogWarningLine " User has $($mbxSendAs.Count) other objects with FullAccess permissions!, they are as follows:"
    If ($mbxSendAs.count -eq 1)
    { LogWarningLine " $($mbxSendAs[0])" }
    Else
    { ForEach ( $a in $mbxSendAs ) { LogWarningLine " $($a)" } }
   }

   ### Instantiate the object that holds command parameters for the migration batch and add to the commandData array
   If ($targetMDB)
   {
    $acct = $($try.Value)+"\"+$($try.key)
    $userLDAPObject = getUserObject ( $acct )
    $commandObject = new-object System.Object
    $commandObject | add-member -MemberType NoteProperty -Name EmailAddress -Value $($userLDAPObject.mail)
    $commandObject | Add-Member -MemberType NoteProperty -Name TargetDatabase -Value $targetMDB
    $commandObject | Add-Member -MemberType NoteProperty -Name samAccountName -Value $($try.key)
    #$commandObject | Add-Member -MemberType NoteProperty -Name migrationEndPoint -Value $migEndPoint
    $commandData += $commandObject
   }
   else
   {
    LogErrorLine "User account: $($try.key) omitted from batch! -->  Target MDB: $($targetMDB)"
   }
  }
 }

  ### All users should now be processed. Build temporary file names
  ### Output the contents of the $CommandData array into a temporary CSV file based on which migration endpoints has been chosen
  Foreach ($row in $commandData )
  {
   $filename = $($env:TEMP)+"\"+$($env:USERNAME)+$($todayTime)+$($Rnd)+$($row.migrationEndPoint)+".csv"
   $row | Export-csv -LiteralPath $filename -Encoding UTF8 -NoTypeInformation -Append
   If (!($($row.migrationEndPoint) -in $mArray)) { $mArray += $($row.migrationEndPoint) }
  }
  Foreach ($migPoint in $mArray )
  {
   $filename = $($env:TEMP)+"\"+$($env:USERNAME)+$($todayTime)+$($Rnd)+$($migPoint)+".csv"
   If ($confirm)
   {
    $operatorAcctEmail = $null
    If ($NotifyME)
    {
     $opAcct = $env:USERDOMAIN+"\"+$env:USERNAME
     $operatorAcctEmail = (getUserObject ( $opAcct )).mail
     Try {
          New-MigrationBatch -AutoStart  -CSVData ([System.IO.File]::ReadAllBytes( $filename )) -local -Name "$($env:USERNAME)$($todayTime)$($Rnd)$($migPoint)" -AutoComplete:$autoComplete -BadItemLimit $BadItemLimit -AllowUnknownColumnsInCsv:$true -NotificationEmails $operatorAcctEmail
          LogLine "New-MigrationBatch -AutoStart -CSVData ([System.IO.File]::ReadAllBytes( $filename )) -local -Name $($env:USERNAME)$($todayTime)$Rnd$($migPoint) -AutoComplete:$AutoComplete -BadItemLimit $BadItemLimit -AllowUnknownColumnsInCsv:$true -NotificationEmails $operatorAcctEmail"
         }
     Catch { $Error[0] }


    }
   Else
   {
    Try {
         New-MigrationBatch -AutoStart -CSVData ([System.IO.File]::ReadAllBytes( $filename )) -local -Name "$($env:USERNAME)$($todayTime)$($Rnd)$($migPoint)" -AutoComplete:$autoComplete -BadItemLimit $BadItemLimit -AllowUnknownColumnsInCsv:$true
         LogLine "New-MigrationBatch -AutoStart  -CSVData ([System.IO.File]::ReadAllBytes( $filename )) -Local -Name $($env:USERNAME)$($todayTime)$Rnd$($migPoint) -AutoComplete:$AutoComplete -BadItemLimit $BadItemLimit -AllowUnknownColumnsInCsv:$true"
        }
    Catch { $Error[0] }
   }
   }
  Else
  {
   $operatorAcctEmail = $null
   notepad $filename
   If ($NotifyME)
   {
    $opAcct = $env:USERDOMAIN+"\"+$env:USERNAME
    $operatorAcctEmail = (getUserObject ( $opAcct )).mail
    LogWarningLine "This is the Migration batch command line:"
    LogLine "New-MigrationBatch -AutoStart -CSVData ([System.IO.File]::ReadAllBytes( $filename )) -Name $($env:USERNAME)$($todayTime)$Rnd$($migPoint) -Local -AutoComplete $AutoComplete -BadItemLimit $BadItemLimit -AllowUnknownColumnsInCsv:$true -NotificationEmails $operatorAcctEmail"
   }
  Else
  {
   LogWarningLine "This is the Migration batch command line:"
   LogLine "New-MigrationBatch -AutoStart  -CSVData ([System.IO.File]::ReadAllBytes( $filename )) -Name $($env:USERNAME)$($todayTime)$Rnd$($migPoint) -Local -AutoComplete $AutoComplete -BadItemLimit $BadItemLimit  -AllowUnknownColumnsInCsv:$true"
  }
  }
 }
 }
else
{
 #### FailArray handling goes here !!!!
 #### Error on input, Print this and give information what accounts are wrong.
 LogErrorLine "Warning!! Error on input!"
 Foreach ($o in $failArray.GetEnumerator()) {LogWarningLine "$($o) is not a migrated object!"}
}


#### Stop logging and script
If (($confirm) -and ($mdbSizeHTUpdated))
{
    {
        $b=@{}
        foreach ($a in $mdbSizeHT.GetEnumerator() )
        { $b.add($a.Name, $a.Value) }
    }
    Export-Clixml -Path $MDBSizeHTdataBaseFile -Encoding UTF8 -inputobject $b
}
If ( ($confirm) -and ($mbx2dbUpdated))
{
    Export-Clixml -Path $User2MDBdataBaseFile -Encoding UTF8 -InputObject $UserMBX2DBHash
}
StartStopInfo -sAction "stop"