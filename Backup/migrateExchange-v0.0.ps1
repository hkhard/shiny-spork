<#
    .SYNOPSIS
    #####################################################################
    # Created by Kontract (c) 2012-2014, v1.16
    #  (Stefan.Alkman@kontract.se)
    #  (Hans.Hard@kontract.se)
    #####################################################################	
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	Version 1.15, 30th November, 2015
	
    .DESCRIPTION
	Migrate Exchange mailboxes from Billerud and Korsnas legacy exchange into
	CORP.LAN
      
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
### 0.1 -- * Initial development version
### 0.2 -- * Database choice validation done
### 0.3 -- * CSV-file
### 0.4 -- * XML file databases for # of mailboxes and username to mbx-db mapping
### 0.41 - * Adjusted for all Billerud domains and accept pipeline input for users
### 1.0 -- **RELEASE VERSION!!
### 1.02 - * Added handling of proxyAddress updating for transfer domain
### 1.03 - * Minor adjustments, mainly reordering of logic into functions
### 1.05 - * adjusted to manage logfiles for different users
### 1.06 - * First try at using exchange ps commands to add smtp transfer address to accounts ** Dev version
### 1.10 - * Release version of reworked smtp-transfer handling and also group additions to handle webmail and activesync
### 1.11 - * Added fix to disable automatic address policy on mail-users who have this set
### 1.12 - * Added extra logging of email addresses so that we see in the log file if transfer address is correct
### 1.13 - * Minor logging changes
### 1.14 - * to do: Get-KNMailbox | Get-KNMailboxPermission | where { ($_.AccessRights -eq "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF") }
###                 Get-Mailbox -ResultSize unlimited | Get-ADPermission | Where {$_.ExtendedRights -like “Send-As” -and $_.User -notlike “NT AUTHORIT\SELF” -and $_.Deny -eq $false} | ft Identity,User,IsInherited -AutoSize
### 1.15 - * Adjusted (get-mailbox).count to unlimited resultsize
### 1.16 - * Adjusted calculation of mailbox size to add 25% extra storage space for mailbox growth
### ================
### End History info
####################

[CmdletBinding()]
param(
   [Parameter(Mandatory = $true,Position = 0,valueFromPipeline=$true)] [string[]] $users,
   [switch] $confirm = $false,
   [switch] $IncrementalSync = $false,
   [switch] $autoComplete = $false,
   [string] $inputFile,
   [switch] $NotifyME = $false
   )

####################
# Include files
####################
. \\segavsadm01\scriptd$\_lib\logFunctions.ps1
. \\segavsadm01\scriptd$\_lib\ad.ps1
Import-Module ..\_lib\SIDHistory.psm1
Import-Module ActiveDirectory


$scriptFileName = ($MyInvocation.MyCommand.Name).split(".")[0]
$logFilePath = "\\segavsadm01.corp.lan\scriptd$\_log\"
openLogFile "$logFilePath$(($MyInvocation.MyCommand.name).split('.')[0])-$(get-date -uformat %D)-$env:USERNAME.log"
$MDBdataBaseFile = "\\segavsadm01.corp.lan\scriptd$\migrateExchange\exchangedb.xml"
$User2MDBdataBaseFile = "\\segavsadm01.corp.lan\scriptd$\migrateExchange\user2mailbox.xml"

####################
# Variables
####################
$Key = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43,55,234,155,54,8,5,34,7)
$failArray = @()			## Array to hold users not found as migrated in new domain
$users2migrate = @{} 		## HashTable to hold username, oldDomain whose mailboxes should be migrated into new domain
$commandData = @()          ## Array to hold command parameters for user objects to migrate
$UserMBX2DBHash = @{}       ## Hash table to hold user to database mappings
$mArray = @()               ## Array to hold what Migration endpoints are used by the input user accounts
$mdbHT = @{}
$Rnd = Get-Random -Minimum 10 -Maximum 99 ### Random number 10 to 99
$BadItemLimit = 30
$NumberofAllowedLargeItems = "Unlimited"
$webmailGroupName = "sec.webmail.external"
$activesyncGroupName = "sec.activesync.external"
$mbx2dbUpdated = $false
$mbxHTupdated = $false
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
If ((Get-PSSession | Where-Object {$_.ComputerName -like "webmail.corp*"}).count -lt 1)
{
 Write-Host -ForegroundColor Green "... Please wait... will load Exchange 2013 PowerShell..."
 $session = New-PSSession -ConfigurationName Microsoft.Exchange2013.BK -ConnectionURI https://webmail.corp.lan/Powershell/ -Authentication NegotiateWithImplicitCredential -AllowRedirection
 Write-Host -ForegroundColor Green "... Loading Exchange 2013 PowerShell for BillerudKorsnäs AB ..."
 Import-PSSession $Session -AllowClobber
}
If ((Get-PSSession | Where-Object {$_.ComputerName -like "segavexc05*"}).count -lt 1)
{
 $adminAccount = "korsnet\lyncmigration"
 If (Test-Path "\\segavsadm01.corp.lan\ScriptD$\userMigration\korsnet.txt")
 {
  Write-Host -ForegroundColor Green "... Please wait... will load Exchange 2010 PowerShell..."
  $cryptPASS = Get-Content "\\segavsadm01.corp.lan\ScriptD$\userMigration\korsnet.txt" | ConvertTo-SecureString -Key $Key
  $KorsnasCred = New-Object System.Management.Automation.PsCredential($adminAccount, $cryptPASS)
  $Session = New-PSSession -ConfigurationName Microsoft.Exchange2010.Korsnas -ConnectionURI http://segavexc05.korsnas.se/Powershell/ -Authentication Kerberos -Credential $KorsnasCred
  Write-Host -ForegroundColor Green "... Loading Exchange 2010 PowerShell for Korsnäs ..."
  Import-PSSession $Session -Prefix KN -AllowClobber
 }
}
If ((Get-PSSession | Where-Object {$_.ComputerName -like "bilexhc*"}).count -lt 1)
{
 $adminAccount = "billerud\lyncmigrering"
 If (Test-Path "\\segavsadm01.corp.lan\ScriptD$\userMigration\billerud.txt")
 {
  Write-Host -ForegroundColor Green "... Please wait... will load Exchange 2010 PowerShell..."
  $cryptPASS = Get-Content "\\segavsadm01.corp.lan\ScriptD$\userMigration\billerud.txt" | ConvertTo-SecureString -Key $Key
  $BillerudCred = New-Object System.Management.Automation.PsCredential($adminAccount, $cryptPASS)
  $session = New-PSSession -ConfigurationName Microsoft.Exchange2010.Billerud -ConnectionURI http://bilexhc03.billerud.com/Powershell/ -Authentication Kerberos -Credential $BillerudCred
  Write-Host -ForegroundColor Green "... Loading Exchange 2010 PowerShell for Billerud  ..."
  Import-PSSession $Session -Prefix BD -AllowClobber
 }
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
 Try {$uldp = Get-ADUser -Identity $strUser -properties mail,proxyAddresses -Server segavsdc01.corp.lan -ErrorAction Stop ; $objFound = $true } Catch {$uldp = "" ; $objFound = $false}
 If ($objFound)
 {
  $uname = $uldp.samaccountname
  $SID = ""
  $SID = Get-SIDHistory -SamAccountName $uname
  if ($SID)
  {
   LogInfoLine " $uname is a migrated user with a sidHistory of: $($SID.SID)"
   $oldName = getOldUserNameFromSid ( $($SID).SID )
   $domainPart = $($oldName).split("\")[0]
  }
  Switch ( $domainPart.ToLower() )
  {
   "korsnet"
   {
    Try {$UserMailbox = Get-KNMailbox -identity $uName} Catch {$userMailbox = $null}
   }
   "billerud"
   {
    Try {$UserMailbox = Get-BDMailbox -identity $uName -DomainController sedc05.billerud.com } Catch {$UserMailbox = $null}
   }
   "office"
   {
    Try {$UserMailbox = Get-BDMailbox -identity $uName -DomainController seoffdc01.office.billerud.com} Catch {$UserMailbox = $null}
   }
   "skbl"
   {
    Try {$UserMailbox = Get-BDMailbox -identity $uName -DomainController seskbldc03.skbl.billerud.com} Catch {$UserMailbox = $null}
   }
   "grb"
   {
    Try {$UserMailbox = Get-BDMailbox -identity $uName -DomainController segrbdc01.grb.billerud.com} Catch {$UserMailbox = $null}
   }
   "bthm"
   {
    Try {$UserMailbox = Get-BDMailbox -identity $uName -DomainController ukbthmdc04.bthm.billerud.com} Catch {$UserMailbox = $null}
   }
   "kbg"
   {
    Try {$UserMailbox = Get-BDMailbox -identity $uName -DomainController sekbgdc03.kbg.billerud.com} Catch {$UserMailbox = $null}
   }
  }
  
  If ($UserMailbox)
  {
   try { $script:users2migrate.add($uname, $domainPart) } catch { }
  }
  else
  {
   If (!($proxyAddressIsUpdated)) { LogWarningLine " $($strUser) is not updated with the proper target proxyAddress!" }
   $script:failArray += $strUser
  }
  }
 else { LogWarningLine " $strUser is not a migrated user (lacks sidHistory)" ; LogErrorLine $error[0] ; $script:failArray += $strUser}
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
  "korsnet"
  {
   LogLine " Trying to get mailbox statistics for: $($valPair.Key)"
   $UserMailboxStats = Get-KNMailbox -identity $($valPair).Key | Get-KNMailboxStatistics
  }
  "billerud"
  {
   $UserMailboxStats = Get-BDMailbox -identity $($valPair).Key -DomainController sedc05.billerud.com | Get-BDMailboxStatistics -DomainController sedc05.billerud.com
  }
  "office"
  {
   $UserMailboxStats = Get-BDMailbox -identity $($valPair).Key -DomainController seoffdc01.office.billerud.com | Get-BDMailboxStatistics -DomainController seoffdc01.office.billerud.com
  }
  "skbl"
  {
   $UserMailboxStats = Get-BDMailbox -identity $($valPair).Key -DomainController seskbldc03.skbl.billerud.com | Get-BDMailboxStatistics -DomainController seskbldc03.skbl.billerud.com
  }
  "grb"
  {
   $UserMailboxStats = Get-BDMailbox -identity $($valPair).Key -DomainController segrbdc01.grb.billerud.com | Get-BDMailboxStatistics -DomainController segrbdc01.grb.billerud.com
  }
  "bthm"
  {
   $UserMailboxStats = Get-BDMailbox -identity $($valPair).Key -DomainController ukbthmdc04.bthm.billerud.com | Get-BDMailboxStatistics -DomainController ukbthmdc04.bthm.billerud.com
  }
  "kbg"
  {
   $UserMailboxStats = Get-BDMailbox -identity $($valPair).Key -DomainController sekbgdc03.kbg.billerud.com | Get-BDMailboxStatistics -DomainController sekbgdc03.kbg.billerud.com
  }
 }
 $UserMailboxStats | Add-Member -MemberType ScriptProperty -Name TotalItemSizeInBytes -Value {$this.TotalItemSize -replace "(.*\()|,| [a-z]*\)", ""}
 $UserMailboxStats | Select-Object DisplayName, TotalItemSizeInBytes,@{Name="TotalItemSizeGB"; Expression={[math]::Round( (($_.TotalItemSizeInBytes/1GB) * 1.25),2)}}
 $returnValue = $UserMailboxStats
 $returnValue
}

###########################################################################
# Function findBestFitTargetMDB ($sizeOption, $mdbArray) by Kontract (c)
#  (Hans.Hard@kontract.se)
#
# Inputs are chosen size option and an array of databases for that size
#  returns a text string IDentity for the chosen database.
###########################################################################
Function findBestFitTargetMDB ( $sizeOption, $mdbArray, $userName )
{
 $tempHT = @{}
 ### First check if we have already processed this user...
 If (!( $Script:UserMBX2DBHash.Get_item($userName) -ge 0) )
 {
  ### If user is not procesed, we find the appropriate MBX db
  ForEach ($mdb in $mdbArray)
  {
   $mbxID = $mdb + $sizeOption
   $numberOfMBXinMDB = $Script:mdbHT.get_item($mbxID)
   $tempHT.Add($mbxID, $numberOfMBXinMDB )
  }
  ### Sort all databases on number of mailboxes, we work on [0] which is the DB with least number of mailboxes
  $sortedMDB = $TempHT.GetEnumerator() | sort Value

  ### Store the new user name to database mapping on the XML HashTable
  $Script:UserMBX2DBHash.Add($userName, $sortedMDB[0].Key )
  $Script:mbx2dbUpdated = $true
  ### Store the new number of database mailboxes on the XML HashTable
  $Script:mdbHT.Set_Item( $sortedMDB[0].Key , ($mdbHT.get_item($sortedMDB[0].Key))+1)
  $Script:mbxHTupdated = $true
  ### Return the database name (with the least number of mailboxes) to the calling function
  $sortedMDB[0].Key
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
 ### Get all MBX DB in corp.lan
 $mdbArray = Get-MailboxDatabase
 ForEach ($mdb in $mdbArray)
 {
  $mbxID = $mdb.AdminDisplayName
  $DBExists = $True
  Try { $numberOfMBXinMDB = (get-mailbox -database $mbxID -resultsize Unlimited -ErrorAction SilentlyContinue).Count } Catch {$DBExists = $False  ; LogErrorLine $Error[0] }
  If ((!($numberOfMBXinMDB)) -and ($DBExists)) {$numberOfMBXinMDB = 0}
  ### If we have more mbx:s on physical DB, store that number. Otherwise choose the HT number
  If ( $numberOfMBXinMDB -gt $mdbHT.get_item( $mbxID) ) {$mdbHT.Set_item($mbxID, $numberOfMBXinMDB)
  }
 }
}

#####################################################################
# Function checkandEditExchangeAttribute by Kontract (c)
#  (Hans.Hard@kontract.se)
#
# Returns $true is address already exists or is updated without errors
#  otherwise returns $false
#####################################################################
function checkandEditExchangeAttribute ( $user )
{
 $relaySMTPAlreadyAdded = $False
 $relaySMTPisOK = $false
 ### Get the user SIDHistory
 $userAcct = Get-ADUser -Identity $user -Properties mail,proxyAddresses
 $userSIDHistory = get-sidhistory -SamAccountName $($userAcct.samAccountName)
 $strTemp = $($userAcct.mail).split("@")
 $relaySMTP = "smtp:"+$strTemp[0]+"@"+"corp."+$strTemp[1]
 If ($userSIDHistory)
 {
  foreach($proxyAddress in $($userAcct.proxyAddresses) )
  {
   if ($proxyAddress.tolower() -eq $relaySMTP) { $relaySMTPAlreadyAdded = $true ; $relaySMTPisOK = $true ; LogLine " $($relaySMTP) already added to proxyAddresses." ; ForEach ($a in $($useracct.emailaddresses)) { LogLine " $a" } }
  }
  if (-not $relaySMTPAlreadyAdded)
  {
   Try { If ($Confirm)
   {
    $Email = Get-MailUser -id $user
    $AddressPolicy = $Email.EmailAddressPolicyEnabled
    $Email.EmailAddresses += ($($relaySMTP))
    If ($AddressPolicy) {$AddressPolicy = $False}  ### Reverts address policy to false = never update automatically ###
    Set-MailUser $user -EmailAddressPolicyEnabled $AddressPolicy -DomainController segavsdc01.corp.lan
    Set-MailUser $user -EmailAddresses $Email.EmailAddresses -DomainController segavsdc01.corp.lan
    LogLine " Added $($relaySMTP) to Mailbox for $($user), configured proxy addresses are:"
    $Email = Get-MailUser -id $user -DomainController segavsdc01.corp.lan
    ForEach ($a in $($email.emailaddresses)) { LogLine " $a" }
    $relaySMTPisOK = $true
   }
   Else
   {
    LogWarningLine " Confirm not specified, will not add $($relaySMTP) to user!"
    $relaySMTPisOK = $true
   }
  }
  Catch { LogWarningLine " Warning! could set mailbox relay address!!!" ; LogErrorLine $Error[0] ;  $relaySMTPisOK = $false }
  }
 }
 $relaySMTPisOK
}

#####################################################################
# Function checkandEditGroupMembership by Kontract (c)
#  (Hans.Hard@kontract.se)
#
#  Checks if user is member of the correct groups, if not adds them
#####################################################################
function checkandEditGroupMembership ( $user )
{
 $isWebmailMember = $False
 $isActiveSyncMember = $False
 ### Get the user LDAP object
 $userAcct = Get-ADUser -Identity $user
 $userLDAPobject = getLDAPADObject -objPath $($userAcct.DistinguishedName)
 $colUserIsMemberOf = $userLDAPObject.MemberOf
 foreach($group in $colUserIsMemberOf )
 {
  If ($group.tolower().Contains($($webmailGroupName)))          ## Checks if user is a member of the webmail group
  {
   $isWebmailMember = $true
  }
  If ($group.tolower().Contains($($activesyncGroupName)))          ## Checks if user is a member of the activesync group
  {
   $isActiveSyncMember = $true
  }
 }
 If (!($isWebmailMember))
 {
  LogLine " Adding $($userLDAPObject.userPrincipalName) to webmail AD Group."
  Try {  If ($Confirm) { Add-ADGroupMember $webmailGroupName $($userLDAPObject.samAccountName) ; $isWebmailMember = $true } Else {LogWarningLine " Would have added $($userLDAPObject.samAccountName) to $webmailGroupName."} ; $isWebmailMember = $true } Catch { LogErrorLine $Error[0] }
 }
If (!($isActiveSyncMember))
 {
  LogLine " Adding $($userLDAPObject.userPrincipalName) to activesync AD Group."
  Try {  If ($Confirm) {  Add-ADGroupMember $activesyncGroupName $($userLDAPObject.samAccountName) ; $isActiveSyncMember = $true } Else {LogWarningLine " Would have added $($userLDAPObject.samAccountName) to $activesyncGroupName."} ; $isActiveSyncMember = $true } Catch { LogErrorLine $Error[0] }
 }
 ($isWebmailMember -and $isActiveSyncMember)
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
    "korsnet"
    {
     $LocalmigrationEndPoint = "Korsnas"
     $LocalremoteHostname = "segavexc05.korsnas.se"
    }
    "billerud"
    {
     $LocalmigrationEndPoint = "Billerud"
     $LocalremoteHostname = "bilexhc03.billerud.com"
    }
    "office"
    {
     $LocalmigrationEndPoint = "Billerud"
     $LocalremoteHostname = "bilexhc03.billerud.com"
    }
    "skbl"
    {
     $LocalmigrationEndPoint = "Billerud"
     $LocalremoteHostname = "bilexhc03.billerud.com"
    }
    "grb"
    {
     $LocalmigrationEndPoint = "Billerud"
     $LocalremoteHostname = "bilexhc03.billerud.com"
    }
    "bthm"
    {
     $LocalmigrationEndPoint = "Billerud"
     $LocalremoteHostname = "bilexhc03.billerud.com"
    }
    "kbg"
    {
     $LocalmigrationEndPoint = "Billerud"
     $LocalremoteHostname = "bilexhc03.billerud.com"
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
  "korsnet"
  {
   $tmpArr = @((Get-KNMailbox -identity $strValue | Get-KNMailboxPermission | where { ($_.AccessRights -eq "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF") }).User)
  }
  "billerud"
  {
   $tmpArr = @((Get-BDMailbox -identity $strValue -DomainController sedc05.billerud.com | Get-BDMailboxPermission -DomainController sedc05.billerud.com | where { ($_.AccessRights -eq "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF") }).User)
  }
  "office"
  {
   $tmpArr = @((Get-BDMailbox -identity $strValue -DomainController seoffdc01.office.billerud.com  | Get-BDMailboxPermission -DomainController seoffdc01.office.billerud.com  | where { ($_.AccessRights -eq "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF") }).User)
  }
  "skbl"
  {
   $tmpArr = @((Get-BDMailbox -identity $strValue -DomainController seskbldc03.skbl.billerud.com | Get-BDMailboxPermission -DomainController seskbldc03.skbl.billerud.com | where { ($_.AccessRights -eq "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF") }).User)
  }
  "grb"
  {
   $tmpArr = @((Get-BDMailbox -identity $strValue -DomainController segrbdc01.grb.billerud.com | Get-BDMailboxPermission -DomainController segrbdc01.grb.billerud.com | where { ($_.AccessRights -eq "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF") }).User)
  }
  "bthm"
  {
   $tmpArr = @((Get-BDMailbox -identity $strValue -DomainController ukbthmdc04.bthm.billerud.com | Get-BDMailboxPermission -DomainController ukbthmdc04.bthm.billerud.com | where { ($_.AccessRights -eq "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF") }).User)
  }
  "kbg"
  {
   $tmpArr = @((Get-BDMailbox -identity $strValue -DomainController sekbgdc03.kbg.billerud.com | Get-BDMailboxPermission -DomainController sekbgdc03.kbg.billerud.com | where { ($_.AccessRights -eq "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF") }).User)
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
 ### Get-Mailbox -ResultSize unlimited | Get-ADPermission | Where {$_.ExtendedRights -like “Send-As” -and $_.User -notlike “NT AUTHORIT\SELF” -and $_.Deny -eq $false} | ft Identity,User,IsInherited -AutoSize
 $tmpArr = @()
 Switch ( $strDomain.ToLower() )
 {
  "korsnet"
  {
   $tmpArr = @((Get-KNMailbox -identity $strValue | Get-KNADPermission | where { ($_.ExtendedRights -like "SendAs") -and -not ($_.User -like "NT AUTHORITY\SELF")  -and $_.Deny -eq $false }).User)
  }
  "billerud"
  {
   $tmpArr = @((Get-BDMailbox -identity $strValue -DomainController sedc05.billerud.com | Get-BDADPermission -DomainController sedc05.billerud.com | where { ($_.ExtendedRights -like "SendAs") -and -not ($_.User -like "NT AUTHORITY\SELF")  -and $_.Deny -eq $false }).User)
  }
  "office"
  {
   $tmpArr = @((Get-BDMailbox -identity $strValue -DomainController seoffdc01.office.billerud.com  | Get-BDADPermission -DomainController seoffdc01.office.billerud.com  | where { ($_.ExtendedRights -like "SendAs") -and -not ($_.User -like "NT AUTHORITY\SELF")  -and $_.Deny -eq $false }).User)
  }
  "skbl"
  {
   $tmpArr = @((Get-BDMailbox -identity $strValue -DomainController seskbldc03.skbl.billerud.com | Get-BDADPermission -DomainController seskbldc03.skbl.billerud.com | where { ($_.ExtendedRights -like "SendAs") -and -not ($_.User -like "NT AUTHORITY\SELF")  -and $_.Deny -eq $false }).User)
  }
  "grb"
  {
   $tmpArr = @((Get-BDMailbox -identity $strValue -DomainController segrbdc01.grb.billerud.com | Get-BDADPermission -DomainController segrbdc01.grb.billerud.com | where { ($_.ExtendedRights -like "SendAs") -and -not ($_.User -like "NT AUTHORITY\SELF")  -and $_.Deny -eq $false }).User)
  }
  "bthm"
  {
   $tmpArr = @((Get-BDMailbox -identity $strValue -DomainController ukbthmdc04.bthm.billerud.com | Get-BDADPermission -DomainController ukbthmdc04.bthm.billerud.com | where { ($_.ExtendedRights -like "SendAs") -and -not ($_.User -like "NT AUTHORITY\SELF")  -and $_.Deny -eq $false }).User)
  }
  "kbg"
  {
   $tmpArr = @((Get-BDMailbox -identity $strValue -DomainController sekbgdc03.kbg.billerud.com | Get-BDADPermission -DomainController sekbgdc03.kbg.billerud.com | where { ($_.ExtendedRights -like "SendAs") -and -not ($_.User -like "NT AUTHORITY\SELF")  -and $_.Deny -eq $false }).User)
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
 $mdbHT = Import-Clixml -Path $MDBdataBaseFile
 $UserMBX2DBHash = Import-Clixml -Path $User2MDBdataBaseFile

 # Find number of mailboxes in all databases
 getMailboxCounts

 ######## Main program starts here! Iterate all valid users and perform actions on them
 If ($users2migrate)
 {
  foreach ($try in $users2migrate.getenumerator() )
  {
   $mbxSize = getOldMailboxSize ( $try )
   $compareSize = $($mbxSize).TotalItemSizeGB
   $cSize = $compareSize * 0.8
   LogLine " $($try.Key) has a mailbox size of: $($cSize)GB."
   If  ($compareSize -lt 1) {$MDBSizeOption = "1GB" ; $DBarray = @("MDB01-", "MDB02-", "MDB11-", "MDB12-") }
   If (($compareSize -ge 1) -and ($compareSize -lt 2) ) {$MDBSizeOption = "1-2GB"; $DBarray = @("MDB03-", "MDB13-") }
   If (($compareSize -ge 2) -and ($compareSize -lt 4) ) {$MDBSizeOption = "2-4GB"; $DBarray = @("MDB04-", "MDB14-") }
   If (($compareSize -ge 4) -and ($compareSize -lt 8) ) {$MDBSizeOption = "4-8GB"; $DBarray = @("MDB05-", "MDB15-") }
   If (($compareSize -ge 8) -and ($compareSize -lt 12) ) {$MDBSizeOption = "8-12GB"; $DBarray = @("MDB06-", "MDB16-") }
   If  ($compareSize -ge 12) {$MDBSizeOption = "12GB"; $DBarray = "MDB07-", "MDB17-" }
   
   ### Make sure that the correct target domain exists on the user object
   $proxyAddressIsUpdated = checkandEditExchangeAttribute -user $($try.Key)
   
   ### Make sure that the correct group memberships exists on the user object
   $groupMembershipIsUpdated = checkandEditGroupMembership -user $($try.Key)

   ### Based on mailbox size, find the best target mailbox for that particular user
   $targetMDB = findBestFitTargetMDB -sizeOption $MDBSizeOption -mdbArray $DBArray -userName $($try.Key)
   LogLine " Choosing $($targetMDB) as database for user $($try.Key)"
   
   ### Switch domain to get what Migration endpoint to use
   $migEndPoint = getMigrationEndPoint( $($try.Value).tolower() )

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
   If (($proxyAddressIsUpdated) -and ($targetMDB) -and ($groupMembershipIsUpdated))
   {
    $acct = $($try.Value)+"\"+$($try.key)
    $userLDAPObject = getUserObject ( $acct )
    $commandObject = new-object System.Object
    $commandObject | add-member -MemberType NoteProperty -Name EmailAddress -Value $($userLDAPObject.mail)
    $commandObject | Add-Member -MemberType NoteProperty -Name TargetDatabase -Value $targetMDB
    $commandObject | Add-Member -MemberType NoteProperty -Name samAccountName -Value $($try.key)
    $commandObject | Add-Member -MemberType NoteProperty -Name migrationEndPoint -Value $migEndPoint
    $commandData += $commandObject
   }
   else
   {
    LogErrorLine "User account: $($try.key) omitted from batch! --> ProxyAddress updated: $($proxyAddressIsUpdated) -- Target MDB: $($targetMDB) -- GroupMemberships: $($groupMembershipIsUpdated)"
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
     New-MigrationBatch -AutoStart -TargetDeliveryDomain corp.billerudkorsnas.com -CSVData ([System.IO.File]::ReadAllBytes( $filename )) -Name "$($env:USERNAME)$($todayTime)$($Rnd)$($migPoint)" -SourceEndPoint $migPoint -AllowIncrementalSyncs:$IncrementalSync -BadItemLimit $BadItemLimit -LargeItemLimit $NumberofAllowedLargeItems -AutoComplete:$autoComplete -AllowUnknownColumnsInCsv:$true -NotificationEmails $operatorAcctEmail
    }
   Else
   {
    New-MigrationBatch -AutoStart -TargetDeliveryDomain corp.billerudkorsnas.com -CSVData ([System.IO.File]::ReadAllBytes( $filename )) -Name "$($env:USERNAME)$($todayTime)$($Rnd)$($migPoint)" -SourceEndPoint $migPoint -AllowIncrementalSyncs:$IncrementalSync -BadItemLimit $BadItemLimit -LargeItemLimit $NumberofAllowedLargeItems -AutoComplete:$autoComplete -AllowUnknownColumnsInCsv:$true
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
    LogLine "New-MigrationBatch -AutoStart -TargetDeliveryDomain corp.billerudkorsnas.com -CSVData ([System.IO.File]::ReadAllBytes( $filename )) -Name $($env:USERNAME)$($todayTime)$Rnd$($migPoint) -SourceEndPoint $migPoint -AllowIncrementalSyncs:$IncrementalSync -BadItemLimit $BadItemLimit -LargeItemLimit $NumberofAllowedLargeItems -AutoComplete:$autoComplete -AllowUnknownColumnsInCsv:$true -NotificationEmails $operatorAcctEmail"
   }
  Else
  {
   LogWarningLine "This is the Migration batch command line:"
   LogLine "New-MigrationBatch -AutoStart -TargetDeliveryDomain corp.billerudkorsnas.com -CSVData ([System.IO.File]::ReadAllBytes( $filename )) -Name $($env:USERNAME)$($todayTime)$Rnd$($migPoint) -SourceEndPoint $migPoint -AllowIncrementalSyncs:$IncrementalSync -BadItemLimit $BadItemLimit -LargeItemLimit $NumberofAllowedLargeItems -AllowUnknownColumnsInCsv:$true -AutoComplete:$autoComplete"
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
If ($mbxHTupdated) { Export-Clixml -Path $MDBdataBaseFile -Encoding UTF8 -inputobject  $mdbHT}
If ($mbx2dbUpdated) { Export-Clixml -Path $User2MDBdataBaseFile -Encoding UTF8 -InputObject $UserMBX2DBHash }
#Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange*"} | Remove-PSSession
StartStopInfo -sAction "stop"