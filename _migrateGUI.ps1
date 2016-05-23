<#
    .SYNOPSIS
    #####################################################################
    # Created by Kontract (c) 2012-2015, v1.10
    #  (Stefan.Alkman@kontract.se)
    #  (Hans.Hard@kontract.se)
    #
    #####################################################################	
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	Version 1.10, 12th August, 2015
	
    .DESCRIPTION
    This script show a UI for selection of mailboxes to send into the
    script used for migration of selected mailboxes.

   	.PARAMETER FormerForest
	Parameter to specify which forest to look for migratable mailboxes
        Defaults to 'korsnas.se'

    .PARAMETER FullSync
    Creates internal databases of objects and their USN's.
        This is used for determining which mailboxes shall be displayed
        in the selection list.

    .PARAMETER updateTables
    Updates the internal database of objects and USN's for the specified forest.

    #>

### Version History
### ===============
### 0.1 -- * Initial version
### 1.0 -- * Release version
### 1.1 -- * Added Send confirm checkbox to confirmation dialogue. Added command to actually send selected objects into the migration routine.
### 1.11 - * Cosmetic changes and adaption of command line to migrateExchange-script
### ================
### End History info

[CmdletBinding()]
param(
    [ValidateSet("martinservera.net")]
    [ValidateNotNullOrEmpty()]
    [string] $formerForest="martinservera.se",
    [switch] $fullSync,
    [switch] $updateTables
    )

#List the attributes to syncronize
# usesXML from syncUserAttributes
# \\sthdcsrvb174\script$\_lib\$scriptFileName-userDataTable.xml 
# \\sthdcsrvb174\script$\_lib\$scriptFileName-domainControllerTable.xml
# this ensures syncronisation of all attributes

$attributesToSync = @("mail")
$attributesToSyncSpecial = @("msExchHomeServerName", "mailNickname", "targetAddress", "displayName", "l", "physicalDeliveryOfficeName")

$sendMailLogs = $false
# $sendMailLogsRecipients = @("Stefan.Alkman@billerudkorsnas.com")

####################
# Include files
####################
Unblock-File \\sthdcsrvb174.martinservera.net\script$\_lib\logFunctions.ps1 -Confirm:$false
Unblock-File \\sthdcsrvb174.martinservera.net\script$\_lib\migratedDB.ps1 -Confirm:$false
Unblock-File \\sthdcsrvb174.martinservera.net\script$\_lib\ad.ps1 -Confirm:$false
. \\sthdcsrvb174.martinservera.net\script$\_lib\logFunctions.ps1
. \\sthdcsrvb174.martinservera.net\script$\_lib\migratedDB.ps1
. \\sthdcsrvb174.martinservera.net\script$\_lib\ad.ps1

$realyUpdate = $true
$script:processErrorCount = 0

$userValuesChanged = 0
$verbose = 0
$LastColumnClicked = 0 # tracks the last column number that was clicked
$LastColumnAscending = $false # tracks the direction of the last sort of this column


##############################################################
# Setup varius dataTables

function CreateDomainControllerTable
{
 $dataTable = New-Object System.Data.DataTable("domainControllerTable")

 $col = $dataTable.Columns.Add("domainPart")
 $col.AllowDBNull = $false
 $col.Unique = $true
 $col = $dataTable.Columns.Add("NETBIOSname")
 $col.AllowDBNull = $false
 $col.Unique = $true
 $dataTable.PrimaryKey = $col
 [void] $dataTable.Columns.Add("domainController")
 [void] $dataTable.Columns.Add("highestCommittedUSN")
 [void] $dataTable.Columns.Add("updatedHighestCommittedUSN")
 [void] $dataTable.Columns.Add("checkDomainForChanges")
 $dataTable.DefaultView.Sort = "NETBIOSname"
 return @(,$dataTable) 
}

function CreateUserTable
{
 $dataTable = New-Object System.Data.DataTable("userTable")

 $col = $dataTable.Columns.Add("adspath")
 $col.AllowDBNull = $false
 $col.Unique = $true
 $col = $dataTable.Columns.Add("distinguishedname")
 $col.AllowDBNull = $false
 $col.Unique = $true
 $col = $dataTable.Columns.Add("objectGUID",[System.GUID])
 $col.AllowDBNull = $false
 $col.Unique = $true
 $dataTable.PrimaryKey = $col
 [void] $dataTable.Columns.Add("DOMAIN")
 [void] $dataTable.Columns.Add("uSNChanged")
 [void] $dataTable.Columns.Add("sidHistory")

 foreach($field in $attributesToSync)
 { [void] $dataTable.Columns.Add($field) }

 foreach($field in $attributesToSyncSpecial)
 {
  if ($field -like "*guid*")
  {
   [void] $dataTable.Columns.Add($field,[System.GUID]) 
  }
  else
  {
   [void] $dataTable.Columns.Add($field) 
  }
 }
 [void] $dataTable.Columns.Add("sAMAccountName")
 $col = $dataTable.Columns.Add("objectSID")
 $col.AllowDBNull = $false
 $col.Unique = $true

 [void] $dataTable.Columns.Add("userAccountControl", [uint64])

 $dataTable.DefaultView.Sort = "objectGUID"
 return @(,$dataTable) 
}



##################################################################################################### OK
# CreateInitialDomainControllerTable(domainTable)
#  Create a initial table with domaincontrollers, domain names a.s.o. are only used initialty and
#  are replaced by XML file after initial run.

function CreateInitialDomainControllerTable($dataTable)
{
 [void] $dataTable.Rows.Add("DC=martinservera,DC=net","MARTINSERVERA","sthdcsrvb170.martinservera.net", 0, $false)
}

##################################################################################################### OK
# UpdateUSN(domainTable)
#
# Update current calue in $domainTable

function UpdateUSN($domainTable)
{
 foreach ($row in $domainTable.rows)
 {
  $row.HighestCommittedUSN = $row.updatedHighestCommittedUSN
 }
}

##################################################################################################### OK
# UpdateUSNFromDCS(domainTable)
#
# Updates USN (update sequence number) in DomainTable. The value are used to search for changes
# sins last check
#
# If USN for domainconroller have increased checkDomainForChanges column are set to $true and 
# updateHighestCommittedUSN set to current value. (After script are finished UpdateUSN are called
# to update HighestCommittedUSN with the updateHighestCommittedUSN i everything are working OK).

function UpdateUSNFromDCS($domainTable)
{
 foreach ($row in $domainTable.rows)
 {
  $RootDSE = [ADSI]"LDAP://$($row.domaincontroller)/RootDSE"
  
  [long] $highestCommittedUSN = "$($RootDSE.highestCommittedUSN)"
  if ($highestCommittedUSN -eq 0)
  {
   Write-Error "Error connecting to domaincontroller $($row.domaincontroller) - aborting"
   exit -1
  }
  if (($row.highestCommittedUSN -eq $false) -or ($row.highestCommittedUSN -lt $highestCommittedUSN))
  {
   $row.checkDomainForChanges = $true 
   $row.updatedHighestCommittedUSN = $highestCommittedUSN
  }
  else
  {
   $row.checkDomainForChanges = $false
  }
 }
}

function GetPropertiesInStringRepresentation($property)
{
 $values = ""
 foreach($propertyValue in $property)
 {
  switch ($propertyValue.GetType().toString())
  {
   "System.Byte[]" {
     # If length 16 chars - convert to string GUID
     if ($result.properties[$propertyName][0].count -eq 16)
     {
      $value = (new-object -TypeName System.Guid -ArgumentList (,$propertyValue))
     }
     else
     {
      # If of other length - try to convert to string SID
      $value = (New-Object System.Security.Principal.SecurityIdentifier($propertyValue,0)).Value
     }
    }
    default { $value = $propertyValue }
   }
   if ($values -eq "") { $values += $value } else { $values += "�$value" } 
  }
 return $values
}

#####################################################################################################
# getObjectsToProcess(domainPart, domainController, startUSN, userTable, filterBase)
#
# Updates USN (update sequence number) in DomainTable. The value are used to search for changes
# sins last check
#
# If USN for domainconroller have increased checkDomainForChanges column are set to $true and 
# ighestCommittedUSN set to current value

function getObjectsToProcess($domainPart, $domainController, $startUSN = $false, $dataTable, $valuesChanged, $filterBase = "(objectcategory=person)(objectClass=User)")
{
 # Initialize Directory Searcher
 $DSearcher = New-Object System.DirectoryServices.DirectorySearcher (New-Object System.DirectoryServices.DirectoryEntry("LDAP://$domainController/$domainPart"))
 $DSearcher.PageSize = 100
 $DSearcher.SearchScope = "Subtree"
 [void]$DSearcher.PropertiesToLoad.Add("distinguishedName")

 # Add all properties that are included the supplied DataTable to be loaded from domain controller
 foreach($column in $dataTable.Columns)
 { [void]$DSearcher.PropertiesToLoad.Add("$column") }

 # Setup search filter - if no startUSN are specified all objects will be returned.
 if ($startUSN)
 { $DSearcher.Filter = "(&(uSNChanged>=$startUSN)$filterBase)" }
 else
 { $DSearcher.Filter = "(&$filterBase)" }

 # Retrive all results
 $results = $DSearcher.FindAll()

 if ($results.count -ne 0)
 {
  foreach($result in $results)
  {
   $newRow = $false
   $changedValues = $false

   # Identify if the object already extists in userTable
   $objectGUID = (new-object -TypeName System.Guid -ArgumentList (,$result.properties["objectGUID"][0]))
   $row = $dataTable.Rows.Find($objectGUID)

   if ($row -eq $null)
   {
    # ObjectGUID not found - new row created
    $Row = $dataTable.NewRow()
    $newRow = $true
   }

   $Row.Domain = $DomainPart
   
   foreach($propertyName in $result.properties.PropertyNames)
   {
    $propValue = $result.properties[$propertyName]

    $values = GetPropertiesInStringRepresentation -property $propValue

    if ($Row.$propertyName -ne $values)
    {
     try
     {
      $Row.$propertyName = $values
     }
     catch
     {
      LogErrorLine "$propertyName cant be set on $($row.sAMAccountName)"
     }
     $changedValues = $true
    }
   }
   if ($newRow -and $changedValues)
   {
    try
    {
     $dataTable.Rows.Add($row)
    }
    catch
    {
     LogErrorLine "Error adding row $($result.path)"
     $script:processErrorCount += 1
    }
   }
   if ($changedValues) { $valuesChanged +=1 }
  }
 }
 LogLine "  - number of objects identified with changes $domainPart are $($results.count), number changed synked objects $valuesChanged"
}

function GetPropertiesInStringRepresentation($property)
{
 $values = ""
 foreach($propertyValue in $property)
 {
  switch ($propertyValue.GetType().toString())
  {
   "System.Byte[]" {
     # If length 16 chars - convert to string GUID
     if ($result.properties[$propertyName][0].count -eq 16)
     {
      $value = (new-object -TypeName System.Guid -ArgumentList (,$propertyValue))
     }
     else
     {
      # If of other length - try to convert to string SID
      $value = (New-Object System.Security.Principal.SecurityIdentifier($propertyValue,0)).Value
     }
    }
    default { $value = $propertyValue }
   }
   if ($values -eq "") { $values += $value } else { $values += "�$value" } 
  }
 return $values
}


############################################################################################################################
#
# Function that create/handles selection dialogbox
#

function createSelectionForm($dataTable)
{
 # Set up the environment
 Add-Type -AssemblyName System.Windows.Forms
 $LastColumnClicked = 0 # tracks the last column number that was clicked
 $LastColumnAscending = $false # tracks the direction of the last sort of this column
 
 # Create a form and a ListView
 $script:Form = New-Object System.Windows.Forms.Form
 $script:ListView = New-Object System.Windows.Forms.ListView
 $script:OkButton = New-Object System.Windows.Forms.Button
 $script:CancelButton = New-Object System.Windows.Forms.Button
 $script:Countlabel = New-Object System.Windows.Forms.Label
 $script:Migrlabel =  New-Object System.Windows.Forms.Label
 $script:LineDevider =  New-Object System.Windows.Forms.Label
 $script:MigrationMode = New-Object System.Windows.Forms.ComboBox

 # Configure the form
 $Form.Text = "Mailbox migration selection"
 $Form.Width = 400
 
 $x = $Form.ClientRectangle.Width-100
 $y = $Form.ClientRectangle.Height-28
 $OKButton.Location = New-Object System.Drawing.Size($x , $y  )
 $OKButton.Size = New-Object System.Drawing.Size(75,23)
 $OKButton.Text = "OK"
 $OKButton.Anchor = "Right, Bottom"
 $OKButton.enabled = $false
 $OKButton.Add_Click({$form.DialogResult = "Ok"; $Form.Close()})
 $Form.Controls.Add($OKButton)

 $x = $Form.ClientRectangle.Width-180
 $y = $Form.ClientRectangle.Height-28
 $CancelButton.Location = New-Object System.Drawing.Size($x , $y  )
 $CancelButton.Size = New-Object System.Drawing.Size(75,23)
 $CancelButton.Text = "Cancel"
 $CancelButton.Anchor = "Right, Bottom"
 $CancelButton.Add_Click({$Form.Close()})
 $Form.Controls.Add($CancelButton)


 $x = 145
 $y = $Form.ClientRectangle.Height-60
 $MigrationMode.Location = New-Object System.Drawing.Size($x , $y  )
 $MigrationMode.Width = 120
 $MigrationMode.Height = 30
 $MigrationMode.Anchor = "Left, Bottom"
 $MigrationMode.Items.Add("Incremental sync") | Out-Null
 $MigrationMode.Items.Add("Auto complete") | Out-Null
 $MigrationMode.SelectedIndex = 1
 $Form.Controls.Add($MigrationMode)


 $x = 5
 $dx = $Form.ClientRectangle.Width-10
 $y = $Form.ClientRectangle.Height-37
 $LineDevider.Location = New-Object System.Drawing.Size($x , $y  )
 $LineDevider.Width = $dx 
 $LineDevider.Height = 2
 $LineDevider.Anchor = "Left, Bottom, Right"
 $LineDevider.Text = ""
 $LineDevider.BorderStyle = "Fixed3D"
 $Form.Controls.Add($LineDevider)

 $x = 5
 $y = $Form.ClientRectangle.Height-28
 $Countlabel.Location = New-Object System.Drawing.Size($x , $y  )
 $Countlabel.Width = 200
 $Countlabel.Height = 30
 $Countlabel.Anchor = "Left, Bottom"
 $Countlabel.Text = "No selected users"
 $Form.Controls.Add($Countlabel)

 $x = 5
 $y = $Form.ClientRectangle.Height-58
 $Migrlabel.Location = New-Object System.Drawing.Size($x , $y  )
 $Migrlabel.Width = 140
 $Migrlabel.Height = 30
 $Migrlabel.Anchor = "Left, Bottom"
 $Migrlabel.Text = "Mailbox migration mode"
 $Form.Controls.Add( $Migrlabel)

 $ListView.View = [System.Windows.Forms.View]::Details
 $ListView.Width = $Form.ClientRectangle.Width
 $ListView.Height = $Form.ClientRectangle.Height-65
 $ListView.Anchor = "Top, Left, Right, Bottom"
 $Form.Controls.Add($ListView)

 # Add columns to the ListView
 $ListView.Columns.Add("Username", -2) | Out-Null
 $ListView.Columns.Add("Display Name") | Out-Null
 $ListView.Columns.Add("City") | Out-Null
 $ListView.Columns.Add("Office") | Out-Null
 $ListView.Columns.Add("Mailadress") | Out-Null
 $ListView.Columns.Add("AD Path") | Out-Null

 foreach($Row in $RowsToProcess)
 {
  $newRow = New-Object System.Windows.Forms.ListViewItem($row.sAMAccountName)
  $newRow.Subitems.Add($row.displayName.toString()) | Out-Null
  $newRow.Subitems.Add($row.l.toString()) | Out-Null
  $newRow.Subitems.Add($row.physicalDeliveryOfficeName.toString()) | Out-Null
  $newRow.Subitems.Add($row.mail.toString()) | Out-Null
  $newRow.Subitems.Add($row.distinguishedname) | Out-Null
  $ListView.Items.Add($newRow) | Out-Null
 }

 $ListView.BorderStyle = "Fixed3D"
 $ListView.FullRowSelect = $true
 
 # Set up the event handler
 $ListView.add_ColumnClick({SortListView $_.Column})
 $ListView.add_SelectedIndexChanged({SelectionChanged $.SelectedItems })
 return $form
} 

#####################################################################################
# Enable/disbaled the OK button

function SelectionChanged
{
 param([parameter(Position=0)]$selectedItems)
 
 if ($ListView.selectedItems.count -gt 0)
 {
  $OKButton.enabled = $true 
  $Countlabel.Text = $ListView.selectedItems.count.toString() + " selected users"
 }
 else
 {
  $OKButton.enabled = $false 
  $Countlabel.Text = "No selected users"
 } 
}

#####################################################################################
# Event handler for sorting
#
function SortListView
{
 param([parameter(Position=0)][UInt32]$Column)
 $Numeric = $true # determine how to sort
 
 if($Script:LastColumnClicked -eq $Column)
 {
  $Script:LastColumnAscending = -not $Script:LastColumnAscending
 }
 else
 {
  $Script:LastColumnAscending = $true
 }

 $Script:LastColumnClicked = $Column
 $ListItems = @(@(@())) # three-dimensional array; column 1 indexes the other columns, column 2 is the value to be sorted on, and column 3 is the System.Windows.Forms.ListViewItem object
 
 foreach($ListItem in $ListView.Items)
 {
  # if all items are numeric, can use a numeric sort
  if($Numeric -ne $false) # nothing can set this back to true, so don't process unnecessarily
  {
   try
   {
    $Test = [Double]$ListItem.SubItems.Text[$Column]
   }
   catch
   {
    $Numeric = $false # a non-numeric item was found, so sort will occur as a string
   }
  }
  $ListItems += ,@($ListItem.SubItems.Text[$Column], $ListItem)
 }
 
 $EvalExpression = {
  if($Numeric)
  { return [Double]$_[0] }
  else
  { return [String]$_[0] }
 }
 
 $ListItems = $ListItems | Sort-Object -Property @{Expression=$EvalExpression; Ascending=$Script:LastColumnAscending}
 
 $ListView.BeginUpdate()
 $ListView.Items.Clear()

 foreach($ListItem in $ListItems)
 {
  $ListView.Items.Add($ListItem[1])
 }
 $ListView.EndUpdate()
}

############################################################################################################################
#
# Function that creates aproval dialog box
#

function createAprovalForm($text)
{
 # Set up the environment
 Add-Type -AssemblyName System.Windows.Forms
 
 # Create a form and a ListView
 $script:Form = New-Object System.Windows.Forms.Form
 $script:TextBox = New-Object System.Windows.Forms.richTextBox
 $script:ConfirmCheckbox = New-Object System.Windows.Forms.CheckBox
 $script:OkButton = New-Object System.Windows.Forms.Button
 $script:CancelButton = New-Object System.Windows.Forms.Button
 $script:LineDevider =  New-Object System.Windows.Forms.Label

 # Configure the form
 $Form.Text = "Please confirm"
 $Form.Width = 400
 
 $x = $Form.ClientRectangle.Width-375
 $y = $Form.ClientRectangle.Height-28
 $ConfirmCheckBox.Location = New-Object System.Drawing.Size($x , $y  )
 $ConfirmCheckBox.Size  = New-Object System.Drawing.Size(150,23)
 $ConfirmCheckBox.Text = "Send Confirm migration!"
 $ConfirmCheckBox.Anchor = "Left, Bottom"
 $Form.Controls.Add($ConfirmCheckBox)

 $x = $Form.ClientRectangle.Width-100
 $y = $Form.ClientRectangle.Height-28
 $OKButton.Location = New-Object System.Drawing.Size($x , $y  )
 $OKButton.Size = New-Object System.Drawing.Size(75,23)
 $OKButton.Text = "Ok"
 $OKButton.Anchor = "Right, Bottom"
 $OKButton.Add_Click({$form.DialogResult = "Ok"; $Form.Close()})
 $Form.Controls.Add($OKButton)

 $x = $Form.ClientRectangle.Width-180
 $y = $Form.ClientRectangle.Height-28
 $CancelButton.Location = New-Object System.Drawing.Size($x , $y  )
 $CancelButton.Size = New-Object System.Drawing.Size(75,23)
 $CancelButton.Text = "Cancel"
 $CancelButton.Anchor = "Right, Bottom"
 $CancelButton.Add_Click({$Form.Close()})
 $Form.Controls.Add($CancelButton)

 $x = 5
 $dx = $Form.ClientRectangle.Width-10
 $y = $Form.ClientRectangle.Height-37
 $LineDevider.Location = New-Object System.Drawing.Size($x , $y  )
 $LineDevider.Width = $dx 
 $LineDevider.Height = 2
 $LineDevider.Anchor = "Left, Bottom, Right"
 $LineDevider.Text = ""
 $LineDevider.BorderStyle = "Fixed3D"
 $Form.Controls.Add($LineDevider)

 $TextBox.Width = $Form.ClientRectangle.Width
 $TextBox.Height = $Form.ClientRectangle.Height-65
 $TextBox.Anchor = "Top, Left, Right, Bottom"
 $TextBox.MultiLine = $true
 $TextBox.ReadOnly = $true
 $TextBox.ScrollBars = "Vertical"
 $TextBox.WordWrap = $false
 $TextBox.SelectionTabs = [int[]] 160
 $TextBox.TabStop = $false
 $TextBox.Text = $Text

 $Form.Controls.Add($TextBox)

 $ListView.BorderStyle = "Fixed3D"
 
 # Set up the event handler
 return $form
} 

 

######################################################################################################
# MAIN program

$scriptFileName = ($MyInvocation.MyCommand.Name).split(".")[0]
$logFilePath = "\\sthdcsrvb174.martinservera.net\script$\_log\"

$pathToDomainControllerTable = "\\sthdcsrvb174.martinservera.net\script$\_lib\$scriptFileName-domainControllerTable.xml"
$pathToUserDataBase = "\\sthdcsrvb174.martinservera.net\script$\_lib\$scriptFileName-userDataTable.xml"

# Opening of logfile
openLogFile "$logFilePath$(($MyInvocation.MyCommand.name).split('.')[0])-$(get-date -uformat %D)-$env:USERNAME.log"
# openLogFile($logFileName)
$timestamp = Get-Date -UFormat %HH-%MM-%SS
LogLine "============ $($timestamp) ============"

# Add recipients that should be emailed changes
if ( $sendMailLogs ) { setRecipients $sendMailLogsRecipients }
$script:fromAddress = "Usermig-" + ($MyInvocation.MyCommand.Name).split(".")[0]+"@martinservera.se"

# Setup dataTable columns.
$domainTable = CreateDomainControllerTable
$userTable = CreateUserTable

LogLine "Phase 1 - Checking domain controller USN"

# Load domainTable from XML file if existent and if not new columns are identified - otherwise create new table from script

if ((Test-Path -path $pathToDomainControllerTable) -and (-not $fullSync))
{
 LogLine " - Importing domain controller information from $pathToDomainControllerTable"
 [void] $domainTable.ReadXml($pathToDomainControllerTable)
}
else
{
 # If no file exists - create table
 LogLine " - Creating new domaintable ($pathToDomainControllerTable not found or new columns to process)"
 $fullSync = $true
 CreateInitialDomainControllerTable -dataTable $domainTable
}

# Create hashTable for checking for updates
$domainUSN = @{}

foreach($Row in $domainTable.Rows)
{ $domainUSN.Add($Row.DomainPart,$Row.highestCommittedUSN) }

LogLine " - retriving current USN values from domain controllers"
UpdateUSNFromDCS -domainTable $domainTable

#######################################################################################################
# Retrieve all changed users

LogLine "Phase 2 - Importing user information from domains"

if ((Test-Path -path $pathToUserDataBase) -and (-not $fullSync) )
{
 LogLine " - Importing usertable from $pathToUserDataBase"
 [void] $userTable.ReadXml($pathToUserDataBase)
}
else
{
 $fullSync = $true
}
$userTable.DefaultView.Sort = "objectGUID"

foreach($domainRow in $domainTable.Rows)
{
 if (-not $fullsync)
 {
  if ($domainRow.checkDomainForChanges -eq $true)
  {
   LogLine " - Importing diff user changes from $($domainRow.domainPart)"
   getObjectsToProcess -domainPart $domainRow.domainPart -domainController $domainrow.domainController -startUSN $domainRow.highestCommittedUSN -dataTable $userTable -valuesChanged $userValuesChanged 
  }
  else
  {
   LogLine " - Skipping unchanged domain $($domainRow.domainPart)"
  }
 }
 else
 {
  LogLine " - Issuing full user import from $($domainRow.domainPart)"
  getObjectsToProcess -domainPart $domainRow.domainPart -domainController $domainrow.domainController -dataTable $userTable -ValuesChanged $userValuesChanged 
 }
}

if ($script:processErrorCount -ne 0) 
{
 LogErrorLine "Script ended with errors ($processErrorCount) - $pathToDomainControllerTable not updated"
 LogInfoLine "Please re-run the script with -fullSync:$true"
 exit -1
}
else
{
 if ($fullSync -or $updateTables)
 {
  LogLine " AD update ended without major error - update XML files" 

  UpdateUSN -domainTable $domainTable
  [void] $domainTable.WriteXml($pathToDomainControllerTable)
  [void] $userTable.WriteXml($pathToUserDataBase)
 }
}


####################################################################################################
# 

LogLine "Phase 3 - Produce dialog"

$processIngCount = $0

# Return all rows for domain corp.lan that are migrated

if ($formerForest -eq "martinservera.se")
{ $targetAddressFilter = "%@martinservera.se" }

$RowsToProcess = $userTable.Select("DOMAIN='DC=martinservera,DC=net' AND mailNickName is not null AND msExchHomeServerName like '%STHDCSRV18%' AND mail like '$targetAddressFilter'")

LogLine "Phase 4 - Open selection dialog"

$newForm = createSelectionForm($RowsToProcess)
$result = $newForm.showDialog()

####################################################################################################
#
# Show a confirmation page 
#
$userNamesOfAccountsToBeMigrated = @()

$migrationType = "incrementalsync"

if ($result -eq "OK")
{
 if ($MigrationMode.SelectedIndex -eq 1)
 { $migrationType = "autocomplete" }

 $usersToMigrate = @()

 $message = "Please confirm action`r`n - Migration of $($ListView.SelectedItems.count) mailboxes`r`n - Using $migrationType`r`n`r`n"
 $message += "Username`tDisplayname`r`n---------------------------------------------------------------------------------------------`r`n"
 if ($ListView.SelectedItems.count -gt 0)
 {
  foreach($selection in $ListView.SelectedItems)
  { 
   $message += " $($ListView.items[$selection.index].Text)`t$($ListView.items[$selection.index].Subitems[1].Text)`r`n"
   $usersToMigrate += $ListView.items[$selection.index].Text

   $userNamesOfAccountsToBeMigrated += $ListView.items[$selection.index].Text
  } 
 }
 
 $newForm = createAprovalForm($message)
 $result = $newForm.showDialog()
 if ($result -eq "OK")
 {
  LogLine "Running migration for $($userNamesOfAccountsToBeMigrated.count)"
  # All usersnames are in the $user$userNamesOfAccountsToBeMigrated array
  # $migrationType = "autocomplete" or "incrementalsync"
  If ($migrationType.ToLower() -eq "autocomplete") { $migrationTypeIncremental = $False ; $migrationTypeAutoComplete = $true }
  Else { $migrationTypeIncremental = $True ; $migrationTypeAutoComplete = $False }
  $userNamesOfAccountsToBeMigrated | & '\\sthdcsrvb174.martinservera.net\Script$\migrateExchange\migrateExchange.ps1' -autoComplete:$migrationTypeAutoComplete -confirm:$($ConfirmCheckbox.Checked)
 }
 else
 {
  LogLine "Cancel selected on confirm dialogue, stopping..."
 }
}
else
{
 LogLine "Cancel selcted on selection dialogue, stopping...r"
}
StopAndReport
