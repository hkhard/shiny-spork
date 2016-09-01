### Connect to Exchange 2016
Function connect-exchange
{
If ((Get-PSSession | Where-Object {$_.ComputerName -like "sthdcsrvb*"}).count -lt 1)
 {
  Write-Host -ForegroundColor Green "... Please wait... will load Exchange 2016 PowerShell..."
  $session = New-PSSession -ConfigurationName Microsoft.Exchange2016.MS -ConnectionURI https://sthdcsrvb153.martinservera.net/Powershell/ -Authentication NegotiateWithImplicitCredential -AllowRedirection
  Write-Host -ForegroundColor Green "... Loading Exchange 2016 PowerShell for Martin & Servera AB ..."
  Import-PSSession $Session -AllowClobber
 }
}

### Connect to s4b
Function Connect-s4b
{
If ((Get-PSSession | Where-Object {$_.ComputerName -like "sthdcsrv8*"}).count -lt 1)
 {
 Write-Host -ForegroundColor Green "... Please wait... will load Skype for Business PowerShell..."
 $session = New-PSSession -ConnectionURI "https://sthdcsrv83.martinservera.net/OcsPowershell" -Authentication NegotiateWithImplicitCredential
 Write-Host -ForegroundColor Green "... Loading Skype for Business PowerShell for Martin & Servera AB ..." 
 Import-PSSession $session -AllowClobber
 }
}
### Main
Connect-exchange
Connect-s4b