### Connect to Exchange 2016
If ((Get-PSSession | Where-Object {$_.ComputerName -like "sthdcsrvb*"}).count -lt 1)
 {
  Write-Host -ForegroundColor Green "... Please wait... will load Exchange 2016 PowerShell..."
  $session = New-PSSession -ConfigurationName Microsoft.Exchange2016.MS -ConnectionURI https://sthdcsrvb153.martinservera.net/Powershell/ -Authentication NegotiateWithImplicitCredential -AllowRedirection
  Write-Host -ForegroundColor Green "... Loading Exchange 2016 PowerShell for Martin & Servera AB ..."
  Import-PSSession $Session -AllowClobber
 }