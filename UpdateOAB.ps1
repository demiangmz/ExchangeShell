#----------------------------------------------------------------------------------
# Script:  UpdateAddressLists.ps1
# Author:  Demian Gomez
# Date:  04/08/2013
# Version: 1.0
#-----------------------------------------------------------------------------------

param(
	[Parameter(Position=0)]
	$ExchangeVersion = ('Exchange2010', 'Exchange2013', 'Exchange2016', 'Exchange2019')
  )


#Global variable to enable Catch statement of "Object not found" exception.
$ErrorActionPreference = "Stop"

#Aux variable declaration
$EventNotFound = $False
$i = 1

#----------------
# Script START
#----------------

Write-Host "Updating Global Address List..." -foregroundcolor yellow
Update-GlobalAddressList "Default Global Address List" | Out-Null
Write-Host "Updating Offline Address Book (it may take up to 5 min)..." -foregroundcolor yellow
Get-OfflineAddressBook | Update-OfflineAddressBook | Out-Null
Start-Sleep 60

If ($ExchangeVersion -eq 'Exchange2010') {

#Loop that checks for Event ID 9107 presence in Application log, if found exits loop, if not found after 5 min it also exits loop.
Do {

Try  {
       Get-EventLog -Log Application -Source MSExchangeSA -After (Get-Date).AddMinutes(-5) | Where {$_.eventID -eq 9107} | Out-Null
  }
Catch { 
	$EventNotFound = $True
	i++
	Start-Sleep 60
  } 
}

While (($EventNotFound) -and (i -lt 5))

# IF the loop ended after 5 min, meaning OAB generation process wasn't successful, output shows this
If ($EventNotFound) { 
	Write-Host "Update-OfflineAddressBook process failed!" -foregroundcolor red
	Write-Host "Check Event Viewer Application log for error events with MSExchangeSA source to start troubleshooting process."
	Write-Host "You can also run Update-OfflineAddressBook with -verbose switch to get detailed errors"
	} 
# ELSE Event ID 9107 was found it means OAB was rebuilt successfully, so we update OAB files on every CAS server.
	Else {
	Write-Host "Updating OAB Files on each CAS Server…" -foregroundcolor yellow
	Get-ClientAccessServer | Update-FileDistributionService -Type OAB $CAS.name
	}
}

Write-Host "Update Process Completed OK!" -foregroundcolor green
Pause
#----------------
# Script END
#----------------
