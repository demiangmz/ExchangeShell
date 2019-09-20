<#
.SYNOPSIS
  Script used to set holidays for users' mailboxes in both Exchange Online (Office 365) and Exchange On-premises

.DESCRIPTION
  <Brief description of script>

.PARAMETER <Parameter_Name>
    <Brief description of parameter input required. Repeat this attribute if required>
.INPUTS
  <Inputs if any, otherwise state None>
.OUTPUTS
  <Outputs if any, otherwise state None - example: Log file stored in C:\Windows\Temp\<name>.log>
.NOTES
  Version:        1.1
  Author:         Demian Gomez
  Creation Date:  6/26/2019
  Purpose/Change: Added FreeBusy changes based on location. Minor script changes like use of switch variables.

  Version:        1.0
  Author:         Demian Gomez
  Creation Date:  1/4/2019
  Purpose/Change: Initial script development
  
.EXAMPLE
  <Example goes here. Repeat this attribute for more than one example>
#>

#---------------------------------------------------------[Initialisations]----------------------------------------------------------------------------------------------------
param(       
        [Parameter(Position=0, Mandatory=$true)]
        [ValidateScript({Test-Path $_ -PathType 'Leaf'})] 
        [string]$HolidaysCSVPath,

        [Parameter(Position=1, Mandatory=$false)]
        [ValidateScript({Test-Path $_ -PathType 'Leaf'})] 
        [string]$EWSLibraryPath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll",

        [Parameter(Position=2, Mandatory=$false)]
        [ValidateSet("Exchange2010", "Exchange2010_SP1","Exchange2010_SP2","Exchange2013","Exchange2013_SP1")]
        [String]$EWSVersion = "Exchange2010_SP2",

        [Parameter(Position=3, Mandatory=$false)]
        [switch]$ExchangeOnline,

        [Parameter(Position=4, Mandatory=$false)]
        [switch]$ExchangeOnPremises,

        [Parameter(Position=5, Mandatory=$false)]
        [switch]$RemoveExistingHolidays
)



#Set Error Action to Silently Continue
$ErrorActionPreference = "Stop"

#Dot Source required Function Libraries
#Example: . "C:\Scripts\Functions\Logging_Functions.ps1"

#----------------------------------------------------------[Declarations]-------------------------------------------------------------------------------------------------------

#Script Version
$sScriptVersion = "1.0"

#Log File Info
<#$sLogPath = "C:\Temp"
$sLogName = "UpdateHolidays.log"
$sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName
#>
#-----------------------------------------------------------[Functions]----------------------------------------------------------------------------------------------------------


Function Create-EWSServiceConnection {

param (
        [Parameter(Position=0, Mandatory=$true)]
        $Identity,
        
        [Parameter(Position=1, Mandatory=$false)]
	    [System.Management.Automation.PSCredential]$Credential,

        [Parameter(Position=2, Mandatory=$false)]
        [bool]$Impersonate = $true,

        [Parameter(Position=3, Mandatory=$false)]
        [ValidateSet("Exchange2010_SP1","Exchange2010_SP2","Exchange2013","Exchange2013_SP1")]
        [String]$EWSVersion = 'Exchange2010_SP2',

        [Parameter(Position=4, Mandatory=$false)]
        [String]$EWSUrl

        )

process {

        #Create the ExchangeService object
        $EWSService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList ([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$EWSVersion)

        #If Credential parameter used, set the credentials on the $service object
        if($Credential) {
                $EWSService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $Credential.UserName, $Credential.GetNetworkCredential().Password
        }

        #If EWSUrl parameter not used, locate the end-point using autoDiscover
        if(!$EWSUrl) {
                $EWSService.AutodiscoverUrl($Identity, {$true})
        }
        else {
                $EWSService.Url = New-Object System.Uri -ArgumentList $EWSUrl
        }        

        #If Impersonation parameter used, impersonate the user
        if($Impersonate) {
            $ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress),$Identity
            $EWSService.ImpersonatedUserId = $ImpersonatedUserId
        }
        
        return $EWSService
    }

} 

Function Remove-ExistingHolidays {

param(
        [Parameter(Position=0, Mandatory=$true)]
        $Identity,

        [Parameter(Position=1, Mandatory=$True)]
	    [Microsoft.Exchange.WebServices.Data.ExchangeService]
        $EWSService
       )


process {

        #Bind EWS to Calendar
        $Calendar = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$Identity)
        $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::IsAllDayEvent,$True)
        $ViewSettings = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000);
        $RemovedItems = $Null
        $Count = $Null

        
        #Start Removal Process
        Write-Host "Working on Calendar folder of "$Identity
        
        Do{
	        $CalendarItems = $EWSService.FindItems($Calendar,$SearchFilter,$ViewSettings)
	        $Count+=$CalendarItems.Items.Count
	        Write-Host "Working on items" $ViewSettings.Offset "to "$Count "..."
	        Foreach($Item in $CalendarItems.Items)
		        {
		        If($Item.Categories -eq "Holiday")
			        {
                     Write-Host "Holiday $($Item.Subject) Removed!" -ForegroundColor Gray
			        $RemovedItems+=@($Item)
			        $Item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
			        }
		        }
	        $ViewSettings.Offset+=$CalendarItems.Items.Count;
	        }
        
        While($CalendarItems.MoreAvailable)
		
        #Report and export
        Write-Host $RemovedItems.Count "Holiday Items were removed!" -ForegroundColor Green
        
    }
}

function New-Holiday {    
            
param(
        [Parameter(Position=0, Mandatory=$true)]
        $Subject,

        [Parameter(Position=1, Mandatory=$true)]
        $Date,

        [Parameter(Position=2, Mandatory=$True)]
	    [String]
	    $Location,

        [Parameter(Position=3, Mandatory=$True)]
	    [Microsoft.Exchange.WebServices.Data.LegacyFreeBusyStatus]
	    $FreeBusyStatus,
        
        [Parameter(Position=4, Mandatory=$True)]
	    [Microsoft.Exchange.WebServices.Data.ExchangeService]
        $EWSService
        )

process {
        
          
        #Configure the start and end time for this all day event
        $start = (get-date $date)
        $end = $start.addhours(24)
        

        #Create and save the appointment
        $appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment -ArgumentList $EWSService
        $appointment.Subject = $Subject
        $appointment.Start = $Start
        $appointment.End = $End
        $appointment.IsAllDayEvent = $true
        $appointment.IsReminderSet = $false
        $appointment.Categories.Add('Holiday')
        $appointment.Location = $Location
        $appointment.LegacyFreeBusyStatus = $FreeBusyStatus
        $appointment.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)
        }
}



#-----------------------------------------------------------[Main Script]---------------------------------------------------------------------------------------------------------

#Load the EWS Assembly
Add-Type -Path $EWSLibraryPath

#Load Holidays from CSV file
$Holidays = Import-Csv $HolidaysCSVPath

If ($ExchangeOnline) {
    Try {

    #EXO
    Start-Transcript -path "C:\Temp\HolidaysOutputEXO.txt" -Force
    
    #Connect to O365 and get employees' mailboxes
    $Credential = get-Credential -Message "Exchange Online"
    $EWSUrl = [system.URI] "https://outlook.office365.com/ews/exchange.asmx"
    Connect-EXOnline -Credential $Credential
    $Mailboxes = get-mailbox | ? {($_.ExchangeUserAccountControl -ne "AccountDisabled") -and ($_.RecipientType -eq "UserMailbox")}

    #Logic to filter location-based holidays and set FreeBusyStatus acordingly
    $Free = [Microsoft.Exchange.WebServices.Data.LegacyFreeBusyStatus]::Free
    $OOF  = [Microsoft.Exchange.WebServices.Data.LegacyFreeBusyStatus]::OOF
        
    $AREmployees = Get-Group "EmployeesARG" | select-object -ExpandProperty members
    $USEmployees = Get-Group "EmployeesUSA" | select-object -ExpandProperty members
        
    #Add Holidays to Each mailbox based on location. If user is from AR then add ARG Holidays as OOF and USA Holidays as Free, if user is from USA then do the same but in reverse.
    ForEach ($mbx in $Mailboxes) {
        $EWSService = Create-EWSServiceConnection -Identity $mbx.primarySmtpAddress -Credential $Credential -Impersonate $true -EWSUrl $EWSUrl -EWSVersion $EWSVersion
            
        If ($RemoveExistingHolidays) {
            Write-Host -ForegroundColor Yellow "Removing holidays from $mbx.Name Calendar..."
            Remove-ExistingHolidays -Identity $mbx.PrimarySmtpAddress -EWSService $EWSService
        }
	    Write-Host -ForegroundColor Yellow "Adding holidays to $mbx Calendar based on user location..."
	    
        If ($AREmployees.Contains($mbx.Identity)) { 
        
        ForEach ($Holiday in $Holidays) {
		    If ($Holiday.Location -eq "Argentina") {
                New-Holiday -Subject $holiday.HolidayName -Date $holiday.Date -Location $holiday.Location -EWSService $EWSService -FreeBusyStatus $OOF
		        Write-Host -ForegroundColor Green "> Added $($Holiday.holidayName)! (ARG)"
            }
            Else {
                
                New-Holiday -Subject $holiday.HolidayName -Date $holiday.Date -Location	$holiday.Location -EWSService $EWSService -FreeBusyStatus $Free
		        Write-Host -ForegroundColor Green "> Added $($Holiday.holidayName)! (USA)"
            }
	    }
        }

        ElseIf ($USEmployees.Contains($mbx.Identity)) { 

        ForEach ($Holiday in $Holidays) {
		    If ($Holiday.Location -eq "United States") {
                $FreeBusyStatus = [Microsoft.Exchange.WebServices.Data.LegacyFreeBusyStatus]::OOF
                New-Holiday -Subject $holiday.HolidayName -Date $holiday.Date -Location $holiday.Location -EWSService $EWSService -FreeBusyStatus $OOF
		        Write-Host -ForegroundColor Green "> Added $($Holiday.holidayName)! (USA)"
            }
            Else {
                $FreeBusyStatus = [Microsoft.Exchange.WebServices.Data.LegacyFreeBusyStatus]::Free
                New-Holiday -Subject $holiday.HolidayName -Date $holiday.Date -Location $holiday.Location -EWSService $EWSService -FreeBusyStatus $Free
		        Write-Host -ForegroundColor Green "> Added $($Holiday.holidayName)! (ARG)"
            }
	    }
        } 
        Else {
            Write-Host -ForegroundColor Red "ERROR: Employee Location could not be determined, please check membership of Employees or Contractors groups (AR or US) in O365!"
            }

    }
    }

    Catch {
            Write-Host $_.Exception.Message 
          }

    Finally {
            Disconnect-EXOnline
            Stop-Transcript
            $Log = Get-Content C:\Temp\HolidaysOutputEXO.txt
            $Log > C:\Temp\HolidaysOutputEmeriosEXO.txt
    }
}

If ($ExchangeOnPremises) {
    Try {
    
    #Exchange On Premises
    Start-Transcript -Path "C:\Temp\HolidaysOutput.txt" -Force
    
    $Credential = get-Credential -Message "Exchange OnPremises"
    Connect-EXOnPremises -Credential $Credential
    #$Mailboxes = Get-mailbox dgomez
    $Mailboxes = Get-Mailbox -OrganizationalUnit domain.com/Employees | ? {$_.HiddenFromAddressListsEnabled -ne "true"} | Sort-Object

    ForEach ($mbx in $Mailboxes) {
        $EWSService = Create-EWSServiceConnection -Identity $mbx.primarySmtpAddress -Credential $Credential -Impersonate $true -EWSVersion $EWSVersion
        Write-Host "*************************************************************"    
        If ($RemoveExistingHolidays) {
            Remove-ExistingHolidays -Identity $mbx.PrimarySmtpAddress -EWSService $EWSService
	         Write-Host -ForegroundColor Yellow "Adding new holidays to $mbx Calendar..."
        }
	        
        ForEach ($Holiday in $Holidays) {
		    New-Holiday -Subject $holiday.HolidayName -Date $holiday.Date -Location $holiday.Location -EWSService $EWSService
		    Write-Host -ForegroundColor Green "> Added $($Holiday.holidayName)!"
	    }
    }
    }
    Catch {
            Write-Host $_.Exception.Message 
           }
    Finally {
        Disconnect-EXOnPremises
        Stop-Transcript
        $Log = Get-Content C:\Temp\HolidaysOutput.txt
        $Log > C:\Temp\HolidaysOutput.txt
    }
}
