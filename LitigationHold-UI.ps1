# Author: Eyal Doron 
# Version: 1.1.1
# Last Modified Date: 24/10/2019  
# Last Modified By: timour.varrasse@atea.com
# TODO : Fix exchange connection module. I use my own.


#------------------------------------------------------------------------------
# PowerShell Functions
#------------------------------------------------------------------------------
Function Write-ToLOGFfile()
{
	param( $LOGFileEntry )
	"$LOGFileEntry" | Out-File $LOGFilePath -Append
}

function checkConnection ()
{
		$a = Get-PSSession
		if ($a.ConfigurationName -ne "Microsoft.Exchange")
		{
			
			write-host     'You are not connected to Exchange Online PowerShell ;-(         ' 
			write-host      'Please connect using the Menu option 1) Login to Office 365 + Exchange Online using Remote PowerShell        '
			#Read-Host "Press Enter to continue..."
			Add-Type -AssemblyName System.Windows.Forms
			[System.Windows.Forms.MessageBox]::Show("You are not connected to Exchange Online PowerShell ;-( `nSelect menu 1 to connect `nPress OK to continue...", "o365info.com PowerShell script", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
			Clear-Host
			break
		}
}



Function DisconnectExchangeOnline ()
{
Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"} | Remove-PSSession

}


Function Set-AlternatingRows {
       <#
       
       #>
    [CmdletBinding()]
       Param(
             [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [string]$Line,
       
           [Parameter(Mandatory=$True)]
             [string]$CSSEvenClass,
       
        [Parameter(Mandatory=$True)]
           [string]$CSSOddClass
       )
       Begin {
             $ClassName = $CSSEvenClass
       }
       Process {
             If ($Line.Contains("<tr>"))
             {      $Line = $Line.Replace("<tr>","<tr class=""$ClassName"">")
                    If ($ClassName -eq $CSSEvenClass)
                    {      $ClassName = $CSSOddClass
                    }
                    Else
                    {      $ClassName = $CSSEvenClass
                    }
             }
             Return $Line
       }
}




#------------------------------------------------------------------------------
# Genral
#------------------------------------------------------------------------------
$FormatEnumerationLimit = -1
$Date = Get-Date
$Datef = Get-Date -Format "\Da\te dd-MM-yyyy \Ti\me H-mm" 
#------------------------------------------------------------------------------
$lineH = "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
$lineH1 = "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"

$line  = "--------------------------------------------------------------------------------------------------"
$line1 = "-------------------------------------------------------------------------------------------------------------------------------------"


#------------------------------------------------------------------------------

# PowerShell console window Style
#------------------------------------------------------------------------------

$pshost = get-host
$pswindow = $pshost.ui.rawui

	$newsize = $pswindow.buffersize
	
	if($newsize.height){
		$newsize.height = 3000
		$newsize.width = 150
		$pswindow.buffersize = $newsize
	}

	$newsize = $pswindow.windowsize
	if($newsize.height){
		$newsize.height = 50
		$newsize.width = 150
		$pswindow.windowsize = $newsize
	}

#------------------------------------------------------------------------------
# HTML Style start 
#------------------------------------------------------------------------------
$Header = @"
<style>
Body{font-family:segoe ui,arial;color:black; }

H1 {font-size: 26px; font-weight:bold;width: 70% text-transform: uppercase; color: #0000A0; background:#2F5496 ; color: #ffffff; padding: 10px 10px 10px 10px ; border: 3px solid #00B0F0;}
H2{ background:#F2F2F2 ; padding: 10px 10px 10px 10px ; color: #013366; margin-top:35px;margin-bottom:25px;font-size: 22px;padding:5px 15px 5px 10px; }

.TextStyle {font-size: 26px; font-weight:bold ; color:black; }

TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 5px;border-style: solid;border-color: #d1d3d4;background-color:#0072c6 ;color:white;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}

.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }

.o365info {height: 90px;padding-top:5px;padding-bottom:5px;margin-top:20px;margin-bottom:20px;border-top: 3px dashed #002060;border-bottom: 3px dashed #002060;background: #00CCFF;font-size: 120%;font-weight:bold;background:#00CCFF url(http://o365info.com/wp-content/files/PowerShell-Images/o365info120.png) no-repeat 680px -5px;
}

</style>

"@

$EndReport = "<div class=o365info>  This report was created by using <a href= http://o365info.com target=_blank>o365info.com</a> PowerShell script </div>"
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------




#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# Exchange Online | Manage Litigation Hold - PowerShell - Script menu
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

$Loop = $true
While ($Loop)
{
    write-host 
    write-host +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    write-host   "Exchange Online | Manage Litigation Hold - PowerShell - Script menu"
    write-host +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    write-host
	write-host -ForegroundColor white  '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor DarkCyan     'Connect Exchange Online using Remote PowerShell        ' 
    write-host -ForegroundColor white  '----------------------------------------------------------------------------------------------' 
	write-host -ForegroundColor Yellow ' 1) Login to Exchange Online using Remote PowerShell ' 
    write-host
	write-host -ForegroundColor green '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor Blue      'SECTION A: Assign Litigation Hold to Mailbox        ' 
    write-host -ForegroundColor green '----------------------------------------------------------------------------------------------' 
    write-host                                              ' 2)  Assign Litigation Hold | Single Mailbox '
	write-host                                              ' 3)  Assign Litigation Hold + define time range | Single Mailbox  '
	write-host -ForegroundColor green '----------------------------------------------------------------------------------------------' 
	write-host                                              ' 4)  Assign Litigation Hold | ALL USER Mailboxes witout Litigation Hold (BULK mode) '
	write-host                                              ' 5)  Assign Litigation Hold based on a Property (Department) | ALL USER Mailboxes (BULK mode)  '
	write-host                                              ' 6)  Assign Litigation ALL USER Mailboxes that dont have Litigation Hold |(BULK mode)  '
	write-host -ForegroundColor green  '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor DarkGreen 'SECTION B:  Display + Export information about Litigation Hold   ' 
    write-host -ForegroundColor green  '----------------------------------------------------------------------------------------------' 
    write-host                                              ' 7)  Display information about Litigation Hold | Single Mailbox  '
	write-host                                              ' 8)  Export information about Litigation Hold | ALL USER Mailboxes   '
	write-host                                              ' 9)  Export information about Litigation Hold | USERS Mailboxes which have a Litigation Hold enabled '
	write-host                                              ' 10) Export information about Litigation Hold | USERS Mailboxes Mailboxes witout Litigation Hold '
	write-host -ForegroundColor green  '----------------------------------------------------------------------------------------------' 
	write-host                                              ' 11) Export information about Litigation Hold | Information about the recoverable items folder  '
	write-host -ForegroundColor green  '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor DarkRed   'SECTION C:  Remove Litigation Hold  ' 
    write-host -ForegroundColor green  '----------------------------------------------------------------------------------------------' 
    write-host                                              ' 12) Remove (Disable) Litigation Hold | Single Mailbox '
    write-host                                              ' 13) Remove (Disable) Litigation Hold based on a criterion (Department) | ALL USER Mailboxes (Bulk mode) '
    write-host                                              ' 14) Remove (Disable) Litigation Hold | ALL USER Mailboxes WITH Litigation Hold enabled (BULK mode) '
	write-host -ForegroundColor green  '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor DarkCyan 'End of PowerShell - Script menu ' 
    write-host -ForegroundColor green  '----------------------------------------------------------------------------------------------' 
	write-host -ForegroundColor Yellow            "20)  Disconnect PowerShell session" 
    write-host
    write-host -ForegroundColor Yellow            "21)  Exit (Or use the keyboard combination - CTRL + C)" 
    write-host

    $opt = Read-Host "Select an option [1-21]"
    write-host $opt
    switch ($opt) 


{



		
1
{

#####################################################################
# Connect Exchange Online using Remote PowerShell
#####################################################################

# == Section: General information ===

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------                                                                
write-host  -ForegroundColor white  	'To be able to use the PowerShell menus in the script,  '
write-host  -ForegroundColor white  	'you will need to Login to Exchange Online using Remote PowerShell. '
write-host  -ForegroundColor white  	'In the credentials windows that appear,   '
write-host  -ForegroundColor white  	'provide your Office 365 Global Administrator credentials.  '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------  
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host

DisconnectExchangeOnline

## This is my alias to connect to exchange @Atea with MFA. You can use your own/Will fix later.
exchange 



# EXO | Exchange Online Recipient type - Uers object 
#------------------------------------------------------
# EXO | Exchange Online mailbox type
#------------------------------------------------------
# All Exchange Online USER Mailboxes 
$global:GetMBXUser     =  Get-MailBox -ResultSize Unlimited -Filter '(RecipientTypeDetails -eq "UserMailbox")' | Sort-Object -Property displayname


#------------------------------------------------------------------------------
# Object Properties Array 
#------------------------------------------------------------------------------

# Report included fields
$global:Array1  = "DisplayName","Alias","LitigationHoldEnabled","LitigationHoldDate", "LitigationHoldOwner","LitigationHoldDuration ","RecipientType","RecipientTypeDetails"

$global:Array2  = "DisplayName","Alias","LitigationHoldEnabled","LitigationHoldDate", "LitigationHoldOwner","LitigationHoldDuration"


}






	
#=========================================================================================
# SECTION A: Assign Litigation Hold to Mailbox 
#=========================================================================================


2
{


#####################################################################
# Assign Litigation Hold | Single Mailbox
#####################################################################

# == Section: General information === 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		----------------------------------------------------------------------------                                                             
write-host  -ForegroundColor white  	'This option will: '
write-host  -ForegroundColor white  	'Assign Litigation Hold | Single Mailbox. '
write-host  -ForegroundColor white		--------------------------------------------------------------------  
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Set-Mailbox <Mailbox> -LitigationHoldEnabled $True '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host


# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 1 parameter:"  
write-host
write-host -ForegroundColor Yellow	"1.  Mailbox name  "  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Target mailbox"    
write-host -ForegroundColor Yellow	"For example:  Bob@o365info.com"
write-host
$Alias  = Read-Host "Type the mailbox name "
write-host
write-host

# == Section: Display Information  ===	

write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue   Display information about: "$Alias".ToUpper() LitigationHold => BEFORE the update
write-host -------------------------------------------------------------------------------------------------

Get-Mailbox $Alias | Select-Object $global:Array2  |  Out-String  

write-host -------------------------------------------------------------------------------------------------

# == Section: PowerShell Command  ===	
Set-Mailbox $Alias -LitigationHoldEnabled $True



# == Section: Display Information  ===	



write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Magenta  Display information about: "$Alias".ToUpper() LitigationHold  => AFTER the update
write-host -------------------------------------------------------------------------------------------------

Get-Mailbox $Alias | Select-Object $global:Array2  |  Out-String  

write-host -------------------------------------------------------------------------------------------------


# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}



3
{


#####################################################################
# Assign Litigation Hold + define time range | Single Mailbox
#####################################################################

# == Section: General information === 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		----------------------------------------------------------------------------                                                             
write-host  -ForegroundColor white  	'This option will: '
write-host  -ForegroundColor white  	'Assign Litigation Hold + define time range | Single Mailbox. '
write-host  -ForegroundColor white		--------------------------------------------------------------------  
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Set-Mailbox <Mailbox> -LitigationHoldEnabled $True -LitigationHoldDuration <Time Range >'
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host


# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 2 parameters:"  
write-host
write-host -ForegroundColor Yellow	"1.  Mailbox name  "  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Target mailbox"    
write-host -ForegroundColor Yellow	"For example:  Bob@o365info.com"
write-host
$Alias  = Read-Host "Type the mailbox name "
write-host
write-host

write-host -ForegroundColor Yellow	"2.  Time Range  "  
write-host -ForegroundColor Yellow	"The Second parameter is the time range for the Litigation Hold"   
write-host -ForegroundColor Yellow	"for example - 2555 for time range of 7 years"   
write-host
$Time =  Read-Host "Type the time range for the Litigation Hold"

# == Section: Display Information   ===	

write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue   Display information about: "$Alias".ToUpper() LitigationHold => BEFORE the update
write-host -------------------------------------------------------------------------------------------------
Get-Mailbox $Alias | Select-Object $global:Array2  |  Out-String  
write-host -------------------------------------------------------------------------------------------------

# == Section: PowerShell Command  ===	
Set-Mailbox $Alias -LitigationHoldEnabled $True -LitigationHoldDuration $Time



# Display Information  ===	



write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Magenta  Display information about: "$Alias".ToUpper() LitigationHold  => AFTER the update
write-host -------------------------------------------------------------------------------------------------
Get-Mailbox $Alias | Select-Object $global:Array2  |  Out-String  
write-host -------------------------------------------------------------------------------------------------


# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}




	
		
4
{


###################################################################################################################################
# Assign Litigation Hold | ALL USER Mailboxes (Bulk mode)
###################################################################################################################################

checkConnection
# General information 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		----------------------------------------------------------------------------  
write-host  -ForegroundColor white  	"Assign Litigation Hold | ALL USER Mailboxes (Bulk mode) "
write-host  -ForegroundColor white		----------------------------------------------------------------------------  
write-host  -ForegroundColor white  	"This menu option, will enable Exchange Litigation Hold for each existing Exchange Online    "
write-host  -ForegroundColor white  	"User mailbox (Enable Audit in Bulk mode) Which doesn’t have Litigation Hold enabled.    "
write-host  -ForegroundColor white		--------------------------------------------------------------------  
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host

#---------------------------------------------------------------------------------------------------
# Define variables
#---------------------------------------------------------------------------------------------------
# Define date format variable
$Datef = Get-Date -Format "\Da\te dd-MM-yyyy \Ti\me H-mm" 

#---------------------------------------------------------------------------------------------------
# Define variables that contain the folders names 
#---------------------------------------------------------------------------------------------------

# C:\INFO\Litigation Hold informaiton
$A20 =  "C:\INFO\Litigation Hold informaiton - $Datef"
# 1. Report mailboxes which the Litigation Hold where enabled for them
$A21 =  "$A20\1. Report mailboxes which the Litigation Hold where enabled for them"
# 2. All mailboxes with Litigation Hold
$A22 =  "$A20\2. All mailboxes with Litigation Hold"

#---------------------------------------------------------------------------------------------------
# Create folders Structure that contain the exported information to TXT, CSV and HTML files
#---------------------------------------------------------------------------------------------------

#  C:\INFO\Litigation Hold informaiton -<Date>
if (!(Test-Path -path $A20))
{New-Item $A20 -type directory}
# 1. Report mailboxes which the Litigation Hold where enabled for them
if (!(Test-Path -path $A21))
{New-Item $A21 -type directory}
# 2. All mailboxes with Litigation Hold
if (!(Test-Path -path $A22))
{New-Item $A22 -type directory}

#---------------------------------------------------------------------------------------------------
# Define various parameter that are related to the Log file
#---------------------------------------------------------------------------------------------------
# Define variables for the LOG file name
$LOGFileName = "Log File.TXT"
$LOGFilePath = "$A21\$LOGFileName"
# Define variables for the LOG file header text 
write-ToLOGFfile $lineH1
Write-ToLOGFfile "Enable Litigation Hold for All Exchange Online USER mailboxes without Litigation Hold  | $Datef"
write-ToLOGFfile $lineH1
#---------------------------------------------------------------------------------------------------
# Define the Array 
#---------------------------------------------------------------------------------------------------
# Define the Array of Exchange Online mailboxes | All USER mailboxes
$GetMBXUser  =  Get-MailBox -ResultSize Unlimited -Filter '(RecipientTypeDetails -eq "UserMailbox")' | Sort-Object -Property displayname
# Define the Array of Exchange Online mailboxes | All USER mailboxes with LitigationHold Enabled
$AllMailboxesWithLitigationHold = $GetMBXUser      | Where-Object {$_.LitigationHoldEnabled -eq $True} | Sort-Object -Property displayname
# Define the Array of Exchange Online mailboxes | All USER mailboxes with LitigationHold Disabled
$AllMailboxesWithOUTLitigationHold = $GetMBXUser   | Where-Object {$_.LitigationHoldEnabled -eq $False} | Sort-Object -Property displayname

# Count the number of Exchange Online USER mailboxes
$GetMBXUserCount = ($GetMBXUser).count

# Count the number of Exchange Online USER mailboxes with Litigation Hold  Enabled
$AllMailboxesWithLitigationHoldCount = ($AllMailboxesWithLitigationHold).count

# Count the number of Exchange Online USER mailboxes with OUT Litigation Hold 
$AllMailboxesWithOUTLitigationHoldCount = ($AllMailboxesWithOUTLitigationHold).count



# Display information about: Exchange Online USER mailboxes and Litigation Hold
write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor blue  "Display information about: Exchange Online USER mailboxes and Litigation Hold"
write-host -------------------------------------------------------------------------------------------------
write-host -------------------------------------------------------------------------------------------------
write-host "There are: " -nonewline; write-host "$GetMBXUserCount" -ForegroundColor Yellow -BackgroundColor blue -nonewline; write-host " Exchange Online USER mailboxes " 
write-host
write-host "There are: " -nonewline; write-host "$AllMailboxesWithLitigationHoldCount" -ForegroundColor Yellow -BackgroundColor blue -nonewline; write-host " Exchange Online USER mailboxes with Litigation Hold Enabled " 
write-host
write-host "There are: " -nonewline; write-host "$AllMailboxesWithOUTLitigationHoldCount" -ForegroundColor Yellow -BackgroundColor blue -nonewline; write-host " Exchange Online USER mailboxes with OUT Litigation Hold " 
write-host -------------------------------------------------------------------------------------------------
write-host
write-host

#---------------------------------------------------------------------------------------------------
# PowerShell Command
#---------------------------------------------------------------------------------------------------

$AllMailboxes = $AllMailboxesWithOUTLitigationHold

# Document the updates by writing information to LOG file 
write-ToLOGFfile $line1
write-ToLOGFfile "There are: $GetMBXUserCount Exchange Online USER mailboxes " 
write-host
write-ToLOGFfile "There are: $AllMailboxesWithLitigationHoldCount Exchange Online USER mailboxes with Litigation Hold Enabled" 
write-host
write-ToLOGFfile "There are: $AllMailboxesWithOUTLitigationHoldCount Exchange Online USER mailboxes WITH OUT Litigation Hold" 
write-ToLOGFfile $line1


ForEach ($Mailbox in $AllMailboxes)
{

# User \ Mailbox identity | Single member from the Array     

$ID1 = $Mailbox.Displayname

# Enable Litigation Hold  on The specified Exchange mailbox (Array member)  
Set-Mailbox $ID1  -LitigationHoldEnabled $True
#  Display progress bar information on the PowerShell console
Write-Progress -Activity "Enable Litigation Hold on - Exchange $ID1 Mailbox"


# Document the updates by writing information to LOG file 
write-ToLOGFfile $line1
Write-ToLOGFfile "Litigation Hold was enabled for - $ID1 mailbox"
write-ToLOGFfile $line1


# Display information about the action that was performed
write-host
write-host -ForegroundColor white		---------------------------------------------------------------------------- 
write-host -ForegroundColor white	"Litigation Hold  was enabled for - " -nonewline; write-host $ID1 -ForegroundColor white -BackgroundColor Darkgreen -nonewline; write-host " mailbox " -ForegroundColor white  
write-host -ForegroundColor white		----------------------------------------------------------------------------  
write-host

}



#---------------------------------------------------------------------------------------------------
# Export Exchange Online Mailboxes Litigation Hold settings to files
#---------------------------------------------------------------------------------------------------

$GetMBXUser =  $null

# Define the Array of Exchange Online mailboxes | All USER mailboxes
$GetMBXUser  =  Get-MailBox -ResultSize Unlimited -Filter '(RecipientTypeDetails -eq "UserMailbox")' | Sort-Object -Property displayname

#---------------------------------------------------------------------------------------------------
# Define variables that contain the files names 
#---------------------------------------------------------------------------------------------------
$FileName1 =   "Exchange Online Mailboxes Litigation Hold settings"



### TXT ####
$GetMBXUser | Select-Object $global:Array1  | Format-List | Out-File $A22\"$FileName1.txt" -Encoding UTF8
##########

### CSV ####
$GetMBXUser | Select-Object $global:Array1  | Export-CSV $A22\"$FileName1.CSV" –NoTypeInformation -Encoding utf8
##########

### HTML ####
$GetMBXUser | Select-Object $global:Array1  | ConvertTo-Html  -post $EndReport -head $Header -Body  "<H1>$FileName1 | $Datef </H1>"  | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-File $A22\"$FileName1.html"
##########




# START User notification about the location of the exported files
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		----------------------------------------------------------------------------                                                             
write-host  -ForegroundColor white  	"The information about Exchange Online mailboxes that their Litigation Hold was enabled, was written to a LOG file "
write-host  -ForegroundColor white  	"You can find the LOG file in the following path:  "
write-host  -BackgroundColor DarkGreen 	$A20   
write-host  -ForegroundColor white		--------------------------------------------------------------------  
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo   
# END User notification about the location of the exported files

# Empty the content of the variable

$LOGFilePath =  $null
$Datef =  $null
$AllMailboxes =  $null
$ID1 =  $null
$GetMBXUser =  $null
$AllMailboxes =  $null
$AllMailboxesWithLitigationHold =  $null
$AllMailboxesWithOUTLitigationHold  =  $null


# End the Command
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}






5
{


#####################################################################
# Assign Litigation Hold based on a criterion (Department) | ALL USER Mailboxes (Bulk mode)
#####################################################################

# == Section: General information === 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		----------------------------------------------------------------------------                                                             
write-host  -ForegroundColor white  	'This option will: '
write-host  -ForegroundColor white  	'Assign Litigation Hold based on a criterion (Department) | ALL USER Mailboxes (Bulk mode). '
write-host  -ForegroundColor white		--------------------------------------------------------------------  
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Get-Recipient -RecipientTypeDetails UserMailbox -ResultSize unlimited -Filter (Department -eq "<Department>") | Set-Mailbox -LitigationHoldEnabled $True'
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host


# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 1 parameter:"  
write-host
write-host -ForegroundColor Yellow	"1.  Department name  "  
write-host -ForegroundColor Yellow	"Provide the Department name"    
write-host -ForegroundColor Yellow	"For example:  HR"
write-host
$Department = Read-Host "Type the Department name  "
write-host
write-host


# == Section: Display Information about xxxx  ===	


$DepartmentUsers = Get-user | Where-Object {($_.RecipientTypeDetails -eq 'UserMailbox') -and ($_.Department -eq $Department)}


write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue    Display information users from : "$Department".ToUpper() 
write-host -------------------------------------------------------------------------------------------------

$DepartmentUsers  |  Out-String  

write-host -------------------------------------------------------------------------------------------------
write-host
write-host
write-host


write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue    Display information about: Litigation Hold Settings => BEFORE the update
write-host -------------------------------------------------------------------------------------------------

$DepartmentUsers  | Select-Object $global:Array2 |  Out-String  

write-host -------------------------------------------------------------------------------------------------

# == Section: PowerShell Command  ===	
$DepartmentUsers  | Set-Mailbox -LitigationHoldEnabled $True



# == Section: Display Information about xxxx  ===	



write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Magenta  Display information about: Litigation Hold Settings  => AFTER the update
write-host -------------------------------------------------------------------------------------------------

$DepartmentUsers   | Select-Object $global:Array2 |  Out-String  

write-host -------------------------------------------------------------------------------------------------


# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}




6
{


#####################################################################
# Assign Litigation ALL USER Mailboxes that dont have Litigation Hold | (Bulk mode)
#####################################################################

# == Section: General information === 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		----------------------------------------------------------------------------                                                             
write-host  -ForegroundColor white  	'This option will: '
write-host  -ForegroundColor white  	' Assign Litigation ALL USER Mailboxes that dont have Litigation Hold | (Bulk mode). '
write-host  -ForegroundColor white		--------------------------------------------------------------------  
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Get-Mailbox | Where {$_.LitigationHoldEnabled -match "False"} | ForEach-Object {$Identity = $_.alias; Set-Mailbox -Identity $Identity -LitigationHoldEnabled $True } '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host



# == Section: Display Information about xxxx  ===	

write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue   Display information about: Litigation Hold Settings => BEFORE the update
write-host -------------------------------------------------------------------------------------------------

$global:GetMBXUser  | Where-Object {$_.LitigationHoldEnabled -match "False"}  | Select-Object Name,LitigationHold* |  Out-String  

write-host -------------------------------------------------------------------------------------------------

# == Section: PowerShell Command  ===	
$global:GetMBXUser  | Where-Object {$_.LitigationHoldEnabled -match "False"} | ForEach-Object {$Identity = $_.alias; Set-Mailbox -Identity $Identity -LitigationHoldEnabled $True } 



# == Section: Display Information about xxxx  ===	



write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Magenta  Display information about: Litigation Hold Settings => AFTER the update
write-host -------------------------------------------------------------------------------------------------

$global:GetMBXUser | Where-Object {$_.LitigationHoldEnabled -match "False"} | Select-Object Name,LitigationHold* |  Out-String  

write-host -------------------------------------------------------------------------------------------------


# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}







		
#===============================================================================================================
# SECTION B:  Display + Export information about Litigation Hold
#===============================================================================================================



7
{


#####################################################################
# Display information about Litigation Hold | Single Mailbox
#####################################################################

# == Section: General information === 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		----------------------------------------------------------------------------                                                             
write-host  -ForegroundColor white  	'This option will: '
write-host  -ForegroundColor white  	'Display information about Litigation Hold | Single Mailbox. '
write-host  -ForegroundColor white		--------------------------------------------------------------------  
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Get-Mailbox <Mailbox> | Select-Object Name,LitigationHold* '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host


# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 1 parameter:"  
write-host
write-host -ForegroundColor Yellow	"1.  Mailbox name  "  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Target mailbox"    
write-host -ForegroundColor Yellow	"For example:  Bob@o365info.com"
write-host
$Alias  = Read-Host "Type the mailbox name "
write-host
write-host

# == Section: Display Information  ===	

write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue   Display information about: "$Alias".ToUpper() LitigationHold 
write-host -------------------------------------------------------------------------------------------------

Get-Mailbox $Alias | Select-Object Name,LitigationHold* |  Out-String  

write-host -------------------------------------------------------------------------------------------------




# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}




8
{

##########################################################################################################
#  Export information about Litigation Hold | ALL USER Mailboxes
##########################################################################################################
checkConnection
# == Section: General information ===

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                              
write-host  -ForegroundColor white		Introduction                                                                                          
write-host  -ForegroundColor white		--------------------------------------------------------------------------------------                                                              
write-host  -ForegroundColor white  	'This option will:  '
write-host  -ForegroundColor white  	'Export information about Litigation Hold | ALL USER Mailboxes '
write-host  -ForegroundColor white		--------------------------------------------------------------------------------------  
write-host  -ForegroundColor white  	'NOTE - The export command will:   '
write-host  -ForegroundColor white  	'1. Create a folder named INFO in drive C: '
write-host  -ForegroundColor white  	'2. Save all of the exported information to diffrent file formats: TXT,CSV and HTML '
write-host  -ForegroundColor white  	'The files will be saved in the follwoing path:' -NoNewline;Write-Host 'C:\INFO\Litigation Hold informaiton\1. ALL USER Mailboxes' -ForegroundColor Yellow  -BackgroundColor DarkGreen							 
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host

# == Section: Export Data to Files ===	

#---------------------------------------------------------------------------------------------------
# Export information to Files
#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
# Create folders Structure that contain the exported information to TXT, CSV and HTML files
#---------------------------------------------------------------------------------------------------


#------------------------------------------------------------------------------
# Folder and path Structure
#------------------------------------------------------------------------------

$A20 =  "C:\INFO\Litigation Hold informaiton"
$A21 =  "$A20\1. ALL USER Mailboxes"


#---------------------------------------------------------------------------------------------------
# Create the required folders in Drive C:\INFO
#---------------------------------------------------------------------------------------------------

# C:\INFO\Litigation Hold informaiton
if (!(Test-Path -path $A20))
{New-Item $A20 -type directory}


# 1. ALL USER Mailboxes
if (!(Test-Path -path $A21))
{New-Item $A21 -type directory}

# == Section: Export data to files  ===

### TXT ####
$global:GetMBXUser  | Select-Object $global:Array1  | Format-List | Out-File $A21\"Litigation Hold - Settings.txt" -Encoding UTF8
##########

### CSV ####
$global:GetMBXUser  | Select-Object $global:Array1  | Export-CSV $A21\"Litigation Hold - Settings.CSV" –NoTypeInformation -Encoding utf8
##########

### HTML ####
$global:GetMBXUser  | Select-Object $global:Array1  | ConvertTo-Html  -post $EndReport -head $Header -Body  "<H1>Litigation Hold - Settings | $Datef </H1>"  | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-File $A21\"Litigation Hold - Settings.html"
##########

	
# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}
			
	



9
{

##########################################################################################################
#  Display + Export information about Litigation Hold | Mailboxes which have a Litigation Hold enabled
##########################################################################################################
checkConnection
# == Section: General information ===

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                              
write-host  -ForegroundColor white		Introduction                                                                                          
write-host  -ForegroundColor white		--------------------------------------------------------------------------------------                                                              
write-host  -ForegroundColor white  	'This option will:  '
write-host  -ForegroundColor white  	'Export information about Litigation Hold | USERS Mailboxes which have a Litigation Hold enabled '
write-host  -ForegroundColor white		--------------------------------------------------------------------------------------  
write-host  -ForegroundColor white  	'NOTE - The export command will:   '
write-host  -ForegroundColor white  	'1. Create a folder named INFO in drive C: '
write-host  -ForegroundColor white  	'2. Save all of the exported information to diffrent file formats: TXT,CSV and HTML '
write-host  -ForegroundColor white  	'The files will be saved in the follwoing path:' -NoNewline;Write-Host 'C:\INFO\Litigation Hold informaiton\2. USER Mailboxes - Litigation Hold enabled' -ForegroundColor Yellow  -BackgroundColor DarkGreen							 
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host

# == Section: Export Data to Files ===	

#---------------------------------------------------------------------------------------------------
# Export information to Files
#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
# Create folders Structure that contain the exported information to TXT, CSV and HTML files
#---------------------------------------------------------------------------------------------------


#------------------------------------------------------------------------------
# Folder and path Structure
#------------------------------------------------------------------------------

$A20 =  "C:\INFO\Litigation Hold informaiton"
$A21 =  "$A20\2. USER Mailboxes - Litigation Hold enabled"


#---------------------------------------------------------------------------------------------------
# Create the required folders in Drive C:\INFO
#---------------------------------------------------------------------------------------------------

# C:\INFO\Litigation Hold informaiton
if (!(Test-Path -path $A20))
{New-Item $A20 -type directory}


# 2. USER Mailboxes - Litigation Hold enabled
if (!(Test-Path -path $A21))
{New-Item $A21 -type directory}

# == Section: Export data to files  ===

### TXT ####
$global:GetMBXUser | Where-Object {$_.LitigationHoldEnabled -eq $True}  | Select-Object $global:Array1  | Format-List | Out-File $A21\"USER Mailboxes - Litigation Hold - Enabled.txt" -Encoding UTF8
##########

### CSV ####
$global:GetMBXUser | Where-Object {$_.LitigationHoldEnabled -eq $True}   | Select-Object $global:Array1  | Export-CSV $A21\"USER Mailboxes - Litigation Hold - Enabled.CSV" –NoTypeInformation -Encoding utf8
##########

### HTML ####
$global:GetMBXUser | Where-Object {$_.LitigationHoldEnabled -eq $True}    | Select-Object $global:Array1  | ConvertTo-Html  -post $EndReport -head $Header -Body  "<H1>USER Mailboxes - Litigation Hold - Enabled | $Datef </H1>"  | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-File $A21\"USER Mailboxes - Litigation Hold - Enabled.html"
##########

	
# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}





10
{

##########################################################################################################
#  Export information about User mailboxes which doesnt have Litigation Hold
##########################################################################################################
checkConnection
# == Section: General information ===

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                              
write-host  -ForegroundColor white		Introduction                                                                                          
write-host  -ForegroundColor white		--------------------------------------------------------------------------------------                                                              
write-host  -ForegroundColor white  	'This option will:  '
write-host  -ForegroundColor white  	'Export information about User mailboxes which doesnt have Litigation Hold '
write-host  -ForegroundColor white		--------------------------------------------------------------------------------------  
write-host  -ForegroundColor white  	'NOTE - The export command will:   '
write-host  -ForegroundColor white  	'1. Create a folder named INFO in drive C: '
write-host  -ForegroundColor white  	'2. Save all of the exported information to diffrent file formats: TXT,CSV and HTML '
write-host  -ForegroundColor white  	'The files will be saved in the follwoing path:' -NoNewline;Write-Host 'C:\INFO\Litigation Hold informaiton\3. USER Mailboxes - NO Litigation Hold' -ForegroundColor Yellow  -BackgroundColor DarkGreen							 
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host

# == Section: Export Data to Files ===	

#---------------------------------------------------------------------------------------------------
# Export information to Files
#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
# Create folders Structure that contain the exported information to TXT, CSV and HTML files
#---------------------------------------------------------------------------------------------------


#------------------------------------------------------------------------------
# Folder and path Structure
#------------------------------------------------------------------------------

$A20 =  "C:\INFO\Litigation Hold informaiton"
$A21 =  "$A20\3. USER Mailboxes - NO Litigation Hold"


#---------------------------------------------------------------------------------------------------
# Create the required folders in Drive C:\INFO
#---------------------------------------------------------------------------------------------------

# C:\INFO\Litigation Hold informaiton
if (!(Test-Path -path $A20))
{New-Item $A20 -type directory}


# 3. USER Mailboxes - NO Litigation Hold
if (!(Test-Path -path $A21))
{New-Item $A21 -type directory}

# == Section: Export data to files  ===

### TXT ####
$global:GetMBXUser | Where-Object {$_.LitigationHoldEnabled -eq $False}  | Select-Object $global:Array1  | Format-List | Out-File $A21\"USER Mailboxes - NO Litigation Hold.txt" -Encoding UTF8
##########

### CSV ####
$global:GetMBXUser | Where-Object {$_.LitigationHoldEnabled -eq $False}  | Select-Object $global:Array1  | Export-CSV $A21\"USER Mailboxes - NO Litigation Hold.CSV" –NoTypeInformation -Encoding utf8
##########

### HTML ####
$global:GetMBXUser | Where-Object {$_.LitigationHoldEnabled -eq $False}   | Select-Object $global:Array1  | ConvertTo-Html  -post $EndReport -head $Header -Body  "<H1>USER Mailboxes - NO Litigation Hold | $Datef </H1>"  | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-File $A21\"USER Mailboxes - NO Litigation Hold.html"
##########

	
# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}
			
	


11
{

##########################################################################################################
#  Export information about Litigation Hold | Information about the recoverable items folder
##########################################################################################################
checkConnection
# == Section: General information ===

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                              
write-host  -ForegroundColor white		Introduction                                                                                          
write-host  -ForegroundColor white		--------------------------------------------------------------------------------------                                                              
write-host  -ForegroundColor white  	'This option will:  '
write-host  -ForegroundColor white  	'Export information about Litigation Hold | Information about the recoverable items folder '
write-host  -ForegroundColor white		--------------------------------------------------------------------------------------  
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Get-Mailbox -ResultSize Unlimited -Filter {LitigationHoldEnabled -eq $true} | Get-MailboxFolderStatistics –FolderScope RecoverableItems | Format-Table Identity,FolderAndSubfolderSize '
write-host  -ForegroundColor white		--------------------------------------------------------------------------------------  
write-host  -ForegroundColor white  	'NOTE - The export command will:   '
write-host  -ForegroundColor white  	'1. Create a folder named INFO in drive C: '
write-host  -ForegroundColor white  	'2. Save all of the exported information to diffrent file formats: TXT,CSV and HTML '
write-host  -ForegroundColor white  	'The files will be saved in the follwoing path:' -NoNewline;Write-Host 'C:\INFO\Litigation Hold informaiton\4. information about the Recoverable items folder' -ForegroundColor Yellow  -BackgroundColor DarkGreen							 
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host

# == Section: Export Data to Files ===	

#---------------------------------------------------------------------------------------------------
# Export information to Files
#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
# Create folders Structure that contain the exported information to TXT, CSV and HTML files
#---------------------------------------------------------------------------------------------------


#------------------------------------------------------------------------------
# Folder and path Structure
#------------------------------------------------------------------------------

$A20 =  "C:\INFO\Litigation Hold informaiton"
$A21 =  "$A20\4. information about the Recoverable items folder"


#---------------------------------------------------------------------------------------------------
# Create the required folders in Drive C:\INFO
#---------------------------------------------------------------------------------------------------

# C:\INFO\Litigation Hold informaiton
if (!(Test-Path -path $A20))
{New-Item $A20 -type directory}


# 3. USER Mailboxes - NO Litigation Hold
if (!(Test-Path -path $A21))
{New-Item $A21 -type directory}

# == Section: Export data to files  ===

### TXT ####
$global:GetMBXUser | Where-Object {$_.LitigationHoldEnabled -eq $True}  | Get-MailboxFolderStatistics –FolderScope RecoverableItems |  Select-Object Identity,FolderAndSubfolderSize | Out-File $A21\"USER Mailboxes - information about the Recoverable items folder.txt" -Encoding UTF8
##########

### CSV ####
$global:GetMBXUser | Where-Object {$_.LitigationHoldEnabled -eq $True}  | Get-MailboxFolderStatistics –FolderScope RecoverableItems |  Select-Object Identity,FolderAndSubfolderSize | Export-CSV $A21\"USER Mailboxes - information about the Recoverable items folder.CSV" –NoTypeInformation -Encoding utf8
##########

### HTML ####
$global:GetMBXUser | Where-Object {$_.LitigationHoldEnabled -eq $True}  | Get-MailboxFolderStatistics –FolderScope RecoverableItems |  Select-Object Identity,FolderAndSubfolderSize | ConvertTo-Html  -post $EndReport -head $Header -Body  "<H1>USER Mailboxes - information about the Recoverable items folder | $Datef </H1>"  | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-File $A21\"USER Mailboxes - information about the Recoverable items folder.html"
##########

	
# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}
			
		




#=========================================================================================
# SECTION C:  Remove Litigation Hold 
#=========================================================================================


12
{


#####################################################################
# Remove Litigation Hold | Single Mailbox
#####################################################################

# == Section: General information === 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		----------------------------------------------------------------------------                                                             
write-host  -ForegroundColor white  	'This option will: '
write-host  -ForegroundColor white  	'Remove Litigation Hold | Single Mailbox. '
write-host  -ForegroundColor white		--------------------------------------------------------------------  
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Set-Mailbox <Mailbox> -LitigationHoldEnabled $False '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host


# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 1 parameter:"  
write-host
write-host -ForegroundColor Yellow	"1.  Mailbox name  "  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Target mailbox"    
write-host -ForegroundColor Yellow	"For example:  Bob@o365info.com"
write-host
$Alias  = Read-Host "Type the mailbox name "
write-host
write-host

# == Section: Display Information  ===	

write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue   Display information about: "$Alias".ToUpper() LitigationHold => BEFORE the update
write-host -------------------------------------------------------------------------------------------------

Get-Mailbox $Alias | Select-Object Name,LitigationHold* |  Out-String  

write-host -------------------------------------------------------------------------------------------------

# == Section: PowerShell Command  ===	
Set-Mailbox $Alias -LitigationHoldEnabled $False



# == Section: Display Information  ===	



write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Magenta  Display information about: "$Alias".ToUpper() LitigationHold  => AFTER the update
write-host -------------------------------------------------------------------------------------------------

Get-Mailbox $Alias | Select-Object Name,LitigationHold* |  Out-String  

write-host -------------------------------------------------------------------------------------------------


# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}



13
{


#####################################################################
# Remove Litigation Hold based on a criterion (Department) | ALL USER Mailboxes (Bulk mode)
#####################################################################

# == Section: General information === 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		----------------------------------------------------------------------------                                                             
write-host  -ForegroundColor white  	'This option will: '
write-host  -ForegroundColor white  	'Remove Litigation Hold based on a criterion (Department) | ALL USER Mailboxes (Bulk mode). '
write-host  -ForegroundColor white		--------------------------------------------------------------------  
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Get-Mailbox -ResultSize Unlimited -Filter {RecipientTypeDetails -eq "UserMailbox"} | Set-Mailbox -LitigationHoldEnabled $False'
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host


# == Section: Display Information   ===	

write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue   Display information about: Litigation Hold Settings => BEFORE the update
write-host -------------------------------------------------------------------------------------------------

$global:GetMBXUser  | Select-Object Name,LitigationHold* |  Out-String  

write-host -------------------------------------------------------------------------------------------------

# == Section: PowerShell Command  ===	
$global:GetMBXUser  | Set-Mailbox -LitigationHoldEnabled $False



# == Section: Display Information   ===	



write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Magenta   Display information about: Litigation Hold Settings  => AFTER the update
write-host -------------------------------------------------------------------------------------------------

$global:GetMBXUser  | Select-Object Name,LitigationHold* |  Out-String  

write-host -------------------------------------------------------------------------------------------------


# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}



		
14
{


###################################################################################################################################
# Remove (Disable) Litigation Hold | ALL USER Mailboxes WITH Litigation Hold enabled (BULK mode)
###################################################################################################################################

checkConnection
# General information 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		----------------------------------------------------------------------------  
write-host  -ForegroundColor white  	"Remove (Disable) Litigation Hold | ALL USER Mailboxes WITH Litigation Hold enabled (BULK mode) "
write-host  -ForegroundColor white		----------------------------------------------------------------------------  
write-host  -ForegroundColor white  	"This menu option, will Remove (Disable) Exchange Litigation Hold for each existing Exchange Online    "
write-host  -ForegroundColor white  	"User mailbox Which have Litigation Hold enabled.    "
write-host  -ForegroundColor white		--------------------------------------------------------------------  
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host

#---------------------------------------------------------------------------------------------------
# Define variables
#---------------------------------------------------------------------------------------------------
# Define date format variable
$Datef = Get-Date -Format "\Da\te dd-MM-yyyy \Ti\me H-mm" 

#---------------------------------------------------------------------------------------------------
# Define variables that contain the folders names 
#---------------------------------------------------------------------------------------------------

# C:\INFO\Litigation Hold informaiton
$A20 =  "C:\INFO\Litigation Hold informaiton - $Datef"
#1. Report - mailboxes which Litigation Hold where Removed (Disabled)
$A21 =  "$A20\1. Report - mailboxes which Litigation Hold where Removed (Disabled)"
# 2. All mailboxes with Litigation Hold
$A22 =  "$A20\2. All mailboxes with Litigation Hold"

#---------------------------------------------------------------------------------------------------
# Create folders Structure that contain the exported information to TXT, CSV and HTML files
#---------------------------------------------------------------------------------------------------

#  C:\INFO\Litigation Hold informaiton -<Date>
if (!(Test-Path -path $A20))
{New-Item $A20 -type directory}
# 1. Report - mailboxes which Litigation Hold where Removed (Disabled)
if (!(Test-Path -path $A21))
{New-Item $A21 -type directory}
# 2. All mailboxes with Litigation Hold
if (!(Test-Path -path $A22))
{New-Item $A22 -type directory}

#---------------------------------------------------------------------------------------------------
# Define various parameter that are related to the Log file
#---------------------------------------------------------------------------------------------------
# Define variables for the LOG file name
$LOGFileName = "Log File.TXT"
$LOGFilePath = "$A21\$LOGFileName"
# Define variables for the LOG file header text 
write-ToLOGFfile $lineH1
Write-ToLOGFfile "Remove (Disable) Litigation Hold for All Exchange Online USER mailboxes with Litigation Hold  | $Datef"
write-ToLOGFfile $lineH1
#---------------------------------------------------------------------------------------------------
# Define the Array 
#---------------------------------------------------------------------------------------------------
# Define the Array of Exchange Online mailboxes | All USER mailboxes
$GetMBXUser  =  Get-MailBox -ResultSize Unlimited -Filter '(RecipientTypeDetails -eq "UserMailbox")' | Sort-Object -Property displayname
$AllMailboxesWithLitigationHold = $GetMBXUser      | Where-Object {$_.LitigationHoldEnabled -eq $True} 
$AllMailboxesWithOUTLitigationHold = $GetMBXUser   | Where-Object {$_.LitigationHoldEnabled -eq $False} | Sort-Object -Property displayname

# Count the number of Exchange Online USER mailboxes
$GetMBXUserCount = ($GetMBXUser).count

# Count the number of Exchange Online USER mailboxes with Litigation Hold Enabled
$AllMailboxesWithLitigationHoldCount = ($AllMailboxesWithLitigationHold).count

# Count the number of Exchange Online USER mailboxes with OUT Litigation Hold 
$AllMailboxesWithOUTLitigationHoldCount = ($AllMailboxesWithOUTLitigationHold).count



# Display information about: Exchange Online USER mailboxes and Litigation Hold
write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor blue  "Display information about: Exchange Online USER mailboxes and Litigation Hold"
write-host -------------------------------------------------------------------------------------------------
write-host -------------------------------------------------------------------------------------------------
write-host "There are: " -nonewline; write-host "$GetMBXUserCount" -ForegroundColor Yellow -BackgroundColor blue -nonewline; write-host " Exchange Online USER mailboxes " 
write-host
write-host "There are: " -nonewline; write-host "$AllMailboxesWithLitigationHoldCount" -ForegroundColor Yellow -BackgroundColor blue -nonewline; write-host " Exchange Online USER mailboxes with Litigation Hold Enabled " 
write-host
write-host "There are: " -nonewline; write-host "$AllMailboxesWithOUTLitigationHoldCount" -ForegroundColor Yellow -BackgroundColor blue -nonewline; write-host " Exchange Online USER mailboxes with OUT Litigation Hold " 
write-host -------------------------------------------------------------------------------------------------
write-host
write-host

#---------------------------------------------------------------------------------------------------
# PowerShell Command
#---------------------------------------------------------------------------------------------------

$AllMailboxes = $AllMailboxesWithLitigationHold

# Document the updates by writing information to LOG file 
write-ToLOGFfile $line1
write-ToLOGFfile "There are: $GetMBXUserCount Exchange Online USER mailboxes " 
write-host
write-ToLOGFfile "There are: $AllMailboxesWithLitigationHoldCount Exchange Online USER mailboxes with Litigation Hold Enabled" 
write-host
write-ToLOGFfile "There are: $AllMailboxesWithOUTLitigationHoldCount Exchange Online USER mailboxes WITH OUT Litigation Hold" 
write-ToLOGFfile $line1


ForEach ($Mailbox in $AllMailboxes)
{

# User \ Mailbox identity | Single member from the Array     

$ID1 = $Mailbox.Displayname

# Enable Litigation Hold on The specified Exchange mailbox (Array member)  
Set-Mailbox $ID1  -LitigationHoldEnabled $False
# Display progress bar information on the PowerShell console
Write-Progress -Activity "Remove (Disable) Litigation Hold on - Exchange $ID1 Mailbox"


# Document the updates by writing information to LOG file 
write-ToLOGFfile $line1
Write-ToLOGFfile "Litigation Hold was Remove (Disable) for - $ID1 mailbox"
write-ToLOGFfile $line1


# Display information about the action that was performed
write-host
write-host -ForegroundColor white		---------------------------------------------------------------------------- 
write-host -ForegroundColor white	"Litigation Hold  was Removed (Disabled) for - " -nonewline; write-host $ID1 -ForegroundColor white -BackgroundColor Darkgreen -nonewline; write-host " mailbox " -ForegroundColor white  
write-host -ForegroundColor white		----------------------------------------------------------------------------  
write-host

}


#---------------------------------------------------------------------------------------------------
# Export Exchange Online Mailboxes Litigation Hold settings to files
#---------------------------------------------------------------------------------------------------

$GetMBXUser =  $null

# Define the Array of Exchange Online mailboxes | All USER mailboxes
$GetMBXUser  =  Get-MailBox -ResultSize Unlimited -Filter '(RecipientTypeDetails -eq "UserMailbox")' | Sort-Object -Property displayname



#---------------------------------------------------------------------------------------------------
# Define variables that contain the files names 
#---------------------------------------------------------------------------------------------------
$FileName1 =   "Exchange Online Mailboxes - Litigation Hold settings"


### TXT ####
$GetMBXUser | Select-Object $global:Array1  | Format-List | Out-File $A22\"$FileName1.txt" -Encoding UTF8
##########

### CSV ####
$GetMBXUser | Select-Object $global:Array1  | Export-CSV $A22\"$FileName1.CSV" –NoTypeInformation -Encoding utf8
##########

### HTML ####
$GetMBXUser | Select-Object $global:Array1  | ConvertTo-Html  -post $EndReport -head $Header -Body  "<H1>$FileName1 | $Datef </H1>"  | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-File $A22\"$FileName1.html"
##########




# START User notification about the location of the exported files
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		----------------------------------------------------------------------------                                                             
write-host  -ForegroundColor white  	"The information about Exchange Online mailboxes that their Litigation Hold was Remove (Disable), was written to a LOG file "
write-host  -ForegroundColor white  	"You can find the LOG file in the following path:  "
write-host  -BackgroundColor DarkGreen 	$A20   
write-host  -ForegroundColor white		--------------------------------------------------------------------  
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo   
# END User notification about the location of the exported files

# Empty the content of the variable

$LOGFilePath =  $null
$Datef =  $null

$AllMailboxes =  $null

$ID1 =  $null
$GetMBXUser =  $null
$AllMailboxes =  $null
$AllMailboxesWithLitigationHold =  $null
$AllMailboxesWithOUTLitigationHold  =  $null


# End the Command
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}





	
#<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
# Section END -Disconnect PowerShell session 
#<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


	
20
{

##########################################
# Disconnect PowerShell session  
##########################################


write-host -ForegroundColor Yellow Choosing this option will Disconnect the current PowerShell session 

Disconnect-ExchangeOnline 

write-host
write-host

#———— Indication ———————

if ($lastexitcode -eq 0)
{
write-host -------------------------------------------------------------
write-host "The command complete successfully !" -ForegroundColor Yellow
write-host "The PowerShell session is disconnected" -ForegroundColor Yellow
write-host -------------------------------------------------------------
}
else

{
write-host "The command Failed :-(" -ForegroundColor red

}

#———— End of Indication ———————


}




21
{

##########################################
# Exit  
##########################################


$Loop = $true
Exit
}

}


}