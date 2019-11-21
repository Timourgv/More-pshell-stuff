#Script to connect to Exchange Online with MFA
#The Exchange Online PowerShell Module needs to be installed from the Exchange Admin Center > Hybrid

#Cleanup any old sessions
get-pssession | Where-Object {$_.ComputerName -eq "outlook.office365.com" -and $_.State -eq "Broken"} | remove-pssession

#Check if the Exchange Online PowerShell Module is installed
$EXO = Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse
if (-Not $EXO) {
    Write-warning "Exchange Online Module is not installed, exiting."
}
else {
    Write-Host "Exchange Online Module is installed"



if (Get-Module -Name "Microsoft.Exchange.Management.ExoPowershellModule") {
    Write-Host "Exchange Online Module already loaded"
}
else {
    Write-Host "Exchange PowerShell Module is not loaded, loading..."

    Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse).FullName |?{$_ -notmatch "_none_"} | select -First 1)
    $EXOSession = New-ExoPSSession
    Import-PSSession $EXOSession
}

if (get-pssession | Where-Object {$_.ComputerName -eq "outlook.office365.com" -and $_.State -eq "Opened"}) {
    Write-Host "Exchange Online PowerShell session already exists"
}
else {
Write-host "Creating new Exchange Online PowerShell session"
$EXOSession = New-ExoPSSession
Import-PSSession $EXOSession -AllowClobber
}

}