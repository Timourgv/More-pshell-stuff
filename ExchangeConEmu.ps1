# ==============================================================================
# Connect to Exchange Online with MFA, using the modern ExchangeOnlineManagement
# module. Replaces the retired ClickOnce "Exchange Online PowerShell Module" +
# New-ExoPSSession approach this script originally used.
# ==============================================================================

#Requires -Version 5.1

# Clean up any leftover EXO sessions from previous runs
Get-PSSession | Where-Object { $_.ComputerName -eq 'outlook.office365.com' } | Remove-PSSession

# Make sure the module is available; install it for the current user if it's missing
if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) {
    Write-Host "ExchangeOnlineManagement module not found, installing for current user..." -ForegroundColor Yellow
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
}

Import-Module ExchangeOnlineManagement

# Opens the MFA sign-in prompt and connects
Connect-ExchangeOnline -ShowBanner:$false

Write-Host "Connected to Exchange Online." -ForegroundColor Green

# When you're done:  Disconnect-ExchangeOnline
