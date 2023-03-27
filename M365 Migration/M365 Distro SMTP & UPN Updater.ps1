<#
.SYNOPSIS
Distribution Group Primary SMTP Address Updater
Version 1.0
Created by Nick Brizzolara
Created on 3/27/2023


.DESCRIPTION
This script is used to update the Primary SMTP addresses & UPN's of distribution groups post M365 migration.

You must provide a CSV containting the following schema fields: oldidentity,newidentity,oldemail,newemail

This script uses the following powershell modules: ExchangeOnlineManagement

This script assumes you have a C:\Temp directory, if you do not either create one or repoint the logging variables.

#>

#Static Variables
#This will only work if the login is NOT using MFA, if it is then remove the credential switch on connect-exchangeonline.
$M365Credentials = Get-Credential
#Make sure to update the $CSV variable to the correct path, may be better to just read-host for a path in the future.
$CSV = Import-Csv "C:\temp\test-smtp.csv"

#Log file location setup | possibly read-host for a log path in the future?
$dateandtime = Get-Date -Format "yyyy_MM_dd_HH_mm"
$logfilelocation = "C:\Temp\Distro_SMTP_Updater_" + $dateandtime + '.log'
$logfilelocationfailures = "C:\Temp\Distro_SMTP_Updater_failures_" + $dateandtime + '.log'

#Logging Function
Function Write-LogMessage ([string]$message) {
    $messageToLog = (Get-Date -Format "yyyyMMdd-HH_mmss") + ": " + $message
    $messageToLog | Out-File -Append -FilePath $logFileLocation
}

#Logging Function for failures
Function Write-LogMessageFailure ([string]$message) {
    $messageToLog = (Get-Date -Format "yyyyMMdd-HH_mmss") + ": " + $message
    $messageToLog | Out-File -Append -FilePath $logfilelocationfailures
}

#Exchange Online Connection
If (Get-InstalledModule 'ExchangeOnlineManagement'){
    Write-Host -ForegroundColor DarkGreen "Exchange Online Management PS Module is already installed, connecting to Exchange Online."
    try {
        #Attempt to connect to Exchange Online if PS module is already installed.
        Import-Module -Name 'ExchangeOnlineManagement'
        Connect-ExchangeOnline -Credential $M365Credentials
        Write-Host -ForegroundColor DarkGreen "Connected to Exchange Online."
        Write-LogMessage("Connected to Exchange Online.")
}
    catch {
        #If you fail to connect to Exchange Online write to the log and exit script.
        Write-Host -ForegroundColor DarkRed "Failed to connect to Exchange Online, please check your credentials and try again."
        Write-LogMessage("Failed to connect to Exchange Online, please check your credentials and try again.")
        Exit
}
    } else {
        try{
            #If ExchangeOnlineManagement PS Module is not installed, try to install it.
            Write-Host -ForegroundColor DarkRed "Exchange Online Management PS Module is not installed, installing Exchange Online Management module."
            Install-Module -Name 'ExchangeOnlineManagement' -Force
            Write-Host -ForegroundColor DarkGreen "Installed ExchangeOnlineManagement Powershell Module."
            Write-LogMessage("Installed ExchangeOnlineManagement Powershell Module.")
            try {
                #If ExchangeOnlineManagement PS Module installs from the try, attempt to connect to Exchange Online.
                Import-Module -Name 'ExchangeOnlineManagement'
                Connect-ExchangeOnline -Credential $M365Credentials
                Write-Host -ForegroundColor DarkGreen "Connected to Exchange Online."
                Write-LogMessage("Connected to Exchange Online.")
            }
            catch {
                #If you fail to connect to Exchange Online write to the log and exit script.
                Write-Host -ForegroundColor DarkRed "Failed to connect to Exchange Online, please check your credentials and try again."
                Write-LogMessage("Failed to connect to Exchange Online, please check your credentials and try again.")
                Exit
            }
        }
        catch{
            #If you fail to install ExchangeOnlineManagement PS Module write to log and exit script.
            Write-Host -ForegroundColor DarkRed "Failed to install ExchangeOnlineManagement Powershell Module, please attempt a manual installation and try running the script again."
            Write-LogMessage("Failed to install ExchangeOnlineManagement Powershell Module, please attempt a manual installation and try running the script again.")
            Exit
        }
    }

    foreach ($distro in $csv) 
    {
            try{
            #Set PrimarySMTPAddress for the current object in loop & log operation to log file.
            Write-Host "Setting primary smtp address for" $distro.oldidentity "from" $distro.OldEmail "to" $distro.newemail
            Set-DistributionGroup -Identity $distro.oldidentity -PrimarySmtpAddress $distro.NewEmail -ErrorAction Stop
            $messagesmtp = "Modifying primary SMTP address for $($distro.oldidentity) and setting it to $($distro.newemail)"
            Write-LogMessage($messagesmtp)
            
        }
            catch{
            #Log failure operation to log file & continue to UPN update section.
            $FailureMessage = "There was an issue while upadting the Primary SMTP address of $($distro.oldidentity), please manually modify the SMTP settings of this user"
            Write-Warning $FailureMessage
            Write-LogMessageFailure ($FailureMessage)
        }

    }
    