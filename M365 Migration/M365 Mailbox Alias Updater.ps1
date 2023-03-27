<#
.SYNOPSIS
M365 Mailbox Alias Addtion tool
Version 1.0
Created by Nick Brizzolara
Created on 3/27/2023

.DESCRIPTION
This script is used to add additional aliases to mailboxes post M365 migration cutover.

You must provide a CSV containting the following schema fields: Identity,Aliases

!!Identity must be the FQDN of the obeject "test@domain.com"!!
!!Aliases must be formatted as such: Alias1@domain.com,Alias2.domain.com,Alias3@domain.com!!

This script uses the following powershell modules: ExchangeOnlineManagement

This script assumes you have a C:\Temp directory, if you do not either create one or repoint the logging variables.

#>
#Static Variables

#This will only work if the login is NOT using MFA, if it is then remove the credential switch on connect-exchangeonline.
$M365Credentials = Get-Credential
#Make sure to update the $CSV variable to the correct path, may be better to just read-host for a path in the future.
$CSV = Import-Csv "C:\temp\aliastest.csv"
#Log file location setup | possibly read-host for a log path in the future?
$dateandtime = Get-Date -Format "yyyy_MM_dd_HH_mm"
$logfilelocation = "C:\Temp\M365_Alias_Adder_" + $dateandtime + '.log'
$logfilelocationfailures = "C:\Temp\M365_Alias_Adder_failures_" + $dateandtime + '.log'

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

    foreach ($mailbox in $csv) 
    {
            try{
            #Splits the aliases into seperate objects using ',' as the split character.
            $formattedaliases = $mailbox.aliases -split ','
            Write-Host "Adding aliases for mailbox $($mailbox.identity)"
            #Adds the formatted aliases to the current mailbox in the loop.
            Set-mailbox -Identity $mailbox.identity -EmailAddresses @{add=$formattedaliases} -ErrorAction Stop
            $messagesmtp = "Adding aliases for mailbox $($mailbox.identity)"
            Write-LogMessage($messagesmtp)
            
        }
            catch{
            #Log failure operation to log file & continue to UPN update section.
            $FailureMessage = "There was an issue while updating the aliases for  $($mailbox.identity), please manually modify the SMTP settings of this user"
            Write-Warning $FailureMessage
            Write-LogMessageFailure ($FailureMessage)
        }

    }
    