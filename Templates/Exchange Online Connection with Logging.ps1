<#
.SYNOPSIS
Exchange Online Connection & Error Checking with Logging - Template
Created by Nick Brizzolara
Created on 12/27/2022
#>

#Static Variables
$M365Credentials = Get-Credential

#Log file location setup
$dateandtime = Get-Date -Format "yyyy_MM_dd_HH_mm"
$logfilelocation = "C:\Temp\Template_" + $dateandtime + '.log'

#Logging Function
Function Write-LogMessage ([string]$message) {
    $messageToLog = (Get-Date -Format "yyyyMMdd-HH_mmss") + ": " + $message
    $messageToLog | Out-File -Append -FilePath $logFileLocation
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
