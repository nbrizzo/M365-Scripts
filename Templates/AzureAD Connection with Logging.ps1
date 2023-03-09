<#
.SYNOPSIS
AzureAD Connection & Error Checking with Logging - Template
Created by Nick Brizzolara
Created on 12/27/2022

!! This Powershell Module is slated for depreciation by Microsoft, consider using the "Microsoft.Graph" Module for future scripts. !!
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


#AzureAD Connection
    If (Get-InstalledModule 'AzureAD'){
        Write-Host -ForegroundColor DarkGreen "Azure Active Directory PS Module is already installed, connecting to AzureAD."
        try {
            #Attempt to connect to AzureAD if PS module is already installed.
            Import-Module -Name 'AzureAD'
            Connect-AzureAD -Credential $M365Credentials
            Write-Host -ForegroundColor DarkGreen "Connected to AzureAD."
            Write-LogMessage("Connected to AzureAD.")
    }
        catch {
            #If you fail to connect to AzureAD write to the log and exit script.
            Write-Host -ForegroundColor DarkRed "Failed to connect to AzureAD, please check your credentials and try again."
            Write-LogMessage("Failed to connect to AzureAD, please check your credentials and try again.")
            Exit
    }
        } else {
            try{
                #If AzureAD PS Module is not installed, try to install it.
                Write-Host -ForegroundColor DarkRed "Azure Active Directory PS Module is not installed, installing AzureAD module."
                Install-Module -Name 'AzureAD' -Force
                Write-Host -ForegroundColor DarkGreen "Installed AzureAD Powershell Module."
                Write-LogMessage("Installed AzureAD Powershell Module.")
                try {
                    #If AzureAD PS Module installs from the try, attempt to connect to AzureAD.
                    Import-Module -Name 'AzureAD'
                    Connect-AzureAD -Credential $M365Credentials
                    Write-Host -ForegroundColor DarkGreen "Connected to AzureAD."
                    Write-LogMessage("Connected to AzureAD.")
                }
                catch {
                    #If you fail to connect to AzureAD write to the log and exit script.
                    Write-Host -ForegroundColor DarkRed "Failed to connect to AzureAD, please check your credentials and try again."
                    Write-LogMessage("Failed to connect to AzureAD, please check your credentials and try again.")
                    Exit
                }
            }
            catch{
                #If you fail to install AzureAD PS Module write to log and exit script.
                Write-Host -ForegroundColor DarkRed "Failed to install AzureAD Powershell Module, please attempt a manual installation and try running the script again."
                Write-LogMessage("Failed to install AzureAD Powershell Module, please attempt a manual installation and try running the script again.")
                Exit
            }
        }
    
