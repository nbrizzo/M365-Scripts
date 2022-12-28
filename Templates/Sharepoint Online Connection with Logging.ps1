<#
.SYNOPSIS
Sharepoint Online Connection & Error Checking with Logging - Template
Created by Nick Brizzolara
Created on 12/27/2022

!! If using Powershell Core update the "Import-Module" to use the "-UseWindowsPowershell" switch or this will not work properly !!
!! Check Powershell edition by running the following command: "$psversiontable.psedition" !!
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
        
#Sharepoint Connection
        If (Get-InstalledModule 'Microsoft.Online.SharePoint.PowerShell'){
            Write-Host -ForegroundColor DarkGreen "Microsoft.Online.SharePoint.PowerShell PS Module is already installed, connecting to Sharepoint Online."
            try {
                #Attempt to connect to Sharepoint Online if PS module is already installed.
                #If you get an error "Connect-SPOService: No valid OAuth 2.0 authentication session exists" you are likely using Powershell 7, use the -UseWindowsPowershell switch, PS7 still doesn't import this module correctly.
                Import-Module -Name 'Microsoft.Online.SharePoint.PowerShell' 
                #Import-Module -Name 'Microsoft.Online.SharePoint.PowerShell' -UseWindowsPowershell
                Connect-AzAccount -Credential $M365Credentials
                Write-Host -ForegroundColor DarkGreen "Connected to Sharepoint Online."
                Write-LogMessage("Connected to Sharepoint Online.")
        }
            catch {
                #If you fail to connect to Exchange Online write to the log and exit script.
                Write-Host -ForegroundColor DarkRed "Failed to connect to Sharepoint Online, please check your credentials and verify what version of Powershell you are using. If you are using Powershell 7 please review the comments within the script in the Sharepoint Connection section."
                Write-LogMessage("Failed to connect to Sharepoint Online, please check your credentials and verify what version of Powershell you are using. If you are using Powershell 7 please review the comments within the script in the Sharepoint Connection section.")
                Exit
        }
            } else {
                try{
                    #If Microsoft.Online.SharePoint.PowerShell PS Module is not installed, try to install it.
                    Write-Host -ForegroundColor DarkRed "Microsoft.Online.SharePoint.PowerShell PS Module is not installed, installing Microsoft.Online.SharePoint.PowerShell module."
                    Install-Module -Name 'Microsoft.Online.SharePoint.PowerShell' -Force
                    Write-Host -ForegroundColor DarkGreen "Installed Microsoft.Online.SharePoint.PowerShell Powershell Module."
                    Write-LogMessage("Installed Microsoft.Online.SharePoint.PowerShell Powershell Module.")
                    try {
                        #If Microsoft.Online.SharePoint.PowerShell PS Module installs from the try, attempt to connect to Sharepoint Online.
                        Import-Module -Name 'Microsoft.Online.SharePoint.PowerShell'
                        Connect-SPOService -Url $SPOURL -Credential $M365Credentials
                        Write-Host -ForegroundColor DarkGreen "Connected to Sharepoint Online."
                        Write-LogMessage("Connected to Sharepoint Online.")
                    }
                    catch {
                        #If you fail to connect to Sharepoint Online write to the log and exit script.
                        Write-Host -ForegroundColor DarkRed "Failed to connect to Sharepoint Online, please check your credentials and verify what version of Powershell you are using. If you are using Powershell 7 please review the comments within the script in the Sharepoint Connection section."
                        Write-LogMessage("Failed to connect to Sharepoint Online, please check your credentials and verify what version of Powershell you are using. If you are using Powershell 7 please review the comments within the script in the Sharepoint Connection section.")
                        Exit
                    }
                }
                catch{
                    #If you fail to install AzureAD PS Module write to log and exit script.
                    Write-Host -ForegroundColor DarkRed "Failed to install Microsoft.Online.SharePoint.PowerShell Powershell Module, please attempt a manual installation and try running the script again."
                    Write-LogMessage("Failed to install Microsoft.Online.SharePoint.PowerShell Powershell Module, please attempt a manual installation and try running the script again.")
                    Exit
                }
            }
        