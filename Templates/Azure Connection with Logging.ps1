<#
.SYNOPSIS
Azure Connection & Error Checking with Logging - Template
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

#Azure Connection
If (Get-InstalledModule 'Az'){
    Write-Host -ForegroundColor DarkGreen "Az PS Module is already installed, connecting to Azure."
    try {
        #Attempt to connect to Azure if PS module is already installed.
        Import-Module -Name 'Az'
        Connect-AzAccount -Credential $M365Credentials
        Write-Host -ForegroundColor DarkGreen "Connected to Azure."
        Write-LogMessage("Connected to Azure.")
}
    catch {
        #If you fail to connect to Azure write to the log and exit script.
        Write-Host -ForegroundColor DarkRed "Failed to connect to Azure, please check your credentials and try again."
        Write-LogMessage("Failed to connect to Azure, please check your credentials and try again.")
        Exit
}
    } else {
        try{
            #If Az PS Module is not installed, try to install it.
            Write-Host -ForegroundColor DarkRed "Az PS Module is not installed, installing Az module."
            Install-Module -Name 'Az' -Force
            Write-Host -ForegroundColor DarkGreen "Installed Az Powershell Module."
            Write-LogMessage("Installed Az Powershell Module.")
            try {
                #If Az PS Module installs from the try, attempt to connect to Azure.
                Import-Module -Name 'Az'
                Connect-AzAccount -Credential $M365Credentials
                Write-Host -ForegroundColor DarkGreen "Connected to Azure."
                Write-LogMessage("Connected to Azure.")
            }
            catch {
                #If you fail to connect to Azure write to the log and exit script.
                Write-Host -ForegroundColor DarkRed "Failed to connect to Azure, please check your credentials and try again."
                Write-LogMessage("Failed to connect to Azure, please check your credentials and try again.")
                Exit
            }
        }
        catch{
            #If you fail to install Az PS Module write to log and exit script.
            Write-Host -ForegroundColor DarkRed "Failed to install Az Powershell Module, please attempt a manual installation and try running the script again."
            Write-LogMessage("Failed to install Az Powershell Module, please attempt a manual installation and try running the script again.")
            Exit
        }
    }
