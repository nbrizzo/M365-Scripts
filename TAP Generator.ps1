<#
.SYNOPSIS
Temporary Access Pass Generator
Created by Nick Brizzolara
Created on 5/15/2023

.DESCRIPTION
Used to create temporary access passes for users listed in a CSV.


#>

#Static Variables
$Userlist = Import-CSV ""
$TAPExportList = "C:\Temp\TAPExport.csv"

#Log file location setup
$dateandtime = Get-Date -Format "yyyy_MM_dd_HH_mm"
$logfilelocation = "C:\Temp\Template_" + $dateandtime + '.log'

#Logging Function
Function Write-LogMessage ([string]$message) {
    $messageToLog = (Get-Date -Format "yyyyMMdd-HH_mmss") + ": " + $message
    $messageToLog | Out-File -Append -FilePath $logFileLocation
}



#Microsoft Graph Connection
    If (Get-InstalledModule 'Microsoft.Graph'){
        Write-Host -ForegroundColor DarkGreen "Microsoft.Graph Module already installed, connecting to Microsoft Graph."
        try {
            #Attempt to connect to Microsoft.Graph if PS module is already installed.
            Import-Module -Name 'Microsoft.Graph.Authentication'
            Connect-MgGraph -Scopes "User.ReadWrite.All, Policy.Read.All, Policy.ReadWrite.AuthenticationMethod, UserAuthenticationMethod.ReadWrite.All"
            Select-mgProfile -Name 'Beta'
            Write-Host -ForegroundColor DarkGreen "Connected to Microsoft Graph."
            Write-LogMessage("Connected to Microsoft Graph.")
    }
        catch {
            #If you fail to connect to Microsoft Graph write to the log and exit script.
            Write-Host -ForegroundColor DarkRed "Failed to connect to Microsoft Graph, please check your credentials and try again."
            Write-LogMessage("Failed to connect to Microsoft Graph, please check your credentials and try again.")
            Exit
    }
        } else {
            try{
                #If Microsoft Graph PS Module is not installed, try to install it.
                Write-Host -ForegroundColor DarkRed "Microsoft Graph PS Module is not installed, installing Microsoft.Graph module."
                Install-Module -Name 'Microsoft.Graph' -Force
                Write-Host -ForegroundColor DarkGreen "Installed Microsoft.Graph Powershell Module."
                Write-LogMessage("Installed Microsoft.Graph Powershell Module.")
                try {
                    #If Microsoft.Graph PS Module installs from the try, attempt to connect to AzureAD.
                    Import-Module -Name 'Microsoft.Graph.Authentication'
                    Connect-MgGraph -Scopes "User.ReadWrite.All, Policy.Read.All, Policy.ReadWrite.AuthenticationMethod, UserAuthenticationMethod.ReadWrite.All"
                    Select-mgProfile -Name 'Beta'
                    Write-Host -ForegroundColor DarkGreen "Connected to Microsoft Graph."
                    Write-LogMessage("Connected to Microsoft Graph.")
                }
                catch {
                    #If you fail to connect to Microsoft Graph write to the log and exit script.
                    Write-Host -ForegroundColor DarkRed "Failed to connect to Microsoft Graph, please check your credentials and try again."
                    Write-LogMessage("Failed to connect to Microsoft Graph, please check your credentials and try again.")
                    Exit
                }
            }
            catch{
                #If you fail to install Microsoft.Graph PS Module write to log and exit script.
                Write-Host -ForegroundColor DarkRed "Failed to install Microsoft.Graph Powershell Module, please attempt a manual installation and try running the script again."
                Write-LogMessage("Failed to install Microsoft.Graph Powershell Module, please attempt a manual installation and try running the script again.")
                Exit
            }
        }


#Loop through the provided CSV of usernames and generate a TAP per user.
ForEach ($user in $Userlist) {

$CurrentTAP = New-MgUserAuthenticationTemporaryAccessPassMethod -Userid $user.username -IsUsable -lifetimeinminutes 480 -erroraction SilentlyContinue | Select-Object LifeTimeinMinutes,TemporaryAccessPass
write-host "Creating a temporary access pass for $($user.username)"

#Creates a custom object array and adds username, keylife and TAP to custom array object
$TAPList += @(
    [PSCustomObject]@{
        Username = $user.username
        KeyLifeTime = $CurrentTAP.LifeTimeinMinutes
        TAP = $CurrentTap.TemporaryAccessPass
    }
)

}

#Export list to CSV
$TAPList | Export-Csv $TAPExportList
