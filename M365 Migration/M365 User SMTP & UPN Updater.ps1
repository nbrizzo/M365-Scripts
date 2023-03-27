<#
.SYNOPSIS
Primary SMTP Address & UPN Updater
Version 2.0
Created by Nick Brizzolara
Created on 3/23/2023
Last updated: 3/25/2023

.DESCRIPTION
This script is used to update the Primary SMTP addresses & UPN's of user & shared mailboxes post M365 migration.

You must provide a CSV containting the following schema fields: oldidentity,newidentity,oldemail,newemail

This script uses the following powershell modules: Microsoft.Graph, Az, ExchangeOnlineManagement

This script assumes you have a C:\Temp directory, if you do not either create one or repoint the logging variables.

#>

#Static Variables
$M365Credentials = Get-Credential
#Make sure to update the $CSV variable to the correct path, may be better to just read-host for a path in the future.
$CSV = Import-Csv "C:\temp\test-smtp.csv"

#Log file location setup | possibly read-host for a log path in the future?
$dateandtime = Get-Date -Format "yyyy_MM_dd_HH_mm"
$logfilelocation = "C:\Temp\SMTP_and_UPN_Rename_" + $dateandtime + '.log'
$logfilelocationfailures = "C:\Temp\SMTP_and_UPN_Rename_failures_" + $dateandtime + '.log'

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

#Function to connect to Microsoft.Graph with user credentials as a switch.
function Connect-MgGraphViaCred {
    <#
    .SYNOPSIS
    Provided by: https://doitpsway.com/how-to-connect-to-the-microsoft-graph-api-using-saved-user-credentials
    Function for connecting to the Microsoft Graph using given credentials.
    This option is unavailable with official Connect-MgGraph command.

    .DESCRIPTION
    Function for connecting to the Microsoft Graph using given credentials.
    This option is unavailable with official Connect-MgGraph command.

    .PARAMETER credential
    Credential object.

    .PARAMETER tenant
    (optional) Azure tenant name or id.

    .EXAMPLE
    $cred = Get-Credential
    Connect-MgGraphViaCred -credential $cred
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential] $credential,

        [string] $tenant = $_tenantDomain
    )

    # connect to Azure using credentials
    $param = @{
        Credential = $credential
        Force      = $true
    }
    if ($tenant) { $param.tenant = $tenant }
    $null = Connect-AzAccount @param

    # retrieve token for MSGraph
    $token = (Get-AzAccessToken -ResourceTypeName MSGraph -ErrorAction Stop).token

    # convert token string to securestring if new version of Connect-MgGraph is used
    if ((Get-Help Connect-MgGraph -Parameter accesstoken).type.name -eq "securestring") {
        $token = ConvertTo-SecureString $token -AsPlainText -Force
    }
    # use token for connecting to Microsoft Graph
    $null = Connect-MgGraph -AccessToken $token -ErrorAction Stop
}

#Microsoft Graph Connection
    If (Get-InstalledModule 'Microsoft.Graph'){
        Write-Host -ForegroundColor DarkGreen "Microsoft.Graph Module already installed, connecting to Microsoft Graph."
        try {
            #Attempt to connect to Microsoft.Graph if PS module is already installed.
            Import-Module -Name 'Microsoft.Graph.Authentication'
            Connect-MgGraphViaCred -Credential $M365Credentials
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
                    Connect-MgGraphViaCred -Credential $M365Credentials
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



foreach ($user in $csv) 
{
        try{
        #Set PrimarySMTPAddress for the current object in loop & log operation to log file.
        Write-Host "Setting primary smtp address for" $user.oldidentity "from" $user.OldEmail "to" $user.NewEmail
        Set-Mailbox -identity $user.oldidentity -WindowsEmailAddress $user.NewEmail -ErrorAction Stop
        $messagesmtp = "Modifying primary SMTP address for $($user.oldidentity) and setting it to $($user.NewEmail)"
        Write-LogMessage($messagesmtp)
        
    }
        catch{
        #Log failure operation to log file & continue to UPN update section.
        $FailureMessage = "There was an issue while upadting the Primary SMTP address of $($user.oldidentity), please manually modify the SMTP settings of this user"
        Write-Warning $FailureMessage
        Write-LogMessageFailure ($FailureMessage)
    }
        try{
        #Set UPN for the current object in loop & log operation to log file.
        Write-Host "Setting UPN for" $user.oldidentity "to" $user.newidentity
        Update-MGUser -UserID $user.oldidentity -userprincipalname $user.newidentity -ErrorAction Stop
        $messageupn = "Modifying primary UPN for $($user.oldidentity) and setting it to $($user.newidentity)"
        Write-LogMessage($messageupn)

    }
        catch{
        #Log failure operation & continue to next object in loop
        $FailureMessage = "There was an issue while updating the UPN of $($user.oldidentity), please manually modify the UPN of this user"
        Write-Warning $FailureMessage
        Write-LogMessageFailure ($FailureMessage)
        }
}


#Simple CSV loop if this script is broken
#$CSV| Where-Object {
#
#   Write-Host "Setting primary smtp address for" $_.oldidentity "from" $_.OldEmail "to" $_.NewEmail
#    Set-Mailbox -identity $_.oldidentity -WindowsEmailAddress $_.NewEmail
#    $messagesmtp = "Modifying primary SMTP address for $($_.oldidentity) and setting it to $($_.NewEmail)"
#    Write-LogMessage($messagesmtp)
#    
#    Write-Host "Setting UPN for" $_.oldidentity "to" $_.newidentity
#    Update-MGUser -UserID $_.oldidentity -userprincipalname $_.newidentity
#    $messageupn = "Modifying primary UPN for $($_.oldidentity) and setting it to $($_.newidentity)"
#    Write-LogMessage($messageupn)
#    
#    }