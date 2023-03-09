<#
.SYNOPSIS
Microsoft Graph Connection & Error Checking with Logging - Template
Created by Nick Brizzolara
Created on 3/8/2023
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

function Connect-MgGraphViaCred {
    <#
    .SYNOPSIS
    Provided by: Provided by: https://doitpsway.com/how-to-connect-to-the-microsoft-graph-api-using-saved-user-credentials
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
    