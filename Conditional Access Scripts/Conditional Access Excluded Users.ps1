<#
.SYNOPSIS
Conditional Access Excluded User List Retrival
Created by Nick Brizzolara
Created on 12/19/2024

.DESCRIPTION
Used to generate a list of by-name excluded users from conditional access policies. 
This script is only needed when nested groups are not used to handle conditional access exclusions.

#>

Connect-MGGraph -Scopes "Policy.Read.All", "User.Read.All"

#Define Conditional Access Policy ID
#If unsure just run Get-MGConditionalAccessPolicy by itself and find the ID of the policy you require, place it here.
$CAPolicyID = "Policy ID"

#Define Static path for file output
$OutFilePath = "C:\Temp\ExcludedMFAUsers.CSV"

#Set variable to the exact conditional access policy we will be grabbing the list of exluded users from.
$CAPolicy = Get-MGConditionalAccessPolicy -ConditionalAccessPolicyID $CAPolicyID

#Generate a variable that holds the ID of all excluded users.
$ExcludedUsers = $CAPolicy.conditions.users.excludeusers

#Create an empty array to store all user details inside, this is required and script will fail without a defined and empty array.
$UserDetails = @()

#Run a loop through all of the user ID's stored in $ExcludedUsers, populate the empty array with the user detail information.
foreach ($userid in $ExcludedUsers) {


$user = get-mguser -userid $userid

#Add specific details per user into the array.
$UserDetails += [PSCustomObject]@{
    DisplayName = $User.DisplayName
    UserPrincipalName = $User.UserPrincipalName
    Id = $User.id
}
}

$UserDetails | Export-CSV -Path $OutFilePath
