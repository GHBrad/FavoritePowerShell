#Created by Brad Johnson
#Created 2/20/2019
#Purpose: Update/correct the User Principal Name in Azure manually

#Prerequisites: msoidcli_64.msi (Microsoft Online Services Sign-In Assistant for IT Professionals RTW), MSOnline module (install-module MSOnline), AzureAD (install-module AzureAD)

Install-Module AzureAD -Force

$Cred = $host.ui.PromptForCredential("Enter your email address and domain password")

Connect-MsolService -credential $Cred

Connect-AzureAD -credential $Cred

$ObjectID = (Get-AzureADUser -searchstring (Read-Host "Enter the username")).ObjectID

Set-MsolUserPrincipalName -NewUserPrincipalName (Read-Host "Enter the email address") -ObjectID $ObjectID
