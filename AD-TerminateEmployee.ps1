#11/29/2018 9:00 AM - Brad Johnson
#This script will prompt for the user to be terminated, gather and export their AD group memberships to a network location, remove membership of those groups, move the user to 
#the Terminated Employees OU and export their email to a network share.

Import-Module activedirectory

#Create Exchange PowerShell session
$UserCredential = Get-Credential 
$ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<EXCHANGE SERVER>/PowerShell -Authentication Kerberos -Credential $UserCredential
Import-PSSession $ExchSession -DisableNameChecking -AllowClobber | Out-Null

#Prompt for user to be terminated
$TerminateUser = Read-Host -Prompt 'Input the username of the employee being terminated. Example: bsaget'

#Export the email to a network share
New-MailboxExportRequest -Mailbox $TerminateUser -FilePath "\\<NETWORK SHARE>\exported email\'$TerminateUser.pst"

#Gather list of AD groups the user is a member of
$ADGroups = Get-ADPrincipalGroupMembership $TerminateUser | where {$_.Name -ne "Domain Users"} 

#Remove from AD groups and export groups to txt file
$ADGroups | Export-Csv "\\<NETWORK SHARE>\TerminatedEmployeeGroupMembership\'$TerminateUser.txt"
Remove-ADPrincipalGroupMembership -Identity $TerminateUser -MemberOf $ADGroups -Confirm:$false

#Move user to Terminated Employees OU
Get-ADUser $TerminateUser | Move-ADObject -TargetPath "<TARGET OU>"

#Ends the Exchange session
Remove-PSSession -Session $ExchSession