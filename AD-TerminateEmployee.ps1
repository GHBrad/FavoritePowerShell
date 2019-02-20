#11/29/2018 9:00 AM - Brad Johnson
#This script will prompt for the user to be terminated, gather and export their AD group memberships to a network location, remove membership of those groups, move the user to 
#the Terminated Employees OU and export their email to a network share.

#Prompt the user to run this script as their domain admin account
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Make sure to launch PowerShell as your domain admin account in order to run this script",0,"ATTENTION",0x1)

Import-Module activedirectory

#Create Exchange PowerShell session
$UserCredential = Get-Credential 
$ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<EXCHANGE SERVER URL>/PowerShell -Authentication Kerberos -Credential $UserCredential
Import-PSSession $ExchSession -DisableNameChecking -AllowClobber | Out-Null

#Prompt for user to be terminated
$TerminatedUser = Read-Host -Prompt 'Input the username of the employee being terminated. Example: bsaget'

#Export the email to a network share then move the export to their H: drive
Get-MailboxExportRequest | Where-Object Status -eq "Completed" | Remove-MailboxExportRequest -Confirm:$False
$RequestID = (New-MailboxExportRequest -Mailbox $TerminatedUser -FilePath "\\<NETWORK SHARE>\$TerminatedUser.pst").RequestGuid.Guid

$strStatus = (Get-MailboxExportRequest -Identity $RequestID).Status
DO {
    $strStatus = (Get-MailboxExportRequest -Identity $RequestID).Status
    Start-Sleep 10
    Clear-Host
    $strCurrDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "Status as of ${strCurrDateTime}: ${strStatus}"
} UNTIL ($strStatus -eq "Completed" -or $strStatus -eq "Failed")

IF ($StrStatus -eq "Completed") {
    New-Item -ItemType Directory -Path "\\<NETWORK SHARE>\$TerminatedUser" -Name "Mailbox Export" -Force
    Move-Item "\\<NETWORK SHARE>\$TerminatedUser.pst" -Destination "\\<NETWORK SHARE>\$TerminatedUser\Mailbox Export\" -Force
}

IF ($StrStatus -eq "Failed") {
    Write-Host "Mailbox export failed. You can try running this script again or investigate."
}

#Gather list of AD groups the user is a member of
$ADGroups = Get-ADPrincipalGroupMembership $TerminatedUser | where {$_.Name -ne "Domain Users"} 

#Remove from AD groups and export groups to txt file
$ADGroups | Export-Csv "\\<NETWORK SHARE\TerminatedEmployeeGroupMembership\'$TerminatedUser.txt"
Remove-ADPrincipalGroupMembership -Identity $TerminatedUser -MemberOf $ADGroups -Confirm:$false

#Move user to Terminated Employees OU
Get-ADUser $TerminatedUser | Move-ADObject -TargetPath "OU=Terminated Employees,<OU PATH>"

#Move home folder
Move-Item "\\<NETWORK SHARE>\$TerminatedUser" -Destination "\\<NETWORK SHARE>\" -Force

#Ends the Exchange session
Remove-PSSession -Session $ExchSession
