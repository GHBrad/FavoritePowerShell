#Created by Brad Johnson
#Date created 12/3/2018
#Purpose: create an Active Directory and associated Exchange account for <DEPARTMENT>.

#ATTENTION! Before running this script, make sure to update NewADUsers.csv in \\<NETWORK SHARE>

Import-Module activedirectory

#Instruct user to update the source spreadsheet
Write-Host -nonewline "ATTENTION! Before running this script, make sure to update NewADUsers.csv in \\<NETWORK SHARE> 
Continue? (Type 'Y' to continue, 'N' to cancel): " -ForegroundColor Red -BackgroundColor White
$response = read-host
if ( $response -ne "Y" ) { exit }

#Prompt user for credentials and capture the input into a variable
$UserCredential = $host.ui.PromptForCredential("Need credentials", "Enter your Domain Admin credentials.", "<DOMAIN>\", "NetBiosUserName")

#Create Exchange PowerShell session
$ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<EXCHANGE SERVER>/PowerShell -Authentication Kerberos -Credential $UserCredential
Import-PSSession $ExchSession -DisableNameChecking -AllowClobber | Out-Null

#Import updated CSV file
$NewADUsers = Import-Csv "\\<NETWORK SHARE>\NewADUsers.csv"

#Loop through each row containing user details in the CSV file
foreach ($User in $NewADUsers) {
    #Read user data from each field in each row and assign the data to a variable as below
    $DisplayName = $User.FirstName + " " + $User.LastName 
    $FirstName = $User.FirstName
    $LastName = $User.LastName
    $Password = (ConvertTo-SecureString $User.Password -AsPlainText -Force)
    $Manager = $User.Manager
    $UserName = "$($FirstName.Substring(0,1))$LastName"

    #Check if an existing user already has the first initial/last name username taken
    Write-Verbose -Message "Checking if [$($Username)] is available"
    if (Get-ADUser -Filter "Name -eq '$Username'") {
        Write-Warning -Message "The username [$($Username)] is not available. Checking alternate..."
        ## If so, check to see if the first initial/middle initial/last name is taken.
        $Username = "$($FirstName.SubString(0, 1))$MiddleInitial$LastName"
        if (Get-ADUser -Filter "Name -eq '$Username'") {
            throw "No acceptable username schema could be created"
        }
        else {
            Write-Verbose -Message "The alternate username [$($Username)] is available."
        }
    }
    else {
        Write-Verbose -Message "The username [$($Username)] is available"
    }
    
    #Create the user now that variables have been defined
    New-ADUser -City: "<CITY>" -Company:"<COMPANY>" -Country:"<COUNTRY>" -Department:"<DEPARTMENT>" -DisplayName:"$DisplayName" -GivenName:"$FirstName" -HomeDirectory:"\\<NETWORK SHARE>\$UserName" -HomeDrive:"<DRIVE LETTER>:" -Name "$DisplayName" -Path:"<ORGANIZATIONAL UNIT>" -PostalCode:"<ZIP CODE>" -SamAccountName:"$UserName" -Server:"<DOMAIN CONTROLLER>" -State:"<STATE>" -StreetAddress:"<STREET ADDRESS>" -Surname:"$LastName" -Title:"<COMPANY TITLE>" -Type:"user" -UserPrincipalName:"$UserName@<DOMAIN>" -Manager:"$Manager"

    #Set password
    Set-ADAccountPassword -Identity $UserName -NewPassword:$Password -Reset:$true -Server:"<DOMAIN CONTROLLER>"

    #Set group membership
    Add-ADPrincipalGroupMembership -Identity "$UserName" -MemberOf "<MEMBEROF GROUPS>" -Server:"<DOMAIN CONTROLLER>"

    #Set authentication
    Set-ADAccountControl -AccountNotDelegated:$false -AllowReversiblePasswordEncryption:$false -CannotChangePassword:$false -DoesNotRequirePreAuth:$false -Identity:"CN=$DisplayName,<ORGANIZATIONAL UNIT>" -PasswordNeverExpires:$false -Server:"<DOMAIN CONTROLLER>" -UseDESKeyOnly:$false

    #Prompt user to change password at logon
    Set-ADUser -ChangePasswordAtLogon:$true -Identity "$UserName" -Server:"<DOMAIN CONTROLLER>" -SmartcardLogonRequired:$false

    #Syncronize DC info
    Invoke-Command -ComputerName <DOMAIN CONTROLLER 1>,<DOMAIN CONTROLLER 2> -Credential $UserCredential -ScriptBlock { & 'C:\windows\System32\repadmin.exe' /syncall }

    #Create mailbox
    Enable-Mailbox -Identity $UserName -Alias $UserName | Out-Null

    #Send new hire email
    Send-MailMessage -From <EMAIL ADDRESS> -to <HR EMAIL> -Cc <OTHER RELEVANT PARTIES> -Subject "Login credentials for $DisplayName" -Body "Username: $UserName
Password: $Password" -SmtpServer <EXCHANGE SERVER>
}

Remove-PSSession -Session $ExchSession
