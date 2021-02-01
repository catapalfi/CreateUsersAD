#Start script

Start-Transcript -Path C:\Scripts\AccountCreation\Logs\Create_Users_transcript.log -Append -IncludeInvocationHeader
function Write-Log {
    param (
        $message,
        [ValidateSet('Info', 'Warning', 'Error')]
        $level = 'Info'
    )
    $MsgColors = @{'Info' = 'Gray'; 'Warning' = 'Yellow'; 'Error' = 'Red' }
    Write-Output "$(get-date -UFormat '%Y%m%d-%H%M%S')::$level:: $message" | Out-File -FilePath "C:\Scripts\AccountCreation\Logs\$(Get-Date -UFormat '%Y%m%d')_Create_Users_transcript.log" -Append
    Write-Host "$(get-date -UFormat '%Y%m%d-%H%M%S')::$level:: $message" -ForegroundColor $MsgColors[$level]
}
Write-Log "Script start."

# Generating random password
Write-Log "Generating random password"
function Get-RandomPassword {
    Param(
        [Parameter(mandatory=$true)]
        [int]$Length
 )
    Begin{
        if($Length -lt 4){
        End
        }
        $Numbers = 1..9
        $LettersLower = 'abcdefghijklmnopqrstuvwxyz'.ToCharArray()
        $LettersUpper = 'ABCEDEFHIJKLMNOPQRSTUVWXYZ'.ToCharArray()
        $Special = '!@#$%^&()=+[{}]/?<>*'.ToCharArray()
 
        #For the 4 character types (upper, lower, numerical, and special)
        $N_Count = [math]::Round($Length*.2)
        $L_Count = [math]::Round($Length*.4)
        $U_Count = [math]::Round($Length*.2)
        $S_Count = [math]::Round($Length*.2)
    }
    Process{
        $Pwd = $LettersLower | Get-Random -Count $L_Count
        $Pwd += $Numbers | Get-Random -Count $N_Count
        $Pwd += $LettersUpper | Get-Random -Count $U_Count
        $Pwd += $Special | Get-Random -Count $S_Count

        #If the password length isn't long enough (due to rounding),
        #add X special characters, where X is the difference between
        #the desired length and the current length.
        if($Pwd.length -lt $Length){
            $Pwd += $Special | Get-Random -Count ($Length - $Pwd.length)
        }

        #Lastly, grab the $Pwd string and randomize the order
        $Pwd = ($Pwd | Get-Random -Count $Length) -join ""
    }
    End{
        $Pwd
    }
}

# Create the header for each csv file in the Drop folder
Write-Log "Creating the header for each csv file in the Drop folder"
$Path = "C:\Scripts\AccountCreation\Drop\"
$Files = Get-ChildItem -Path $Path

Foreach ($File in $Files)
{
    $FilePath = $Path+$File.Name
    $Header = "FirstName,LastName,DisplayName,UserName,UPN,Email,extensionAttribute1,AccountDescription,AccountOU,employeeType,AccountOwner,OwnerEmail,OwnerDisplayName,RequestorEmail”
    $User = "C:\Scripts\AccountCreation\Create\"+$File
    $Header | Out-File $User
    $O365Output = Get-Content -Path $FilePath
    $O365Output | Out-File $User -Append
    $U = Import-CSV -Path $User

    # Define user properties from CSV
    $FirstName=$U.FirstName
    $LastName=$U.LastName
    $DisplayName = $U.DisplayName
    $UserName = $U.UserName
    $UPN = $U.UPN
    $Email=$U.Email
    $extensionAttribute1=$U.extensionAttribute1
    $AccountDescription = $U.AccountDescription
    $OU = $U.AccountOU
    $employeeType=$U.employeeType
    $AccountOwner=$U.AccountOwner
    $OwnerEmail=$U.OwnerEmail
    $OwnerDisplayName=$U.OwnerDisplayName
    $RequestorEmail=$U.RequestorEmail
    $PWD= Get-RandomPassword -Length 16
    Write-Log "Defined user properties from CSV $UPN with account owner $AccountOwner and owner $OwnerEmail"

    # Defining the Account Organizaiton Unit
    $OwnerDN = Get-ADUser -Filter 'UserPrincipalName -eq $AccountOwner' -Server domain.local:3268 | Select-Object -ExpandProperty DistinguishedName
    $OwnerDC=$OwnerDN.Substring($OwnerDN.IndexOf("DC="))
    $AccountOU= $OU +","+ $OwnerDC
    Write-Log "Definined account OU: $AccountOU" -level Info

    # Defining the User domain Organizaiton Unit
    $OwnerDN = Get-ADUser -Filter 'UserPrincipalName -eq $AccountOwner' -Server domain.local:3268
    $Userdomain=$OwnerDN.DistinguishedName.Split(",") | ? {$_ -like "DC=*"}
    $Userdomain = $userdomain.Substring(3)
    $UserdomainDN=$userdomain -join "."
    Write-Log "Defined user domain: $UserdomainDN" -level Info

    # Defining the parameters for user creation from CSV
    $Parameters = @{
        GivenName=$FirstName
        Surname = $LastName
        Name=$DisplayName
        DisplayName = $DisplayName
        SamAccountName = $UserName
        UserPrincipalName = $UPN
        EmailAddress = $Email
        Description = $AccountDescription
        Path = $AccountOU
        AccountPassword = (ConvertTo-SecureString $pwd -AsPlainText -Force)
        Enabled = $true
        ChangePasswordAtLogon = $true
        PasswordNeverExpires = $false
        Server=$userdomainDN
    }
    Write-Log "Defined the user parameters from CSV"

    # Create new user in AD with the defined parameters
    try {
        if (!(Get-ADUser -Filter {samaccountname -eq "$Username"}))
        {
    New-ADUser @Parameters -OtherAttributes @{EmployeeType=$employeeType; extensionAttribute1=$extensionAttribute1; extensionAttribute2=$employeeType}
    Write-Log "Created user $DisplayName with UPN: $UPN " -level Info
    # If the user creation is successfull write it in the logs
    }

    else {
    Write-Log "Samaccount for username [$($Displayname)] already exists" -level Error
    # If the user already exists write it in the logs, move file and send failure email to requestor
    Move-Item -Path $user -Destination C:\Scripts\AccountCreation\Error
    Write-Log "Processed file moved to the Error folder" -level Info

    # Removing processed file from the Create folder write it in the logs
    Remove-Item -Path $FilePath
    Write-Log "Processed user file removed from the Create folder" -level Info
    Send-MailMessage -To "$RequestorEmail", "it.support@domain.com" -From "noreply@domain.com" -Subject "The account already exists!" -Body "The user $username already exists in AD. Please fill in another form with a different account name." -SmtpServer "smtp.domain.local"

    Stop-Transcript

    Exit
    }
 }

 catch {
    Write-Log "Can't create user [$($Displayname)] : $_" -level Error
    # If the user cannot be created write it in the logs, move file and send email to IT

    Move-Item -Path $user -Destination C:\Scripts\AccountCreation\Error
    Write-Log "Processed file moved to the Error folder" -level Info
    
    # Removing processed file from the Create folder write it in the logs
    Remove-Item -Path $FilePath
    Write-Log "Processed User file removed from the Create folder" -level Info
    Send-MailMessage -To "it.support@domain.com" -From "no-reply@domain.com" -Subject "User creation failed" -Body "Please find attached log file for AD account creation process. User creation failed, check logs." -Attachments "C:\Create_Users\Logs\$(Get-Date -UFormat '%Y%m%d')_Create_Users_transcript.log" -SmtpServer "smtp.domain.local"

    Stop-Transcript
    
    Exit
 }

    # Sending email message to requestor with credentials
    $Body=
    (Get-Content C:\Scripts\AccountCreation\EmailTemplate\NewADUser_ForEmployee.txt -Raw) | Foreach-Object {
        $_ -replace '{givenName} {sn},', $OwnerDisplayName `
        -replace '{userPrincipalName}', $UPN `
        -replace '{password}', $pwd `
        -replace '{email}', $Email `
        -replace '{accountName}', $UPN
 }
 
    Send-MailMessage -To "$OwnerEmail" -From "no-reply@domain.com" -Subject "New AD account created" -Body $Body -BodyAsHtml -SmtpServer "smtp.domain.local"
    Write-Log "Sending email message to requestor with credentials" -level Info

    # Moving processed file to the Archive folder write it in the logs
    Move-Item -Path $user -Destination C:\Scripts\AccountCreation\Archive
    Write-Log "Processed User file moved to the Archive folder" -level Info

    # Removing processed file from the Create folder write it in the logs
    Remove-Item -Path $FilePath
    Write-Log "Processed User file removed from the Create folder" -level Info

    Start-Sleep -Seconds 10
}
Write-Log "Sending email with logs to IT"
Send-MailMessage -To "it.support@domain.com" -From "no-reply@domain.com" -Subject "New AD account created" -Body "Please find attached log file for AD account creation process.Please also try to examine the transcript log file in case some errors might not have been caught." -Attachments "C:\Scripts\AccountCreation\Logs\$(Get-Date -UFormat '%Y%m%d')_Create_Users_transcript.log" -SmtpServer "smtp.domain.local"

Stop-Transcript