# Runs as admin
Write-Host "Checking for elevation... "  
$CurrentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent()) 
if (($CurrentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) -eq $false) 
{ 
    $ArgumentList = "-noprofile -noexit -file `"{0}`" -Path `"$Path`" -MaxStage $MaxStage" 
    If ($ValidateOnly) { $ArgumentList = $ArgumentList + " -ValidateOnly" } 
    If ($SkipValidation) { $ArgumentList = $ArgumentList + " -SkipValidation $SkipValidation" } 
    If ($Mode) { $ArgumentList = $ArgumentList + " -Mode $Mode" } 
    Write-Host "elevating" 
    Start-Process powershell.exe -Verb RunAs -ArgumentList ($ArgumentList -f ($myinvocation.MyCommand.Definition)) -Wait 
    Exit 
}  
Write-Host "in admin mode.."
## Functions, messages, and modules
$FormatEnumerationLimit = -1
$Shell = New-Object -ComObject "WScript.Shell"
$Shell.Popup("All user information is read from C:\scripts\ActiveDirectory\CSVs\user-creation.csv. You'll be prompted to enter user information manually if the CSV is empty (fine if creating a single user). If you need to create multiple users, pause here and enter all information into the CSV (notepad has to run as admin to edit). Titles and departments are case sensitive (IE: A&A, Senior Manager). Click OK once complete.", 0, "Notice", 0) > $null
$Shell.Popup("Ensure you have a copy of the passwords written down somewhere for user profile setup. The CSV file is reset upon completion of the script. Click OK to acknowledge.", 0, "Notice", 0) > $null
$Shell.Popup("You'll be prompted to log into 365 multiple times for different modules used by this process. Completing connection to all modules can take some time and it may look like the script hung. Give it a bit.", 0, "Notice", 0) > $null
Write-Host "Importing Modules" -ForegroundColor Cyan
Import-Module AzureAD
Import-Module importexcel
Import-Module ExchangeOnlineManagement
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Users.Actions 
Import-Module Microsoft.Graph.Identity.DirectoryManagement
Import-Module MicrosoftTeams
Write-Host "Modules imported" -ForegroundColor Cyan
Function Start-Countdown{
    <#
    .SYNOPSIS
        Provide a graphical countdown if you need to pause a script for a period of time
    .PARAMETER Seconds
        Time, in seconds, that the function will pause
    .PARAMETER Messge
        Message you want displayed while waiting
    .EXAMPLE
        Start-Countdown -Seconds 30 -Message Please wait while Active Directory replicates data...
    .NOTES
        Author:            Martin Pugh
        Twitter:           @thesurlyadm1n
        Spiceworks:        Martin9700
        Blog:              www.thesurlyadmin.com
       
        Changelog:
           2.0             New release uses Write-Progress for graphical display while couting
                           down.
           1.0             Initial Release
    .LINK
        http://community.spiceworks.com/scripts/show/1712-start-countdown
    #>
    Param(
        [Int32]$Seconds = 10,
        [string]$Message = "Pausing for 10 seconds..."
    )
    ForEach ($Count in (1..$Seconds))
    {   Write-Progress -Id 1 -Activity $Message -Status "Exiting" -PercentComplete (($Count / $Seconds) * 100)
        Start-Sleep -Seconds 1
    }
    Write-Progress -Id 1 -Activity $Message -Status "Completed" -PercentComplete 100 -Completed
}
## Start script
# Starting variables
## This is for checking if it's running on a DC or on a local machine for testing ##
$file = if ($env:COMPUTERNAME -eq "DomainController") {"C:\scripts\ActiveDirectory\CSVs\user-creation.csv"
}elseif($env:COMPUTERNAME -eq "LocalComputer"){"Fill in path"
}
# Log into Microsoft
Connect-AzureAD > $null
Connect-ExchangeOnline -ShowBanner:$false
Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.ReadWrite.All" > $null
Connect-MicrosoftTeams > $null
# Checks if there are any E3 or E5 licenses available before proceeding
Write-Host "Getting available 365 licenses" -ForegroundColor Cyan
Start-Sleep -Seconds 3
## The license check checks for a few different SKUs used by the original script, but only assigns E5 licenses below for this example ##
$licenses = Get-MGSubscribedSku | Where-Object {$_.skupartnumber -eq "ENTERPRISEPREMIUM" -or $_.skupartnumber -eq "ENTERPRISEPACK" -or $_.skupartnumber -eq "EXCHANGESTANDARD"} | ForEach-Object {
    [PSCustomObject]@{
        License = $_.skupartnumber
        Active = $_.consumedunits
        Total = $_.prepaidunits.enabled
    }
}
$zerocheck = ($licenses | Measure-Object -Property total -sum | Select-Object -ExpandProperty sum) - ($licenses | Measure-Object -Property active -sum | Select-Object -ExpandProperty sum)
if ($zerocheck -eq 0){
    Start-Countdown -Seconds 15 -Message "Script will now exit. There are no E3 or E5 licenses available. Reach out to management to provision the necessary licenses and run the script again."
    Stop-Process -ID $PID
}else{
    $licenses | ForEach-Object {
        foreach ($license in $_.license){
            foreach ($active in $_.active){
                foreach ($total in $_.total){
                    if ($license -eq "ENTERPRISEPREMIUM"){
                        $available = $total - $active
                        if ($available -eq 0){
                            Write-Host "There are no E5 licenses available" -ForegroundColor Magenta
                        }else{
                            Write-Host "Available E5 licenses: $available" -ForegroundColor Yellow
                        }
                    }elseif($license -eq "ENTERPRISEPACK"){
                        $available = $total - $active
                        if ($available -eq 0){
                            Write-Host "There are no E3 licenses available" -ForegroundColor Magenta
                        }else{
                            Write-Host "Available E3 licenses: $available" -ForegroundColor Yellow
                        }
                    }elseif($license -eq "EXCHANGESTANDARD"){
                        $available = $total - $active
                        if ($available -eq 0){
                            Write-Host "There are no Exchange Online licenses available" -ForegroundColor Magenta
                        }else{
                            Write-Host "Available Exchange Online licenses: $available" -ForegroundColor Yellow
                        }
                    }
                }
            }
        }
    }
}
do {
    $continue = Read-Host -Prompt "Are enough licenses available for you to continue? [Yes/No]"
    if ($continue -eq "yes" -or $continue -eq "no"){
        break
    }else{
        Write-Host "Invalid answer, try again" -ForegroundColor Magenta
    }
}until($continue -eq "Yes" -or $continue -eq "No")
if ($continue -eq "No"){
    Start-Countdown -Seconds 15 -Message "Script will now exit. Reach out to management to provision the necessary licenses and run the script again."
    Stop-Process -ID $PID
}
$reportemail = Read-Host -Prompt "Enter your email address for results report sent at the end of the script"
# Checks if CSV is populated and prompts for user information if empty
$csvcheck = Import-Csv $file
if ($csvcheck.firstname -eq ""){
    Write-Host "CSV file is empty, fill in user information below" -Foregroundcolor Cyan
    $csv = Import-Csv $file
    $csv.FirstName = Read-Host -Prompt "Enter new user's first name"
    $csv.LastName = Read-Host -Prompt "Enter new user's last name"
    $csv.Password = Read-Host -Prompt "Enter new user's password"
    $csv.Office = Read-Host -Prompt "Enter new user's main office" 
    $csv.Department = Read-Host -Prompt "Enter new user's department"
    $csv.Title = Read-Host -Prompt "Enter new user's title"
    $csv.Firstname = ((Get-Culture).TextInfo).ToTitleCase(($csv.Firstname))
    $csv.Lastname = ((Get-Culture).TextInfo).ToTitleCase(($csv.Lastname))
    $csv.Department = ((Get-Culture).TextInfo).ToTitleCase(($csv.department))
    $csv.Title = ((Get-Culture).TextInfo).ToTitleCase(($csv.Title))  
    $csv | Export-Csv $file -NoTypeInformation
}
# Gets disabled users for reactivation check and checks if an account exists for any of the users being created
$active = @(Get-ADUser -Filter * -SearchBase "Fill in OU where active users are stored" | Select-Object -ExpandProperty samaccountname)
$disabled = @(Get-ADUser -Filter * -SearchBase "Fill in OU for where disabled users are stored" | Select-Object -ExpandProperty samaccountname)
# Import CSV and create accounts
$csv = Import-Csv $file
$CSV | ForEach-Object {
    $FirstName = $_.firstname
    $LastName = $_.lastname
    $username = ($FirstName[0]+$LastName).ToLower()
    $DisplayName = $FirstName +" "+ $LastName
    $Password = ($_.password | ConvertTo-SecureString -AsPlainText -Force)
    $Office = $_.office
    $upn = "$username@domain.com"
    $Department = $_.department
    $Title = $_.title
    $DefaultGroups = @("Fill in default groups")
    ## Use the following if you have more than user OU for different locations or designations. The logic can be duplicated with additional elseif statements for more OUs ##
    if ($office -eq "Office1"){
        $ADPath = "Fill in OU for Office1"
        $Office1Groups = "Add any specific groups for these users"
        $DefaultGroups += $Office1Groups
    }elseif($office -eq "Office2"){
        $ADPath = "Fill in OU for Office2"
        $Office2Groups = @("Add any specific groups for these users")
        $DefaultGroups += $Office2Groups
    }
    ## Quick check for any existing active accounts with the same username ##
    if ($active -contains $username){
        Write-Host "An account for $DisplayName already exists. Skipping" -ForegroundColor Yellow
    }
    ## This if check is how my org identifies disabled users and prompts to reactivate if the same username is found, can be adjusted to your needs ##
    if ($disabled -contains $username+"_old" -or $disabled -contains $username) {
        $disabledcheck = @(Write-Host "A previously created account was found for $displayname. Reactivate? [Yes/No]:" -ForegroundColor Yellow -NoNewline;Read-Host)
        if ($disabledcheck -eq "yes"){
            Write-Host "Reactivating account for $displayname" -Foregroundcolor Cyan
            ## Org specific; termed users email addresses are aliased to a main DisabledUsers account and the account's email/samaccountname is appended with _old. Remove as needed ##
            Write-Host "Removing alias from DisabledUsers account" -Foregroundcolor Cyan
                set-aduser disabledusers -remove @{proxyaddresses = "smtp:$username@domain.com"} -Verbose
            Write-Host "Removing '_old' from previously disabled account" -Foregroundcolor Cyan
                Get-ADUser $username"_old" | set-aduser -UserPrincipalName $username"@domain.com" -Verbose
                Get-ADUser $username"_old" | set-aduser -SamAccountName $username -Verbose
                Enable-ADAccount $username   
            Write-Host "Moving account for $displayname to the correct OU and setting applicable groups" -Foregroundcolor Cyan
                if ($office -eq "Office1"){
                    Move-ADObject (Get-ADUser $username).objectguid -TargetPath "Fill in OU for Office1" -Verbose
                }
                if ($office -eq "Office2"){
                    Move-ADObject (Get-ADUser $username).objectguid -TargetPath "Fill in OU for Office2"
                }
                foreach ($defaultgroup in $defaultgroups){
                    Add-ADGroupMember -Identity $defaultgroup -Members $username -Verbose
                }
                ## Sets default groups for different departments/titles. This logic can be duplicated with additional if statements ##
                if ($department -eq "Department1") {
                    Add-ADGroupMember -Identity "Group for all department1 users" -Members $username -Verbose
                    if ($title -eq "Associate"){
                        Add-ADGroupMember -Identity "Group for all department1 associates" -Members $username -Verbose
                    }
                }
                if ($department -eq "Department2"){
                    if ($title -eq "Associate"){
                        Add-ADGroupMember -Identity "Group for department2 associates" -Members $username -Verbose
                    }
                }
                ## This might not be needed, but it's for a group needed by everyone except a particular department. Remove as needed ##
                if ($department -ne "Department1"){
                    Add-ADGroupMember -Identity "Group for everyone except department1" -Members $username -Verbose
                }
            Write-Host "Setting new password and user attributes" -Foregroundcolor Cyan
                Set-ADAccountPassword -Identity (Get-ADUser $username).objectguid -reset -NewPassword $Password -Verbose
                Set-ADUser $username -Replace @{"department" = "$department"} -Verbose
                set-aduser $username -Replace @{"title" = "$title"} -Verbose
        }else{
            Start-Countdown -Seconds 15 -Message "Script will now exit. A new account cannot be made with the same username of an existing account. Please address existing duplicate for $displayname and try again."
            Stop-Process -ID $PID
        }
    }else{
        Write-Host "Creating account for $displayname" -ForegroundColor Cyan
        $NewUserParams = @{
            Path = $ADPath
            GivenName = $FirstName
            Surname = $LastName
            Name = $displayname
            Displayname = $displayname
            UserPrincipalName = $upn
            SamAccountName = $username
            Department = $Department
            Title = $Title
            EmailAddress = $upn
            AccountPassword = $Password
            Enabled = 1
            ChangePasswordAtLogon = 0
        }
        New-Aduser @NewUserParams -Verbose
        foreach ($defaultgroup in $defaultgroups){
            Add-ADGroupMember -Identity $defaultgroup -Members $username -Verbose
        }
        ## Sets default groups for different departments/titles. This logic can be duplicated with additional if statements ##
        if ($department -eq "Department1") {
            Add-ADGroupMember -Identity "Group for all department1 users" -Members $username -Verbose
            if ($title -eq "Associate"){
                Add-ADGroupMember -Identity "Group for all department1 associates" -Members $username -Verbose
            }
        }
        if ($department -eq "Department2"){
            if ($title -eq "Associate"){
                Add-ADGroupMember -Identity "Group for department2 associates" -Members $username -Verbose
            }
        }
        ## This might not be needed, but it's for a group needed by everyone except a particular department. Remove as needed ##
        if ($department -ne "Department1"){
            Add-ADGroupMember -Identity "Group for everyone except department1" -Members $username -Verbose
        }
    }
}
# Loop to run ADSync until it reports a successful sync
Write-Host "Running ADSync" -ForegroundColor Cyan
    do {
        Start-ADSyncSyncCycle -PolicyType Delta -OutVariable ADSyncResult -ErrorAction SilentlyContinue > $null
        Start-Sleep -Seconds 5
        if (($ADSyncResult | Out-String) -match "Success"){
            break
        }
    }until(($ADSyncResult | Out-String) -match "Success")
# Loop to get check that each new user is in Azure AD
Write-Host "Searching for new users in 365. This will take a few minutes" -Foregroundcolor Cyan
$csv | ForEach-Object {
    $FirstName = $_.firstname
    $LastName = $_.lastname
    $username = ($FirstName[0]+$LastName).ToLower()
    $DisplayName = $FirstName +" "+ $LastName
    $upn = "$username@domain.com"
    do{
        Get-AzureADUser -All $true | Where-Object {$_.userprincipalname -match $upn} -ErrorAction SilentlyContinue
    }until($null -ne (Get-AzureADUser -All $true | Where-Object {$_.userprincipalname -match $upn}))
}
# New loop to configure user in 365 after pause for ADSync to complete
$csv | ForEach-Object {
    $FirstName = $_.firstname
    $LastName = $_.lastname
    $username = ($FirstName[0]+$LastName).ToLower()
    $DisplayName = $FirstName +" "+ $LastName
    Write-Host "Assigning E5 license to $DisplayName." -Foregroundcolor Cyan
        $upn = Get-ADUser $username | Select-Object -ExpandProperty userprincipalname
        $userID = Get-MGUser -Filter "startswith(userprincipalname, '$upn')" | Select-Object -ExpandProperty ID
        $E5 = Get-MGSubscribedSku | Where-Object {$_.skupartnumber -eq "ENTERPRISEPREMIUM"} | Select-Object -ExpandProperty SkuID
        Update-MgUser -UserId $userID -UsageLocation US -Verbose
        Set-MgUserLicense -UserId $userID -AddLicenses @{SkuID = $E5} -RemoveLicenses @() -Verbose
    Write-Host "Searching for mailbox created for $DisplayName. This will take a few minutes" -Foregroundcolor Cyan
        do {
            $mailbox = Get-Mailbox -Identity $upn -ea 0
            if ($null -ne $mailbox){
                break
            }
        }until($null -ne $mailbox)
    Write-Host "Found mailbox for $DisplayName" -Foregroundcolor Cyan
        ## Org specific;, but we set our termed users as shared mailboxes so this part sets a reactivated user back to a standard mailbox. Remove as needed ##
        if ($mailbox.recipienttypedetails -eq "SharedMailbox"){
            Write-Host "Setting shared mailbox for $displayname back to standard" -Foregroundcolor Cyan
            Set-Mailbox $upn -Type Regular -verbose
        }
    Write-Host "Setting calendar permissions for $displayname" -Foregroundcolor Cyan
        ## Org specific;; we set all users calendars as read-only so everyone can view availability. Remove as needed ##
        $usermailbox = $mailbox.PrimarySmtpAddress
        Set-MailboxFolderPermission -identity $usermailbox":\Calendar" -User Default -AccessRights Reviewer -Verbose
    }
## Org specific; we use Teams for our phone system and this part queries for the assigned Teams number to add to AD. Does not add the Teams number as it needs to be added manually through our provider. Remove/modify as needed ##
## Loop is now separate so that all users are licensed first before querying for Teams numbers ##
$csv | ForEach-Object {
    $FirstName = $_.firstname
    $LastName = $_.lastname
    $username = ($FirstName[0]+$LastName).ToLower()
    $DisplayName = $FirstName +" "+ $LastName
    Write-Host "Searching for number assigned to $DisplayName for up to 3 minutes. Please wait" -Foregroundcolor Cyan
    $timeout = New-TimeSpan -Seconds 180
    $endtime = (Get-Date).add($timeout)
    do{
        $teamsdirectget = Get-CSOnlineUser -Identity $username | Select-Object -ExpandProperty LineURI -ea 0
        if ($null -ne $teamsdirectget){
            break
        }
    }until($null -ne $teamsdirectget -or (Get-Date) -gt $endtime)
    if ($null -eq $teamsdirectget){
        Write-Host "No Teams number found for $DisplayName after 3 minutes. Moving to next steps. Phone attribute needs to be set manually" -ForegroundColor Magenta
    }else{Write-Host "Number found for $DisplayName, setting phone attribute" -Foregroundcolor Cyan
        ## Formats the number as 123.456.7890
        $teamsdirect = $teamsdirectget.substring(6).insert(3,'.').insert(7,'.')
        Set-ADUser $username -OfficePhone $teamsdirect
    }
}
## Generates report and emails to whoever ran the script. Add email settings if you want to keep this ##
## This is for checking if it's running on a DC or on a local machine for testing ##
$exportpath = if ($env:COMPUTERNAME -eq "DomainController") {"C:\scripts\ActiveDirectory\Reports"
}elseif($env:COMPUTERNAME -eq "LocalComputer"){"Fill in path on local computer"
}
$date = Get-Date -Format M-d-yyyy
$reportuser = Get-ADUser ($reportemail.split("@")[0]) | Select-Object -ExpandProperty samaccountname
$exportfile = "$exportpath\$reportuser-$date-UserCreation.xlsx"
Write-Host "Generating results report" -ForegroundColor Cyan
$report = $csv | ForEach-Object{
    $FirstName = $_.firstname
    $LastName = $_.lastname
    $username = ($FirstName[0]+$LastName).ToLower()
    $groups = (@(Get-ADPrincipalGroupMembership -Identity $username | Select-Object -ExpandProperty Name | Sort-Object) -Join ",'" -Replace "'","" | Out-String).trim()
    $officeget = Get-ADUser $username -Properties * | Select-Object -ExpandProperty distinguishedname
    ## Office can be expanded with additional elseif statements for more locations ##
    $office = if ($officeget -match "Office1"){"Office1"}elseif($officeget -match "Office2"){"Office2"}
    $license = Get-MgUserLicenseDetail -UserId (Get-MGUser -Filter "startswith(userprincipalname, '$username')" | Select-Object -ExpandProperty ID) | Where-Object {$_.skupartnumber -eq "ENTERPRISEPREMIUM" -or $_.SkuPartNumber -eq "ENTERPRISEPACK"} | Select-Object -ExpandProperty skupartnumber
    $teamsnumber = Get-CSOnlineUser -Identity $username | Select-Object -ExpandProperty LineURI
    [PSCustomObject]@{
        Name = Get-ADUser -Identity $username | Select-Object -ExpandProperty name
        Username = $username
        Office = $office
        Department = Get-ADUser -Identity $username -Properties * | Select-Object -ExpandProperty Department
        Title = Get-ADUser -Identity $username -Properties * | Select-Object -ExpandProperty Title
        Groups = $groups
        ## Licenses can be expanded with additional elseif statements for reporting on more license types ##
        License = if ($license -eq "ENTERPRISEPREMIUM"){"E5"}else{"E3"}
        Mailbox = Get-Mailbox -Identity $username | Select-Object -ExpandProperty RecipientTypeDetails
        CalendarPermissionsSet = if ((Get-MailboxFolderPermission $username":\Calendar").user -match "Default" -and (Get-MailboxFolderPermission $username":\Calendar").accessrights -match "Reviewer"){"Yes"}else{"No"}
        TeamsNumberAssigned = if ($null -ne $teamsnumber){$teamsnumber.substring(6).insert(3,'.').insert(7,'.')}else{"No"}
    }
} 
$output = $report | Export-Excel -Path $exportfile -Passthru
Set-ExcelRange -Worksheet $output.Sheet1 -Range "A1:K1" -HorizontalAlignment Center
Set-ExcelRange -Worksheet $output.Sheet1 -Range "A:E" -AutoSize
Set-ExcelRange -Worksheet $output.Sheet1 -Range "F:F" -Width 106.43
Set-ExcelRange -Worksheet $output.Sheet1 -Range "F:F" -WrapText
Set-ExcelRange -Worksheet $output.Sheet1 -Range "G:K" -AutoSize
Close-ExcelPackage $output
Write-Host "Emailing report to $reportemail" -ForegroundColor Cyan
$params = @{ 
    To = $reportemail
    From = "UserCreate@domain.com" 
    sub = "User Creation Report" 
    body = "See attached file"
    Attachment = $exportfile
    BodyAsHTML = $true 
    SMTPServer = "Insert SMTP server" 
}
Send-MailMessage @params
# Resets CSV file
Write-Host "Resetting CSV file" -Foregroundcolor Cyan
$Shell.Popup("Double check the CSV file isn't still open, otherwise PS isn't able to reset it. Click OK once complete.", 0, "Notice", 0) > $null
    $csv | Select-Object * -ExcludeProperty * | Export-Csv $file -NoTypeInformation
    $headers = @(
        "Firstname"
        "Lastname"
        "Password"
        "Office"
        "Department"
        "Title"
    )
    {} | Select-Object $headers | Export-Csv $file -NoTypeInformation
Remove-Variable csvcheck
Remove-Variable csv
Write-Host "Disconnecting from MS Online" -Foregroundcolor Cyan
    Disconnect-AzureAD
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MicrosoftTeams
    Disconnect-MgGraph
Write-Host "Done!" -ForegroundColor Green
