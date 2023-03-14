<#
Author: Walter Muncaster
Date: 2/1/2021
Purpose: Check Workday users for ON LEAVE status and update status for corresponding user object in Active Directory. 
Scan AD for users that have not signed in for 30+ days. If not On Leave, send warning email to users manager, Ent Apps, Security, IT.
For users that have not signed in for 55 days, send offboarding email to stakeholders.
For any users that have not signed in for 60+ days, disable and move AD object.
#>


Import-module ActiveDirectory
[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"


# Function: Send Email Alert for admins to review potential issues in code/etc
function AdminAlertEmail($Body){
	if ($Body -notlike "*Could not create SSL/TLS secure channel.*") {
		write-host "Alerting Admin"
		$From = $env:COMPUTERNAME + $env:SENDERDOMAIN
		$To = $env:ALERTSENDTO
		$SMTPHost = $env:SMTP
		$Subject = "AD Stale Account Issue on $env:COMPUTERNAME"
		Send-MailMessage -SmtpServer $SMTPHost -From $From -To $To -Body $Body -Subject $Subject
	}
}


# Function: Send Slack Channel Alert for admins to review potential issues in code/groups/etc
function AdminAlertSlack($Body){
	if($Body -notlike "*Could not create SSL/TLS secure channel.*") {
		$JSONStart = "{`"text`": `""
		$JsonEnd = "`"}"
		$FullBody = $jsonstart + $Body + $jsonend
		Invoke-WebRequest -Method POST -Uri $env:SLACKURI -ContentType application/json -Body $FullBody
	}
}


#Function: Compose 30 Day Email and return email object
function Compose30DayEmail($30DayUserList,$SendToList) {
    
    $From = $env:FROM 
    $Subject = "Inactive Account Monitor"
    $HTML = $30DayUserList | % {$_} | select @{n="Name";e={$_.CN}},@{n="Email";e={$_.Mail}},Title,@{n="Last Sign-In Date";e={$_.LastLogonDate}},@{n="Manager Name";e={$_.ManagerName[0]}},@{n="Employee Type";e={$_.EmployeeType}},@{n="Disable Date";e={$_.disableDate[0]}} | Sort-Object -Property "Last Sign-In Date" | ConvertTo-HTML | Out-String
    $HTML = $HTML.replace("<table>", '<table cellpadding="10">')
    $Body = @"
        <p><h4>Hi All, <br>Please review the information below.</br></h4></p>
        <p>You may manage one of the following user accounts which are being flagged as inactive.</p>
        <p>The user(s) below should sign into <a href="https://okta.com/">okta.com</a> to prevent being disabled and offboarded. If the user is no longer Employeed please IMMEDIATELY submit a Service-Now ticket.</p>
        <p><u>Accounts inactive for 30+ days:</u></p>
        <p>$HTML</p>
        <p>If you have any questions or issues with the above information please reach out to IT.</p>
        <p>Thank you!</p>
"@
    $FullMessage = New-Object System.Net.Mail.MailMessage $From, $SendToList, $Subject, $Body
    $FullMessage.IsBodyHtml = $True

    return $FullMessage

}


#Function: Compose 55 Day Email and return email object
function Compose55DayEmail($55DayUserList,$SendToList) {
    
    $From = $env:FROM 
    $Subject = "Action Required: Inactive Account Monitor"
    $HTML = $55DayUserList | % {$_} | select @{n="Name";e={$_.CN}},@{n="Email";e={$_.Mail}},Title,@{n="Last Sign-In Date";e={$_.LastLogonDate}},@{n="Manager Name";e={$_.ManagerName[0]}},@{n="Employee Type";e={$_.EmployeeType}},@{n="Disable Date";e={$_.disableDate[0]}} | Sort-Object -Property "Last Sign-In Date" | ConvertTo-HTML | Out-String
    $HTML = $HTML.replace("<table>", '<table cellpadding="10">')
    $Body = @"
    <p><h4>Hi All, <br>Please review the information below.</br></h4></p>
    <p>You may manage one of the following user accounts which WILL BE OFFBOARDED due to inactivity.</p>
    <p>The below user(s) should sign into <a href="https://okta.com/">okta.com</a> today to prevent being disabled and offboarded. If the user is no longer Employed please IMMEDIATELY submit a Service-Now ticket.</p>
    <p><u>Accounts inactive for 55+ days:</u><br>These accounts are scheduled to be deactivated.</br></p>
    <p>$HTML</p>
    <p>If you have any questions or issues with the above information please reach out to IT.</p>
    <p>Thank you!</p>
"@
    $FullMessage = New-Object System.Net.Mail.MailMessage $From, $SendToList, $Subject, $Body
    $FullMessage.IsBodyHtml = $True

    return $FullMessage

}

# Get list of active employees from Workday and store in $WDUsers
$WorkdayURL = $env:WORKDAYURL
$WDUsername = $env:WORKDAYUSERNAME
$KeyFile = $env:KEYFILE
$Password = Get-Content $KeyFile | ConvertTo-SecureString
$WebClient = new-object System.Net.WebClient
$WebClient.Credentials = new-object System.Net.NetworkCredential($WDUsername, $Password)
$WebPage = $webclient.DownloadString($WorkdayURL) | Out-File $env:TMP\pscsv.csv
$WDUsers = Import-Csv $env:TMP\pscsv.csv

# Ensure Workday User data is populated. Alert & Exit if not
if ($WDUsers.count -lt 2000) {
    $WDUserCount = $WDUsers.count
    $ErrorMessage = "Insufficient Workday User data for AD Stale Account Automation. Workday User Count: $WDUserCount"
    AdminAlertEmail $ErrorMessage
    AdminAlertSlack $ErrorMessage
    Exit
}

$WDUsers | % {
    # For each Workday user, check 'On Leave' status and update user object in AD if True
    if ($_.on_leave -eq "1") {
        Set-ADUser -identity $_.samaccountname -Add @{'extensionAttribute7'="On Leave"} -ErrorAction SilentlyContinue
        Set-ADUser -identity $_.samaccountname -Replace @{'extensionAttribute7'="On Leave"}
    }
}

# Sleep 60 so that Active Directory user object updates have time to take effect
sleep 60

$30DayInactiveUsers = [System.Collections.ArrayList]::new()
$55DayInactiveUsers = [System.Collections.ArrayList]::new()
$31DaysAgo = (get-date).adddays(-31)
$56DaysAgo = (get-date).adddays(-56)
$61DaysAgo = (get-date).adddays(-61)
$OrgUnits = $env:ORGUNITS


$OrgUnits | % {

    # Collect users that have not signed in for 30+ days from specified OU & store in $ADUsersPerOU
    $ADUsersPerOU = Get-ADUser -SearchBase $_ -Filter {Enabled -eq $true -and lastlogondate -le $31DaysAgo} -properties lastlogondate,title,manager,extensionattribute7,cn,mail,employeeType

    $ADUsersPerOU | % {
        
        # Determine if user is out of scope for automation
        if (($_.extensionAttribute7 -eq "Delayed") -or ($_.extensionAttribute7 -eq "Processed by Automation Script") -or ($_.extensionAttribute7 -eq "On Leave") -or ($_.extensionAttribute7 -eq "Managed By Sailpoint")) {
            write-host "Skipping!"
        }

        else {
            
            # Set employee type field
            $_.employeeType = switch -Wildcard ($_) {
                {$_.DistinguishedName -like "*TEMP*"} {"Temp Employee"; Break}
                {$_.employeeType -like "*CONSULTANT*"} {"Consultant"; Break}
                {$_.employeeType -like "*CONTRACTOR*"} {"Contractor"; Break}
                default {"Full Time Employee"}
            }

            $lastLogonDate = $_.LastLogonDate
            $_.DisableDate = ($lastLogonDate).AddDays(60) | Get-Date -format "MM-dd-yyyy"

            if (!$_.Manager) {
                $_.ManagerName = "No Manager Listed"
            }
            else {
                $ManagerInfo = Get-ADObject -Identity $_.Manager -Properties mail,DisplayName | select mail,DisplayName
                $_.ManagerName = $ManagerInfo.DisplayName
                if ($ManagersEmail) {
                    $ManagersEmail += "," + $ManagerInfo.mail
                } else {
                    $ManagersEmail = $ManagerInfo.mail
                }
                
            }

            # Last logon date for user is 30 days ago
            if ($lastLogonDate -eq $31DaysAgo) {
                $30DayInactiveUsers.add($_) > $null
            }
            
            # Last logon date for user is 55 days ago
            elseif ($lastLogonDate -eq $56DaysAgo) {
                $55DayInactiveUsers.add($_) > $null
            }

            # Last logon date for user is 60+ days ago. Disable & move user object
            if ($lastLogonDate -le $61DaysAgo) {
                Set-ADUser -identity $_.samaccountname -Enabled $false
                Move-adobject -Identity $_.DistinguishedName -TargetPath $env:ARCHIVEDOU
            }
	    
        }       
    }
}

$EntAppsEmail = $env:ENTAPPSEMAIL
$SecurityEmail = $env:SECURITYEMAIL
$ITServicesEmail = $env:ITEMAIL
$SMTPHost = $env:SMTP
$EmailClient = New-Object System.Net.Mail.SmtpClient $SMTPHost

# Compose & send email for users inactive 30 days
if ($30DayInactiveUsers) {
    $30SendTo = "$EntAppsEmail,$SecurityEmail,$ManagersEmail"
    $30DayMessage = Compose30DayEmail $30DayInactiveUsers $30SendTo
    $EmailClient.Send($30DayMessage)
}

# Compose & send email for users inactive 55 days
if ($55DayInactiveUsers) {
    $55SendTo = "$EntAppsEmail,$SecurityEmail,$ManagersEmail,$ITServicesEmail"
    $55DayMessage = Compose55DayEmail $55DayInactiveUsers $55SendTo
    $EmailClient.Send($55DayMessage)
}
