<#
Author: Walter Muncaster
Purpose: Generate detailed CSV report of all Active Directory distribution groups.
Returned Output: PrimarySMTPAddress,Name,DisplayName,Alias,WhenCreated,WhenChanged,
RequireSenderAuthenticationEnabled,AllowedSenders,ManagedBy,Members
#>


Import-Module ActiveDirectory


$ErrorActionPreference = "Stop" 

# Function: Return user email from AD user name
function GetUserEmail($User) {
    $UserEmail = $null
    
    try {$UserEmail = get-recipient -identity $User | select -expandproperty primarysmtpaddress}
    catch {}
                
    if (!$UserEmail) {
        try {$UserEmail = get-azureaduser -filter "startswith(DisplayName,'$User')" | select -ExpandProperty userprincipalname}
        catch {}
    }

    if (!$UserEmail) {
        try {$UserEmail = get-aduser -Identity $User | select -ExpandProperty userprincipalname}
        catch {}
    }

    return $UserEmail
}


# Establish connection w/ O365 env
$Creds = Get-Credential
$O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $Creds -Authentication Basic -AllowRedirection
Connect-MsolService -Credential $Creds -ErrorVariable ConnectingMSOLServiceError
Import-PSSession $O365Session
Connect-AzureAd -Credential $Creds

$DistGroups = Get-DistributionGroup -filter {(recipienttype -eq "MailUniversalDistributionGroup")} -ResultSize unlimited | select PrimarySMTPAddress,managedby,name,displayname,alias,RequireSenderAuthenticationEnabled,acceptmessagesonlyfromsendersormembers,whenchanged,whencreated
$Counter = 1
$TotalCount = $DistGroups.count

$DistGroups | % {

    $Managers = [System.Collections.ArrayList]::new()
    $AllowedSenders = [System.Collections.ArrayList]::new()
    $Members = [System.Collections.ArrayList]::new()

    $GroupMembers = Get-DistributionGroupMember -identity $_.primarysmtpaddress -ResultSize unlimited | select PrimarySMTPAddress,name

    # For each member in group, check for email, otherwise lookup email and append to $Members list
    if ($GroupMembers) {
        foreach ($Member in $GroupMembers) {
            if ($Member.primarysmtpaddress) {
                $Member.add($member.primarysmtpaddress) > $null

            } else {
                $Email = GetUserEmail($Member.name)
                if ($Email) {
                    $Members.add($Email) > $null
                } else {
                    $Members.add($Member.name) > $null
                }
            }
        }
    }

    # For each Manager of group, check for email, otherwise lookup email and append to $Managers list
    if ($_.ManagedBy) {
        foreach ($Manager in $_.ManagedBy) {
            $Email = GetUserEmail($Manager)
            if ($Email) {
                $Managers.add($Email) > $null
            } else {
                $Managers.add($Manager) > $null
            }

        }
        
    }

    # For each Allowed Sender of group, check for email, otherwise lookup email and append to $AllowedSenders list
    if ($_.AcceptMessagesOnlyFromSendersOrMembers) {
        foreach ($Sender in $_.AcceptMessagesOnlyFromSendersOrMembers) {
            $Email = GetUserEmail($Sender)
            if ($Email) {
                $AllowedSenders.add($Email) > $null
            } else {
                $AllowedSenders.add($Sender) > $null
            }
        }
    }

    $_ | select PrimarySMTPAddress,Name,DisplayName,Alias,WhenCreated,WhenChanged,RequireSenderAuthenticationEnabled,@{n="AcceptMessagesOnlyFromSendersOrMembers";e={$AllowedSenders -join ";"}},@{n="ManagedBy";e={$Managers -join ";"}},@{n="Members";e={$Members -join ";"}} | export-csv "$env:EXPORTPATH\distributionGroups.csv" -append
    write-host "$Counter of $TotalCount"
    $Counter++
}
