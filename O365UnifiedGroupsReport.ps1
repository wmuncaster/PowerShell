<#
Author: Walter Muncaster
Purpose: Generate detailed CSV report of all O365 unified groups.
Returned Output: PrimarySMTPAddress,Name,DisplayName,RequireSenderAuthenticationEnabled,
AcceptMessagesOnlyFromSendersOrMembers,ManagedBy,Members
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

$UnifiedGroups = Get-UnifiedGroup | select PrimarySMTPAddress,managedby,name,displayname,RequireSenderAuthenticationEnabled,AcceptMessagesOnlyFromSendersOrMembers
$Counter = 1
$TotalCount = $UnifiedGroups.count

$UnifiedGroups | % {

    $Members = Get-UnifiedGroupLinks -identity $_.primarysmtpaddress -linktype members | select -ExpandProperty windowsliveid
    $Managers = [System.Collections.ArrayList]::new()
    
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
    
    $_ | select PrimarySMTPAddress,Name,DisplayName,RequireSenderAuthenticationEnabled,AcceptMessagesOnlyFromSendersOrMembers,@{n="ManagedBy";e={$Managers -join ";"}},@{n="Members";e={$Members -join ";"}} | export-csv "$env:EXPORTPATH\unifiedGroups.csv" -append
    Write-Host "$Counter of $TotalCount"
    $Counter++
} 
