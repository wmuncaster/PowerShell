<#
Author: Walter Muncaster
Purpose: Generate detailed CSV report on all User Mailboxes in Office365 env.
Output returned: 
DisplayName, UserPrincipalName, PrimarySize, PrimarySizeBytes, PrimaryItemCount, 
ArchiveSize, ArchiveSizeBytes, ArchiveItemCount, ArchiveExpand, LitigationHoldEnabled, AccountEnabled
#>

# Establish connection w/ O365 env
$Creds = Get-Credential
$O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $Creds -Authentication Basic -AllowRedirection
Connect-MsolService -Credential $Creds -ErrorVariable ConnectingMSOLServiceError
Import-PSSession $O365Session
Connect-AzureAd -Credential $Creds

# Get all User Mailboxes in O365 env
Write-Host "Getting Users..."
$Users = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize:Unlimited
$FileName = (Get-Date).ToString('MM-dd-yyyy_hh-mm-ss')
$Counter = 1
$TotalCount = $Users.count

$Users | % {

    try { $AccountEnabled = (Get-AzureADUser -ObjectId $_.UserPrincipalName).AccountEnabled }
    catch { $AccountEnabled = $null }
    
    $MailboxStats = Get-MailboxStatistics $_.UserPrincipalName
    $PrimarySizeBytes = ($MailboxStats.TotalItemSize).ToString().Split("(")[-1].Split()[0]

    if ($_.ArchiveStatus -eq "Active") {
        $MailBoxStatsArchived = Get-MailboxStatistics -archive $_.UserPrincipalName
        $ArchiveSize = $MailBoxStatsArchived.TotalItemSize
        $ArchiveSizeBytes = ($MailBoxStatsArchived.TotalItemSize).ToString().Split("(")[-1].Split()[0]
        $ArchiveItemCount = $MailBoxStatsArchived.ItemCount
        $ArchiveExpanded = $_.AutoExpandingArchiveEnabled

    }
    else {
        $ArchiveSize = "No Archive"
        $ArchiveSizeBytes = "No Archive"
        $ArchiveItemCount = "No Archive"
        $ArchiveExpanded = "No Archive"
    }

    $_ | select DisplayName,UserPrincipalName,@{n="PrimarySize";e={$MailboxStats.TotalItemSize}},@{n="PrimarySizeBytes";e={$PrimarySizeBytes}},@{n="PrimaryItemCount";e={$MailboxStats.ItemCount}},@{n="ArchiveSize";e={$ArchiveSize}},@{n="ArchiveSizeBytes";e={$ArchiveSizeBytes}},@{n="ArchiveItemCount";e={$ArchiveItemCount}},@{n="ArchiveExpanded";e={$ArchiveExpanded}},LitigationHoldEnabled,IsDirSynced,@{n="AccountEnabled";e={$AccountEnabled}} | export-csv -path "c:\scripts\output\DSS-HULU-Mig-Count_$FileName.csv" -append
    write-host "$Counter of $TotalCount completed."
    $Counter++

} 
