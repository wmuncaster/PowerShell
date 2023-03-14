<#
Author: Walter Muncaster
Purpose: Generate detailed CSV report on all Shared Mailboxes in Office365 env.
Output returned: 
DisplayName, UserPrincipalName, PrimarySize, PrimarySizeBytes, PrimaryItemCount, 
ArchiveSize, ArchiveSizeBytes, ArchiveItemCount, ArchiveExpand, LitigationHoldEnabled
#>

# Establish connection w/ O365 env
$Creds = Get-Credential
$O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $Creds -Authentication Basic -AllowRedirection
Connect-MsolService -Credential $Creds -ErrorVariable ConnectingMSOLServiceError
Import-PSSession $O365Session

# Get all Shared Mailboxes in O365 env
Write-Host "Getting Mailboxes..."
$Mailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited          
$FileName = (Get-Date).ToString('MM-dd-yyyy_hh-mm-ss')
$Counter = 1
$TotalCount = $Mailboxes.count

$Mailboxes | % {
    
    try { $MailboxStats = Get-MailboxStatistics $_.UserPrincipalName }
    catch { $MailboxStats = $null }

    if ($MailboxStats) {
        $PrimarySize = $MailboxStats.TotalItemSize
        $PrimarySizeBytes = ($PrimarySize).ToString().Split("(")[-1].Split()[0]
        $PrimaryItemCount = $MailboxStats.ItemCount
    }
    else {
        $PrimarySize = "No Mailbox"
        $PrimarySizeBytes = "No Mailbox"
        $PrimaryItemCount = "No Mailbox"
    }

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

    $_ | select DisplayName,UserPrincipalName,@{n="PrimarySize";e={$PrimarySize}},@{n="PrimarySizeBytes";e={$PrimarySizeBytes}},@{n="PrimaryItemCount";e={$PrimaryItemCount}},@{n="ArchiveSize";e={$ArchiveSize}},@{n="ArchiveSizeBytes";e={$ArchiveSizeBytes}},@{n="ArchiveItemCount";e={$ArchiveItemCount}},@{n="ArchiveExpanded";e={$ArchiveExpanded}},LitigationHoldEnabled | export-csv -path "$env:FILEPATH\$FileName.csv" -append
    write-host "$Counter of $TotalCount completed."
    $Counter++

} 
