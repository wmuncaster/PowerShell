<#
Author: Walter Muncaster
Purpose: Generate CSV report on all Mobile Devices that are connected to O365 env.
Filter for devices that have synced within last 90 days and operating system is 'iOS' or 'Android'.
Returned output: UserPrincipalName, ActiveMobileDeviceCount, TotalMobileDeviceCount, LastSyncDates, DeviceOS
#>

$ErrorActionPreference="Stop"

# Establish connection w/ O365 env
$Creds = Get-Credential
$O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $Creds -Authentication Basic -AllowRedirection
Connect-MsolService -Credential $Creds -ErrorVariable ConnectingMSOLServiceError
Import-PSSession $O365Session

# Get all enabled user objects in AD
Write-Host "Getting Users..."
$ADUsers = Get-ADUser -filter {enabled -eq $true} -Searchbase $env:SEARCHBASE
$UsersCount = $ADUsers.count
$Counter = 1
$TodaysDate = Get-Date -Format "MM-dd-yyyy"
$90DaysAgo = (Get-Date).AddDays(-90)

$ADUsers | % {
    
    Write-Host "Processing $Counter of $UsersCount"
    $Counter++
    
    # Collect synced device data for user, skip iteration if user not found
    try { $AllDeviceData = Get-MobileDeviceStatistics -Mailbox $_.SamAccountName }
    catch { return }
    
    $RelevantDeviceData = $AllDeviceData | ? {($_.LastSuccessSync -ge $90DaysAgo) -and ($_.DeviceOS -like "*iOS*" -or $_.DeviceOS -like "*Android*")}
    $TotalDeviceCount = @($AllDeviceData | ? {$_.DeviceOS -like "*iOS*" -or $_.DeviceOS -like "*Android*"}).count
    $ActiveDeviceCount = @($RelevantDeviceData).count
    $RelevantDeviceData | % {$_ | Add-Member -NotePropertyName LastSuccessSyncFormatted -NotePropertyValue $_.LastSuccessSync.ToString("MM/dd/yyyy")}
    $LastSyncDates = ($RelevantDeviceData.LastSuccessSyncFormatted -join ";")
    $DeviceOS = ($RelevantDeviceData.DeviceOS -join ";")

    Write-Host "Writing to file!"

    $_ | select UserPrincipalName,@{n="Active Mobile Device Count";e={$ActiveDeviceCount}},@{n="Total Mobile Device Count";e={$TotalDeviceCount}},@{n="Last Sync Dates Within 90 Days";e={$LastSyncDates}},@{n="Device OS";e={$DeviceOS}} | Export-Csv -path "$env:FILEPATH\ConnectedDevices_$TodaysDate.csv" -append

}
