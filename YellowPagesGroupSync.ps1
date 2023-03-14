<#
Author: Walter Muncaster
Desc: Query Yellowpages for active services & then create/update Active Directory Distribution Groups. 
(Member Addition/Removal, Service/Group Deprecation, Group Creation, Group Attribute Updates)
#>

Import-Module ActiveDirectory
[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"

$logName = (get-date -format "MM-dd-yyyy") + "-log.txt"
$log = $env:LOGLOCATION + $logName
$defaultMembers = $env:DEFAULTMEMBERS

function AdminAlert($body){
 if ($body -notlike "*Could not create SSL/TLS secure channel.*") {
   write-host "Alerting Admin"
   $from = $env:COMPUTERNAME
   $to = $env:TO
   $smtphost = $env:SMTPHOST
   $timeout = "60"
   $subject = "YellowPages Group Automation Issue on $env:computername"
   Send-MailMessage -SmtpServer $smtphost -From $from -To $to -Body $body -Subject $subject
 }
}

#Slack Channel admin alert for Slack admins to review potential issues in code/groups/etc
function AdminAlertSlack($body){
 if($body -notlike "*Could not create SSL/TLS secure channel.*") {
   $JSONStart = "{`"text`": `""
   $JsonEnd = "`"}"
   $fullbody = $jsonstart + $body + $jsonend
   Invoke-WebRequest -Method POST -Uri $env:SLACKCHANNEL -ContentType application/json -Body $fullbody
 }
}



#Get all YellowPage Service data - $YPData
$YPData = Invoke-RestMethod -Method GET -Uri $env:YP
$ypcount = $YPData.count
Write-output "$(Get-Date);Processing YP Data $ypcount" > $log

#Get all current AD groups - $ADGroups
$ADGroups = Get-ADGroup -Filter * -SearchBase $env:GROUPSEARCHBASE -Properties extensionAttribute2,extensionAttribute15,members,description
$adgroupscount = $ADGroups.count
Write-output "$(Get-Date);Processing AD Groups $adgroupscount" >> $log

#Get all AD Users in "UserBase" - $ADUsers
$ADUsers = Get-ADUser -filter {enabled -eq $true} -Searchbase $env:USERSEARCHBASE
$aduserscount = $ADUsers.count
Write-output "$(Get-Date);Processing AD Users $aduserscount" >> $log


#Ensure YPData, ADGroups, ADUsers have valid data. If not, send email/Slack alert to admins & exit.
if (($YPData.count -lt "2000") -or ($ADGroups.count -lt "2000") -or ($ADUsers.count -lt "2500")) {
 $errorMessage = "Insufficient data for YellowPages group automation. YellowPages Data Count: " + $YPData.count + ", AD Groups Count: " + $ADGroups.count + ", AD Users Count: " + $ADUsers.count
 AdminAlert($errorMessage)
 AdminAlertSlack($errorMessage)
 exit
}


# Create list of current AD Groups(hashtables) w/ members(samAccountName), name, YP ID, description, expiry - $ADGroupsWithMembers
Write-output "$(Get-Date);Creating & Processing ADGroupsWithMembers" >> $log
$ADGroupsWithMembers = [System.Collections.ArrayList]::new()
$ADGroups | % {
 $error.clear()
 $errNum=0

 $groupInfo = @{"name"=$_.name; "id"=$_.extensionAttribute2; "expiry"=$_.extensionAttribute15; "desc"=$_.description; "members"=[System.Collections.ArrayList]::new()}

 foreach ($member in $_.members) {
   if ($member -notin $defaultMembers) {
     foreach ($user in $ADUsers) {
       if ($user.DistinguishedName -eq $member) {
         $groupInfo["members"].Add($user.SamAccountName) > $null
         break
       }
     }
   }
 }

 $ADGroupsWithMembers.Add($groupInfo) > $null

 while ($errNum -lt ($error.count)) {
   write-output "$(Get-Date);$error[$errnum]" >> $log
   $errNum ++
 }
}


#Create hashtable of current YP Service IDs w/ members(samAccountName), group description - $YPServices
Write-output "$(Get-Date);Creating & Processing YPServices" >> $log
$YPServices = @{}
$YPData | % {
 if (-not $_.archived) {
   $members = [System.Collections.ArrayList]::new()

   $members.Add($_.owner) > $null
   if ($_.secondary_owner) {
     $members.Add($_.secondary_owner) > $null
   }

   if ($_.dependents) {
     foreach ($dependent in $_.dependents) {
       foreach ($service in $YPData) {
         if ($service._id -eq $dependent) {
           $members.Add($service.owner) > $null
           if ($service.secondary_owner) {
             $members.Add($service.secondary_owner) > $null
           }
           break
         }
       }
     }
   }

   $serviceData = @{"members"=$members | select -unique; "desc"=$_.name + " - " + $env:YPPREFIX + $_._id}
   $YPServices.Add($_._id, $serviceData)
 }
}


# Create list of Groups to Create (IDs) - $groupsToCreate
$groupsToCreate = $YPServices.keys | ? {$_ -notin $ADGroups.extensionAttribute2} | select $_

# Iterate through groupsToCreate & create group, add members, set properties
Write-output "$(Get-Date);Creating & Processing groupsToCreate" >> $log
$groupsToCreate | % {
 $error.clear()
 $errNum=0
 $serviceId = $_
 $groupName = "YP-" + $_ + "-clients"

 # create group & set properties
 Write-output "$(Get-Date);Attempting to create group $_" >> $log
 New-ADGroup -name $groupName -groupCategory distribution -groupScope global -description $YPServices[$_]["desc"] -displayName $groupName -otherAttributes @{"extensionAttribute2"=$_;"mail"=($groupName);"info"="Automated YP Distro Groups";"msExchHideFromAddressLists"=$true;"msExchRequireAuthToSendTo"=$false} -path $env:DESTINATION

 # add members
 try {
   Write-output "$(Get-Date);Attempting to add array of members to group $_" $YPServices[$_]["members"] >> $log
   Add-ADGroupMember -identity $groupName -members $YPServices[$_]["members"]
 }
 catch {
   Write-output "$(Get-Date);Could not add entire array of members, will attempt to add members individually." >> $log
   foreach ($member in $YPServices[$serviceId]["members"]) {
     Write-output "$(Get-Date);Attempting to add $member to $serviceId." >> $log
     try {Add-ADGroupMember -identity $groupName -members $member}
     catch {Write-Output "$(Get-Date);Could not add $member. User does not exist." >> $log}
   }
 }

 # add default members (NOC Dist List & Slack Alert Channel)
 Set-ADGroup -identity $groupName -add @{"member"=$defaultMembers}

 while ($errNum -lt ($error.count)) {
   write-output "$(Get-Date);$error[$errnum]" >> $log
   $errNum ++
 }
}


# Iterate through ADGroupsWithMembers (current AD groups) & verify the service is currently active in YP.
# If so, update group members & description in AD. If not, set 2 week expiry & delete group from AD.
Write-output "$(Get-Date);Updating groups in ADGroupsWithMembers" >> $log
$ADGroupsWithMembers | % {
 $error.clear()
 $errNum=0
 $groupName = "YP-" + $_.id + "-clients"
 $serviceId = $_.id

 # service is active in YP
 if ($serviceId -in $YPServices.keys) {
   Write-output "$(Get-Date);YP Service $serviceId is active. Updating members & attributes." >> $log

   $groupMembers = $_.members
   $serviceMembers = $YPServices[$serviceId]["members"]
   $membersToAdd = $serviceMembers | ? {$_ -notin $groupMembers}
   $membersToRemove = $groupMembers | ? {$_ -notin $serviceMembers}

   # add members to AD group
   if ($membersToAdd) {
     Write-output "$(Get-Date);Attempting to add array of members to group $serviceId" $membersToAdd >> $log
     try {Add-ADGroupMember -identity $groupName -members $membersToAdd}
     catch {
       Write-output "$(Get-Date);Could not add entire array of members, will attempt to add members individually." >> $log
       foreach ($member in $membersToAdd) {
         Write-output "$(Get-Date);Attempting to add $member to $serviceId." >> $log
         try {Add-ADGroupMember -identity $groupName -members $member}
         catch {Write-Output "$(Get-Date);Could not add $member. User does not exist." >> $log}
       }
     }
   }

   # remove members from AD group
   if ($membersToRemove) {
     Write-output "$(Get-Date);Attempting to remove array of members from group $serviceId" $membersToRemove >> $log
     try {Remove-ADGroupMember -identity $groupName -members $membersToRemove -confirm:$false}
     catch {Write-Output "$(Get-Date);Could not remove full array of members." >> $log}
   }

   # update group description
   if ($_.desc -ne $YPServices[$_.id]["desc"]) {
     Write-output "$(Get-Date);Updating description of group $serviceId.">> $log
     try {Set-ADGroup -identity $groupName -replace @{"description"=$YPServices[$_.id]["desc"]}}
     catch {Write-Output "$(Get-Date);Could not update group description." >> $log}
   }
 }

 # service is inactive in YP
 else {

   # expiry present, if <= today; delete group
   if ($_.expiry) {
     $today = get-date -format "MM/dd/yyyy"
     if (($today -eq $_.expiry) -or ($_.expiry -lt $today)) {
       Write-output "$(Get-Date);YP Service $serviceId is NOT active. Expiration date is <= today. Will attempt to delete AD group." >> $log
       try {Remove-ADGroup -identity $groupName -confirm:$false}
       catch {Write-Output "$(Get-Date);Could not delete AD group." >> $log}
     }
     else {
       Write-output "$(Get-Date);YP Service $serviceId is NOT active. Expiration date is not today. Skipping AD group." >> $log
     }
   }

   # add expiry to AD group
   else {
     Write-output "$(Get-Date);YP Service $serviceId is NOT active. Setting Expiry date 2 weeks from today." >> $log
     $expiryDate = (get-date).addDays(14) | get-date -format "MM/dd/yyyy"
     try {Set-ADGroup -identity $groupName -replace @{"extensionAttribute15"=$expiryDate}}
     catch {Write-Output "$(Get-Date);Could not set Expiry date for AD group." >> $log}
   }
 }

 while ($errNum -lt ($error.count)) {
   write-output "$(Get-Date);$error[$errnum]" >> $log
   $errNum ++
 }
}
