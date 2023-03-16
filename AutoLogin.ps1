
<#
Author: 		Walter Muncaster
Purpose: 		Set autologin items for login without credentials 
				Optional immediate forced reboot
#>


$dommainname=""
$username="administrator"
$password="yourPasswordHere"
$immediatereboot = "N"

$regapproot = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
 
Set-ItemProperty -Path $regapproot -name "DefaultDomainName"-type string -value $dommainname 
Set-ItemProperty -Path $regapproot -name "DefaultUserName" 	-type string -value $username
Set-ItemProperty -Path $regapproot -name "DefaultPassword" 	-type string -value $password
Set-ItemProperty -Path $regapproot -name "AutoAdminLogon" 	-type string -value "1"
 
if ( $immediatereboot -eq "Y") { 
    Restart-Computer -Force 
}

    

