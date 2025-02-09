﻿#Peter Endacott 2018
#Active Directory User Creation Automation Script
#Used with AutoTech

#set execution policy
Set-ExecutionPolicy RemoteSigned

#Load "System.Web" assembly in PowerShell console 
#[Reflection.Assembly]::LoadWithPartialName("System.Web")

#Import AD Module
import-module activedirectory

#Import Quest Active Roles Module
Add-PSSnapin Quest.ActiveRoles.ADManagement

#Import Exchange Mangement Tools
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin

#Calling GeneratePassword Method 
#[System.Web.Security.Membership]::GeneratePassword(10,0)

###########################

##Debug Switches

$ir = 1 #Input Read
$uc = 1 #User Creation
$ea = 1 #Email Alerts
$ua = 1 #User Availibility Check
$mb = 1 #Mailbox Creation
$cc = 1 #Contact Creation

###########################

##Variable Store

#$at = #AllTasks
#$un = #UserName
#$fn = #FullName
#$fc = #FullName Counter
#$1n = #FirstName
#$2n = #LastName
#$2i = #LastName Initial
#$iv = #Initial Variable
#$uc = #Username Check
#$tu = #Template User
#$tc = #Template Counter
#$pa = #User Path (Container)
#$rp = [System.Web.Security.Membership]::GeneratePassword(10,0)#Random Password
#$ec = email csv
#$mdb = Mailbox database
$domain = "contoso.com\" #Default Domain
#$ymb = Yes Mailbox
#$exea = External Email Address
#$ecc = Email Cycle Counter
$path1 = "\\fileshare\profile$\" #Default Profle location
$path2 = "\\fileshare\home$\" #Default Home Drive location


###########################

##Password Generator

$PArray = 'Orange','Apple','Raspberry','Strawberry','Grape','Loganberry','Lemon','Cherry'

##Email Settings

$fromAddr = "AutoTech@contoso.com" # Enter the FROM address for the e-mail alert
$toAddr = "is@contoso.com" # Enter the TO address for the e-mail alert
$smtpsrv = "172.0.0.10" # Enter the FQDN or IP of a SMTP relay

###########################

##Input Read
$date = Get-Date -Format ddMMyy
$path = "C:\AutoTech\ActiveDirectory\Validate\Files\Passed\UserCheckPassed$date.csv"
if ([System.IO.File]::Exists($path)){
$at = Import-Csv "C:\AutoTech\ActiveDirectory\Validate\Files\Passed\UserCheckPassed$date.csv" -Header WORKORDERID,VALUE1,VALUE2,VALUE3,VALUE4
$wo = @()
$wo += $at.WORKORDERID
$fn = @()
$fn += $at.VALUE1	#Read Full Name from Input
$tu = @()
$tu += $at.VALUE2	#Read Template User from Input
$exea = @()
$exea += $at.VALUE3	#Read email address from Input


##Counter Reset
$tc = 0
$fc = 0
$ecc = 0
}Else
{
Exit
}

##User Create

if ($ua = 1){
$wo | ForEach-Object{
get-aduser $tu[$tc]												#Get Template User
$pa = (Get-AdUser $tu[$tc]).distinguishedName.Split(',',3)[2]	#Get Template User OU Data

$iv = 1
$1n = ($fn[$fc].split(' ')[0])	#Split first name from full name
$2n = ($fn[$fc].split(' ')[1])	#Split second name from full name
$2i = $2n.substring(0,$iv)	#Split second initial from second name
$un = "$1n$2i"				#create username from first name and second intial value


$uc = Get-ADUser -LDAPFilter "(SAMAccountName=$un)"			#Check UserName Availability
If ($uc -ne $null){ #02/08/2018 You switched this line from a $Null to a 0. Might resolve the 7 times cycling
Do{
						#$uc = Get-ADUser -LDAPFilter "(SAMAccountName=$un)"
						$iv = $iv+1
						$2i = $2n.substring(0,$iv)
						$un = "$1n$2i"
						$uc = Get-ADUser -LDAPFilter "(SAMAccountName=$un)"
						}
						Until ($uc -eq $null -or $iv -eq "7")
						}
						
$PValue1 = $PArray[(Get-Random -Maximum ([array]$PArray).count)]
$PValue2 = Get-Random -Minimum 100 -Maximum 999
#			
$rp = "$PValue1$PValue2"		#[System.Web.Security.Membership]::GeneratePassword(10,0)
New-ADUser -SamAccountName $un -Instance "$tu[$tc]" -path $pa -Name "$2n, $1n" -AccountPassword (ConvertTo-SecureString -AsPlainText "$rp" -Force) -ProfilePath "\\fileserver\profile$\$un\" -Enabled $true -ChangePasswordAtLogon $true -GivenName "$1n" -SurName "$2n" -ScriptPath "drivemappings2.bat" -HomeDrive "H" -HomeDirectory "\\fileserver\home$\$un\" -UserPrincipalName "$un@contoso.com"
Start-Sleep 30 #Delay to allow New AD User object to propogate
(get-qaduser $tu[$tc]).memberof | Add-QADGroupMember -Member "$un"

#Create Network Folders and Apply Permissions (Added 17/12/2018 for explicit folder creation)

#Profile Drive
$pathUser1 = "$path1$un\"
New-Item $pathUser1 -ItemType Directory
$FullControl = [System.Security.AccessControl.FileSystemRights]"FullControl"
$userR = New-Object System.Security.Principal.NTAccount($un)
$type = [System.Security.AccessControl.AccessControlType]::Allow
$inheritanceFlag = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit, ObjectInherit"
$propagationFlag = [System.Security.AccessControl.PropagationFlags]::None
$accessControlEntryR = New-Object System.Security.AccessControl.FileSystemAccessRule @($userR, $FullControl, $inheritanceFlag, $propagationFlag, $type)
$objACL = Get-ACL $pathUser1
$objACL.AddAccessRule($accessControlEntryR)
Set-ACL $pathUser1 $objACL

#Home Drive
$pathUser2 = "$path2$un\"
New-Item $pathUser2 -ItemType Directory
$FullControl = [System.Security.AccessControl.FileSystemRights]"FullControl"
$userR = New-Object System.Security.Principal.NTAccount($un)
$type = [System.Security.AccessControl.AccessControlType]::Allow
$inheritanceFlag = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit, ObjectInherit"
$propagationFlag = [System.Security.AccessControl.PropagationFlags]::None
$accessControlEntryR = New-Object System.Security.AccessControl.FileSystemAccessRule @($userR, $FullControl, $inheritanceFlag, $propagationFlag, $type)
$objACL = Get-ACL $pathUser2
$objACL.AddAccessRule($accessControlEntryR)
Set-ACL $pathUser2 $objACL

#Mailbox creation

if ($2i -match "[a-m]")
{
$mdb = "link.contoso.com\Users1 Mailbox Database"
}
else
{
$mdb = "link.contoso.com\Users2 Mailbox Database"
}
Start-Sleep 20 #Delay to allow User object to propogate to all DC's

enable-mailbox -Identity $domain$un  -database $mdb

#OLDMETHOD$ymb = get-mailbox -Identity $domain$un | Select-Object PrimarySmtpAddress

#$SMTPResult = Get-ADUser -Identity $un -Properties ProxyAddresses | select -ExpandProperty ProxyAddresses | ? {$_ -clike "SMTP:*"}
#$ymb = $SMTPResult[5..40] -join ''

##Create contact from External Email (VF or TP)

New-MailContact -ExternalEmailAddress $exea[$ecc] -Name "$1n $2n VF" -OrganizationalUnit "contoso.com/Employees/"

##Set Forwarding from contoso.com to external

Start-Sleep 20 #Delay to allow New Contact object to propogate

Set-Mailbox $domain$un -ForwardingAddress $exea[$ecc]

###################


$date = Get-Date -Format ddMMyy

$wo[$fc],$fn[$fc],$un,$rp,$exea[$ecc] -join ',' | Out-File -FilePath "C:\AutoTech\ActiveDirectory\Create\Files\Output$date.csv" -Append;
$fc = $fc + 1
$tc = $tc + 1
$ecc = $ecc + 1
}}
##########################

##Email Alert
$ec = Import-Csv "C:\AutoTech\ActiveDirectory\Create\Files\Output$date.csv" -Header Req,FullN,UserN,Pword,Email

if ($ea -eq 1 -and $ec -ne $null){

$date2 = Get-Date -DisplayHint Date

$body = @("
<center><table border=1 width=50% cellspacing=0 cellpadding=8 bgcolor=Black cols=5>
<tr bgcolor=White><td>Request ID</td><td>Name</td><td>Username</td><td>Password</td><td>Email</tr></td>")

$i = 0

do {
if($i % 2){$body += "<tr bgcolor=#D2CFCF><td>$($ec[$i].Req)</td><td>$($ec[$i].FullN)</td><td>$($ec[$i].UserN)</td><td>$($ec[$i].Pword)</td><td>$($ec[$i].Email)</td></tr>";$i++}
else {$body += "<tr bgcolor=#EFEFEF><td>$($ec[$i].Req)</td><td>$($ec[$i].FullN)</td><td>$($ec[$i].UserN)</td><td>$($ec[$i].Pword)</td><td>$($ec[$i].Email)</td></tr>";$i++}
}
while ($ec[$i] -ne $null)

$body += "</table></center>"

Send-MailMessage -To $toAddr -From $fromAddr -Subject "Info: $($ec.Count) User Accounts Created on $date2" -Body "$body" -SmtpServer $smtpsrv -BodyAsHtml
}
######
