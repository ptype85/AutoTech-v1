﻿#Peter Endacott 2018
#Import AD Password Reset Requests from Service Desk
#Used with AutoTech

#set execution policy
#Set-ExecutionPolicy RemoteSigned

$date = Get-Date -Format ddMMyyhhmmss
invoke-sqlcmd -inputfile "C:\AutoTech\ActiveDirectory\PasswordReset\PResetImp.sql" -ServerInstance "dbserver.contoso.com" | Export-CSV "C:\AutoTech\ActiveDirectory\PasswordReset\Files\Input\RequestImport$date.csv" -NoTypeInformation

$Dircheck = Get-ChildItem C:\AutoTech\ActiveDirectory\PasswordReset\Files\Input | Measure-Object
if ($Dircheck -ne 0) {

##Action Variable
$ua = 1
$ea = 1

##Password Generator

$PArray = 'Orange','Apple','Raspberry','Strawberry','Grape','Loganberry','Lemon','Cherry'

##Email Settings

$fromAddr = "ServiceDesk@contoso.com" # Enter the FROM address for the e-mail alert
$smtpsrv = "172.0.0.10" # Enter the FQDN or IP of a SMTP relay

##Input Read
$at = Import-Csv "C:\AutoTech\ActiveDirectory\PasswordReset\Files\Input\RequestImport$date.csv"
$wo = @()
$wo += $at.WORKORDERID
$un = @()
$un += $at.VALUE1

##Counter Reset
$tc = 0

if ($ua = 1){
$wo | ForEach-Object{
$cu = $un[$tc]
		$userTest = get-aduser $cu
		if ($userTest -eq $null){
		$rp = "INVALID USER"
		$ue = "User not found"
		$wo[$tc],$cu,$rp,$ue  -join ',' | Out-File -FilePath "C:\AutoTech\ActiveDirectory\PasswordReset\Files\Input\Fail\Output$date.csv" -Append;
		$tc ++
		}
	else
	{
		$ru = Import-Csv C:\AutoTech\ActiveDirectory\Config\RestrictedUN.csv #Load restricted User List
		if ($ru -match $cu){
		$rp = "INVALID USER"
		$ue = "Restricted user"
		$wo[$tc],$cu,$rp,$ue  -join ',' | Out-File -FilePath "C:\AutoTech\ActiveDirectory\PasswordReset\Files\Input\Fail\Output$date.csv" -Append;
		$tc ++
		}

	
	
$ud = get-aduser $cu -Properties EmailAddress
$ue = $ud.EmailAddress

$PValue1 = $PArray[(Get-Random -Maximum ([array]$PArray).count)]
$PValue2 = Get-Random -Minimum 100 -Maximum 999
$rp = "$PValue1$PValue2"

Set-ADAccountPassword -Identity $cu -reset -NewPassword (ConvertTo-SecureString -AsPlainText "$rp" -Force)

$wo[$tc],$cu,$rp,$ue  -join ',' | Out-File -FilePath "C:\AutoTech\ActiveDirectory\PasswordReset\Files\Input\Output$date.csv" -Append;

$tc ++
}
#########Email Notifications

##7Zip Settings

$process = "c:\Program Files\7-Zip\7z.exe"

##InputRead
$notiRead = Import-Csv "C:\AutoTech\ActiveDirectory\PasswordReset\Files\Input\Output$date.csv" -Header Req,UserN,Pword,Email
$req = @()
$req += $notiRead.Req
$Usn = @()
$Usn += $notiRead.UserN
$Pwd = @()
$Pwd += $notiRead.Pword
$Eml = @()
$Eml += $notiRead.Email

$ot = 0

## Create Var's per Workorder
$req | ForEach-Object{
$Request = $req[$ot]
$UserName = $usn[$ot]
$Password = $pwd[$ot]
$UEmailAd = $eml[$ot]

##Create Text File
$FileContent = @"
Password: $Password
"@
$FileContent | Out-File -FilePath "C:\AutoTech\ActiveDirectory\PasswordReset\Files\Txt\VDICredentials$Request.txt" -Encoding ASCII
Start-Sleep -Seconds 20

##Create Zip File
$sourceFile = "C:\AutoTech\ActiveDirectory\PasswordReset\Files\Txt\VDICredentials$Request.txt"
$destinationFile = "C:\AutoTech\ActiveDirectory\PasswordReset\Files\Zip\$Request.zip"
$rp = [System.Web.Security.Membership]::GeneratePassword(10,0)
$ZipPassword = $rp
Start-Process $process -ArgumentList "a $destinationFile $sourceFile -p$ZipPassword"
Start-Sleep -Seconds 20

##Send Zip File

if ($ea -eq 1 -and $req[$ot] -ne $null){

$toAddr = $UEmailAd
$body = @"
Your VDI Password has been reset. Please find your Virtual Desktop password in the attached Zip file.
<BR>
<BR>
For Security purposes, the Zip file is encrypted with a password. You will recieve this password on a separate email shortly.
<BR>
<BR>
If you did not request this password reset, please notify ServiceDesk@contoso.com.
<BR>
<BR>
Thank you,

"@

Send-MailMessage -To $toAddr -From $fromAddr -Subject "$Request - VDI Password Reset" -Body $body -SmtpServer $smtpsrv -BodyAsHtml -Attachments $destinationFile

$Request,$UserName,$Password,$UEmailAd -join ',' | Out-File -FilePath "C:\AutoTech\ActiveDirectory\PasswordReset\Files\Email\EmailAlertOutput$date.csv" -Append;
}
Start-Sleep -Seconds 20

##Send Zip Password

$body2 = @"
The password for zip file conatining your Virtual Desktop password is as follows.
<BR>
<BR>
<b>$rp</b> - If copying this text, please ensure that the space is removed.
<BR>
<BR>
Thank you.

"@

Send-MailMessage -To $toAddr -From $fromAddr -Subject "$Request - VDI Password reset" -Body $body2 -SmtpServer $smtpsrv -BodyAsHtml

$ot ++
}

}
}
else
{
exit
}
}
else
{
Exit
}
