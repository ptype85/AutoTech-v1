﻿#Peter Endacott 2018
#User Credentials Sender Script
#Used with AutoTech

#set execution policy
Set-ExecutionPolicy RemoteSigned

#Load "System.Web" assembly in PowerShell console 
[Reflection.Assembly]::LoadWithPartialName("System.Web")

##Debug Switches

$ir = 1 #Input Read
$ea = 1 #Email Alerts

##Email Settings

$fromAddr = "servicedesk@contoso.com" # Enter the FROM address for the e-mail alert
#$toAddr = "pe@contoso.com" # Enter the TO address for the e-mail alert
$smtpsrv = "172.0.0.10" # Enter the FQDN or IP of a SMTP relay

##7Zip Settings

$process = "c:\Program Files\7-Zip\7z.exe"

##Input Read
$date = Get-Date -Format ddMMyy
$path = "C:\AutoTech\ActiveDirectory\Create\Files\Output$date.csv"
if ([System.IO.File]::Exists($path)){
$at = Import-Csv C:\AutoTech\ActiveDirectory\Create\Files\Output$date.csv  -Header Req,FullN,UserN,Pword,Email

$wo = @()
$wo += $at.req
$fn = @()
$fn += $at.FullN
$un = @()
$un += $at.UserN
$pw = @()
$pw += $at.Pword
$em = @()
$em += $at.Email
}else
{
Exit
}

##Counter Reset

$ct = 0

## Create Var's per Workorder
$wo | ForEach-Object{
$UserName = $un[$ct]
$Password = $pw[$ct]
$FileName = $wo[$ct]
$FullName = $fn[$ct]
$UEmailAd = $em[$ct]

##Create Text File
$FileContent = @"
UserName: $UserName

Password: $Password
"@
$FileContent | Out-File -FilePath "C:\AutoTech\ActiveDirectory\Notify\Files\Txt\VDICredentials$FileName.txt" -Encoding ASCII
Start-Sleep -Seconds 20

##Create Zip File
$sourceFile = "C:\AutoTech\ActiveDirectory\Notify\Files\Txt\VDICredentials$FileName.txt"
$destinationFile = "c:\AutoTech\ActiveDirectory\Notify\Files\Zip\$FileName.zip"
$rp = [System.Web.Security.Membership]::GeneratePassword(10,0)
$ZipPassword = $rp
Start-Process $process -ArgumentList "a $destinationFile $sourceFile -p$ZipPassword"
Start-Sleep -Seconds 20

##Send Zip File

if ($ea -eq 1 -and $wo[$ct] -ne $null){

$toAddr = $UEmailAd
$body = @"
Dear $FullName,
<BR>
<BR>
Please find your credentials in the attached Zip file.
<BR>
<BR>
For Security purposes, the Zip file is encrypted with a password. You will recieve this password on a separate email shortly.
<BR>
<BR>
If you encounter any issues with using your credentials, please contact ServiceDesk@contoso.com, quoting your Request ID, $FileName.
<BR>
<BR>
In order to access the virtual desktop, you will need the VMWare Horizon software to be installed on your machine. This can be requested from "https://myitshop.internal.contoso.com/Horizon-View-Client-4-3" 
<BR>
<BR>
Thank you,

"@



Send-MailMessage -To $toAddr -From $fromAddr -Subject "$FullName User Credentials" -Body $body -SmtpServer $smtpsrv -BodyAsHtml -Attachments $destinationFile

$FullName,$UserName,$FileName,$ZipPassword,$UEmailAd -join ',' | Out-File -FilePath "C:\AutoTech\ActiveDirectory\Notify\Files\Email\EmailAlertOutput$date.csv" -Append;
}
Start-Sleep -Seconds 20

##Send Zip Password

$body2 = @"
Dear $FullName,
<BR>
<BR>
The password for zip file conatining your Desktop credentials is as follows.
<BR>
<BR>
<b>$rp</b> - If copying this text, please ensure that the space is removed.
<BR>
<BR>
Thank you.

"@

Send-MailMessage -To $toAddr -From $fromAddr -Subject "$FullName User Credentials" -Body $body2 -SmtpServer $smtpsrv -BodyAsHtml

$ct = $ct + 1
}
