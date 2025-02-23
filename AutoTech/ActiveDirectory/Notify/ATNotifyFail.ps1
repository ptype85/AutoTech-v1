﻿#Peter Endacott 2018
#User Creation ailiure Script
#Used with AutoTech

#set execution policy
Set-ExecutionPolicy RemoteSigned

##Debug Switches

$ir = 1 #Input Read
$ua = 1 #Script Action

##Email Settings

$fromAddr = "ServiceDesk@contoso.com" # Enter the FROM address for the e-mail alert
$smtpsrv = "172.0.0.10" # Enter the FQDN or IP of a SMTP relay

##Input Read
$date = Get-Date -Format ddMMyy
$path = "C:\AutoTech\ActiveDirectory\Validate\Files\Failed\UserCheckFailed$date.csv"
if ([System.IO.File]::Exists($path)){
$at = Import-Csv "C:\AutoTech\ActiveDirectory\Validate\Files\Failed\UserCheckFailed$date.csv"  -Header Req,UserN,ResU,UVal,EmExi,VMres,FullN,Email

$wo = @()
$wo += $at.req
$fn = @()
$fn += $at.FullN
$un = @()
$un += $at.UserN
$ru = @()
$ru += $at.ResU
$uv = @()
$uv += $at.UVal
$ex = @()
$ex += $at.EmExi
$vr = @()
$vr += $at.VMres
$em = @()
$em += $at.Email

}else
{
Exit
}

##Counter Reset
$woc = 0 #Work Order Count
$fnc = 0 #Full Name Count
$unc = 0 #Template User Name Count
$ruc = 0 #Restricted User Name Count
$uvc = 0 #Template User Validation Count
$exc = 0 #Email Validation Count
$vrc = 0 #VM Pool Count
$emc = 0 #Email Count

if ($ua = 1){
$wo | ForEach-Object{

if ($ru[$uvc] -eq "0"){
$woe = $wo[$woc]
$toAddr = $em[$emc]
$une = $un[$unc]
$date2 = Get-Date -DisplayHint Date
Send-MailMessage -To $toAddr -From $fromAddr -Subject "Your new user request $woe has been put on hold." -Body "Hello,<BR><BR>Your request, <b>$woe</b>, for a new user on the platform has been placed on hold. This is because the template user ID which was supplied, <b>$une</b>, is invalid. Please ask the requester to navigate to:<BR><BR>http://servicedesk.yestelco.com:8082/WorkOrder.do?woMode=viewWO&woID=$woe<BR><BR>If you encounter any difficulties, please contact <a href=""mailto:ServiceDesk@contoso.com"">Service Desk.</a><BR><BR>Your request will be put on hold for 10 days. If there is no response within this time, it will be closed.<BR><BR>Regards,<BR>AutoTech" -SmtpServer $smtpsrv -BodyAsHtml
}

if ($uv[$ruc] -eq "0"){
$woe = $wo[$woc]
$toAddr = $em[$emc]
$une = $un[$unc]
$date2 = Get-Date -DisplayHint Date
Send-MailMessage -To $toAddr -From $fromAddr -Subject "Your new user request $woe has been put on hold." -Body "Hello,<BR><BR>Your request, <b>$woe</b>, for a new user on the platform has been placed on hold. This is because the template user ID which was supplied, <b>$une</b>, is a restriced account. Please ask the requester to navigate to:<BR><BR>http://servicedesk.yestelco.com:8082/WorkOrder.do?woMode=viewWO&woID=$woe<BR><BR>If you encounter any difficulties, please contact <a href=""mailto:ServiceDesk@contoso.com"">Service Desk.</a><BR><BR>Your request will be put on hold for 10 days. If there is no response within this time, it will be closed.<BR><BR>Regards,<BR>AutoTech" -SmtpServer $smtpsrv -BodyAsHtml
}

if ($ex[$exc] -eq "0"){
$woe = $wo[$woc]
$toAddr = $em[$emc]
$date2 = Get-Date -DisplayHint Date
Send-MailMessage -To $toAddr -From $fromAddr -Subject "Your new user request $woe has been put on hold." -Body "Hello,<BR><BR>Your request, <b>$woe</b>, for a new user on the platform has been placed on hold. This is because the email address which was supplied, <b>$toAddr</b>, already exists in the environment. This would usually suggest that an account already exists and may be disabled. Please ask the requester to navigate to:<BR><BR>http://servicedesk.yestelco.com:8082/WorkOrder.do?woMode=viewWO&woID=$woe<BR><BR>If you encounter any difficulties, please contact <a href=""mailto:ServiceDesk@contoso.com""> Service Desk.</a><BR><BR>Your request will be put on hold for 10 days. If there is no response within this time, it will be closed.<BR><BR>Regards,<BR>AutoTech" -SmtpServer $smtpsrv -BodyAsHtml
}

$woc = $woc + 1
$fnc = $fnc + 1
$unc = $unc + 1
$ruc = $ruc + 1
$uvc = $uvc + 1
$exc = $exc + 1
$vrc = $vrc + 1
$emc = $emc + 1
}
}
