﻿#Peter Endacott 2018
#Active Directory Validation User Script
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
$eal = 1 #Email Alerts
$ua = 1 #User Availibility Check
$re = 1 #Restricted User Check

###########################

##Variable Store

#$at = #AllTasks
$domain = "contoso.com\"
#$ru = Restricted User List
#$uv = User validation Result
#$rr = Restircted User Result
#$VPResult = Virtual Pool Result

##Email Settings

$fromAddr = "AutoTech@contoso.com" # Enter the FROM address for the e-mail alert
$toAddr = "is@contoso.com" # Enter the TO address for the e-mail alert
$smtpsrv = "172.0.0.10" # Enter the FQDN or IP of a SMTP relay

###########################

##Input Read
$date = Get-Date -Format ddMMyy
$path = "C:\AutoTech\ActiveDirectory\Import\Files\Open\ImportUsers.csv"
if ((Get-Content "C:\AutoTech\ActiveDirectory\Import\Files\Open\RequestImport$date.csv") -ne $Null) {
$at = Import-Csv C:\AutoTech\ActiveDirectory\Import\Files\Open\RequestImport$date.csv #-Header WORKORDERID,VALUE1,VALUE2,VALUE3,VALUE4,VALUE5
$ru = Import-Csv C:\AutoTech\ActiveDirectory\Config\RestrictedUN.csv #Load restricted User List
$wo = @()
$wo += $at.WORKORDERID
$tu = @()
$tu += $at.VALUE5	#Read Template User from Input
$ue = @()
$ue += $at.VALUE2
$VEnt = Import-Csv C:\AutoTech\ActiveDirectory\VMWare\Files\VMEntReport.csv
$GNO = $VEnt.displayname -replace "contoso.com\\*" #Removing contoso.com from import
$FullName = @()
$FullName += $at.Value1

##Counter Reset
$tc = 0
$rc = 0
$vc = 0
$ec = 0
$pc = 0
$uv = (1..$wo.Count)
$rr = (1..$wo.Count)
$mv = (1..$wo.Count)
$VPResult = (1..$wo.Count)
}
Else {
$date2 = Get-Date -DisplayHint Date
Send-MailMessage -To $toAddr -From $fromAddr -Subject "Info: No New User Requests to process on $date2" -Body "No Accounts waiting to process.<BR><BR>I'm so bored..." -SmtpServer $smtpsrv -BodyAsHtml
Exit
}

##Template User Validity Check

if ($ua = 1){
$wo | ForEach-Object{
$us = $tu[$tc]
$result = get-user $domain$us
if (!$result) {$uv[$rc] = 0}
Else {$uv[$rc] = 1}

##Check for restricted Template Accounts
#if ($re = 1){
#$wo | ForEach-Object{
if ($ru -match $tu[$tc]) {$rr[$vc] = 0} 
Else {$rr[$vc] = 1}

##Check if supplied email exists as a mailbox or contact
$ea = $ue[$tc]
$mb = Get-Mailbox $ea
$mc = Get-MailContact $ea
if ((!$mc) -and (!$mb)) {$mv[$ec] = 1} else {$mv[$ec] = 0}

##Check VM Pool Capacity
$VAG = Get-ADPrincipalGroupMembership $tu[$tc] | select name | Where-Object {$_.name -like '*VPS VDI*'} | Where-Object {$_.name -ne 'VPS VDI Desktop Users (SG)'} | Out-String -stream  | select-object -index 3 #Find a VDI Security Group
$NVAG = $VAG.trim() #Remove white space from string
if (!$VAG) {$VPResult[$pc] = "No Virtual Pool Group found"}
Else
{
if ($GNO -like $NVAG) {
$GM = Get-ADGroupMember -Identity $nvag -Recursive | %{Get-ADUser -Identity $_.distinguishedname -Properties Enabled } | where { $_.Enabled -eq $True } #Pulls enabled AD users from VDI Group
$GC = $GM.count #Count of enabled users in VDI Group
$data = import-csv C:\AutoTech\ActiveDirectory\VMWare\Files\VMEntReport.csv | Where-Object {$_."displayName" -eq "$domain$nvag"}
$pdata = import-csv C:\AutoTech\ActiveDirectory\VMWare\Files\VMPoolReport.csv | Where-Object {$_."pool_id" -eq $data.pool_id }
$AIP = $pdata.maximumCount
$poolname = $data.pool_id
if ($pdata.maximumCount -gt $gc) {$VPResult[$pc] = "Virtual machines availabile in pool $POOLNAME. $GC out of $AIP machines are in use "} Else {$VPResult[$pc] = "Insufficent machine availability in $POOLNAME. $GC out of $AIP machines are in use."}} 
Else {$VPResult[$pc] = "VDI Security Group not associated with Pool"}
}
##########################

$date = Get-Date -Format ddMMyy

If (($rr[$vc]) -eq 0 -or ($uv[$rc]) -eq 0 -or ($mv[$ec]) -eq 0){
$wo[$tc],$tu[$tc],$uv[$tc],$rr[$vc],$mv[$ec],$VPResult[$pc],$FullName[$tc],$ue[$tc] -join ',' | Out-File -FilePath "C:\AutoTech\ActiveDirectory\Validate\Files\Failed\UserCheckFailed$date.csv" -Append;
}
Else
{
$wo[$tc],$FullName[$tc],$tu[$tc],$ue[$tc],$VPResult[$pc] -join ',' | Out-File -FilePath "C:\AutoTech\ActiveDirectory\Validate\Files\Passed\UserCheckPassed$date.csv" -Append;
}

$tc = $tc + 1
$rc = $rc + 1
$vc = $vc + 1
$ec = $ec + 1
$pc = $pc + 1
$VAG = $null
}}

##Email Alert
$ef = Import-Csv "C:\AutoTech\ActiveDirectory\Validate\Files\Passed\UserCheckPassed$date.csv" -Header Req,FullN,TemplateU,Email,VMResult
$fu = Import-Csv "C:\AutoTech\ActiveDirectory\Validate\Files\Failed\UserCheckFailed$date.csv" -Header Req,FullN,TemplateU,Email,VMResult
$cof = $fu.Req.Count
$cop = $ef.Count

if ($eal -eq 1){

$date2 = Get-Date -DisplayHint Date

$body = @("
Hello,
<BR>
<BR>
The following users have passed validation and will be created by Auto Tech at 15:00 today.
<BR>
<BR>
<B>$cof</B> new user requests have failed validation and will be notified.
<BR>
<BR>
<center><table border=1 width=50% cellspacing=0 cellpadding=8 bgcolor=Black cols=5>
<tr bgcolor=White><td>Request ID</td><td>Name</td><td>Template</td><td>VM Pool Result</tr></td>")

$i = 0

do {
if($i % 2){$body += "<tr bgcolor=#D2CFCF><td>$($ef[$i].Req)</td><td>$($ef[$i].FullN)</td><td>$($ef[$i].TemplateU)</td><td>$($ef[$i].VMResult)</td></tr>";$i++}
else {$body += "<tr bgcolor=#EFEFEF><td>$($ef[$i].Req)</td><td>$($ef[$i].FullN)</td><td>$($ef[$i].TemplateU)</td><td>$($ef[$i].VMResult)</td></tr>";$i++}
}
while ($ef[$i] -ne $null)

$body += "</table></center>"

Send-MailMessage -To $toAddr -From $fromAddr -Subject "Info: $($cop) New User Requests Passed Validation $date2" -Body "$body" -SmtpServer $smtpsrv -BodyAsHtml
} 
