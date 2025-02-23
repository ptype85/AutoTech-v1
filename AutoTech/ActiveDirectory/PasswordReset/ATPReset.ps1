﻿#Peter Endacott 2018
#Active Directory User Password Reset Automation Script
#Used with AutoTech

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

##Password Generator

$PArray = 'Orange','Apple','Raspberry','Strawberry','Grape','Loganberry','Lemon','Cherry'

##Email Settings

$fromAddr = "AutoTech@contoso.com" # Enter the FROM address for the e-mail alert
$toAddr = "pe@contoso.com" # Enter the TO address for the e-mail alert
$smtpsrv = "172.0.0.10" # Enter the FQDN or IP of a SMTP relay

###########################

##Input Read
$date = Get-Date -Format ddMMyy
$path = "C:\AutoTech\ActiveDirectory\PasswordReset\Files\Input\.csv"
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
