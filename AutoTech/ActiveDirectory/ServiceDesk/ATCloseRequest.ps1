﻿#Peter Endacott 2018
#Request Management - Close Request
#Used with AutoTech

#set execution policy
Set-ExecutionPolicy RemoteSigned

##Input Read
$date = Get-Date -Format ddMMyy
$path = "C:\AutoTech\ActiveDirectory\Notify\Files\Email\EmailAlertOutput$date.csv"
if ([System.IO.File]::Exists($path)){
$date = Get-Date -Format ddMMyy
$at = Import-Csv "C:\AutoTech\ActiveDirectory\Notify\Files\Email\EmailAlertOutput$date.csv"  -Header FullN,UserN,ReqID,NE,Email
$wo = @()
$wo += $at.ReqID
$fn = @()
$fn += $at.FullN
$em = @()
$em += $at.Email}
Else {
Exit
}

##API Settings
$ApiKey = "..."
$SdpUri = "http://servicedesk.contoso.com:8082"

##Counter Set
$LC = 0

$wo | ForEach-Object{
$urlwo = $wo[$LC]
$Uri = $SdpUri + "/sdpapi/request/$urlwo"
$resolution = "Test Successful"


   
            $Parameters = @{
            "operation" = @{
                "details" = @{
                   "closeAccepted" = "Accepted";
                   "closeComment" = "Account Creation Completed Successfully"
                    }
                }
            } 
        
       
        $input_data = $Parameters | ConvertTo-Json -Depth 50
        $Uri = $Uri + "?format=json&OPERATION_NAME=CLOSE_REQUEST&INPUT_DATA=$input_data&TECHNICIAN_KEY=$ApiKey"
        $result = Invoke-RestMethod -Method POST -Uri $Uri -OutFile "C:\Test\APIoutput.xml"
        $result 
$LC = $LC + 1
}
