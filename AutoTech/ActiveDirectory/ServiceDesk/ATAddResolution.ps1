﻿#Peter Endacott 2018
#Request Management- Add Resolution
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
$ResFN = $fn[$LC]
$ResEM = $em[$LC]


$Uri = $SdpUri + "/sdpapi/request/$urlwo/resolution"
$resolution = "User account creation for $ResFN completed by AutoTech successfully. Credentials have been sent to $ResEM"


   
            $Parameters = @{
            "operation" = @{
                "details" = @{
                    "resolution" = @{
                       "resolutiontext" = $resolution
                    }
                }
            } 
        }
       
        $input_data = $Parameters | ConvertTo-Json -Depth 50
        $Uri = $Uri + "?format=json&OPERATION_NAME=ADD_RESOLUTION&INPUT_DATA=$input_data&TECHNICIAN_KEY=$ApiKey"
        $result = Invoke-RestMethod -Method POST -Uri $Uri -OutFile "C:\Test\APIoutput.xml"
        $result 
 
$LC = $LC + 1
}
