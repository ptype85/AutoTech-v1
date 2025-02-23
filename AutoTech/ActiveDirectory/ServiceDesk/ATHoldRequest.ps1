﻿#Peter Endacott 2018
#Request Management - Put Request on Hold
#Used with AutoTech

#set execution policy
Set-ExecutionPolicy RemoteSigned

##Input Read
$date = Get-Date -Format ddMMyy
$path = "C:\AutoTech\ActiveDirectory\Validate\Files\Failed\UserCheckFailed$date.csv"
if ([System.IO.File]::Exists($path)){
$date = Get-Date -Format ddMMyy
$at = Import-Csv "C:\AutoTech\ActiveDirectory\Validate\Files\Failed\UserCheckFailed$date.csv"  -Header ReqID,UserN,ResU,UVal,EmExi,VMRes
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


   
            $Parameters = @{
            "operation" = @{
                "details" = @{
                   "status" = "On Hold";
                    }
                }
            } 
        
       
        $input_data = $Parameters | ConvertTo-Json -Depth 50
        $Uri = $Uri + "?format=json&OPERATION_NAME=EDIT_REQUEST&INPUT_DATA=$input_data&TECHNICIAN_KEY=$ApiKey"
        $result = Invoke-RestMethod -Method POST -Uri $Uri -OutFile "C:\AutoTech\Test\APIoutput.xml"
        $result 
$LC = $LC + 1
}
