#Peter Endacott 2018
#Request Management - Resolve Request with no action for 10 days
#Used with AutoTech

#set execution policy
Set-ExecutionPolicy RemoteSigned

##Input Read
$date = Get-Date -Format ddMMyy
$path = "C:\AutoTech\ActiveDirectory\ServiceDesk\Files\OnHold\RequestOnHold$date.csv"
if ([System.IO.File]::Exists($path)){
$date = Get-Date -Format ddMMyy
$at = Import-Csv "C:\AutoTech\ActiveDirectory\ServiceDesk\Files\OnHold\RequestOnHold$date.csv"
$wo = @()
$wo += $at.WORKORDERID
$lu = @()
$lu += $at.LAST_TECH_UPDATE
$fn = @()
$fn += $at.VALUE1
}
Else {
Exit
}

##API Settings
$ApiKey = "..."
$SdpUri = "http://servicedesk.contoso.com:8082"

##Counter Set
$LC = 0
$epoch = [int64](([datetime]::UtcNow)-(get-date "1/1/1970")).TotalMilliseconds
$responseperiod = ($epoch - "864000000")

$wo | ForEach-Object{
if (($lu[$LC]) -lt $responseperiod) {
$urlwo = $wo[$LC]
$Uri = $SdpUri + "/sdpapi/request/$urlwo"
$resolution = "Test Successful"


   
            $Parameters = @{
            "operation" = @{
                "details" = @{
                   "closeAccepted" = "Accepted";
                   "closeComment" = "Account Creation Unsuccessful"
                    }
                }
            } 
        
       
        $input_data = $Parameters | ConvertTo-Json -Depth 50
        $Uri = $Uri + "?format=json&OPERATION_NAME=CLOSE_REQUEST&INPUT_DATA=$input_data&TECHNICIAN_KEY=$ApiKey"
        $result = Invoke-RestMethod -Method POST -Uri $Uri -OutFile "C:\Test\APIoutput.xml"
        $result 
$LC = $LC + 1
}
} else
{
$LC = $LC + 1
}
