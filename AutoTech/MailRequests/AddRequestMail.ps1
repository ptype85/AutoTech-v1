#Peter Endacott 2018
#Request Management- Add Request for Email Forms
#Used with AutoTech

#set execution policy
#Set-ExecutionPolicy RemoteSigned

##Input Read
$date = Get-Date -Format ddMMyy
$path = "C:\AutoTech\MailRequests\Files\Output$date.csv"
if ([System.IO.File]::Exists($path)){
$date = Get-Date -Format ddMMyy
$at = Import-Csv "C:\AutoTech\MailRequests\Files\Output$date.csv"  -Header FullN,Email,NE,Template
$tu = @()
$tu += $at.Template
$fn = @()
$fn += $at.FullN
$em = @()
$em += $at.Email
}
Else {
Exit
}

##API Settings
$ApiKey = "..."
$SdpUri = "http://servicedesk.contoso.com:8082"

##Counter Set
$LC = 0

$AT | ForEach-Object{
$templ = $tu[$LC]
$ResFN = $fn[$LC]
$ResEM = $em[$LC]


$Uri = $SdpUri + "/sdpapi/request?"
$resolution = "User account creation request. Details are:Name:$ResFN, Email:$ResEM, Template user:$templ"


   
$Parameters = @{
    "operation"= @{
        "details" = @{
            "subject" ="New starter request $ResFN";
             "description"="$resolution";
              "requester" ="AutoTech";
              "CREATEDBY" = "AutoTech";
               "requesttype"= "Service Request";
                "impact"= "Level 5 - Affects User";
                 "mode"= "E-Mail"; "Urgency"= "3 - Low";
                  "Priority"= "3 - Low";
                   "group"= "ServiceDesk";
                    "technician"= "AutoTech";
                     "level"= "1st Line";
                      "status"= "open";
                       "Category"= "Active Directory";
                        "subcategory"= "User";
                         "Item"= "Create New" 
               }
        }
  }
       
        $input_data = $Parameters | ConvertTo-Json -Depth 50
        $Uri = $Uri + "format=json&OPERATION_NAME=ADD_REQUEST&INPUT_DATA=$input_data&TECHNICIAN_KEY=$ApiKey"
        $result = Invoke-RestMethod -Method POST -Uri $Uri -OutFile "C:\AutoTech\MailRequests\API\Out$LC.json"
        $result 

$APIRead = Get-Content "C:\AutoTech\MailRequests\API\out$lc.json"
$WO = $APIRead[118..122]  -join '';

$wo,$fn[$lc],$em[$lc],$null,$null,$tu[$lc] -join ',' | Out-File -FilePath "C:\AutoTech\ActiveDirectory\Import\Files\Open\RequestImport$date.csv" -Encoding ASCII -Append;
 
$LC = $LC + 1
}




#$SDLink = "http://localhost:8080/"
#$api="sdpapi/request?"
#$format="json"
#$APIKey = "<Your API-key>"

#$inputData = @{
#operation=
#@{details=
#@{subject="Servicedesk Plus MSP example";
#description="Yay! I'am able post a request with PowerShell!";
#requester="John Doe";
#site="Sample Site";
#account="Sample Account"
#}
#}
#} |ConvertTo-Json
#$URI=$SDLink+$api
#$postParams = @{TECHNICIAN_KEY=$($APIKey);data=$($inputData);format=$($format)}
#Invoke-WebRequest -Uri $URI -Body $postParams -Method POST -TimeoutSec 10 
#write-host $SDLink
