#Peter Endacott 2018
#Import Requests On Hold
#Used with AutoTech

#set execution policy
Set-ExecutionPolicy RemoteSigned

$date = Get-Date -Format ddMMyy
invoke-sqlcmd -inputfile "C:\AutoTech\ActiveDirectory\ServiceDesk\GetLastTUpdate.sql" -ServerInstance "wheeljack.yestelco.com" | Export-CSV "C:\AutoTech\ActiveDirectory\ServiceDesk\Files\OnHold\RequestOnHold$date.csv" -NoTypeInformation