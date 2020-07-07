#Peter Endacott 2018
#Import On Hold New User Requests from Service Desk
#Used with AutoTech

#set execution policy
Set-ExecutionPolicy RemoteSigned

$date = Get-Date -Format ddMMyy
invoke-sqlcmd -inputfile "C:\AutoTech\ActiveDirectory\Import\OHSQL.sql" -ServerInstance "dbserver.yestelco.com" | Export-CSV "C:\AutoTech\ActiveDirectory\Import\Files\OnHold\OHRequestImport$date.csv" -NoTypeInformation
