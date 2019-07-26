#Peter Endacott 2018
#Import New User Requests from Service Desk
#Used with AutoTech

#set execution policy
#Set-ExecutionPolicy RemoteSigned

$date = Get-Date -Format ddMMyy
invoke-sqlcmd -inputfile "C:\AutoTech\ActiveDirectory\Import\WJSQL.sql" -ServerInstance "wheeljack.yestelco.com" | Export-CSV "C:\AutoTech\ActiveDirectory\Import\Files\Open\RequestImport$date.csv" -NoTypeInformation