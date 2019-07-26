#Peter Endacott 2018
#Excel Read - Retrieves details from mail submitted forms.
#Used with AutoTech

#set execution policy
#Set-ExecutionPolicy RemoteSigned

##Input read
$bob = Get-ChildItem "C:\AutoTech\MailRequests" -Filter *.xlsm
$date = Get-Date -Format ddMMyy
$US = @()
$US += $bob.Name
$ct = 0

##Find files and move to processed
If ($bob -ne $null){
$US | ForEach-Object {
$ui = $US[$ct]
Rename-Item -Path "C:\AutoTech\MailRequests\$ui" "C:\AutoTech\MailRequests\Processed$ui"
Move-Item -Path "C:\AutoTech\MailRequests\Processed$ui" -Destination "C:\AutoTech\MailRequests\Processed"
$ct = $ct + 1
}
}
else
{
exit
}

##Initiate Excel
$xl = New-Object -COM "Excel.Application"
$ct = 0

##Retrieve details from forms
$US | ForEach-Object {
$ui = $US[$ct]
$wb = $xl.Workbooks.Open("C:\AutoTech\MailRequests\processed\Processed$ui")
$ws = $wb.sheets.Item(1)
$UN = $ws.Cells.Item(8,2).text
$EM = $ws.Cells.Item(11,2).text
$LO = $ws.Cells.Item(14,2).text
$TU = $ws.Cells.Item(17,2).text
$UN,$EM,$LO,$TU -join ',' | Out-File -FilePath "C:\AutoTech\MailRequests\Files\Output$date.csv" -Append;
$ct = $ct + 1
}

##Nuke the living heck out of the Excel process to prevent memory leaks
$xl=New-Object –com Excel.Application
$xl.Quit()
while( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)){}
[VOID][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)
[VOID][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)
[VOID][System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Remove-Variable -Name ws
Remove-Variable -Name wb
Remove-Variable -Name xl
Remove-Variable -Name UN
Remove-Variable -Name EM
Remove-Variable -Name LO
Remove-Variable -Name TU
kill -ProcessName excel