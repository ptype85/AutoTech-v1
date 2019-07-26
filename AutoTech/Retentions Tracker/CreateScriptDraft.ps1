$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandType=[System.Data.CommandType]'StoredProcedure'
$SqlCmd.CommandText = "IssueTracker_User_CreateNewUser"
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SQLServer = "dangermouse.yestelco.com"
$SQLDBName = "retentions"
$uid ="yestelco.com\autotech"
$pwd ="M4nUn1t3D5UcKD0nk3YB4lL5"
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True; User ID = $uid; Password = $pwd;"
$SqlCmd.Connection = $SqlConnection
$ParamUserName = $SqlCmd.Parameters.Add("@UserName", [Data.SqlDbType]::VarChar)
$ParamUserName.Direction = [Data.ParameterDirection]::Input
$ParamUserName.Value = "MickeyM"
$ParamRoleName = $SqlCmd.Parameters.Add("@RoleName", [Data.SqlDbType]::VarChar)
$ParamRoleName.Direction = [Data.ParameterDirection]::Input
$ParamRoleName.Value = "Consultant"
$ParamEmail = $SqlCmd.Parameters.Add("@Email", [Data.SqlDbType]::VarChar)
$ParamEmail.Direction = [Data.ParameterDirection]::Input
$ParamEmail.Value = "Mickey.Mouse@vodafone.com"
$ParamDisplayName = $SqlCmd.Parameters.Add("@DisplayName", [Data.SqlDbType]::VarChar)
$ParamDisplayName.Direction = [Data.ParameterDirection]::Input
$ParamDisplayName.Value = "Mickey Mouse"
$ParamUserPassword = $SqlCmd.Parameters.Add("@UserPassword", [Data.SqlDbType]::VarChar)
$ParamUserPassword.Direction = [Data.ParameterDirection]::Input
$ParamUserPassword.Value = "Lemon567"
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)

$SQLConfirm = "SELECT w.userID From retentions.dbo.IssueTracker_Users w Where w.UserName = 'MickeyM'"
$Result = Invoke-sqlcmd -Query $SQLConfirm -ServerInstance "dangermouse.yestelco.com"


$SqlCmd2 = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd2.CommandType=[System.Data.CommandType]'StoredProcedure'
$SqlCmd2.CommandText = "IssueTracker_Project_AddUserToProject"
$SqlConnection2 = New-Object System.Data.SqlClient.SqlConnection

$SqlConnection2.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True; User ID = $uid; Password = $pwd;"
$SqlCmd2.Connection = $SqlConnection
$ParamUserID =  $SqlCmd2.Parameters.Add("@UserId", [Data.SqlDbType]::VarChar)
$ParamUserID.Direction = [Data.ParameterDirection]::Input
$UserID = $Result[0]

$ParamUserID.Value = "$UserID"
$ParamProjID =  $SqlCmd2.Parameters.Add("@ProjectId", [Data.SqlDbType]::VarChar)
$ParamProjID.Direction = [Data.ParameterDirection]::Input
$ParamProjID.Value = "1"
$SqlAdapter2 = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter2.SelectCommand = $SqlCmd2
$DataSet2 = New-Object System.Data.DataSet
$SqlAdapter2.Fill($DataSet2)