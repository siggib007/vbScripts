Dim oSQLServer, oJobSchedule, chkAuthentication, txtServer, sUsername, sPasswd

Set oSQLServer = CreateObject("SQLDMO.SQLServer")
Set oJobSchedule = CreateObject("SQLDMO.JobSchedule")

chkAuthentication = true
txtServer = "SGB1"

oSQLServer.LoginTimeout = -1 '-1 is the ODBC default (60) seconds
'Connect to the Server
If chkAuthentication Then
  With oSQLServer
  'Use NT Authentication
    .LoginSecure = True
  'Do not reconnect automatically
    .AutoReConnect = False
  'Now connect
    .Connect txtServer
  End With
Else
  With oSQLServer
  'Use SQL Server Authentication
    .LoginSecure = False
  'Do not reconnect automatically
    .AutoReConnect = False
  'Use SQL Security
    .Connect txtServer, sUsername, sPasswd
  End With
End If
wscript.echo "Your Login: " & oSQLServer.Login

Set oJob = oSQLServer.JobServer.Jobs("Backup Hard Drive")
' Set the schedule name.
oJobSchedule.Name = "Single_Execution"

' Indicate a single scheduled execution by using the
' FrequencyType property.
oJobSchedule.Schedule.FrequencyType = 1

' Use the ActiveStartDate and ActiveStartTimeOfDay properties
' to indicate the scheduled execution time for a JobSchedule
' object implementing a single run.
oJobSchedule.Schedule.ActiveStartDate = "19980922"
oJobSchedule.Schedule.ActiveStartTimeOfDay = "130000"

' Optional, but cleaner. Indicated that schedule never expires.
oJobSchedule.Schedule.ActiveEndDate = SQLDMO_NOENDDATE
oJobSchedule.Schedule.ActiveEndTimeOfDay = SQLDMO_NOENDTIME

' Alter the job, adding the new schedule.
oJob.BeginAlter
oJob.JobSchedules.Add oJobSchedule
oJob.DoAlter

If Not oSQLServer Is Nothing Then
	'When done with the connection to SQLServer you must Disconnect
    oSQLServer.DisConnect
End If
  
Set oSQLServer = Nothing
Set oJob = Nothing
Set oJobSchedule = Nothing
