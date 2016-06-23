Option Explicit 
Dim Job, oSQLServer, bUseIntegrated, txtServer, sUsername, sPasswd, LastStatus, CurrStatus, AgentStatus, Step

Set oSQLServer = CreateObject("SQLDMO.SQLServer")

If wscript.arguments.count > 0 Then
	txtServer = wscript.arguments(0)
	wscript.echo Now() & " Connecting to " & txtServer & "..."
Else
	txtServer = ""
	wscript.echo Now() & " Connecting to local machine..."
End If

If wscript.arguments.count > 2 Then
	sUsername = wscript.arguments(1)
	sPasswd = wscript.arguments(2)
	wscript.echo Now() & " Connecting to " & txtServer & " using " & sUsername
Else
	sUsername = ""
	sPasswd = ""
	bUseIntegrated = true
	wscript.echo Now() & " Connecting using current logins"
End If

'txtServer = "gnettools23"

oSQLServer.LoginTimeout = -1 '-1 is the ODBC default (60) seconds
'Connect to the Server
wscript.echo "Attempting to establish connection"
If bUseIntegrated Then
  With oSQLServer
  'Use NT Authentication
    .LoginSecure = True
    .AutoReConnect = False
    .Connect txtServer
  End With
Else
  With oSQLServer
  'Use SQL Server Authentication
    .LoginSecure = False
    .AutoReConnect = False
    .Connect txtServer, sUsername, sPasswd
  End With
End If
wscript.echo "Your Login: " & oSQLServer.Login

wscript.echo "SQLServerAgent Autostart is set to " & oSQLServer.JobServer.AutoStart

Select Case oSQLServer.JobServer.Status
	Case 0 
		AgentStatus = "Unknown"
	Case 1
		AgentStatus = "Running"
	Case 2
		AgentStatus = "Paused"
	Case 3
		AgentStatus = "Stopped"
	Case 4
		AgentStatus = "Starting"
	Case 5
		AgentStatus = "Stopping"
	Case 6
		AgentStatus = "Continuing from paused state"
	Case 7 
		AgentStatus = "Pausing"
	Case Else
		AgentStatus = "Undefined"
End Select 

wscript.echo "SQLServerAgent state is " & AgentStatus


For each job in oSQLServer.JobServer.Jobs
	Select Case job.LastRunOutcome
		Case 0 
			LastStatus = "Failed"
		Case 1
			LastStatus = "Successful"
		Case 3
			LastStatus = "Cancelled"
		Case 4
			LastStatus = "Executing"
		Case 5
			LastStatus = "Uknown"
		Case Else 
			LastStatus = "Undefined"
	End Select
	
	Select Case job.CurrentRunStatus
		Case 3
			CurrStatus = "Waiting for a retry"
		Case 1
			CurrStatus = "Executing"
		Case 4
			CurrStatus = "Not Running"
		Case 7
			CurrStatus = "Done, writing to history log"
		Case 5
			CurrStatus = "Suspended"
		Case 0
			CurrStatus = "Unknown"
		Case 6
			CurrStatus = "Waiting for step to finish"
		Case 2
			CurrStatus = "Waiting for worker thread"
		Case Else
			CurrStatus = "Undefined"
	End Select 
	writelog job.name & vbtab & LastStatus & vbtab & job.LastRunDate & vbtab & CurrStatus
	For each step in job.jobsteps
		writelog step.name & vbtab & step.command
	Next
Next

If Not oSQLServer Is Nothing Then
	'When done with the connection to SQLServer you must Disconnect
    oSQLServer.DisConnect
End If
  
Set oSQLServer = Nothing

Sub writelog (msg)
wscript.echo msg
End Sub 