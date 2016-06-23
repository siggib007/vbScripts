Option Explicit
Dim Locator, conn, strNameSpace, strServer, SQLServerObj, strDomain, uid, pwd

Const wbemImpersonationLevelImpersonate = 3
'On Error Resume Next 
If wscript.arguments.count > 0 Then
	strServer = wscript.arguments(0)
	wscript.echo Now() & " Connecting to " & strServer & "..."
Else
	strServer = ""
	wscript.echo Now() & " Connecting to local machine..."
End If

If wscript.arguments.count > 2 Then
	uid = wscript.arguments(1)
	pwd = wscript.arguments(2)
	wscript.echo Now() & " Connecting to " & strServer & " using " & uid
Else
	uid = ""
	pwd = ""
	wscript.echo Now() & " Connecting using current logins"
End If

strNameSpace = "root\CIMV2"

Set locator=CreateObject("WbemScripting.SwbemLocator")
If strServer = "" Then
	Set conn=locator.connectserver (, strNameSpace)
	If err.number <> 0 Then 
		wscript.echo "unable to connect"
		wscript.quit
	End If
	conn.security_.impersonationlevel = wbemImpersonationLevelImpersonate 
Else
	Set conn=locator.connectserver (strServer, strNameSpace,uid,pwd)
	If err.number <> 0 Then 
		wscript.echo "unable to connect"
		wscript.quit
	End If
	conn.security_.impersonationlevel = wbemImpersonationLevelImpersonate 
End If
Set SQLServerObj=conn.get("Win32_Service.name=""MSSQLServer""")
wscript.echo "Starting..."
wscript.echo sqlserverobj.startservice
'wscript.echo "SQL Server service is " & sqlserverobj.state
'wscript.echo "SQL Server Service Autostart set to " & sqlserverobj.startmode
wscript.echo ""
wscript.echo Now() & " Script Complete."
