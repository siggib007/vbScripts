Option Explicit
Dim NetAdaptorObj,ServerObj,NIC, NICIP, NICConfigObj, ipaddr, x, DS, tz, MemKB
Dim server,NICIPObj, SObj, strDomain,uid,pwd, OSObj, OS, QFE, QFEObj
Dim Locator, conn, strNameSpace, strServer,SQLServerObj
Const wbemImpersonationLevelImpersonate = 3

On Error Resume Next 
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
	Select Case err.number
		Case -2147023174 
			wscript.echo Now() & " Unable to connect to local server"
			wscript.quit
		Case -2147024891
			wscript.echo Now() & " Access denied while attempting to connect to local server"
			wscript.quit
		Case 0 
			wscript.echo Now() & " Successfully connected to the local server"
			wscript.echo ""
		Case Else
			wscript.echo Now() & " " & err.number & " : " & err.source & " : " & err.description
			wscript.quit
	End Select
	conn.security_.impersonationlevel = wbemImpersonationLevelImpersonate 
	Set SObj = conn.execquery("select * from win32_ComputerSystem")
	For each server in sobj
		strServer = server.name
	Next
Else
	Set conn=locator.connectserver (strServer, strNameSpace,uid,pwd)
		Select Case err.number 
		Case -2147023174 
			wscript.echo Now() & " Unable to connect to server " & strServer
			wscript.quit
		Case -2147024891
			wscript.echo Now() & " Access denied while attempting to connect to " & strServer
			wscript.quit
		Case 0 
			wscript.echo Now() & " Successfully connected to the server " & strServer
			wscript.echo ""
		Case Else
			wscript.echo Now() & " " & err.number & " : " & err.source & " : " & err.description
			wscript.quit
	End Select
	conn.security_.impersonationlevel = wbemImpersonationLevelImpersonate 
End If

Set NetAdaptorObj = conn.execquery("select * from win32_networkadapter where adaptertype = ""Ethernet 802.3""")

Set ServerObj=conn.get("Win32_ComputerSystem.name=""" & strServer & """")
memkb=CInt(ServerObj.totalphysicalmemory/(1024*1024))
tz=ServerObj.CurrentTimeZone/60
strDomain = serverobj.domain
if ServerObj.DaylightInEffect then 
	DS=""
else
	DS="not"
end if

wscript.echo ServerObj.name & " belongs to the " & strdomain & " domain and has " & memkb & " MB of physical memory."
wscript.echo "Has a timezone of GMT " & ServerObj.currenttimezone/60 & " and day light savings is " & ds & " in effect."
'wscript.quit
wscript.echo "It has the following Ethernet 802.3 adaptors"
For each NIC in NetAdaptorObj
'On error resume next
	set NICConfigObj=conn.get("Win32_NetworkAdapterconfiguration.index=" & NIC.deviceid)
	wscript.echo NIC.Name & "; MAC: " & NIC.MACAddress 
	'wscript.echo UBound(nicconfigobj.ipaddress)+1 & " IP addresses bound"
	For x=0 to UBound(nicconfigobj.ipaddress)
		wscript.echo "IP: " & NICConfigObj.IPaddress(x) & "; DHCP: " & NICConfigObj.DHCPEnabled
	Next
Next
wscript.echo "Operating system Info"
Set OSObj = conn.execquery("select * from win32_OperatingSystem")
For each OS in OSObj
	Wscript.echo OS.caption & " Build: " & os.buildnumber & " " & os.buildtype
	wscript.echo "SP " & os.servicepackmajorversion & "." & os.servicepackminorversion
Next
Set QFEObj = conn.execquery("select * from win32_QuickFixEngineering")
wscript.echo "Hotfixes applied"
For each QFE in QFEObj
	wscript.echo QFE.hotfixid & " - Service Pack in effect: " & qfe.servicepackineffect _
		& " - Descr: " & qfe.description
Next
Set SQLServerObj=conn.get("Win32_Service.name=""MSSQLServer""")
wscript.echo ""
wscript.echo "SQL Server service is " & sqlserverobj.state
wscript.echo "SQL Server Service Autostart set to " & sqlserverobj.startmode
wscript.echo ""
wscript.echo Now() & " Script Complete."
'on error goto 0
