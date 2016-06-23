Option Explicit
Dim FileObj, strLine, fso, f, fc, f1, strParts, FolderSpec
Dim DeviceName, Interface, InterfaceIP, InterfaceMask, HSRP, ACLName
Dim DeviceID, IntID, ACLID, linenum, IPType, IPInst, description
Dim cn, rs, strSQL, cmd

Const DBServerName = "satnetengfs01"
Const DefDBName = "ACL"
Const adCmdText = 1 

If WScript.Arguments.Count <> 1 Then 
  WScript.Echo "Usage: configparser inpath"
  WScript.Quit
End If

FolderSpec = WScript.Arguments(0)

wscript.echo Now & " starting processing " & folderspec

Set cn  = CreateObject("ADODB.Connection")
Set rs  = CreateObject("ADODB.Recordset")
Set cmd = CreateObject("ADODB.Command")
Set fso = CreateObject("Scripting.FileSystemObject")
Set f   = fso.GetFolder(folderspec)
Set fc  = f.Files

wscript.echo Now & " initializing variables"

InitVariables

wscript.echo Now & " connecting to database server " & dbservername

cn.Provider = "sqloledb"
cn.Properties("Data Source").Value = DBServerName
cn.Properties("Initial Catalog").Value = DefDBName
cn.Properties("Integrated Security").Value = "SSPI"
cn.Open
Cmd.ActiveConnection = cn
cmd.commandtimeout = 60 
cmd.commandtext = adCmdText
wscript.echo Now & " connection established, cleaning out old data"

strsql = "DELETE FROM ACL.dbo.Interfaces"
Cmd.CommandText = strSQL
Cmd.Execute	
wscript.echo Now & " Interfaces done"
strsql = "DELETE FROM ACL.dbo.Devices"
Cmd.CommandText = strSQL
Cmd.Execute	
wscript.echo  Now & " Devices done"
strsql = "DELETE FROM ACL.dbo.ACLName where iACLID > 0"
Cmd.CommandText = strSQL
Cmd.Execute	
wscript.echo Now & " ACL Names done"
strsql = "DELETE FROM ACL.dbo.IPAddr"
Cmd.CommandText = strSQL
Cmd.Execute	
wscript.echo Now & " IP Addresses done"
strsql = "DELETE FROM ACL.dbo.ACL_Line"
Cmd.CommandText = strSQL
Cmd.Execute	
wscript.echo Now & " ACL Lines done"

wscript.echo Now & " Cleared the old values"

For Each f1 in fc
	strparts = split(f1.name,".")
	If UBound(strparts) = 1 Then 
		DeviceName = strparts(0)
		wscript.echo Now & " Processing " & devicename
		DeviceID = GetID ("Devices","iDeviceID","vcDeviceName",DeviceName)
		Set FileObj = fso.opentextfile(folderspec & "\" & f1.name)
		While not fileobj.atendofstream
			strLine = Trim(FileObj.readline)
			If strLine = "!" Then
				'flush all to DB and clear variables. 
				If Interface <> "" and InterfaceIP <> "" Then 
					'wscript.echo "Writing interface " & interface & " to database"
					'strSQL = "INSERT INTO ACL.dbo.Interfaces (vcIntName, vcIntIP, vcIntSubnet, vcHSRP1, vcHSRP2, iACL_in_ID, iACL_out_ID, iDeviceID) VALUES "
					'strsql = strsql & "('" & interface & "','" & InterfaceIP & "','" & InterfaceMask & "','" & HSRP1 & "','" & HSRP2 & "'," & inACLID & "," & outACLID & "," & deviceid & ")"
					'wscript.echo strsql
					'Cmd.CommandText = strSQL 
					'Cmd.Execute	
				End If		
				InitVariables
			End If 
			strparts = split(strline," ")
			
			If UBound(strparts) > 0 Then
				If IsNumeric(aclname) and strparts(0) <> "access-list" Then InitVariables
				If strparts(0) = "interface" Then
					Interface = strparts(1)
					'intid = getid("Interfaces","iIntID","vcIntName",Interface)				
					'wscript.echo "Found Interface " & interface
				End If 
				If strparts(0) = "description" and Interface <> "" Then
					description = Right(strline,Len(strline)-12)
					description = replace(description,"'","''")				
					'strsql = "update acl.dbo.interfaces set vcdescription = '" & description & "' where iIntID = " & intid
					'wscript.echo "Updating interface with description"
					'wscript.echo strsql
					'Cmd.CommandText = strSQL
					'Cmd.Execute							
				End If				
				If strparts(1) = "address" and Interface <> "" Then
					'wscript.echo "found address number of words " & UBound(strparts)
					'wscript.echo strline
					If UBound(strparts) > 2 Then
						InterfaceIP = strparts(2)
						InterfaceMask = strParts(3)
						IPType = 1 'Primary
						ipinst = 1
						'wscript.echo "Interface " & interface & " has ip of " & InterfaceIP & " and mask of " & InterfaceMask
					End If 
					If UBound(strparts) = 4 Then
						If strparts(4) = "secondary" Then
							IPType = 2 ' Secondary
							IPInst = IPInst + 1
						End If 
					End If 
					If intid = 0 Then
						strSQL = "insert into Interfaces (vcIntName,iDeviceID, vcdescription) values ('" & Interface & "'," & deviceid & ",'" & description & "')"
						'wscript.echo strsql
						Cmd.CommandText = strSQL
						Cmd.Execute	
						strSQL = "select iIntID from Interfaces where vcIntName = '" & Interface & "' and iDeviceID=" & deviceid
						rs.Open strSQL, cn
						If not rs.eof Then 
							intid = rs.fields(0).value
						End If 
						rs.close
					End If 
					strsql = "insert into ACL.dbo.iPAddr (vcIP, vcSubnet, iIPType, iInstance, iInterfaceID) values ('" & _
								InterfaceIP & "','" & interfacemask & "'," & IPType & "," & ipinst & "," & intid & ")"
					Cmd.CommandText = strSQL
					Cmd.Execute
				End If  
				If strparts(1) = "access-group" and Interface <> "" Then
					'wscript.echo "Found Access group with number of words " & UBound(strparts)
					'wscript.echo strline
					If UBound(strparts) = 3 Then
						strparts(2) = replace(strparts(2),"'","''")
						ACLID = GetID("ACLName","iACLID","vcACLName",strparts(2))
						Select Case strparts(3)
							Case "in"
								strsql = "update acl.dbo.interfaces set iACL_in_ID = " & ACLID & " where iIntID = " & intid
							Case "out"
								strsql = "update acl.dbo.interfaces set iACL_out_ID = " & ACLID & " where iIntID = " & intid
						End Select
						'wscript.echo "Updating interface with ACLName"
						'wscript.echo strsql
						Cmd.CommandText = strSQL
						Cmd.Execute							
					End If 	
					ipinst = 1					
				End If 
				
				If strparts(1) = "helper-address" and Interface <> "" Then
					iptype = 4 'Helper
					strsql = "insert into ACL.dbo.iPAddr (vcIP, vcSubnet, iIPType, iInstance, iInterfaceID) values ('" & _
								strparts(2) & "','255.255.255.255'," & IPType & "," & ipinst & "," & intid & ")"
					Cmd.CommandText = strSQL
					Cmd.Execute
				End If 
								
				If strparts(0) = "standby" and Interface <> "" Then
					'wscript.echo "found standby number of words " & UBound(strparts)
					'wscript.echo strline
					If UBound(strparts) = 3 Then
						If strparts(2) = "ip" Then 
							HSRP = strparts(3)
							IPInst = strparts(1)
							IPType = 3 'HSRP
							strsql = "insert into ACL.dbo.iPAddr (vcIP, vcSubnet, iIPType, iInstance, iInterfaceID) values ('" & _
										HSRP & "','255.255.255.255'," & IPType & "," & ipinst & "," & intid & ")"
							Cmd.CommandText = strSQL
							Cmd.Execute
						End If 
					End If 						
				End If 
				If strparts(1) = "access-list" Then
					ACLName = strparts(3)
					ACLName = replace(ACLName,"'","''")
					If UBound(strparts) = 3 Then ACLID = GetID("ACLName","iACLID","vcACLName",ACLName)	
					linenum = 0				
				End If 
				If strparts(0) = "access-list" Then
					ACLName = strparts(1)
					ACLID = GetID("ACLName","iACLID","vcACLName",ACLName)
					linenum = 0
				End If
				If ACLID > 0 and interface = "" and strparts(1) <> "access-list" Then 
					'wscript.echo "writing ACL line to DB"
					If strparts(0) = "permit" or strparts(0) = "deny" or strparts(0) = "remark" or strparts(0) = "access-list" Then
						linenum = linenum + 1
						strline = replace(strline,"'","''")
						strsql = "INSERT INTO ACL.dbo.ACL_Line (iACLid,iDeviceID,iLineNum,vcACL_Line) VALUES (" & _ 
									ACLID & ", " & DeviceID & ", " & linenum & ", '" & strline & "')"
						'wscript.echo strsql
						Cmd.CommandText = strSQL
						Cmd.Execute
					End If 
				End If 	
			End If 
		Wend 
	End If 
Next

Cmd.CommandText = "update acl.dbo.Interfaces set iACL_in_ID = 0 where iACL_in_ID is null"
Cmd.Execute
Cmd.CommandText = "update acl.dbo.Interfaces set iACL_out_ID = 0 where iACL_in_ID is null"
Cmd.Execute

cn.close

Set FileObj    = nothing
Set fc         = nothing
Set f          = nothing
Set fso        = nothing
Set rs         = nothing
Set cn         = nothing	
Set cmd        = nothing


Sub InitVariables ()
	Interface     = ""
	InterfaceIP   = ""
	InterfaceMask = ""
	HSRP          = ""
'	HSRP2         = ""
	ACLName       = ""
	IntID         = 0
	ACLID         = 0
	linenum       = 0
	IPInst        = 1
End Sub

Function GetID( Tablename, IDFieldName, FieldName, Criteria)
Dim strSQL, resultID

	resultID = 0 
	'wscript.echo "InGetID for " & tablename & " looking for ID of " & criteria
	strSQL = "select " & IDFieldName & " from " & Tablename & " where " & FieldName & " = '" & Criteria & "'"
	'wscript.echo strsql
	rs.Open strSQL, cn

	If not rs.eof Then 
		resultID = rs.fields(0).value
	End If 
	rs.close
	
	If resultid < 1 Then 
		strSQL = "insert into " &  Tablename & " (" & FieldName & ") values ('" & Criteria & "')"
		'wscript.echo strsql
		Cmd.CommandText = strSQL
		Cmd.Execute	
		strSQL = "select " & IDFieldName & " from " & Tablename & " where " & FieldName & " = '" & Criteria & "'"
		rs.Open strSQL, cn
		If not rs.eof Then 
			resultID = rs.fields(0).value
		End If 
		rs.close
	End If 
	
	'wscript.echo "GetID = " & resultID
	GetID = resultID
	
End Function 