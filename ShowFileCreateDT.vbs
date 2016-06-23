Const MailServerName = "tk2smtp.phx.gbl"
Const SMTPTimeout = 10
Const FromAddress = "siggib@microsoft.com"
Const ToAddress = "siggib@microsoft.com"
Const CCAddress = ""
Const BCCAddress = " "
Const cdoSendUsingPort = 2
Const cdoNTLM = 2

Set cn  = CreateObject("ADODB.Connection")
Set cmd = CreateObject("ADODB.Command")
Set rs  = CreateObject("ADODB.Recordset")
Set fso = CreateObject("Scripting.FileSystemObject")

wscript.echo "Number of arguments supplied: " & wscript.arguments.count
If wscript.arguments.count > 1 Then
	sTable   = wscript.arguments(0)
	subject = wscript.arguments(1)
Else
	wscript.echo "Supply all arguments"
	wscript.quit(1)
End If


DBServerName = "by2netsql01"     ' -- Replace with db server name
DBName       = "cmdb"     ' -- Replace with db name
UserName     = "ScriptRW"     ' -- Replace with user name
Password     = "thisbites2."     ' -- Replace with pwd

'subject = "OOBFile Update"
'stable = "CMDB.dbo.OOBListLocations"

cn.Provider = "sqloledb"
cn.Properties("Data Source").Value = DBServerName
cn.Properties("Initial Catalog").Value = DBName
cn.Properties("User ID").Value = UserName
cn.Properties("Password").Value = Password
cn.Open

Cmd.ActiveConnection = cn
cmd.CommandTimeout = 60

outstr = ""
cmdtext = "select Path from " & stable & " where path is not null"
rs.open cmdtext, cn
Do until rs.eof
	sPath = rs.fields(0).value
	If fso.fileexists(spath) Then
		dtModified = showfileinfo(spath)
		cmdtext = "update " & stable & " set LastUpdated ='" & dtmodified & "' where path ='" & spath & "'"
		outstr = outstr & "<" & spath & "> modified on: " & dtmodified & vbcrlf
	Else
		cmdtext = "update " & stable & " set LastUpdated = null where path ='" & spath & "'"
		outstr = outstr & "Cannot find <" & spath & ">" & vbcrlf
	End If 
	Cmd.CommandText=cmdtext
	Cmd.Execute    
	rs.movenext
loop

wscript.echo outstr

mysendmail subject, outstr

wscript.echo "Mail sent"

Set fso = nothing
Set cn  = nothing
Set cmd = nothing
Set rs  = nothing

Function ShowFileInfo(filespec)
	
	Set f = fso.GetFile(filespec)
	ShowFileInfo =  f.DateLastModified 

End Function


Sub MySendMail(StrSubject,msg)
	Dim iMsg,iConf,Flds
	
	Set iMsg = CreateObject("CDO.Message") 
	Set iConf = CreateObject("CDO.Configuration") 
	Set Flds = iConf.Fields 
	
	With Flds 
	  .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort 
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MailServerName 
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = SMTPTimeout
	  .Update 
	End With 
	
	With iMsg 
	  Set .Configuration = iConf 
	      .To       = ToAddress
	      .CC	    = CCAddress
	      .bcc      = BCCAddress
	      .From     = FromAddress 
	      .Subject  = StrSubject 
	      .textbody = Msg
	      .Send          
	End With
	
	Set iConf = nothing
	Set iMsg = nothing
End Sub
