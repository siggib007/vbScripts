Option Explicit
Const cdoSendUsingPort = 2
Const cdoNTLM = 2

Dim attach(), cn, rs, fso, conffile, FileObj, strline, objfileout, strparts, FromAddress, ToAddress
Dim CCAddress, Subject, MsgBody, DBServerName, DBName, cmdtext, strout, x, AttachResults, ResultDelim
Dim SQLQueryFile, OutFile, fld, sformat, cmd, SQLTimeOut, SMTPTimeout, MailServerName

ReDim attach(-1)

Set fso = CreateObject("Scripting.FileSystemObject")

If wscript.arguments.count > 0 Then
	ConfFile = wscript.arguments(0)
Else
	wscript.echo "Please provide Configuration file location"
	wscript.quit (1)
End If 

If not fso.fileexists(ConfFile) Then
	wscript.echo "Configuration file name not valid"
	wscript.quit(1)
End If 

wscript.echo "Processing conf file"

Set FileObj = fso.opentextfile(ConfFile)
While not fileobj.atendofstream
	strLine = Trim(FileObj.readline)
	strparts = split(strline,"=")
	Select Case LCase(Trim(strparts(0)))
		Case "fromaddress"
			FromAddress = Trim(strparts(1))
		Case "toaddress"
			ToAddress = Trim(strparts(1))
		Case "ccaddress"
			CCAddress = Trim(strparts(1))
		Case "subject"
			Subject = Trim(strparts(1))
		Case "msgbody"
			MsgBody = Trim(strparts(1))
		Case "attach"
			ReDim preserve attach(UBound(attach)+1)
			attach(UBound(attach)) = Trim(strparts(1))
		Case "dbservername"
			DBServerName = Trim(strparts(1))
		Case "dbname"
			DBName = Trim(strparts(1))
		Case "cmdtext"
			cmdtext = Trim(strparts(1))
		Case "sqlqueryfile"
			SQLQueryFile = Trim(strparts(1))
		Case "outfile"
			OutFile = Trim(strparts(1))
		Case "attachresults"
			AttachResults = Trim(strparts(1))
		Case "resultdelim"
			ResultDelim = Trim(strparts(1))
					Case "format"
			sFormat = LCase(Trim(strparts(1)))
		Case "SQLTimeOut"
			SQLTimeOut = CInt(strparts(1))
		Case "SMTPTimeout"
			SMTPTimeout = CInt(strparts(1))
		Case "MailServerName"
			MailServerName = Trim(strparts(1))
	End Select			
Wend
fileobj.close
Set fileobj = nothing

ResultDelim = replace(ResultDelim,"""", "")

Select Case ResultDelim 
	Case "tab"
		ResultDelim = vbtab
	Case "cr"
		ResultDelim = vbcrlf
End Select

If outfile = "" and LCase(AttachResults) <> "yes" Then
	If not fso.FolderExists("c:\temp") Then
		fso.CreateFolder("c:\temp")
	End If 
	outfile = "c:\temp\tempout.tmp"
End If 

If fso.fileexists(sqlqueryfile) Then
	wscript.echo "reading SQL Query"
	Set FileObj = fso.opentextfile(sqlqueryfile)
	cmdtext = fileobj.readall
	fileobj.close
	Set fileobj = nothing	
End If

If cmdtext <> "" Then
	If InStr(outfile,"\") = 0 or InStr(outfile,".") = 0 Then
		wscript.echo "Invalid output file name " & outfile
		wscript.echo "File name should be a complete path. For example C:\Output.txt"
		Set fso = nothing
		wscript.quit (1)
	End If
	
	wscript.echo "opening connection to SQL Server " & dbservername

	Set cn  = CreateObject("ADODB.Connection")
	Set RS  = CreateObject("ADODB.Recordset")
		
	cn.Provider = "sqloledb"
	cn.Properties("Data Source").Value = DBServerName
	cn.Properties("Initial Catalog").Value = DBName
	cn.Properties("Integrated Security").Value = "SSPI"
	cn.Open

	wscript.echo "Executing SQL query"
	Set objfileout = fso.createtextfile(outfile)
	RS.open cmdtext, cn
	strout = ""
	For Each fld In rs.Fields
		strout = strout & fld.name & ResultDelim
	Next
	objfileout.writeline Left(strout,Len(strout)-1)
	
	wscript.echo "generating results file"
	Do until rs.eof
		strout = ""
		For x = 0 to rs.fields.count - 1
			strout = strout & RS.fields(x).value & ResultDelim
		Next
		objfileout.writeline Left(strout,Len(strout)-1)
		RS.movenext
	loop
	objfileout.close
	Set objfileout = nothing
	rs.close
	If LCase(AttachResults) = "yes" Then
		wscript.echo "attaching results"
		ReDim preserve attach(UBound(attach)+1)
		attach(UBound(attach)) = outfile
	Else
		wscript.echo "putting results in email body"
		Set FileObj = fso.opentextfile(outfile)
		MsgBody = msgbody & vbcrlf & vbcrlf & fileobj.readall
		fileobj.close
		Set fileobj = nothing		
	End If 
	cn.close
	
	Set RS  = nothing
	Set cn  = nothing	
End If 

wscript.echo "Sending mail with Attachment count of " & UBound(attach) + 1
mysendmail subject,msgbody
wscript.echo "Mail sent"

Set fso = nothing


Sub MySendMail(StrSubject,msg)
	Dim iMsg,iConf,Flds, x
	
	Set iMsg = CreateObject("CDO.Message") 
	Set iConf = CreateObject("CDO.Configuration") 
	Set Flds = iConf.Fields 
	
	With Flds 
	  .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort 
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MailServerName 
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = SMTPTimeout
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")= cdoNTLM
	  .Update 
	End With 
	
	With iMsg 
	  Set .Configuration = iConf 
	      .To       = ToAddress
	      .CC	= CCAddress
	      .From     = FromAddress 
	      .Subject  = StrSubject 
	      .textbody = Msg
	      For x = 0 to UBound(attach)
	      	.AddAttachment(attach(x))
	      Next
	      .Send 
	End With
End Sub
