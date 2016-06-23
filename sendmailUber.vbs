Option Explicit
'Const MailServerName = "smtphost.redmond.corp.microsoft.com" 
'Const MailServerName = "satsmtpa01"
'Const MailServerName = "tk2smtp.phx.gbl"
Const cdoSendUsingPort = 2
Const cdoNTLM = 2

Dim attach(), cn, rs, fso, conffile, FileObj, strline, objfileout, strparts, FromAddress, ToAddress
Dim CCAddress, Subject, MsgBody, DBServerName, DBName, cmdtext, strout, x, AttachResults, ResultDelim
Dim SQLQueryFile, OutFile, fld, sformat, cmd, SQLTimeOut, SMTPTimeout, MailServerName, SMTPAuth

ReDim attach(-1)

Set fso = CreateObject("Scripting.FileSystemObject")

If wscript.arguments.count > 0 Then
	ConfFile = wscript.arguments(0)
Else
	wscript.echo "Please provide Configuration file location"
	wscript.quit (1)
End If 

If not fso.fileexists(ConfFile) Then
	wscript.echo "Configuration file " & conffile & " can not be found."
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
		Case "sqltimeout"
			SQLTimeOut = CInt(strparts(1))
		Case "smtptimeout"
			SMTPTimeout = CInt(strparts(1))
		Case "mailservername"
			MailServerName = Trim(strparts(1))
		Case "smtpauth"
			SMTPAuth = LCase(Trim(strparts(1)))
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
	Set cmd = CreateObject("ADODB.Command")
		
	cn.Provider = "sqloledb"
	cn.Properties("Data Source").Value = DBServerName
	cn.Properties("Initial Catalog").Value = DBName
	cn.Properties("Integrated Security").Value = "SSPI"

	On Error Resume Next

	cn.Open
	checkerror "Error while open connection to SQL server."
  	
	On Error GoTo 0
	
	Cmd.ActiveConnection = cn
	cmd.CommandTimeout = SQLTimeOut

	wscript.echo "Executing SQL query"
	
	If sformat = "html" Then
		strparts = split(outfile,".")
		wscript.echo "File Extension is:" & strparts(1) & "!"
		If strparts(1) <> "html" and strparts(1) <> "htm" Then
			outfile = outfile & ".html"
		End If 
	End If 
	
	wscript.echo "cmdtext = " & cmdtext
	Set objfileout = fso.createtextfile(outfile)
	cmd.commandtext = cmdtext

	On Error Resume Next	
	Set RS = cmd.Execute
	checkerror "Error while executing query"
	On Error GoTo 0
	
	If sformat = "html" Then
		FormatResultHTML
	Else 
		FormatResultText
	End If 

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
		If sformat = "html" Then
			msgbody = fileobj.readall
		Else
			MsgBody = msgbody & vbcrlf & vbcrlf & fileobj.readall
		End If 
		fileobj.close
		Set fileobj = nothing		
	End If 
	cn.close
	
	Set cmd = nothing
	Set RS  = nothing
	Set cn  = nothing	
End If 

wscript.echo "Sending mail with Attachment count of " & UBound(attach) + 1 & " through " & MailServerName
mysendmail subject,msgbody
wscript.echo "Mail sent to " & ToAddress & ";" & CCAddress

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
	  If SMTPAuth = "yes" Then .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")= cdoNTLM
	  .Update 
	End With 
	
	With iMsg 
	  Set .Configuration = iConf 
	      .To            = ToAddress
	      .CC	           = CCAddress
	      .From          = FromAddress 
	      .Subject       = StrSubject
	      If sformat     = "html" Then 
	      	.HTMLBody    = msg
	      Else
	      	.textbody = Msg
	      End If
	      On Error Resume next
	      For x = 0 to UBound(attach)
	      	wscript.echo "Now attaching " & attach(x)
	      	.AddAttachment(attach(x))
	      	checkerror "Error while attaching."
	      Next
	      .Send
	      checkerror "Error while sending email." 
	      On Error GoTo 0   
	End With
End Sub

Sub FormatResultText ()

	wscript.echo "generating text results file"

	strout = ""

	For Each fld In rs.Fields
		strout = strout & fld.name & ResultDelim
	Next
	objfileout.writeline Left(strout,Len(strout)-1)
	
	Do until rs.eof
		strout = ""
		For x = 0 to rs.fields.count - 1
			strout = strout & RS.fields(x).value & ResultDelim
		Next
		objfileout.writeline Left(strout,Len(strout)-1)
		RS.movenext
	loop
End Sub

Sub FormatResultHTML()
Dim X, fldValue, strtd
Const Indent = "    "

	wscript.echo "generating HTML results file"
	
	strOut = "<html>" & vbcrlf
	strout = strout & "<head>" & vbcrlf
	strout = strout & "<title></title>" & vbcrlf
	strout = strout & "</head>" & vbcrlf & vbcrlf
	strout = strout & "<body>" & vbcrlf & vbcrlf
	strout = strout & msgbody & "<br>" & vbcrlf
	strout = strout & "<table border=1 cellpadding=5 >" & vbcrlf
  	strout = strout & indent & "<tr>" & vbcrlf
	For Each fld In rs.Fields
		strout = strout & indent & indent & "<td align=center><b>" & fld.name & "</b></td>" & vbcrlf
	Next
  	strout = strout & indent & "</tr>" & vbcrlf
  	strout = strout & indent & "<tr>" & vbcrlf
	While not rs.eof
	  For x=0 to rs.fields.count - 1
	  	Select Case rs.fields(x).type 'What type of column is this?
	  		Case 3 'Int
	  			fldValue = rs.fields(x).value 'FormatNumber (rs.fields(x).value,0,TristateTrue,TristateFalse,TristateTrue)
	  			If IsNull(fldValue) Then fldvalue = 0
	  			strTD= "<td align=right>"
	  		Case 4,5 ' Numeric
		  		fldValue =rs.fields(x).value' FormatPercent (rs.fields(x).value,2,TristateTrue,TristateFalse,TristateTrue)
		  		If IsNull(fldValue) Then fldvalue = 0
		  		strTD= "<td align=right>"
		  	Case 200 'String
		  		fldValue = Trim(rs.fields(x).value)
		  		If fldValue="" or IsNull(fldValue) Then 
		  			'wscript.echo "empty string replaced with space"
		  			fldvalue = "<br>"
		  		End If 
		  		If IsNumeric(fldvalue) Then
		  			strTD= "<td align=right>"
		  		Else
		  			strTD= "<td>"
		  		End If 
		  	Case Else
		  		fldValue = rs.fields(x).value
		  		strTD= "<td>"
		End Select 
		If fldValue = "zTotal" Then fldValue = "Total"
		strout = strout & indent & indent & strTD & fldValue & "</td>" & vbcrlf
	  Next
  	  strout = strout & indent & "</tr>"  & vbcrlf
  	  strout = strout & indent & "<tr>"  & vbcrlf
	  rs.movenext
	Wend
	strout = strout & indent & "</tr>" & vbcrlf
	strout = strout & "</table>" & vbcrlf
	strout = strout & "</body>" & vbcrlf
	strout = strout & "</html>" & vbcrlf	
	objfileout.writeline strout
End Sub

Sub CheckError (strMsg)
	If err.number <> 0 Then
		wscript.echo strmsg & ". Error number:" & err.number & " " & err.Description
		wscript.quit (1)
	End If
End Sub
