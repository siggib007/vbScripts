Option Explicit
Dim FileObj, strLine, fso, f, fc, f1, strOut, strParts, FolderSpec, strOutFileName
Dim objFileOut, strFileNameParts, x, iLimit, strReport

Const MailServerName = "smtphost.redmond.corp.microsoft.com" 
'Const MailServerName = "Tk2smtp2.phx.gbl" 
Const SMTPTimeout = 10
Const FromAddress = """Siggi Bjarnason"" <siggib@microsoft.com>"
Const ToAddress = """Internal Netpro"" <inetpro@microsoft.com>"
Const CCAddress = ""
Const Subject = "IPC Queue report" 
Const cdoSendUsingPort = 2
Const cdoNTLM = 2

Const strFileNameCriteria = "_sh_ipc_queue"


If WScript.Arguments.Count <> 2 Then 
  WScript.Echo "Usage: parser inpath limit"
  WScript.Quit
End If

FolderSpec = WScript.Arguments(0)
iLimit = CInt(WScript.Arguments(1))
strreport = Now & " Starting analyzing " & folderspec & " limiting to " & ilimit & vbcrlf
strreport = strreport & String(65,"-") & vbcrlf

'wscript.echo strreport
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder(folderspec)
Set fc = f.Files
For Each f1 in fc
	If f1.name <> strOutFileName AND InStr(f1.name,strFileNameCriteria) > 0 Then
		strFileNameParts = split(f1.name,"_")
		Set FileObj = fso.opentextfile(folderspec & "\" & f1.name)
		While not fileobj.atendofstream
			strLine = Trim(FileObj.readline)
			If strline <> "" Then
				strparts = split(strline," ")
				If IsNumeric(strparts(2)) Then
					If CInt(strparts(2)) > ilimit Then 
						strOut = strFileNameParts(0) & "," & strparts(2) & ","
						For x=3 to UBound(strparts) 
							strout = strout & strparts(x) & " " 
						Next
					End If 
				End If  
			End If 
		Wend 
		If strOut <> "" Then 
			strreport = strreport & strout & vbcrlf
		End If 
		FileObj.close
		strOut = ""
	End If
Next

Set FileObj = nothing
Set fc = nothing
Set f = nothing
Set fso = nothing

wscript.echo strreport
wscript.echo Now & " Analysis complete"

mysendmail subject,strreport
wscript.echo "Mail sent"


Sub MySendMail(StrSubject,msg)
	Dim iMsg,iConf,Flds
	
	Set iMsg = CreateObject("CDO.Message") 
	Set iConf = CreateObject("CDO.Configuration") 
	Set Flds = iConf.Fields 
	
	With Flds 
	  .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort 
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MailServerName 
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = SMTPTimeout
	  .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")= cdoNTLM
	  .Update 
	End With 
	
	With iMsg 
	  Set .Configuration = iConf 
	      .To       = ToAddress
	      .CC	= CCAddress
	      .From     = FromAddress 
	      .Subject  = StrSubject 
	      .textbody = Msg
	      .Send 
	End With
End Sub
