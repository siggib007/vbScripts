Option Explicit
Dim LogFileObj, strLine, fso, Strlastfile, array1, array2,tmpFileObj

Const ForReading = 1
Const strLogFileName = "DumpFileLog.txt"
Const strTmpFileName = "Checklog.tmp"
Const strComplete = "Job Complete"
Const Subject = "Dump File Generation has failed to execute"
Const MsgBody = "Please fix"
Const SMTPTimeout = 10
Const FromAddress = "ghtools@microsoft.com"
Const ToAddress = "siggib@microsoft.com"
Const CCAddress = ""'"2062951027@mobile.att.net"
Const cdoSendUsingPort = 2
Const MailServerName = "smarthost.dns.microsoft.com" 

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.fileexists(strtmpfilename) Then
	Set tmpFileObj = fso.opentextfile(strtmpfilename)
	If not tmpfileobj.atendofstream Then
		strlastFile = tmpfileobj.readline
	Else
		strlastfile = "blank" & vbtab & "file"

	End If 
	tmpfileobj.close
Else
	strlastfile = "blank" & vbtab & "file"
End If
'wscript.echo strlastfile
array2 = split(strlastfile,vbtab)

If fso.FileExists(strlogfilename) Then
	Set LogFileObj = fso.OpenTextFile(strlogFileName, ForReading)
	While not LogFileObj.atendofstream
		strLine = LogFileObj.readline
	Wend

Else
	strline= Now & vbtab & "Logfile " & strlogfilename & " not found."
	wscript.echo "Logfile " & strlogfilename & " not found"
	'wscript.quit
End If 


array1=split(strline,vbtab)

If (Trim(array1(0)) = Trim(array2(0))) or (Trim(array1(1)) <> strComplete) Then
	wscript.echo subject 
	wscript.echo strline & vbcrlf
	mysendmail subject,strline
	wscript.echo "Mail has been sent"
Else
	Set tmpfileobj=fso.createtextfile(strtmpfilename,true) 
	tmpfileobj.writeline strline
	tmpfileobj.close
	wscript.echo "Everything seems OK"
End If 

  
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
	      .CC	= CCAddress
	      .From     = FromAddress 
	      .Subject  = StrSubject 
	      .textbody = Msg
	      .Send 
	End With
End Sub
