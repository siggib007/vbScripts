option explicit
Dim LogFileObj, strLine, fso, Strlastfile, array1, array2,tmpFileObj,logObj

Const ForReading = 1
Const strLogFileName = "e:\inv\wmi\checklog.log"
Const strTmpFileName = "e:\inv\wmi\Checkloglog.tmp"
Const strComplete = "Job Complete.xxx"
Const Subject = "Checklog not functioning"
Const MsgBody = "Please fix"
Const SMTPTimeout = 10
Const FromAddress = "ghtools@microsoft.com"
Const ToAddress = "siggib@microsoft.com"
Const CCAddress = ""
Const cdoSendUsingPort = 2
Const MailServerName = "smarthost.dns.microsoft.com" 
Const ForAppending = 8 

Set fso = CreateObject("Scripting.FileSystemObject")

Set logObj = fso.opentextfile("e:\inv\wmi\checkloglog.log",ForAppending,true)
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
'writelog strlastfile
array2 = split(strlastfile,vbtab)
If Trim(array2(1)) = strComplete Then
	writelog "Done"
	wscript.quit
End If

If fso.FileExists(strlogfilename) Then
	Set LogFileObj = fso.OpenTextFile(strlogFileName, ForReading)
Else
	writelog "Logfile " & strlogfilename & " not found"
	wscript.quit
End If 

While not LogFileObj.atendofstream
	strLine = LogFileObj.readline
Wend

array1=split(strline,vbtab)

If (Trim(array1(0)) = Trim(array2(0))) and (Trim(array1(1)) <> strComplete) Then
	writelog strline
	writelog "CheckLog doesn't seem to be running."
	mysendmail subject,strline
	Set tmpfileobj=fso.createtextfile(strtmpfilename,true) 
	tmpfileobj.writeline "Checklog not Running" & vbtab & strComplete
	tmpfileobj.close
	writelog "Mail sent"
Else
	writelog "Everything seems OK"
	Set tmpfileobj=fso.createtextfile(strtmpfilename,true) 
	tmpfileobj.writeline strline
	tmpfileobj.close
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

Sub writelog (msg)
	
	wscript.echo msg
	logObj.writeline Now & vbtab & msg

End Sub