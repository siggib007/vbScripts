option explicit
Dim LogFileObj, strLine, fso, Strlastfile, array1, array2,tmpFileObj

Const ForReading = 1
Const strLogFileName = "LabServerLog.txt"
Const strTmpFileName = "Checklog.tmp"
Const strComplete = "Job Complete."
Const strProfileName = "Siggi Bjarnason"
Const strSubject = "WMI Collector is hung"
Const strMessage = "Please fix"
Const strRecpt = "Siggib;2062951027@mobile.att.net"

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
If Trim(array2(1)) = strComplete Then
	wscript.echo "Done"
	wscript.quit
End If

If fso.FileExists(strlogfilename) Then
	Set LogFileObj = fso.OpenTextFile(strlogFileName, ForReading)
Else
	wscript.echo "Logfile " & strlogfilename & " not found"
	wscript.quit
End If 

While not LogFileObj.atendofstream
	strLine = LogFileObj.readline
Wend

array1=split(strline,vbtab)

If (Trim(array1(0)) = Trim(array2(0))) and (Trim(array1(1)) <> strComplete) Then
	wscript.echo "Collector is hung, please fix"
	wscript.echo strline
	'mysendmail strprofilename,strrecpt,strsubject,strmessage
	Set tmpfileobj=fso.createtextfile(strtmpfilename,true) 
	tmpfileobj.writeline "Collector Hung" & vbtab & strComplete
	tmpfileobj.close
Else
	wscript.echo "Everything seems OK"
	Set tmpfileobj=fso.createtextfile(strtmpfilename,true) 
	tmpfileobj.writeline strline
	tmpfileobj.close
End If 

  
Sub MySendMail(profilename,strrecipient,subject,msg)
Dim objSession, oInbox, colMessages, oMessage, colRecipients,recipientarray,recipient

	Set objSession = CreateObject("MAPI.Session")
	objSession.Logon profilename
	
	Set oInbox = objSession.Inbox
	Set colMessages = oInbox.Messages
	Set oMessage = colMessages.Add()
	Set colRecipients = oMessage.Recipients
	
	recipientarray = split (strrecipient,";")
	For each recipient in recipientarray
		colRecipients.Add recipient
		colRecipients.Resolve
	Next
	
	oMessage.Subject = subject
	oMessage.Text = msg
	oMessage.Send
	
	objSession.Logoff
	Set objSession = nothing

End Sub 
