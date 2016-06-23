strfrom="Siggib@microsoft.com"
strto="siggib@microsoft.com"
strSubject="CDONTS.NewMail Testing"
strbody="This is a different sendmail test"
Set objSendMail = CreateObject("CDONTS.NewMail")
If Err.Number Then
	WScript.Echo "CDONTS not installed, error: " & CStr(Err.Number)
	WScript.Quit(1)
End If

objSendMail.From = strFrom
objSendMail.To = strTo
objSendMail.Subject = strSubject
objSendMail.Body = strBody
objSendMail.Send
