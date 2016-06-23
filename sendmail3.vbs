On Error Resume Next
Set objNewMail = CreateObject("CDONTS.NewMail") 
objNewMail.Send "a-siggib@microsoft.com", "a-siggib@microsoft.com", "Hello", _ 
                  "I sent this in 3 statements!", 0 ' low importance 
If Err.Number Then
	wscript.echo "Error: " & CStr(err.number) & " : " & err.description & " Occurred"
End If 
Set objNewMail = Nothing ' canNOT reuse it for another messag