Option Explicit
Dim HTTP, fso, outfileobj
Set fso = CreateObject("Scripting.FileSystemObject")
Set HTTP = CreateObject("Microsoft.XMLHTTP")
If WScript.Arguments.Count <> 2 Then 
  WScript.Echo "Usage: GetURL URL OutputFilename"
  WScript.Quit
End If
Set outfileobj=fso.createtextfile(WScript.Arguments(1),true) 
HTTP.Open "GET", WScript.Arguments(0), False
HTTP.Send
If HTTP.statusText = "OK" Then
  outfileobj.write HTTP.responseText
Else
  Wscript.Echo "Error getting page:" & HTTP.statusText
End If

