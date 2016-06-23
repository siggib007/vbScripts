Option Explicit
Dim HTTP
Set HTTP = CreateObject("Microsoft.XMLHTTP")
If WScript.Arguments.Count <> 1 Then 
  WScript.Echo "Usage: HTTPGet2 URL"
  WScript.Quit
End If
HTTP.Open "GET", WScript.Arguments(0), False
HTTP.Send
If HTTP.statusText = "OK" Then
  wscript.echo HTTP.responseText
Else
  Wscript.Echo "Error getting page:" & HTTP.statusText
End If

