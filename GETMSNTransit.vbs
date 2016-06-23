Option Explicit
Dim HTTP, fso, outfileobj, URL
Set fso = CreateObject("Scripting.FileSystemObject")
Set HTTP = CreateObject("Microsoft.XMLHTTP")
Set outfileobj=fso.createtextfile("c:\temp\msntransit.csv",true) 
URL = "http://eng.jsnet.com/mrtg/billing/transit-costs." & DatePart("yyyy",Now()) & "." & DatePart("m",Now()) & ".csv"
wscript.echo "Fetching file " & url & " and saving it to c:\temp\msntransit.csv"
HTTP.Open "GET", URL, False
HTTP.Send
If HTTP.statusText = "OK" Then
  outfileobj.write HTTP.responseText
  wscript.echo "Success!"
Else
  Wscript.Echo "Error getting page:" & HTTP.statusText
End If

