Option Explicit
Dim HTTP, strURL

Set HTTP = CreateObject("Microsoft.XMLHTTP")
If WScript.Arguments.Count <> 1 Then 
  WScript.Echo "Usage: ideploy ticketid"
  WScript.Quit
End If

If IsNumeric (WScript.Arguments(0)) Then
	If  WScript.Arguments(0) < 9999 Then
	  WScript.Echo "Please provide a valid ticket number"
	  WScript.Quit			
	End If 
Else 
  WScript.Echo "Please provide a valid ticket number"
  WScript.Quit
End If
	
strURL = "http://socticketing/2.0/printable.asp?CTZO=480&TSID=2&TID=" & WScript.Arguments(0)
'wscript.echo "Fectching: " & strurl
HTTP.Open "GET", strURL, False
HTTP.Send
If HTTP.statusText = "OK" Then
  'wscript.echo HTTP.responseText
  ParseHTML(HTTP.responseText)
Else
  Wscript.Echo "Error getting page:" & HTTP.statusText
End If

Public Sub ParseHTML(strResults)
Dim iLoc1, iLoc2, iLoc3, iLoc4, strVlan, strSWPort, i, Dict, DictKeys
	
	Set Dict = CreateObject("Scripting.Dictionary")

	iLoc1 = InStr(strResults,"Master Server Configurations")
	
	If iLoc1 = 0 Then
		wscript.echo "iDeploy section not found"
		wscript.quit
	End If 
	
	Do While iLoc1 > 0 	
		iLoc1 = InStr(iLoc1, strResults,"VLAN:")
		If iloc1 = 0 Then Exit Do
		iLoc1 = InStr(iLoc1, strResults,"<TD")
		iLoc2 = InStr(iLoc1, strResults,">")
		iLoc3 = InStr(iLoc2, strResults,"</TD>" )
		strVlan = Mid(strResults,iLoc2 + 1, iLoc3-iLoc2 - 1)
		strvlan = replace(strVlan, vbcrlf,"")
		strvlan = replace(strvlan, vbtab, "")
		strvlan = Trim(strvlan)
		'For i = 1 to Len(strvlan)
		'	wscript.echo Asc(Mid(strvlan,i,1))
		'Next
		'wscript.echo "Vlan: " & strvlan
		iloc1 = iloc3
		iLoc1 = InStr(iLoc1, strResults,"Switch Port:")
		iLoc1 = InStr(iLoc1, strResults,"</TD>")
		iLoc1 = InStr(iLoc1, strResults,"<TD>")
		iLoc2 = InStr(iLoc1 + 4, strResults,">")
		iLoc3 = InStr(iLoc2, strResults,"</" )
		strSWPort = Mid(strResults,iLoc2 + 1, iLoc3-iLoc2 - 1)
		strSWPort = replace(strSWPort, vbcrlf,"")
		strSWPort = replace(strSWPort, vbtab, "")
		strSWPort = replace(strSWPort, "&nbsp;", "")
		strSWPort = Trim(strSWPort)
		'wscript.echo "1: " & iloc1 & "   2: " & iloc2 & "  3: " & iloc3
		'wscript.echo "Switch Port: " & strSWPort & vbtab & "Vlan: " & strvlan
		iloc1 = iloc3
		If Dict.exists(strvlan) Then
			Dict.item(strvlan) = dict.item(strvlan) & vbcrlf & strSWPort
		Else
			Dict.add strvlan, strSWPort
		End If 		
	Loop 
	DictKeys = Dict.Keys

	For i = 0 to Dict.count -1 
		wscript.echo DictKeys(i) & vbcrlf & (Dict.item(DictKeys(i))) & vbcrlf
	Next

End Sub

Function SortIt(strValue)
  Dim arrUnsorted
  Dim arrSorted
  Dim strSorted
  arrUnsorted = strValue.split(vbCrLf)
  arrSorted = arrUnsorted.sort()
  strSorted = ""
  For I = 0 to UBound(arrSorted)
    strSorted = strSorted & arrSorted(I) & vbCrLf
  Next
  SortIt = strSorted
End Function
