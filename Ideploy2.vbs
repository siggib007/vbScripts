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
Dim iLoc1, iLoc2, iLoc3, iLoc4, strVlan, strSWPort, i, dVlan, dSWPort, DictKeys, iSectStart
Dim VlanArray(), SwitchArray(), x ,y, Dict, PortArray
	
	Set dVlan = CreateObject("Scripting.Dictionary")
	Set dSWPort = CreateObject("Scripting.Dictionary")
	Set Dict = CreateObject("Scripting.Dictionary")
	
	i = 0
	
	ReDim VlanArray(1,i)
	ReDim SwitchArray(2,i)

	iSectStart = InStr(strResults, "Ticket Title:")

	If iSectStart = 0 Then
		wscript.echo strResults
		wscript.quit
	End If 
	
	iSectStart = InStr(strResults,"Master Server Configurations")

	If iSectStart = 0 Then
		wscript.echo "iDeploy section not found"
		wscript.quit
	End If 
	
	iLoc1 = iSectStart
	Do While iLoc1 > 0 	
		ReDim preserve VlanArray(1,i)
		iLoc1 = InStr(iLoc1, strResults,"VLAN:")
		If iloc1 = 0 Then Exit Do
		iLoc1 = InStr(iLoc1, strResults,"<TD")
		iLoc2 = InStr(iLoc1, strResults,">")
		iLoc3 = InStr(iLoc2, strResults,"</TD>" )
		strVlan = Mid(strResults,iLoc2 + 1, iLoc3-iLoc2 - 1)
		strvlan = replace(strVlan, vbcrlf,"")
		strvlan = replace(strvlan, vbtab, "")
		strvlan = Trim(strvlan)
		vlanarray(0,i) = strvlan
		vlanarray(1,i) = iLoc2
		iloc1 = iloc3
		i = i + 1
	Loop
	
	'wscript.echo "i=" & i & "  UBound(VlanArray,2) = " & UBound(VlanArray,2)
	
	'i = i + 1
	ReDim preserve VlanArray(1,i )
	vlanarray(1,i) = Len(strResults)
	vlanarray(0,i) = "EOF"
	'wscript.echo "i=" & i & "  UBound(VlanArray,2) = " & UBound(VlanArray,2)

	
	'For i = 0 to UBound(VlanArray,2) 
	'	wscript.echo VlanArray(0,i) & "    " & VlanArray(1,i)
	'Next
	
	'wscript.echo vbcrlf & vbcrlf
	
	i = 0
	iLoc1 = iSectStart
	Do While iLoc1 > 0 	
		ReDim preserve SwitchArray(2,i)	
		iLoc1 = InStr(iLoc1, strResults,"Switch Port:")
		If iloc1 = 0 Then Exit Do		
		iLoc1 = InStr(iLoc1, strResults,"</TD>")
		iLoc1 = InStr(iLoc1, strResults,"<TD>")
		iLoc2 = InStr(iLoc1 + 4, strResults,">")
		iLoc3 = InStr(iLoc2, strResults,"</" )
		strSWPort = Mid(strResults,iLoc2 + 1, iLoc3-iLoc2 - 1)
		strSWPort = replace(strSWPort, vbcrlf,"")
		strSWPort = replace(strSWPort, vbtab, "")
		strSWPort = replace(strSWPort, "&nbsp;", "")
		strSWPort = Trim(strSWPort)
		PortArray = split(strSWPort, " ")
		'wscript.echo strSWPort & " : " & UBound(PortArray) & " : " & PortArray(0)
		'wscript.echo "1: " & iloc1 & "   2: " & iloc2 & "  3: " & iloc3
		'wscript.echo "Switch Port: " & strSWPort & vbtab & "Vlan: " & strvlan
		SwitchArray(0,i) = strSWPort
		SwitchArray(1,i) = iLoc2
		iloc1 = iloc3
		i = i + 1
	Loop 

	'For i = 0 to UBound(SwitchArray,2) - 1
	'	wscript.echo SwitchArray(0,i) & "    " & SwitchArray(1,i)
	'Next
	
	'wscript.echo vbcrlf & vbcrlf
	
	For x = 0 to UBound(VlanArray,2) - 1
		For y = 0 to UBound(SwitchArray,2) - 1
			If SwitchArray(1,y) > VlanArray(1,x) and SwitchArray(1,y) < VlanArray(1,x+1) Then
				If Dict.exists(VlanArray(0,x)) Then
					'wscript.echo "Adding " & SwitchArray(0,y) & " to existing vlan " & VlanArray(0,x)  & vbcrlf
					'wscript.echo "Previous members: " & vbcrlf & Dict.item(VlanArray(0,x))  & vbcrlf
					Dict.item(VlanArray(0,x)) = dict.item(VlanArray(0,x)) & vbcrlf & SwitchArray(0,y)
				Else
					'wscript.echo "Adding " & SwitchArray(0,y) & " to New vlan " & VlanArray(0,x) & vbcrlf
					Dict.add VlanArray(0,x), SwitchArray(0,y)
				End If 		
			End If 
		Next
	Next

	DictKeys = Dict.Keys

	wscript.echo vbcrlf & vbcrlf
	
	For i = 0 to Dict.count - 1
		wscript.echo DictKeys(i) & vbcrlf & BubbleSort(Dict.item(DictKeys(i))) & vbcrlf
	Next

End Sub


Function BubbleSort(ByVal strValue)
  Dim arrUnsorted
  Dim arrSorted
  Dim bolSorted
  Dim strSorted, i, strTemp

  arrUnsorted = Split(strValue, vbCrLf)

  bolSorted = False

  Do Until bolSorted
    bolSorted = True
    For I = 0 to UBound(arrUnsorted)
      ' Compare this entry to the next entry
      If I < UBound(arrUnsorted) Then
        If arrUnsorted(I+1) < arrUnsorted(i) Then
          strTemp = arrUnsorted(I+1)
          arrUnsorted(I+1) = arrUnsorted(I)
          arrUnsorted(I) = strTemp
          bolSorted = False
        End If
      End If
    Next
  Loop

  strSorted = ""

  For I = 0 to UBound(arrUnsorted)
    strSorted = strSorted & arrUnsorted(I) & vbCrLf
  Next

  BubbleSort = strSorted

End Function
