Option Explicit
Dim fso, strOutFileName, strInFileName, objFileOut, objFileIn
Dim strLine, strLineParts, strCurTimeRange, strOut, strRuleName

strOutFileName = "C:\Users\siggib\Documents\Work Stuff\Denali\CPTC\ASARulesConfig.txt"
strInFileName = "C:\Users\siggib\Documents\Work Stuff\Denali\CPTC\BrickConfigs\RuleSets\RuleSets.csv"

Set fso = CreateObject("Scripting.FileSystemObject")
Set objFileOut = fso.createtextfile(strOutFileName)
If fso.fileexists(strInFileName) Then
	Set objFileIn = fso.opentextfile(strInFileName)
Else
	wscript.echo "Can't find " & strInFileName & ", can't proceed without it. Exiting!!!"
	wscript.quit
End If
While not objFileIn.atendofstream
	strLine = objFileIn.readline
	strLineParts = split (Trim(strLine), ",")
	If strLineParts(0) <> "Rule Sets" Then 
'		If strLineParts(3) = "" Then 
		If strLineParts(2) <> "" Then 
			If strLineParts(3) = "both" Then 
				objFileOut.writeline "access-list " & strLineParts(0) & "-in remark " & strLineParts(2)
				objFileOut.writeline "access-list " & strLineParts(0) & "-out remark " & strLineParts(2)
			Else
				objFileOut.writeline "access-list " & strLineParts(0) & "-" & strLineParts(3) & " remark " & strLineParts(2)
			End If
		End If 
		strOut = "access-list " & strLineParts(0) & " extended " & strLineParts(7)
		If strLineParts(6) = "" Then 
			strOut = strOut & " ip"
		Else
			strOut = strOut & " object-group Srv-" & strLineParts(6)
		End If
		If strLineParts(4) = "" Then 
			strOut = strOut & " any"
		Else
			If IsIPAddr (strLineParts(4)) Then 
				strOut = strOut & " host " & strLineParts(4)
			Else
				strOut = strOut & " object-group HG-" & strLineParts(4)
			End If  
		End If
		If strLineParts(5) = "" Then 
			strOut = strOut & " any"
		Else
			If IsIPAddr (strLineParts(5)) Then 
				strOut = strOut & " host " & strLineParts(5)
			Else
				strOut = strOut & " object-group HG-" & strLineParts(5)
			End If  
		End If
		If strLineParts(8) <> "" Then 
			strOut = strOut & " time-range " & strLineParts(8)
		End If
		If strLineParts(1) = "no" Then 
			strOut = strOut & " inactive "
		End If
		If strLineParts(3) = "both" Then
			strOut = replace(strOut,strLineParts(0),strLineParts(0) & "-in")
			objFileOut.writeline strOut
			strOut = replace(strOut,strLineParts(0) & "-in",strLineParts(0) & "-out")
			objFileOut.writeline strOut
		Else
			strOut = replace(strOut,strLineParts(0),strLineParts(0) & "-" & strLineParts(3))
			objFileOut.writeline strOut
		End If 
	End If
Wend
wscript.echo "Done! Output file " & strOutFileName

Function IsIPAddr(strIP)
Dim strIPParts, iPartCount

	strIPParts = split(strIP, ".")
'	wscript.echo "Checking " & strIP & " to see if it is an IP address"
	iPartCount = UBound(strIPParts)
'	wscript.echo "There are " & iPartCount & " period seperated parts"
'	wscript.echo "is the first part a number? " & IsNumeric(strIPParts(0))
	If IsNumeric(strIPParts(0)) And iPartCount = 3 Then
'		wscript.echo "conclusion, " & strIP & " is an IP address"
		IsIPAddr = True
	Else
'		wscript.echo "conclusion, " & strIP & " is NOT an IP address"
		IsIPAddr = False
	End If
End Function 		
	