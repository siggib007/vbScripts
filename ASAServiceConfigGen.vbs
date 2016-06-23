Option Explicit
Dim fso, strOutFileName, strInFileName, objFileOut, objFileIn
Dim strLine, strLineParts, strCurService, strOut, iRangeParts

strOutFileName = "C:\Users\siggib\Documents\Denali\CPTC\ASAServiceConfig.txt"
strInFileName = "C:\Users\siggib\Documents\Denali\CPTC\AllServices.csv"

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
	If strLineParts(0) <> "Name" Then 
		If strCurService <> strLineParts(0) Then 
			strCurService = strLineParts(0)
			wscript.echo "Creating service " & strCurService
			objFileOut.writeline "object-group service Srv-" & strCurService
			If strLineParts(1) <> "" Then
				objFileOut.writeline " description " & strLineParts(1)
			End If
		End If
		strOut =" service-object " & strLineParts(2)
		If strLineParts(4) <> "" Then 
			strOut = strOut & " source"
			If InStr(strLineParts(4),"-") > 0 Then 
				iRangeParts = split (strLineParts(4),"-")
				strOut = strOut & " range " & iRangeParts(0) & " " & iRangeParts(1)
			Else
				strOut = strOut & " eq " & strLineParts(4)
			End If
		End If
		If strLineParts(3) <> "" Then 
			If InStr(strLineParts(3),"-") > 0 Then 
				iRangeParts = split (strLineParts(3),"-")
				strOut = strOut & " range " & iRangeParts(0) & " " & iRangeParts(1)
			Else
				strOut = strOut & " eq " & strLineParts(3)
			End If
		End If
		objFileOut.writeline strOut
	End If
Wend
wscript.echo "Done! Output file " & strOutFileName
