Option Explicit
Dim fso, strOutFileName, strInFileName, objFileOut, objFileIn, strHost
Dim strLine, strLineParts, strCurHost, strOut, strHostNameOut

strOutFileName = "C:\Users\siggib\Documents\Denali\CPTC\ASAHostConfig.txt"
strInFileName = "C:\Users\siggib\Documents\Denali\CPTC\BrickConfigs\RuleSets\Hosts.csv"

strHostNameOut = ""
strCurHost = ""
strOut = ""

Set fso = CreateObject("Scripting.FileSystemObject")
Set objFileOut = fso.createtextfile(strOutFileName)
If fso.fileexists(strInFileName) Then
	Set objFileIn = fso.opentextfile(strInFileName)
Else
	wscript.echo "Can't find " & strInFileName
	wscript.echo "Can't proceed without it. Exiting!!!"
	wscript.quit
End If
While not objFileIn.atendofstream
	strLine = objFileIn.readline
	strLineParts = split (Trim(strLine), ",")
	If strLineParts(0) <> "Host Group" Then 
		If strCurHost <> strLineParts(0) Then 
			strCurHost = strLineParts(0)
			wscript.echo "Creating Host " & strCurHost
			strOut = strOut & "object-group network HG-" & strLineParts(0) & vbcrlf
			If strLineParts(1) <> "" Then
				strOut = strOut & " description " & strLineParts(1) & vbcrlf
			End If
		End If
		strOut = strOut & " network-object "
		If strLineParts(2) <> "" Or strLineParts(5) <> ""  Then
			If strLineParts(2) = "" then
				strHost = "A-" & strLineParts(3)
			Else 
				strHost = strLineParts(2)
			End If 
			strHostNameOut = strHostNameOut & "name " & strLineParts(3) & " " & strHost
			If strLineParts(5) <> "" Then
				strHostNameOut = strHostNameOut & " description " & strLineParts(5)
			End If 
			strHostNameOut = strHostNameOut & vbcrlf
			If strLineParts(4) = "host" Or strLineParts(4) = "255.255.255.255" Then
				strOut = strOut & "host " & strHost
			Else 
				strOut = strOut & strHost & " " & strLineParts(4)
			End If 
		Else 
			If strLineParts(4) = "host" Or strLineParts(4) = "255.255.255.255" Then
				strOut = strOut & "host " & strLineParts(3)
			Else 
				strOut = strOut & strLineParts(3) & " " & strLineParts(4)
			End If 			
		End If
		strOut = strOut & vbcrlf
	End If
Wend
wscript.echo "writing to Output file " & strOutFileName
objFileOut.writeline strHostNameOut
objFileOut.writeline strOut
wscript.echo "Done!"
