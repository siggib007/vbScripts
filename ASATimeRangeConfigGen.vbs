Option Explicit
Dim fso, strOutFileName, strInFileName, objFileOut, objFileIn
Dim strLine, strLineParts, strCurTimeRange

strOutFileName = "C:\Users\siggib\Documents\Work Stuff\Denali\CPTC\ASATimeRangeConfig.txt"
strInFileName = "C:\Users\siggib\Documents\Work Stuff\Denali\CPTC\BrickConfigs\RuleSets\TimeRules.csv"

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
		If strCurTimeRange <> strLineParts(0) Then 
			strCurTimeRange = strLineParts(0)
			wscript.echo "Creating Time Range " & strCurTimeRange
			objFileOut.writeline "time-range " & strCurTimeRange
		End If
		objFileOut.writeline " periodic " & strLineParts(3) & " " & strLineParts(1) & " to " & strLineParts(2)
	End If
Wend
wscript.echo "Done! Output file " & strOutFileName
