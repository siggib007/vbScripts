Option Explicit
Dim FileObj, strLine, fso, f, fc, f1, strOut, strParts, FolderSpec, strOutFileName
Dim objFileOut, strFileNameParts, strReport, strFileNameCriteria


'Const strFileNameCriteria = "BigIPSN"


If WScript.Arguments.Count <> 3 Then 
  WScript.Echo "Usage: parser inpath criteria outfile"
  WScript.Quit
End If

FolderSpec = WScript.Arguments(0)
strFileNameCriteria = WScript.Arguments(1)
strOutFileName = WScript.Arguments(2)
'strreport = Now & " Starting analyzing " & folderspec 
'strreport = strreport & String(65,"-") & vbcrlf
strreport = "Device,SN" & vbcrlf
wscript.echo Now & " Starting analyzing " & folderspec 
'wscript.echo strreport
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder(folderspec)
Set fc = f.Files
For Each f1 in fc
	If f1.name <> strOutFileName AND InStr(f1.name,strFileNameCriteria) > 0 Then
		strFileNameParts = split(f1.name,"_")
		Set FileObj = fso.opentextfile(folderspec & "\" & f1.name)
		While not fileobj.atendofstream
			strLine = Trim(FileObj.readline)
			If strline <> "" Then
				strparts = split(strline,":")
				If Trim(strparts(0)) = "Appliance SN" Then
						strOut = strFileNameParts(0) & "," & LCase(Trim(strparts(1)))
				End If  
			End If 
		Wend 
		If strOut <> "" Then 
			strreport = strreport & strout & vbcrlf
		End If 
		FileObj.close
		strOut = ""
	End If
Next

Set objFileOut = fso.createtextfile(strOutFileName)

objFileOut.write strreport
objfileout.close

Set FileObj = nothing
Set objfileout = nothing
Set fc = nothing
Set f = nothing
Set fso = nothing

wscript.echo strreport
wscript.echo Now & " Analysis complete"
