Option Explicit
Dim FileObj, strLine, fso, f, fc, f1, strOut, strParts, FolderSpec, strOutFileName, objFileOut, strcriteria,bPrint

' If WScript.Arguments.Count <> 3 Then
' 	wscript.echo "Lists all lines in any files in the specified directory that fall between specified lines"
'   WScript.Echo "Usage: parser criteria inpath outfilename"
'   WScript.Quit
' End If

' FolderSpec = WScript.Arguments(1)
' strOutFileName = WScript.Arguments(2)
' strCriteria = wscript.arguments(0)

FolderSpec = "C:\Users\sbjarna\Documents\IP Projects\Automation\GiACL\OMW-ABF-IN"
strOutFileName = "C:\Users\sbjarna\Documents\IP Projects\Automation\GiACL\OMW-ABF-IN-CDN-DMZ.txt"
bPrint = false

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder(folderspec)
Set objFileOut = fso.createtextfile(strOutFileName)
Set fc = f.Files
' objFileOut.writeline "Device,Line"
For Each f1 in fc
	If f1.name <> strOutFileName Then
		Set FileObj = fso.opentextfile(folderspec & "\" & f1.name)
		objFileOut.writeline "*** [" & Left(f1.name,InStrRev(f1.name,".")-1) & "] ***"
		While not fileobj.atendofstream
			strLine = Trim(FileObj.readline)
			If InStr(strline, "remark") > 0 Then
				If InStr(strline, "DMZ") > 0 or InStr(strline, "Cache") > 0 Then
					bPrint = true
				else
					bPrint = false
				end if
			End If
			if bPrint = true then objFileOut.writeline strline
		Wend
		FileObj.close
	End If
Next

objFileOut.close
Set FileObj = nothing
Set objFileOut = nothing
Set fc = nothing
Set f = nothing
Set fso = nothing