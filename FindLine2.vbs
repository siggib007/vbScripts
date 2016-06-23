Option Explicit
Dim FileObj, strLine, fso, f, fc, f1, strOut, strParts, FolderSpec, strOutFileName, objFileOut, strcriteria

If WScript.Arguments.Count <> 3 Then 
	wscript.echo "Lists all lines in any files in the specified directory where criteria exists anywhere in the line."
  WScript.Echo "Usage: parser criteria inpath outfilename"
  WScript.Quit
End If

FolderSpec = WScript.Arguments(1)
strOutFileName = WScript.Arguments(2)
strCriteria = wscript.arguments(0)

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder(folderspec)
Set objFileOut = fso.createtextfile(strOutFileName)
Set fc = f.Files
objFileOut.writeline "Device,Line"
For Each f1 in fc
	If f1.name <> strOutFileName Then
		Set FileObj = fso.opentextfile(folderspec & "\" & f1.name)
		While not fileobj.atendofstream
			strLine = Trim(FileObj.readline)
			If InStr(strline, strcriteria) > 0 Then
				wscript.echo Left(f1.name,InStrRev(f1.name,".")-1) & "," & strline
				objFileOut.writeline Left(f1.name,InStrRev(f1.name,".")-1) & "," & strline
			End If  
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