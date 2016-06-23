Option Explicit
Dim FileObj, strLine, fso, f, fc, f1, strOut, strParts, FolderSpec, strOutFileName, objFileOut, strFileNameParts, x

If WScript.Arguments.Count <> 2 Then 
  WScript.Echo "Usage: parser inpath outfilename"
  WScript.Quit
End If

FolderSpec = WScript.Arguments(0)
strOutFileName = WScript.Arguments(1)

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder(folderspec)
Set objFileOut = fso.createtextfile(strOutFileName)
Set fc = f.Files
For Each f1 in fc
	If f1.name <> strOutFileName AND InStr(f1.name,"_sh_ipc_queue") > 0 Then
		strFileNameParts = split(f1.name,"_")
		Set FileObj = fso.opentextfile(folderspec & "\" & f1.name)
		'If not fileobj.atendofstream Then FileObj.readline
		While not fileobj.atendofstream
			strLine = Trim(FileObj.readline)
			If strline <> "" Then
				'wscript.echo strline
				strparts = split(strline," ")
				If IsNumeric(strparts(2)) Then
					strOut = f1.DateLastModified & "," & strFileNameParts(0) & "," & strparts(2) & ","
					For x=3 to UBound(strparts) 
						strout = strout & strparts(x) & " " 
					Next
				End If  
			End If 
		Wend 
		If strOut <> "" Then 
			wscript.echo strOut
			objFileOut.writeline strOut
		End If 
		FileObj.close
		strOut = ""
	End If
Next

objFileOut.close
Set FileObj = nothing
Set objFileOut = nothing
Set fc = nothing
Set f = nothing
Set fso = nothing