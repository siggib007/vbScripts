Option Explicit
Dim FileObj, strLine, fso, f, fc, f1, strOut, strParts, FolderSpec, strOutFileName, objFileOut

If WScript.Arguments.Count <> 2 Then 
  WScript.Echo "Usage: configparser inpath, outfilename"
  WScript.Quit
End If

FolderSpec = WScript.Arguments(0)
strOutFileName = WScript.Arguments(1)

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder(folderspec)
Set objFileOut = fso.createtextfile(strOutFileName)
Set fc = f.Files
strOut = "Device,Interface,PortChannelMember"
For Each f1 in fc
	If f1.name <> strOutFileName Then
		Set FileObj = fso.opentextfile(folderspec & "\" & f1.name)
		While not fileobj.atendofstream
			strLine = Trim(FileObj.readline)
			If Left(strline,9) = "interface" Then
				If strOut <> "" Then 
					'wscript.echo strOut
					'objFileOut.writeline strOut
					output strout
				End If 
				strOut = Left(f1.name,InStrRev(f1.name,".")-1) & "," & Right(strline,Len(strline)-10)
			Else
				If Left(strline,13) = "channel-group" Then
					strParts = split(strline," ")
					strOut = strOut & "," & "PC" & strParts(1)
				End If 
			End If  
		Wend 
		If strOut <> "" Then 
			'wscript.echo strOut
			'objFileOut.writeline strOut
			output strout
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

Sub output (strOut)
Dim strparts

strparts = split(strout,",")

If UBound(strparts) > 1 Then
	wscript.echo strout
	objFileOut.writeline strOut
End If 
	
End Sub