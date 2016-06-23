Option Explicit
'Option Compare text

Dim FileObj, strLine, fso, strOut, strParts, FileSpec, strCriteria

If WScript.Arguments.Count <> 2 Then 
  WScript.Echo "Usage: intparser inpathfilename, searchstring"
  WScript.Quit
End If

FileSpec = WScript.Arguments(0)
strCriteria = WScript.Arguments(1)

Set fso = CreateObject("Scripting.FileSystemObject")
strOut = ""
'wscript.echo strOut
Set FileObj = fso.opentextfile(filespec)
While not fileobj.atendofstream
	strLine = Trim(FileObj.readline)
	If Left(strline,9) = "interface" Then
		If strOut <> "" and UBound(split(strout,vbcrlf))>0 Then 
			wscript.echo strOut
			'objFileOut.writeline strOut
		Else
		'	wscript.echo strout & " " &  UBound(split(strout,vbcrlf))
		End If 
		'strOut = Left(FileSpec,InStrRev(FileSpec,".")-1) & "," & Right(strline,Len(strline)-10)
		strOut = strline  'Right(strline,Len(strline)-10)
	Else
		If InStr(LCase(strline),LCase(strcriteria)) > 0 Then
			strOut = strOut & vbcrlf & " " & strline
		End If 
	End If  
Wend 
If strOut <> "" and UBound(split(strout,","))=1 Then 
	wscript.echo strOut
	'objFileOut.writeline strOut
End If 
FileObj.close
strOut = ""

Set FileObj = nothing
Set fso = nothing