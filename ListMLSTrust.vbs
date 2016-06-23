Option Explicit
Dim FileObj, strLine, fso, f, fc, f1, strParts, FolderSpec, strOutFileName, objFileOut
Dim strDesc, strACL, strTrust, strInt 

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
output "Device;Interface;Description;ACL;trustqos"
For Each f1 in fc
	If f1.name <> strOutFileName And (InStr(f1.name,"76e") > 0  Or InStr(f1.name,"6ne") > 0 ) Then
		Set FileObj = fso.opentextfile(folderspec & "\" & f1.name)
		While not fileobj.atendofstream
			strLine = Trim(FileObj.readline)
			If Left(strLine,9) = "interface" Then
				If strInt <> "" Then 
					output strInt & ";" & strDesc & ";" & strACL & ";" & strTrust
				End If 
				strInt = Left(f1.name,InStrRev(f1.name,".")-1) & ";" & Right(strLine,Len(strLine)-10)
				strDesc = ""
				strACL = "" 
				strTrust = "False" 
			End If 
			If Left(strline,11) = "description" Then
				strDesc = Right(strLine,Len(strLine)-11)
			End If 
			If Left(strline,15) = "ip access-group" Then
				strParts = split(strLine," ") 
				strACL = strParts(2) 
			End If  
			If strLine = "mls qos trust dscp" Then
				strTrust = "true"
			End If 
		Wend 
		If strInt <> "" Then 
			output strInt & ";" & strDesc & ";" & strACL & ";" & strTrust
		End If 
		FileObj.close
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

strparts = split(strout,";")

If UBound(strparts) > 1 Then
	wscript.echo strout
	objFileOut.writeline strOut
End If 
	
End Sub