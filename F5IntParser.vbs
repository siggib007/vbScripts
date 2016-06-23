Option Explicit

Dim FileObj, strLine, fso, f, fc, f1, strOut, strParts, FolderSpec, strOutFileName, objFileOut, bRightSection, trunkname, x

If WScript.Arguments.Count <> 2 Then
  WScript.Echo "Usage: " & wscript.scriptname & " inpath, outfilename"
  WScript.Quit
End If

FolderSpec = WScript.Arguments(0)
strOutFileName = WScript.Arguments(1)

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder(folderspec)
Set objFileOut = fso.createtextfile(strOutFileName)
Set fc = f.Files
strOut = "Device,PortChannelName,Interface" '& vbcrlf
For Each f1 in fc
	If f1.name <> strOutFileName Then
		Set FileObj = fso.opentextfile(folderspec & "\" & f1.name)
		While not fileobj.atendofstream
			strLine = Trim(FileObj.readline)
			If strline = "# '/config/bigip_base.conf' file" Then
				bRightSection = True
				'logout "Found right section in " & f1.name
			End If
			If bRightSection Then
				If Left(strline,5) = "trunk" Then
					If strOut <> "" Then
						'wscript.echo strOut
						'objFileOut.writeline strOut
						logout strout
						strout = ""
					End If
					trunkname = Left(f1.name,InStrRev(f1.name,".")-1) & "," & Mid(strline,7,Len(strline)-8)
					'logout "Found trunk: " & trunkname
				Else
					If Left(strline,9) = "interface" Then
						If trunkname <> "" Then
							strParts = split(strline," ")
							'wscript.echo "Found a interface line consisting of " & UBound(strparts) & " parts: " & strline
							'logout "Found a interface line consisting of " & UBound(strparts) & " parts: " & strline
							For x= 1 to UBound(strparts)
								strOut = strOut & trunkname & "," & strparts(x) & vbcrlf
							Next
						Else
							strParts = split(strline," ")
							If strparts(1) <> "mgmt" Then strOut = strOut & Left(f1.name,InStrRev(f1.name,".")-1) & ",n/a," & strparts(1)	& vbcrlf
						End If
					End If
				End If
				If Left(strline,4) = "vlan" Then
					brightsection = false
					'logout "Found vlan section, exiting"
					trunkname = ""
				End If
			End If
		Wend
		If strOut <> "" Then
			'wscript.echo strOut
			logout strOut
		End If
		FileObj.close
		strOut = ""
	End If
Next

objFileOut.close

Set FileObj = fso.opentextfile(strOutFileName)
strline = fileobj.readall
strline = replace(strline,vbcrlf&vbcrlf,vbcrlf)
fileobj.close

Set objFileOut = fso.createtextfile(strOutFileName)
objFileOut.write strline
objfileout.close

Set FileObj = nothing
Set objFileOut = nothing
Set fc = nothing
Set f = nothing
Set fso = nothing

Sub logout(strText)
	wscript.echo strText
	objFileOut.writeline strText
End Sub
