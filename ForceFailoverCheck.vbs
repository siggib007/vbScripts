Option Explicit
Dim FileObj, strLine, fso, f, fc, f1, strOut, strParts, FolderSpec, strOutFileName, objFileOut
Dim bActive, bStandby, strActive, strStandby, bFoundActive, bFoundStandby, bExit

If WScript.Arguments.Count <> 2 Then 
  WScript.Echo "Usage: cscript ForceFailoverCheck.vbs inpath, outfilename"
  WScript.Quit
End If

FolderSpec = WScript.Arguments(0)
strOutFileName = WScript.Arguments(1)

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder(folderspec)
Set objFileOut = fso.createtextfile(strOutFileName)
Set fc = f.Files
strOut = "Device,Status"
wscript.echo strOut
objFileOut.writeline strOut
strout = ""

For Each f1 in fc
	'wscript.echo "variables init"
	bFoundActive = false
	bfoundStandby = false
	bActive = ""
	bStandby = ""
	bExit = False
	
	If f1.name <> strOutFileName Then
		'wscript.echo "Processing " & f1.name
		Set FileObj = fso.opentextfile(folderspec & "\" & f1.name)
		bExit = fileobj.atendofstream
		While not bExit
			strLine = Trim(FileObj.readline)
			'wscript.echo strline
			If strline = "[Failover.ForceActive]" Then  
				bFoundActive = True
			Else
				If bFoundStandby = True and bActive = "" Then 
					bActive = "N/A"
					strActive = "N/A"
					'wscript.echo "Didn't find value entry under ForceActive, bactive: " & bactive
				End If 				
			End If 				
			If strline = "[Failover.ForceStandby]" Then 
				bFoundStandby = True
			Else
				If bFoundStandby = True and bStandby = "" and Left(strline,1) = "[" Then 
					bStandby = "N/A"
					strStandby = "N/A"
					'wscript.echo "Didn't find value entry under ForceStandby, bstandby: " & bstandby
				End If 
			End If 
			If Left(strline,6) = "value=" and bFoundActive = True Then
				'wscript.echo "Found " & strline
				strParts = split(strline,"=")
				If strParts(1) = "enable" or strParts(1) = "disable" Then 
					If bFoundStandby = false Then
						bActive = "OK" 
						strActive = strParts(1)
					Else
						bStandby = "OK"
						strStandby = strParts(1)
					End If 
				Else
					If bFoundStandby = false Then
						bActive = "Bad" 
						strActive = strParts(1)
					Else
						bStandby = "Bad"
						strStandby = strParts(1)
					End If 
				End If 
			End If
												
			If bStandby = "" Then
				bExit = fileobj.atendofstream
			Else
				bExit = True
			End If 
		Wend
	    
	    'wscript.echo "bActive=" & bactive & " bStandby=" & bstandby
		
		If bActive = "OK" and bStandby = "OK" and (strActive <> strStandby) Then
			strOut = Left(f1.name,InStrRev(f1.name,".")-1) & ", PASS"
		Else
			If bFoundActive = True Then strOut = Left(f1.name,InStrRev(f1.name,".")-1) & ", FAILED, Active=" & strActive & " Standby=" & strStandby
		End If 
		 
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