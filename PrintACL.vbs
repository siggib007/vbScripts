Option Explicit
Dim FileObj, strLine, fso, f, fc, f1, strOut, strParts, FolderSpec, strOutFileName, objFileOut, strcriteria, bPrint, strHost, strDescription, strIP, iPos1, iPos2, dictSubnets,strIPParts

' If WScript.Arguments.Count <> 3 Then
' 	wscript.echo "Lists all lines in any files in the specified directory that fall between specified lines"
'   WScript.Echo "Usage: parser criteria inpath outfilename"
'   WScript.Quit
' End If

' FolderSpec = WScript.Arguments(1)
' strOutFileName = WScript.Arguments(2)
' strCriteria = wscript.arguments(0)

FolderSpec = "C:\Users\sbjarna\Documents\IP Projects\Automation\GiACL\OMW-ABF-IN"
strOutFileName = "C:\Users\sbjarna\Documents\IP Projects\Automation\GiACL\OMW-ABF-IN-CDN-DMZ.csv"
bPrint = false
set dictSubnets = CreateObject("Scripting.Dictionary")

InitializeDicts

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder(folderspec)
Set objFileOut = fso.createtextfile(strOutFileName)
Set fc = f.Files
objFileOut.write "Device,Section,Subnets"
For Each f1 in fc
	If f1.name <> strOutFileName Then
		Set FileObj = fso.opentextfile(folderspec & "\" & f1.name)
		strHost = Left(f1.name,InStrRev(f1.name,".")-1)
		While not fileobj.atendofstream
			strLine = Trim(FileObj.readline)
			If InStr(strline, "remark") > 0 Then
				If InStr(strline, "DMZ") > 0 or InStr(strline, "Cache") > 0 Then
					iPos1 = InStr(strLine,"ENGPCI")
					if iPos1 = 0 then
						iPos1 = InStr(strLine,"connecting to ")+14
					end if
					iPos2 = InStr(strLine,"Global")
					if iPos2 = 0 then
						iPos2 = len(strLine) - iPos1
					else
						iPos2 = iPos2 - iPos1
					end if
					strDescription = trim(Mid(strLine,iPos1,iPos2))
					bPrint = true
					objFileOut.write vbcrlf & strHost & "," & strDescription
				else
					bPrint = false
				end if
			else
				if bPrint = true then
					iPos1 = InStr(strLine,"ipv4 any ") + 9
					iPos2 = InStr(strLine,"nexthop1")
					if iPos2 = 0 then
						iPos2 = len(strLine) - iPos1 + 1
					else
						iPos2 = iPos2 - iPos1
					end if
					' objFileOut.writeline "pos1:" & iPos1 & "   pos2:" & iPos2
					strIP = trim(mid(strLine,iPos1,iPos2))
					strIPParts = split (strIP," ")
					if ubound(strIPParts)>0 then
						if dictSubnets.exists(strIPParts(1)) then
							strIP = strIPParts(0) & "/" & dictSubnets.Item(strIPParts(1))
						else
							strIP = strIPParts(0) & "|" & strIPParts(1)
						end if
					end if
					objFileOut.write "," & strIP
				end if
			End If
		Wend
		FileObj.close
	End If
Next
wscript.echo "Done"

objFileOut.close
Set FileObj = nothing
Set objFileOut = nothing
Set fc = nothing
Set f = nothing
Set fso = nothing


sub InitializeDicts
'-------------------------------------------------------------------------------------------------'
' Function InitializeDicts                                                                        '
'                                                                                                 '
' This sub takes no inpput and just loads dictionaries with predefined values                     '
'-------------------------------------------------------------------------------------------------'

dictSubnets.add "0.0.0.0",32
dictSubnets.add "0.0.0.1",31
dictSubnets.add "0.0.0.3",30
dictSubnets.add "0.0.0.7",29
dictSubnets.add "0.0.0.15",28
dictSubnets.add "0.0.0.31",27
dictSubnets.add "0.0.0.63",26
dictSubnets.add "0.0.0.127",25
dictSubnets.add "0.0.0.255",24
dictSubnets.add "0.0.1.255",23
dictSubnets.add "0.0.3.255",22
dictSubnets.add "0.0.7.255",21
dictSubnets.add "0.0.15.255",20
dictSubnets.add "0.0.31.255",19
dictSubnets.add "0.0.63.255",18
dictSubnets.add "0.0.127.255",17
dictSubnets.add "0.0.255.255",16
dictSubnets.add "0.1.255.255",15
dictSubnets.add "0.3.255.255",14
dictSubnets.add "0.7.255.255",13
dictSubnets.add "0.15.255.255",12
dictSubnets.add "0.31.255.255",11
dictSubnets.add "0.63.255.255",10
dictSubnets.add "0.127.255.255",9
dictSubnets.add "0.255.255.255",8
dictSubnets.add "1.255.255.255",7
dictSubnets.add "3.255.255.255",6
dictSubnets.add "7.255.255.255",5
dictSubnets.add "15.255.255.255",4
dictSubnets.add "31.255.255.255",3
dictSubnets.add "63.255.255.255",2
dictSubnets.add "127.255.255.255",1
dictSubnets.add "255.255.255.255",0

end sub