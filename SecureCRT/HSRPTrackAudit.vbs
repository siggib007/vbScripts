#$language = "VBScript"
#$interface = "1.0"

'|----------------------------------------------------------------------------------------------------------|
'|  Author: Siggi Bjarnason                                                                                 |
'|  Authored: 06/23/16                                                                                      |
'|  Copyright: Siggi Bjarnason 2016                                                                         |
'|----------------------------------------------------------------------------------------------------------|

Option Explicit
dim strInFile, strOutFile, dictSubnets, objVlanDict, strOutVlanFile, strDebugOutFile

' User Spefified values, specify values here per your needs
strInFile        = "C:\Users\sbjarna\Documents\IP Projects\Automation\TrackingAudit\AllNX7K.csv" ' Input file, comma seperated. First value device name, first line header
strOutFile       = "C:\Users\sbjarna\Documents\IP Projects\Automation\TrackingAudit\AllNX7K-Results.csv" ' The name of the output file, CSV file listing results
strOutVlanFile   = "C:\Users\sbjarna\Documents\IP Projects\Automation\TrackingAudit\AllNX7K-Vlan-Results.csv" ' The name of the output file, CSV file listing results
strDebugOutFile  = "C:\Users\sbjarna\Documents\IP Projects\Automation\TrackingAudit\Debug.txt" ' File to write debug info to.
const Timeout    = 5    ' Timeout in seconds for each command, if expected results aren't received withing this time, the script moves on.


'Nothing below here is user configurable proceed at your own risk.
set dictSubnets = CreateObject("Scripting.Dictionary")
set objVlanDict = CreateObject("Scripting.Dictionary")

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8

	dim strParts, strLine, objFileIn, objFileOut, objFileVlan, host, ConCmd, fso, nError, strErr, strIPAddr, strAvailTrack
	dim strOut, strOutPath, iResponse, strLineParts, strCommand, strComment, strTemp, bCont, strVlan, strTrack, strVlanComment
	dim bTrack, bLo101, iVlanCount, strTempParts, objDebugOut

	InitializeDicts

	strOutPath = left (strOutFile, InStrRev (strOutFile,"\"))

	' Creating a File System Object to interact with the File System
	Set fso = CreateObject("Scripting.FileSystemObject")

	strOut = ""
	if not fso.FileExists(strInFile) Then
		msgbox "Input file " & strInFile & " not found, exiting"
		exit sub
	end if

	if not fso.FolderExists(strOutPath) then
		CreatePath (strOutPath)
		strOut = strOut & vbcrlf & """" & strOutPath & """ did not exists so I created it" & vbcrlf
	end if
	if strOut <> "" then
		msgbox strOut
	end if

	crt.screen.synchronous = true
	crt.screen.IgnoreEscape = True

	'Opening both intput and output files
	set objFileOut = fso.OpenTextFile(strOutFile, ForWriting, True)
	set objDebugOut = fso.OpenTextFile(strDebugOutFile, ForWriting, True)
	set objFileVlan = fso.OpenTextFile(strOutVlanFile, ForWriting, True)
	Set objFileIn  = fso.OpenTextFile(strInFile, ForReading, false)

	objFileOut.writeline "primaryIPAddress,hostName,AvailTrack,Lo101,SVIcount,comment"
	objFileVlan.writeline "primaryIPAddress,hostName,Vlan,Track,comment"

	'Skip over the first header line
	strLine = objFileIn.readline
	'Start Looping through the input file
	While not objFileIn.atendofstream
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = trim(replace(strParts(0),"""",""))
		strIPAddr = trim(replace(strParts(1),"""",""))

		strComment = ""
		strAvailTrack = ""
		strVlan = ""
		objVlanDict.RemoveAll

		If crt.Session.Connected Then
			crt.Session.Disconnect
		end if

		ConCmd = "/SSH2 /ACCEPTHOSTKEYS "  & host
		on error resume next
		crt.Session.Connect ConCmd
		on error goto 0
		If crt.Session.Connected Then
			crt.Screen.Synchronous = True
			crt.Screen.WaitForString "#",Timeout
			nError = Err.Number
			strErr = Err.Description
			If nError <> 0 Then
				result = "Error " & nError & ": " & strErr
			end if
			objDebugOut.writeline "Connected to " & host & " at " & now()
			crt.Screen.Send("term len 0" & vbcr)
			crt.Screen.WaitForString "#",Timeout
			strCommand = "show interface Lo101"
			crt.Screen.Send(strCommand & vbcr)
			iResponse = crt.Screen.WaitForStrings ("loopback101 ","Invalid","#",Timeout)
			select case iResponse
				case 0
					strComment = strComment & "Timeout on Lo101;"
					bLo101 = false
				case 1
					' Found loopback101
					bLo101 = True
				case 2
					' strComment = strComment & "No Lo101;"
					bLo101 = false
				case 3
					strComment = strComment & "Couldn't find Lo101;"
					bLo101 = false
				case else
					msgbox "at Lo101, Unexpected choice #" & iResponse
			end select
			' crt.Screen.Send(vbcr)
			crt.Screen.WaitForString "#",Timeout
			strCommand = "show track brief"
			crt.Screen.Send(strCommand & vbcr)
			iResponse = crt.Screen.WaitForStrings ("Last Change","#",Timeout)
			select case iResponse
				case 0
					strComment = strComment & "Timeout on show track;"
					' msgbox "Timeout on show track"
					bcont=false
					bTrack=false
				case 1
					bCont=True
					bTrack=True
					' msgbox "found header line"
					strAvailTrack = ""
				case 2
					' strComment = strComment & "No track;"
					strAvailTrack = "none"
					' msgbox "No tracking"
					bcont=false
					bTrack=false
				case else
					msgbox "at show track, Unexpected choice #" & iResponse
			end select
			bTrack = false
			do while bcont
				iResponse=crt.Screen.WaitForStrings (vbcrlf, "#", Timeout)
				select case iResponse
					case 0
						strComment = strComment & "Timeout on show track loop;"
						exit do
					case 1
						strTemp=trim(crt.Screen.Readstring (" ","#",Timeout))
						if crt.Screen.MatchIndex = 1 then
							strAvailTrack = strAvailTrack & strTemp & " "
							bTrack = True
						else
							if crt.Screen.MatchIndex=0 then
								strComment = strComment & "Timeout on reading show track loop;"
							end if
							exit do
						end if
					case 2
						' Found prompt, done
						exit do
					case else
						msgbox "at show track loop, Unexpected choice #" & iResponse
						exit do
				end select
			loop
			' crt.Screen.WaitForString "#",Timeout
			strCommand = "show hsrp brief"
			crt.Screen.Send(strCommand & vbcr)
			bCont=True
			do while bcont
				iResponse=crt.Screen.WaitForStrings ("lan", "# ", Timeout)
				objDebugOut.writeline "WaitForStrings results:" & iResponse
				select case iResponse
					case 0
						strComment = strComment & "Timeout on show HSRP Loop;"
						objDebugOut.writeline strComment
						exit do
					case 1
						strTemp=trim(crt.Screen.Readstring (" ","#",Timeout))
						objDebugOut.writeline "read: '" & strTemp & "'"
						objDebugOut.writeline "MatchIndex:" & crt.Screen.MatchIndex
						if crt.Screen.MatchIndex = 1 then
							strTemp = "Vlan" & strTemp
							objDebugOut.writeline "Parsed line: '" & strTemp & "'"
							if not objVlanDict.exists(strTemp) then
								objVlanDict.add strTemp, ""
							end if
						else
							if crt.Screen.MatchIndex=0 then strComment = strComment & "Timeout on reading show HSRP loop;"
							if crt.Screen.MatchIndex=2 then strComment = strComment & "Found prompt reading show HSRP loop;"
							objDebugOut.writeline strComment
							exit do
						end if
					case 2
						' Found prompt, done
						exit do
					case else
						msgbox "at show hsrp loop, Unexpected choice #" & iResponse
						exit do
				end select
			loop
			iVlanCount = objVlanDict.count
			if right(strComment,1)=";" then
				strComment = left(strComment,len(strComment)-1)
			end if
			strAvailTrack = trim(strAvailTrack)
			for each strVlan in objVlanDict
				strTrack = ""
				strCommand = "show running-config interface " & strVlan
				crt.Screen.Send(strCommand & vbcr)
				do while true
					iResponse = crt.Screen.WaitForStrings (" track ","#",Timeout)
					select case iResponse
						case 0
							strVlanComment = strVlanComment & "Timeout on show Vlan details;"
							exit do
						case 1
							strTemp=trim(crt.Screen.Readstring (" decrement","#", vbcrlf, Timeout))
							select case crt.Screen.MatchIndex
								case 0
									strVlanComment = strVlanComment & "Timeout on reading show Vlan details;"
									exit do
								case 1,3
									strTrack = strTrack & strTemp & " "
								case 2
									exit do
									'found prompt so we're done
								case else
									msgbox "at show svi, Unexpected choice #" & iResponse
									exit do
							end select
						case 2
							' Found prompt, done.
							exit do
						case else
							msgbox "at reading svi, Unexpected choice #" & iResponse
					end select
				loop
				strTrack = trim(strTrack)
				objFileVlan.writeline strIPAddr & "," & host & "," & strVlan & "," & strTrack & ","	& strComment
			next
			objFileOut.writeline strIPAddr & "," & host & "," & strAvailTrack & "," & bLo101 & ","	& iVlanCount & "," & strComment
		else
			nError = crt.GetLastError
			strErr = crt.GetLastErrorMessage
			objFileOut.write strIPAddr & "," & host & ",Not Connected,Error " & nError & ": " & strErr
		end if
	wend

	objFileOut.close
	objFileIn.close
	objFileVlan.close
	objDebugOut.close
	Set objFileIn   = Nothing
	Set objFileOut  = Nothing
	set objFileVlan = Nothing
	set objDebugOut = Nothing

	Set fso = Nothing
	If crt.Session.Connected Then
		crt.Session.Disconnect
	end if

	msgbox "All Done, Cleanup complete"

end sub

Function CreatePath (strFullPath)
'-------------------------------------------------------------------------------------------------'
' Function CreatePath (strFullPath)                                                               '
'                                                                                                 '
' This function takes a complete path as input and builds that path out as nessisary.             '
'-------------------------------------------------------------------------------------------------'
dim pathparts, buildpath, part, fso

Set fso = CreateObject("Scripting.FileSystemObject")

	pathparts = split(strFullPath,"\")
	buildpath = ""
	for each part in pathparts
		if buildpath<>"" then
			if buildpath = "\" then
				buildpath = buildpath & part
			else
				buildpath = buildpath & "\" & part
			end if
			if not fso.FolderExists(buildpath) then
				fso.CreateFolder(buildpath)
			end if
		else
			if part="" then
				buildpath = "\"
			else
				buildpath = part
			end if
		end if
	next
end function

sub InitializeDicts
	dictSubnets.add "255.255.255.255", "/32"
	dictSubnets.add "255.255.255.254", "/31"
	dictSubnets.add "255.255.255.252", "/30"
	dictSubnets.add "255.255.255.248", "/29"
	dictSubnets.add "255.255.255.240", "/28"
	dictSubnets.add "255.255.255.224", "/27"
	dictSubnets.add "255.255.255.192", "/26"
	dictSubnets.add "255.255.255.128", "/25"
	dictSubnets.add "255.255.255.0", "/24"
	dictSubnets.add "255.255.254.0", "/23"
	dictSubnets.add "255.255.252.0", "/22"
	dictSubnets.add "255.255.248.0", "/21"
	dictSubnets.add "255.255.240.0", "/20"
	dictSubnets.add "255.255.224.0", "/19"
	dictSubnets.add "255.255.192.0", "/18"
	dictSubnets.add "255.255.128.0", "/17"
	dictSubnets.add "255.255.0.0", "/16"
	dictSubnets.add "255.254.0.0", "/15"
	dictSubnets.add "255.252.0.0", "/14"
	dictSubnets.add "255.248.0.0", "/13"
	dictSubnets.add "255.240.0.0", "/12"
	dictSubnets.add "255.224.0.0", "/11"
	dictSubnets.add "255.192.0.0", "/10"
	dictSubnets.add "255.128.0.0", "/9"
	dictSubnets.add "255.0.0.0", "/8"
	dictSubnets.add "254.0.0.0", "/7"
	dictSubnets.add "252.0.0.0", "/6"
	dictSubnets.add "248.0.0.0", "/5"
	dictSubnets.add "240.0.0.0", "/4"
	dictSubnets.add "224.0.0.0", "/3"
	dictSubnets.add "192.0.0.0", "/2"
	dictSubnets.add "128.0.0.0", "/1"
end sub