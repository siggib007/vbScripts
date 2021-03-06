#$language = "VBScript"
#$interface = "1.0"

'|----------------------------------------------------------------------------------------------------------|
'|  Author: Siggi Bjarnason                                                                                 |
'|  Authored: 06/23/16                                                                                      |
'|  Copyright: Siggi Bjarnason 2016                                                                         |
'|----------------------------------------------------------------------------------------------------------|

Option Explicit
dim strInFile, strOutFile, AuditCmd, dictSubnets

' User Spefified values, specify values here per your needs
strInFile        = "C:\Users\sbjarna\Documents\IP Projects\Automation\ARGACLs\PHISYOFWG3132.csv" ' Input file, comma seperated. First value device name, first line header
strOutFile       = "C:\Users\sbjarna\Documents\IP Projects\Automation\ARGACLs\PHI-SYO-FWG31-32-Audit.csv" ' The name of the output file, CSV file listing results
const Timeout    = 5    ' Timeout in seconds for each command, if expected results aren't received withing this time, the script moves on.

'Nothing below here is user configurable proceed at your own risk.

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8

	dim strParts, strLine, objFileIn, objFileOut, host, ConCmd, fso, nError, strErr, strInterface, strIPAddr
	dim strOut, strOutPath, strLastHost, strDescription, strIPAddrV6, iResponse, strLineParts

	set dictSubnets = CreateObject("Scripting.Dictionary")

	InitializeDicts

	If crt.Session.Connected Then
		crt.Session.Disconnect
	end if

	strOutPath = left (strOutFile, InStrRev (strOutFile,"\"))

	' Creating a File System Object to interact with the File System
	Set fso = CreateObject("Scripting.FileSystemObject")

	strOut = ""
	strLastHost = ""
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
	Set objFileIn  = fso.OpenTextFile(strInFile, ForReading, false)

	objFileOut.writeline "Router,Interface,Description,IPv4,IPv6"

	'Skip over the first header line
	strLine = objFileIn.readline
	'Start Looping through the input file
	While not objFileIn.atendofstream
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(0)

		strInterface = strParts(1)
		AuditCmd = "show running-config interface " & strInterface

		if strLastHost <> host then
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
				crt.Screen.Send("term len 0" & vbcr)
				crt.Screen.WaitForString "#",Timeout
			end if
		end if

		If crt.Session.Connected Then
			crt.Screen.Send(AuditCmd & vbcr)
			iResponse = crt.Screen.WaitForStrings ("description ","%",Timeout)
			if iResponse = 1 then
				strDescription=trim(crt.Screen.Readstring (vbcrlf,Timeout))
				crt.Screen.WaitForStrings "ip address ", "ipv4 address ",Timeout
				strIPAddr=trim(crt.Screen.Readstring (vbcrlf,Timeout))
				strLineParts=split(strIPAddr," ")
				if ubound(strLineParts)=1 then
					if dictSubnets.exists(strLineParts(1)) then
						strIPAddr = strLineParts(0) & dictSubnets.Item(strLineParts(1))
					else
						strIPAddr = strLineParts(0) & "***" & strLineParts(1) & "***"
					end if
				end if
				iResponse = crt.Screen.WaitForStrings ("ipv6 address ","RP/0","#",Timeout)
				if iResponse = 1 then
					strIPAddrV6=trim(crt.Screen.Readstring (vbcrlf,Timeout))
				else
					strIPAddrV6 = ""
				end if
			else
				strDescription = "not found"
				strIPAddr = ""
				strIPAddrV6 = ""
			end if
			objFileOut.writeline host & "," & strInterface & "," & strDescription & "," & strIPAddr & "," & strIPAddrV6
		else
			nError = crt.GetLastError
			strErr = crt.GetLastErrorMessage
			objFileOut.write host & ",Not Connected,Error " & nError & ": " & strErr
		end if
		strLastHost = host
	wend

	objFileOut.close
	objFileIn.close
	Set objFileIn  = Nothing
	Set objFileOut = Nothing

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