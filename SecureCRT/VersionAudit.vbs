#$language = "VBScript"
#$interface = "1.0"

'|----------------------------------------------------------------------------------------------------------|
'|  Author: Siggi Bjarnason                                                                                 |
'|  Authored: 09/07/16                                                                                      |
'|  Copyright: Siggi Bjarnason 2016                                                                         |
'|----------------------------------------------------------------------------------------------------------|

Option Explicit
dim strInFile, strOutFile

' User Spefified values, specify values here per your needs
strInFile        = "C:\Users\sbjarna\Documents\IP Projects\Automation\Version\MissingDevices.csv" ' Input file, comma seperated. First value device name, first line header
strOutFile       = "C:\Users\sbjarna\Documents\IP Projects\Automation\Version\MissingDevicesOut.csv" ' The name of the output file, CSV file listing results
const Timeout    = 5    ' Timeout in seconds for each command, if expected results aren't received withing this time, the script moves on.

'Nothing below here is user configurable proceed at your own risk.

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8

	dim strParts, strLine, objFileIn, objFileOut, host, ConCmd, fso, nError, strErr, strOSType
	dim strOut, strOutPath, strVersion, strHardware

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
	Set objFileIn  = fso.OpenTextFile(strInFile, ForReading, false)

	'Write a header for output file
	objFileOut.writeline "hostName,OSType,Harware,OSVersion"

	'Skip over the first header line
	strLine = objFileIn.readline
	'Start Looping through the input file
	While not objFileIn.atendofstream
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(0)
		strOSType = ""
		strVersion = ""
		strHardware = ""

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
			crt.Screen.Send("term len 0" & vbcr)
			crt.Screen.WaitForString "#",Timeout
			crt.Screen.Send("show ver" & vbcr)
			crt.Screen.WaitForString vbcrlf,Timeout
			strOSType=replace(left(trim(crt.screen.Readstring("(",Timeout)),36),",","")
			if instr(strOSType,"Cisco IOS ") > 0 or instr(strOSType,"Cisco Internetwork Operating System ") > 0 Then
				crt.Screen.WaitForString " Version ",Timeout
				strVersion = replace(crt.screen.Readstring(" ",Timeout),",","")
				if instr(strOSType,"IOS-XE") > 0 Then
					crt.Screen.WaitForString "Next reboot license Level: ",Timeout
					crt.Screen.WaitForString vbcrlf&vbcrlf,Timeout
				else	
					crt.Screen.WaitForString "export@cisco.com."&vbcrlf&vbcrlf,Timeout
				end if
				strHardware = trim(crt.screen.Readstring("(",Timeout))
			end if

			if instr(strOSType,"Cisco Nexus") > 0 Then
				crt.Screen.WaitForString " system:    version ",Timeout
				strVersion = crt.screen.Readstring(vbcrlf,Timeout)
				crt.Screen.WaitForString "Hardware"&vbcrlf,Timeout
				strHardware = trim(crt.screen.Readstring("(",Timeout))
			end if
			objFileOut.writeline host & "," & strOSType & "," & strHardware & "," & strVersion
		else
			nError = crt.GetLastError
			strErr = crt.GetLastErrorMessage
			objFileOut.writeline host & ",Not Connected,Error " & nError & ": " & strErr
		end if
	wend

	objFileOut.close
	objFileIn.close
	Set objFileIn  = Nothing
	Set objFileOut = Nothing

	Set fso = Nothing

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
