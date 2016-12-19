#$language = "VBScript"
#$interface = "1.0"

'|----------------------------------------------------------------------------------------------------------|
'|  Author: Siggi Bjarnason                                                                                 |
'|  Authored: 12/19/16                                                                                      |
'|  Copyright: Siggi Bjarnason 2016                                                                         |
'|----------------------------------------------------------------------------------------------------------|

Option Explicit
dim strInFile, strOutFile

' User Spefified values, specify values here per your needs
strInFile     = "C:\Users\sbjarna\Documents\IP Projects\Automation\GiACL\ARGList.csv"
strOutFile    = "C:\Users\sbjarna\Documents\IP Projects\Automation\GiACL\ARG-ACLs-Brief.csv"


'Nothing below here is user configurable proceed at your own risk.

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8
	Const Timeout = 5

	dim strParts, strLine, objFileIn, objFileOut, host, ConCmd, fso, nError, strErr, strResult, iPrompt
	dim strOut, strOutPath, IPAddr, objACLDict, strACL, strInterface, iLineCount, bBound, strIntDescr
 
	strOutPath = left (strOutFile, InStrRev (strOutFile,"\"))
	Set fso = CreateObject("Scripting.FileSystemObject")
	set objACLDict = CreateObject("Scripting.Dictionary")

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

	' Creating a File System Object to interact with the File System

	set objFileOut = fso.OpenTextFile(strOutFile, ForWriting, True)
	Set objFileIn  = fso.OpenTextFile(strInFile, ForReading, false)

	objFileOut.writeline "primaryIPAddress,hostName,ACL Name,Line Count,Interface, Interface description"
	strLine = objFileIn.readline
	While not objFileIn.atendofstream
		objACLDict.RemoveAll
		strInterface=""
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(0)
		IPAddr = strParts(1)

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
			crt.Screen.Send("show running-config | include ipv4 access-list" & vbcr)
			crt.Screen.WaitForString vbcr,Timeout
			do while true
				iPrompt=crt.Screen.WaitForStrings ("access-list", "#", Timeout)
				select case iPrompt
					case 0
						msgbox "Timeout"
						exit do
					case 2
						exit do
					case 1
						crt.Screen.WaitForString "",Timeout
						strResult=trim(crt.Screen.Readstring (vbcrlf,Timeout))
						if not objACLDict.exists(strResult) then
							objACLDict.add strResult, host
						end if						
					case else
						msgbox "Unexpected choice #" & iPrompt
						exit do
				end select
			loop
			for each strACL in objACLDict
				bBound = false
				crt.Screen.Send("show access-lists " & strACL & " | utility wc lines" & vbcr)
				crt.Screen.WaitForString "GMT" & vbcrlf,Timeout
				iLineCount=trim(crt.Screen.Readstring (vbcrlf,Timeout))
				crt.Screen.WaitForString "#",Timeout
				crt.Screen.Send("show access-lists " & strACL & "  usage pfilter location all " & vbcr)
				do while true
					iPrompt=crt.Screen.WaitForStrings ("Interface :", "#", Timeout)
					select case iPrompt
						case 0
							msgbox "Timeout"
							exit do
						case 2
							if bBound = false Then
								objFileOut.writeline IPAddr & "," & host & "," & strACL & "," & iLineCount & ",unbound" 
							end if
							exit do
						case 1
							strInterface=trim(crt.Screen.Readstring (vbcrlf,Timeout))
							crt.Screen.WaitForString "#",Timeout
							crt.Screen.Send("show interfaces " & strInterface & "  description " & vbcr)
							crt.Screen.WaitForString "up          up   ",Timeout
							strIntDescr=trim(crt.Screen.Readstring (vbcrlf,Timeout))
							objFileOut.writeline IPAddr & "," & host & "," & strACL & "," & iLineCount & "," & strInterface	& "," & strIntDescr
							bBound = True						
						case else
							msgbox "Unexpected choice #" & iPrompt
							exit do
					end select
				loop				
			next 			
			crt.Session.Disconnect
		else
			nError = crt.GetLastError
			strErr = crt.GetLastErrorMessage
			objFileOut.writeline IPAddr & "," & host & ",Not Connected,Error " & nError & ": " & strErr
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
