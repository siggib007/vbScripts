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

	dim strParts, strLine, objFileIn, objFileOut, host, ConCmd, fso, nError, strErr, strResult, iPrompt, strResultParts
	dim strOut, strOutPath, IPAddr, objACLDict, strACL, strInterface, iLineCount, strIntDescr, strIPVer, strInt, x

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
			crt.Screen.Send("show running-config | include access-list" & vbcr)
			crt.Screen.WaitForString vbcr,Timeout
			do while true
				strResult=trim(crt.Screen.Readstring (vbcrlf,"RP/0/RS",Timeout))
				if strResult = "" then
					exit do
				end if
				strResultParts = split (strResult," ")
				' objFileOut.writeline  strResult & " has " & ubound(strResultParts) & " parts"
				if ubound(strResultParts) = 2 then
					if strResultParts(1) = "access-list" then
						strIPVer = strResultParts(0)
						strACL = strResultParts(2)
						if not objACLDict.exists(strACL) then
							objACLDict.add strACL, strIPVer
						end if
					end if
				end if
			loop
			for each strACL in objACLDict
				strInterface = ""
				crt.Screen.Send("show access-lists " & objACLDict(strACL) & " " & strACL & " | utility wc lines" & vbcr)
				crt.Screen.WaitForString "GMT" & vbcrlf,Timeout
				iLineCount=trim(crt.Screen.Readstring (vbcrlf,Timeout))
				crt.Screen.WaitForString "#",Timeout
				crt.Screen.Send("show access-lists " & objACLDict(strACL) & " " & strACL & "  usage pfilter location all " & vbcr)
				do while true
					iPrompt=crt.Screen.WaitForStrings ("Interface :", "#", Timeout)
					select case iPrompt
						case 0
							objFileOut.writeline IPAddr & "," & host & "," & strACL & ",Timeout"
						case 2
							exit do
						case 1
							strInterface=strInterface & "|" & trim(crt.Screen.Readstring (vbcrlf,Timeout))
						case else
							msgbox "Unexpected choice #" & iPrompt
							exit do
					end select
				loop
				if strInterface = "" then
					objFileOut.writeline IPAddr & "," & host & "," & strACL & "," & iLineCount & ",unbound"
				else
					strResultParts = split (strInterface,"|")
					for x=0 to ubound(strResultParts)
						strInt = strResultParts(x)
						if strInt <> "" then
							crt.Screen.Send("show interfaces " & strInt & "  description " & vbcr)
							crt.Screen.WaitForStrings "up          up   ","admin-down  admin-down",Timeout
							strIntDescr=trim(crt.Screen.Readstring (vbcrlf,Timeout))
							objFileOut.writeline IPAddr & "," & host & "," & strACL & "," & iLineCount & "," & strInt & "," & strIntDescr
						end if
					next
				end if
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
