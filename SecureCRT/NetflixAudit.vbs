#$language = "VBScript"
#$interface = "1.0"
Option Explicit
public TestMode
TestMode = False

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8
	Const Timeout = 5
	
	dim strInFile, strOutFile
	
	'|----------------------------------------------------------------------------------------------------------|
	'|  Author: Siggi Bjarnason                                                                                 |
	'|  Authored: 06/23/16                                                                                      |
	'|  Copyright: Siggi Bjarnason 2016                                                                         |
	'|----------------------------------------------------------------------------------------------------------|
	
	' User Spefified values, specify values here per your needs
	
	strInFile    = "C:\Users\sbjarna\Documents\IP Projects\Automation\Netflix\AGC_cleanup_and_add.csv"
	strOutFile   = "C:\Users\sbjarna\Documents\IP Projects\Automation\Netflix\Audit.txt"
	
	'Nothing below here is user configurable proceed at your own risk.
	
	dim strParts, strLine, objFileIn, objFileOut, host, ConCmd, cmd, fso, nError, strErr, strResult, x, strResultParts, strNextHop
	
	crt.screen.synchronous = true
	crt.screen.IgnoreEscape = True
	
	' Creating a File System Object to interact with the File System
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	set objFileOut = fso.OpenTextFile(strOutFile, ForWriting, True)
	Set objFileIn = fso.OpenTextFile(strInFile, ForReading, false)
	strLine = objFileIn.readline
	While not objFileIn.atendofstream
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(1)
		
		If crt.Session.Connected Then
			crt.Session.Disconnect
		end if
		
		ConCmd = "/SSH2 "  & host
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
			for x = 2 to ubound(strParts)-1
				cmd = "show access-lists OMW-ABF-IN | include " & strParts(x) & vbcrlf
				crt.Screen.Send(cmd)
				crt.Screen.WaitForString "[K",Timeout
				strResult=trim(crt.Screen.Readstring ("RP/0/",vbcrlf,Timeout))
				strResultParts = split(strResult," ")
				if ubound(strResultParts) > 5 then
					strNextHop = strResultParts(ubound(strResultParts))
				else
					strNextHop = ""
				end if
				objFileOut.writeline host & "," & strParts(x) & "," & strResult  & "," & strNextHop
				crt.Screen.WaitForString "#", Timeout
			next 
			crt.Session.Disconnect
		else
			nError = crt.GetLastError
			strErr = crt.GetLastErrorMessage
			objFileOut.writeline host & ",Not Connected,Error " & nError & ": " & strErr
		end if
	wend
	
	Set objFileIn  = Nothing
	Set objFileOut = Nothing
	Set fso = Nothing
	
	msgbox "All Done, Cleanup complete"
end sub

