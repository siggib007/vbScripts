#$language = "VBScript"
#$interface = "1.0"

'|----------------------------------------------------------------------------------------------------------|
'|  Author: Siggi Bjarnason                                                                                 |
'|  Authored: 06/23/16                                                                                      |
'|  Copyright: Siggi Bjarnason 2016                                                                         |
'|----------------------------------------------------------------------------------------------------------|

Option Explicit
dim strInFile, strOutFile, strDelOut

' User Spefified values, specify values here per your needs
strInFile    = "C:\Users\sbjarna\Documents\IP Projects\Automation\GiPrefix\ARGList.csv"
strOutFile   = "C:\Users\sbjarna\Documents\IP Projects\Automation\GiPrefix\Audit.csv"


'Nothing below here is user configurable proceed at your own risk.

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8
	Const Timeout = 2
	const VerifyCmd = "show run prefix-set Gi-Out"

	dim strParts, strLine, objFileIn, objFileOut, objDelOut, host, ConCmd, cmd, fso, nError, strErr, strResult, x,y
	dim strResultParts, strNextHop, objDelDICT, strDelList(), strLineNum, strOut, strResLines

	crt.screen.synchronous = true
	crt.screen.IgnoreEscape = True

	' Creating a File System Object to interact with the File System
	Set fso = CreateObject("Scripting.FileSystemObject")

	set objFileOut = fso.OpenTextFile(strOutFile, ForWriting, True)
	Set objFileIn  = fso.OpenTextFile(strInFile, ForReading, false)

	strLine = objFileIn.readline
	While not objFileIn.atendofstream
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(0)

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
			crt.Screen.Send(VerifyCmd & vbcrlf)
			crt.Screen.WaitForString vbcrlf & "  ",Timeout
			strResult=trim(crt.Screen.Readstring (",",vbcrlf,Timeout))
			crt.Session.Disconnect
			objFileOut.writeline host & "," & strResult
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

