#$language = "VBScript"
#$interface = "1.0"

'|----------------------------------------------------------------------------------------------------------|
'|  Author: Siggi Bjarnason                                                                                 |
'|  Authored: 11/08/16                                                                                      |
'|  Copyright: Siggi Bjarnason 2016                                                                         |
'|----------------------------------------------------------------------------------------------------------|

Option Explicit
dim strInFile, strOutFile, strPCFOutFile

' User Spefified values, specify values here per your needs
strInFile     = "C:\Users\sbjarna\Documents\IP Projects\Automation\ARGMap\test.csv"
strOutFile    = "C:\Users\sbjarna\Documents\IP Projects\Automation\ARGMap\testARGFWGPCFMap.csv"
strPCFOutFile = "C:\Users\sbjarna\Documents\IP Projects\Automation\ARGMap\testPCFARGMap.csv"

'Nothing below here is user configurable proceed at your own risk.

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8
	Const Timeout = 2
	const VerifyCmd = "show interfaces description | include up          up"

	dim strParts, strLine, objFileIn, objFileOut, host, ConCmd, fso, nError, strErr, strResult, strPCF
	dim strResultParts, strPCFList, strFWGList, objPCF, objFWG, iPrompt, strTemp, bcont, strTemp2, objPCFOut

	crt.screen.synchronous = true
	crt.screen.IgnoreEscape = True

	' Creating a File System Object to interact with the File System
	Set fso = CreateObject("Scripting.FileSystemObject")

	set objFileOut = fso.OpenTextFile(strOutFile, ForWriting, True)
	set objPCFOut = fso.OpenTextFile(strPCFOutFile, ForWriting, True)
	Set objFileIn  = fso.OpenTextFile(strInFile, ForReading, false)
	set objPCF = CreateObject("Scripting.Dictionary")
	set objFWG = CreateObject("Scripting.Dictionary")
	bcont = True

	objFileOut.writeline "ARG,PCF,FWG"
	objPCFOut.writeline "PCF,ARG Pair"

	strLine = objFileIn.readline
	While not objFileIn.atendofstream
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(0)

		strPCFList=""
		strFWGList=""

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
			crt.Screen.Send(VerifyCmd & vbcr)
			do while true
				iPrompt=crt.Screen.WaitForStrings ("Te0", "Hu0", "#", Timeout)
				select case iPrompt
					case 0
						' msgbox "Timeout"
						exit do
					case 3
						' msgbox "Found prompt"
						exit do
					case 1,2
						strTemp=trim(crt.Screen.Readstring (vbcrlf,Timeout))
						if instr(strTemp,"PCF") > 0 then
							strTemp = mid(strTemp,instr(strTemp,"PCF")-2,8)
							if objPCF.exists(strTemp) then
								if instr(objPCF(strTemp),host)=0 then
									strTemp2 = objPCF(strTemp) & "/" & host
									objPCF(strTemp) = strTemp2
									if strPCFList = "" then
										strPCFList = strPCFList & strTemp
									else
										strPCFList = strPCFList & ";" & strTemp
									end if
								end if
							else
								objPCF.add strTemp, host
								if strPCFList = "" then
									strPCFList = strPCFList & strTemp
								else
									strPCFList = strPCFList & ";" & strTemp
								end if
							end if
						end if
						if instr(strTemp,"FWG") > 0 then
							strTemp = mid(strTemp,instr(strTemp,"FWG"),8)
							if not objFWG.exists(strTemp) then
								objFWG.add strTemp, host
								if strFWGList = "" then
									strFWGList = strFWGList & strTemp
								else
									strFWGList = strFWGList & ";" & strTemp
								end if
							end if
						end if
					case else
						msgbox "Unexpected choice #" & iPrompt
						exit do
				end select
			loop
			objFileOut.writeline host & "," & strPCFList  & "," & strFWGList
			crt.Session.Disconnect
		else
			nError = crt.GetLastError
			strErr = crt.GetLastErrorMessage
			objFileOut.writeline host & ",Not Connected,Error " & nError & ": " & strErr
		end if
	wend

	for each strPCF in objPCF
		objPCFOut.writeline strPCF & "," & objPCF(strPCF)
	next
	objFileOut.close
	objFileIn.close
	Set objFileIn  = Nothing
	Set objFileOut = Nothing

	Set fso = Nothing

	msgbox "All Done, Cleanup complete"

end sub
