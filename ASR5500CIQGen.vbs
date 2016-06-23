Option Explicit
'-------------------------------------------------------------------------------------------------'
' This script will take in two spreadsheets as input. The first a CIQ template the second a       '
' varibles file. It will then duplicate the template CIQ and replace the values in it with        '
' values from the variables file.                                                                 '
'                                                                                                 '
' Author: Siggi Bjarnason                                                                         '
' Date: 02/19/2015                                                                                '
' Usage: parser InFile=inpath template=templatepath CIQ=ciqpath log file hide                     '
'        parser: The name of this script                                                          '
'        InFile: The complete path to the spreadsheet with the site specific data                 '
'        template: The complete path to the CIQ template spreadsheet                              '
'        CIQ: The complete path of where you want the generated CIQ's saved                       '
'        log: flag to indicate detailed log, summary is default                                   '
'        file: flag to log everything to a log file, default is no log file                       '
'        hide: Keep Excel hidden during processing, only applies to the instanse the script starts'
'                                                                                                 '
'  All arguments are optional and the order does not matter                                       '
'  The nessisary data you don't provide will be prompted for                                      '
'  if you provide "help" or "?" as an argument help message will be printed and script will exit  '
'                                                                                                 '
'-------------------------------------------------------------------------------------------------'

'User definable Constants

const DefCIQFolderName = "CIQ\"
const DefTemplateCIQ   = "C:\AutoCIQ\Template_CIQ.xlsx"
const DefInFile        = "C:\AutoCIQ\Master Assignments Bogus Test.xlsx"


'Nothing below here is user configurable proceed at your own risk.

'Variable declaration
Dim strCIQFolder, strScriptFullName, strLogFileName, strInFile
Dim strTemplate, iArg, strParts, strScriptName
Dim strArgParts, strDefCIQFolder, strInput, strTmp
Dim dictSubnets, dictCommon, bLog, bFile, bShowExcel, fso, x, y
Dim objLogOut, app, wbin, wsin, wbout, wsout, wsLog

bLog  = false
bFile = false
bShowExcel = True
strScriptFullName = wscript.ScriptFullName
strParts = split (strScriptFullName, "\")
strScriptName = strParts(ubound(strParts))

' Creating a File System Object to interact with the File System
Set fso = CreateObject("Scripting.FileSystemObject")

'Process command line arguments
' If command line copy the arguments into an array.
 If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
	redim strArgParts(Wscript.Arguments.count -1)
	for iArg = 0 to Wscript.Arguments.Count - 1
		strArgParts(iArg) = WScript.Arguments(iArg)
	next
 Else
 	' If not, use InputBox( ) to prompt for arguments the split into an array
 	strInput = trim(InputBox("Please provide any optional arguments, seperate by space, you'd like. For help type help. Abort will exit","Optional Arguments"))
 	strArgParts = split(strInput, " ")
 End If

for iArg = 0 to ubound(strArgParts)

	strParts = split (strArgParts(iArg),"=")
	select case lcase(strParts(0))
		case "infile"
			strInFile = strParts(1)
			wscript.echo "found an argument for Infile: " & strInFile
		case "template"
			strTemplate = strParts(1)
			wscript.echo "found an argument for template: " & strTemplate
		case "ciq"
			strCIQFolder = strParts(1)
			wscript.echo "found an argument for CIQ Folder: " & strCIQFolder
		case "log"
			bLog = true
			wscript.echo "detailed logging enabled"
		case "file"
			bFile = true
			wscript.echo "log file enabled"
		case "hide"
			bShowExcel = false
			wscript.echo "not showing excel"
		case "help","?"
			PrintHelp
			wscript.quit
		case else
			wscript.echo "Invalid argument " & strArgParts(iArg) & ". try help for valid options"
			wscript.quit
	end select
next

'if log file is enabled create the log file
strLogFileName = Mid(strScriptFullName, 1, InStrRev(strScriptFullName, ".")) & "log"

if bFile then
	wscript.echo "Log file " & strLogFileName
	Set objLogOut = fso.createtextfile(strLogFileName, true)
end if
wscript.echo "starting " & strScriptName &" at " & now

'Validate input
if strInFile = "" then
	strInFile = UserInput("No input file was specified. Please provide input file " & vbcrlf & _
								"complete path or leave blank to use the default of " & DefInFile & vbcrlf & "Your Input:")
	if strInFile = "" then strInFile = DefInFile
end if

if strTemplate = "" then
	strTemplate = UserInput("No input file was specified. Please provide input file " & vbcrlf & _
								"complete path or leave blank to use the default of " & DefTemplateCIQ & vbcrlf & "Your Input:")
	if strTemplate = "" then strTemplate = DefTemplateCIQ
end if

'Validating existance of input file, prompt if not valid
while not fso.FileExists(strInFile)
	strInFile = UserInput("input file " & strInFile & " is not valid. Please provide new one or leave blank to abort:")
	if strInFile = "" then
		wscript.echo "No input provided, aborting"
		wscript.quit
	end if
wend

'Validating existance of template file, prompt if not valid
while not fso.FileExists(strTemplate)
	strTemplate = UserInput("template file " & strTemplate & " is not valid. Please provide new one or leave blank to abort:")
	if strTemplate = "" then
		wscript.echo "No input provided, aborting"
		wscript.quit
	end if
wend

'parse out default CIQ folder from template filename
strDefCIQFolder = Mid(strTemplate, 1, InStrRev(strTemplate, "\")) & DefCIQFolderName

'Prompt if CIQ folder was not provided
if strCIQFolder = "" then
	strCIQFolder = UserInput("No CIQ folder provided. Please provide comlete path where you want the CIQ's stored " & vbcrlf & _
								"or leave blank to use the default of " & strDefCIQFolder & vbcrlf & "Your Input:")
	if strCIQFolder = "" then strCIQFolder = strDefCIQFolder
end if

if right(strCIQFolder,1)<> "\" then
	strCIQFolder = strCIQFolder & "\"
end if

'Validating the existance of the CIQ folder, create if not valid
if not fso.FolderExists(strCIQFolder) then
	fso.CreateFolder(strCIQFolder)
	WriteLog strCIQFolder & " did not exists so I created it"
end if

'Open Excel and get hook into it
on error resume next
Set app = CreateObject("Excel.Application")
If Err.Number <> 0 Then
	WriteLog "Unable to start Excel, probably not installed correctly. Unable to read the CIQ's without it"
	wscript.quit
end if
on error goto 0

'Set the visibility of Excel based on input parameters.
app.visible = bShowExcel

on error resume next
Set wbin = app.Workbooks.Open (strInFile,0,true)
If Err.Number <> 0 Then
	WriteLog "Unable to open input file, received this error when attempting: " &  Err.Description
	WriteLog "Aborting for now. resolve the issue with the input file and try again"
	wscript.quit
end if
on error goto 0
Set wsin = wbin.Worksheets(1)

on error resume next
Set wbout = app.Workbooks.Open (strTemplate,0,true)
If Err.Number <> 0 Then
	WriteLog "Unable to open input file, received this error when attempting: " &  Err.Description
	WriteLog "Aborting for now. resolve the issue with the input file and try again"
	wscript.quit
end if
on error goto 0
Set wsout = wbout.Worksheets("IP")
Set wsLog = wbout.Worksheets("Change Log")
y=2
while wsLog.cells(y,1)<>""
	y=y+1
wend	
'Initializing Dictionaries, aka ordered arrays
set dictSubnets = CreateObject("Scripting.Dictionary")
set dictCommon  = CreateObject("Scripting.Dictionary")

'Load initial values into dictionaries as applicable
InitializeDicts

x=2
Do
	strTmp = CleanStr(wsin.Cells(x,4)) ' start by grabing the site private subnet and start dividing it
	if bLog then WriteLog "Parsing " & CleanStr(wsin.Cells(x,1))
	dictCommon.RemoveAll
	dictCommon.add "GWName",      CleanStr(wsin.Cells(x,1))
	dictCommon.add "SiteName",    CleanStr(wsin.Cells(x,2))
	dictCommon.add "SiteCode",    CleanStr(wsin.Cells(x,3))
	dictCommon.add "SiteID",      CleanStr(wsin.Cells(x,6))
	dictCommon.add "PCFAS",       CleanStr(wsin.Cells(x,7))
	dictCommon.add "IPv6Net",     CleanStr(wsin.Cells(x,8))
	dictCommon.add "IPv6Loop",    CleanStr(wsin.Cells(x,9))
	dictCommon.add "TestPoolNet", CleanStr(wsin.Cells(x,10))
	dictCommon.add "ARG21",       "ARG"&CleanStr(wsin.Cells(x,3))&"21"
	dictCommon.add "ARG22",       "ARG"&CleanStr(wsin.Cells(x,3))&"22"
	
	strParts = split(strTmp,".")
	
	dictCommon.add "GaSubnet",    strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 0
	dictCommon.add "GiSubnet",    strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 8
	dictCommon.add "GnSubnet",    strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 16
	dictCommon.add "LiSubnet",    strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 24
	dictCommon.add "S1uSubnet",   strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 32
	dictCommon.add "BPNet",       strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 40
	dictCommon.add "GiLoIP",      strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 49
	dictCommon.add "S6bLoIP",     strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 50
  dictCommon.add "GxLoIP",      strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 51
  dictCommon.add "GyLoIP",      strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 52
  dictCommon.add "GaLoIP",      strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 53
  dictCommon.add "CaleaLoIP",   strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 54
  dictCommon.add "PCSCFLoIP",   strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 55
  dictCommon.add "S11SGwLoIP",  strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 56
  dictCommon.add "S1uSGwLoIP",  strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 57
  dictCommon.add "GiSgiLoIP",   strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 58
  dictCommon.add "GnPGwLoIP",   strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 59

	strTmp = CleanStr(wsin.Cells(x,5)) ' Now grab the public subnet
	strParts = split(strTmp,".")
  dictCommon.add "GnLoIP",      strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 0
  dictCommon.add "S5S8SGwLoIP", strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 1
  dictCommon.add "S8SGwLoIP",   strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 2
  dictCommon.add "S8PGwLoIP",   strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 3
  dictCommon.add "S4SGwLoIP",   strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 4
  dictCommon.add "S12SGwLoIP",  strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 5
	
	if bLog then WriteLog "Updating template and saving "
	wsout.cells(1,3).value = dictCommon.item("SiteName")
	wsout.cells(1,6).value = dictCommon.item("SiteID")
	wsout.cells(4,3).value = dictCommon.item("GWName")
	wsout.cells(4,7).value = dictCommon.item("PCFAS")
	wsout.cells(6,9).value = dictCommon.item("ARG21")
	wsout.cells(10,9).value = dictCommon.item("ARG22")
	wsout.cells(9,3).value = dictCommon.item("GaSubnet")
	wsout.cells(10,3).value = dictCommon.item("GiSubnet")
	wsout.cells(11,3).value = dictCommon.item("GnSubnet")
	wsout.cells(12,3).value = dictCommon.item("LiSubnet")
	wsout.cells(13,3).value = dictCommon.item("S1uSubnet")
	wsout.cells(14,3).value = dictCommon.item("BPNet")
	wsout.cells(15,3).value = dictCommon.item("IPv6Net")
	wsout.cells(17,3).value = dictCommon.item("S6bLoIP")
	wsout.cells(18,3).value = dictCommon.item("GnLoIP")
	wsout.cells(19,3).value = dictCommon.item("GiLoIP")
	wsout.cells(20,3).value = dictCommon.item("GxLoIP")
	wsout.cells(21,3).value = dictCommon.item("GyLoIP")
	wsout.cells(22,3).value = dictCommon.item("GaLoIP")
	wsout.cells(23,3).value = dictCommon.item("CaleaLoIP")
	wsout.cells(24,3).value = dictCommon.item("PCSCFLoIP")
	wsout.cells(25,3).value = dictCommon.item("S11SGwLoIP")
	wsout.cells(26,3).value = dictCommon.item("S1uSGwLoIP")
	wsout.cells(27,3).value = dictCommon.item("GiSgiLoIP")
	wsout.cells(28,3).value = dictCommon.item("GnPGwLoIP")
	wsout.cells(29,3).value = dictCommon.item("S5S8SGwLoIP")
	wsout.cells(30,3).value = dictCommon.item("S8SGwLoIP")
	wsout.cells(31,3).value = dictCommon.item("S8PGwLoIP")
	wsout.cells(32,3).value = dictCommon.item("S4SGwLoIP")
	wsout.cells(33,3).value = dictCommon.item("S12SGwLoIP")
	wsout.cells(34,3).value = dictCommon.item("IPv6Loop")
	wsout.cells(35,3).value = dictCommon.item("TestPoolNet")
	
	if bLog then WriteLog "Updating change log"

	wsLog.cells(y,1).value = now
	wsLog.cells(y,3).value = "New CIQ for " & dictCommon.item("GWName") & " auto generated by " & strScriptName

	if bLog then WriteLog "Saving "
	wbout.SaveAs strCIQFolder & dictCommon.item("GWName") & "_CIQ.xlsx",,,,,,,2
	if bLog then WriteLog "Done Next one ... "
	x=x+1

Loop Until CleanStr(wsin.Cells(x,1))=""

wbin.close False
wbout.close False
app.quit ' Close Excel

WriteLog "Done. CIQ's can be found in " & strCIQFolder

'Cleanup, close out files, release resources, etc.
if bFile then
	objLogOut.close
	set objLogOut = nothing
end if


Set fso         = nothing
Set wsin        = Nothing
Set wbin        = Nothing
Set wsout       = Nothing
Set wbout       = Nothing
Set app         = Nothing
set dictSubnets = Nothing
set dictCommon  = Nothing

wscript.echo now & vbtab & "All Done, Cleanup complete"

function CleanStr (strMsg)
'-------------------------------------------------------------------------------------------------'
' Function CleanStr (strMsg)                                                                      '
'                                                                                                 '
' This function accepts one input parameter, a string, trims any leading or trailing spaces       '
' as well as any Carrige return or line feed characters from that string then returns it          '
'-------------------------------------------------------------------------------------------------'

	strMsg = replace(strMsg,vbcr,"")
	strMsg = replace(strMsg,vblf,"")
	strMsg = trim(strMsg)
	CleanStr = strMsg
end function

Function WriteLog (strMsg)
'-------------------------------------------------------------------------------------------------'
' Function WriteLog (strMsg)                                                                      '
'                                                                                                 '
' This function accepts one input parameter, a string, and writes it to the screen or a file      '
' based on command line arguments                                                                 '
'-------------------------------------------------------------------------------------------------'

' Check if the script runs in CSCRIPT.EXE, i.e. is being run from command line, if so write to screen, otherwise do nothing
' Need to avoid having the script initiate 100 popups
	If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
		wscript.echo now & vbtab & strMsg
	end if

	'If log file is enabled write to it, otherwise don't
	if bFile then objLogOut.writeline now & vbtab & strMsg

end function

Function UserInput( myPrompt )
' This function prompts the user for some input.
' When the script runs in CSCRIPT.EXE, StdIn is used,
' otherwise the VBScript InputBox( ) function is used.
' myPrompt is the the text used to prompt the user for input.
' The function returns the input typed either on StdIn or in InputBox( ).
' Written by Rob van der Woude
' http://www.robvanderwoude.com
 ' Check if the script runs in CSCRIPT.EXE
 If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
 	' If so, use StdIn and StdOut
	WScript.StdOut.Write myPrompt & " "
 	UserInput = WScript.StdIn.ReadLine
 Else
 	' If not, use InputBox( )
 	UserInput = InputBox( myPrompt )
 End If
End Function

Function AskYesNo (myPrompt)
'-------------------------------------------------------------------------------------------------'
' Function AskYesNo (myPrompt)                                                                    '
'                                                                                                 '
' This function accepts one input parameter, a string, and uses it to prompt the user with a      '
' Yes/No question                                                                                 '
'-------------------------------------------------------------------------------------------------'

Dim strInput
 If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
 	WScript.StdOut.Write myPrompt & " "
 	strInput = WScript.StdIn.ReadLine
 	if left(ucase(strInput),1)="Y" then
 		AskYesNo = "Yes"
 	else
 		AskYesNo = "No"
 	end if
 else
 	strInput = Msgbox(myPrompt, vbYesNo, "Question for you")
	If strInput = vbYes Then
 		AskYesNo = "Yes"
 	else
 		AskYesNo = "No"
 	end if
 end If
End Function

sub InitializeDicts
'-------------------------------------------------------------------------------------------------'
' Function InitializeDicts                                                                        '
'                                                                                                 '
' This sub takes no inpput and just loads dictionaries with predefined values                     '
'-------------------------------------------------------------------------------------------------'

 dictSubnets.add "32", "0.0.0.0"
 dictSubnets.add "31", "0.0.0.1"
 dictSubnets.add "30", "0.0.0.3"
 dictSubnets.add "29", "0.0.0.7"
 dictSubnets.add "28", "0.0.0.15"
 dictSubnets.add "27", "0.0.0.31"
 dictSubnets.add "26", "0.0.0.63"
 dictSubnets.add "25", "0.0.0.127"
 dictSubnets.add "24", "0.0.0.255"
 dictSubnets.add "23", "0.0.1.255"
 dictSubnets.add "22", "0.0.3.255"
 dictSubnets.add "21", "0.0.7.255"
 dictSubnets.add "20", "0.0.15.255"
 dictSubnets.add "19", "0.0.31.255"
 dictSubnets.add "18", "0.0.63.255"
 dictSubnets.add "17", "0.0.127.255"
 dictSubnets.add "16", "0.0.255.255"
 dictSubnets.add "15", "0.1.255.255"
 dictSubnets.add "14", "0.3.255.255"
 dictSubnets.add "13", "0.7.255.255"
 dictSubnets.add "12", "0.15.255.255"
 dictSubnets.add "11", "0.31.255.255"
 dictSubnets.add "10", "0.63.255.255"
 dictSubnets.add "9", "0.127.255.255"
 dictSubnets.add "8", "0.255.255.255"
 dictSubnets.add "7", "1.255.255.255"
 dictSubnets.add "6", "3.255.255.255"
 dictSubnets.add "5", "7.255.255.255"
 dictSubnets.add "4", "15.255.255.255"
 dictSubnets.add "3", "31.255.255.255"
 dictSubnets.add "2", "63.255.255.255"
 dictSubnets.add "1", "127.255.255.255"
end sub

sub PrintHelp
'-------------------------------------------------------------------------------------------------'
' sub PrintHelp                                                                        '
'                                                                                                 '
' This sub takes no inpput and just prints out a help message                                     '
'-------------------------------------------------------------------------------------------------'

dim strHelpMsg

strHelpMsg =              " Usage: " & strScriptName & " InFile=inpath template=templatepath CIQ=ciqpath log file hide" & vbcrlf
strHelpMsg = strHelpMsg & "        " & strScriptName & ": The name of this script" & vbcrlf
strHelpMsg = strHelpMsg & "        InFile: The complete path to the spreadsheet with the site specific data" & vbcrlf
strHelpMsg = strHelpMsg & "        template: The complete path to the CIQ template spreadsheet" & vbcrlf
strHelpMsg = strHelpMsg & "        CIQ: The complete path of where you want the generated CIQ's saved" & vbcrlf
strHelpMsg = strHelpMsg & "        log: flag to indicate detailed log, summary is default" & vbcrlf
strHelpMsg = strHelpMsg & "        file: flag to log everything to a log file, default is no log file" & vbcrlf
strHelpMsg = strHelpMsg & "        hide: Keep Excel hidden during processing, only applies to the instanse the script starts. Default is to show Excel." & vbcrlf
strHelpMsg = strHelpMsg & vbcrlf
strHelpMsg = strHelpMsg & "  All arguments are optional and the order does not matter." & vbcrlf
strHelpMsg = strHelpMsg & "  The nessisary data you don't provide will be prompted for" & vbcrlf
strHelpMsg = strHelpMsg & "  if you provide ""help"" or ""?"" as an argument this message will be printed and script will exit" & vbcrlf
strHelpMsg = strHelpMsg & vbcrlf
strHelpMsg = strHelpMsg & "  Example: cscript " & strScriptName & " InFile=C:\CIQGen\2015GGSNSubnets.xlsx template=C:\CIQ\ASR5500CIQ.xlsx log file" & vbcrlf

 ' Check if the script runs in CSCRIPT.EXE
 If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
 ' If so, use StdIn and StdOut
 WScript.StdOut.Writeline strHelpMsg
Else
 Msgbox strHelpMsg', "Help message"
end if


end sub
