Option Explicit
'-----------------------------------------------------------------------------------------------------------'
' This script will take in two spreadsheets as input. The first a CIQ template the second a                 '
' varibles file. It will then duplicate the template CIQ and replace the values in it with                  '
' values from the variables file.                                                                           '
'                                                                                                           '
' Author: Siggi Bjarnason                                                                                   '
' Date: 02/19/2015                                                                                          '
' Usage: parser InFile=inpath template=templatepath CIQ=ciqpath log file hide                               '
'        parser:    The name of this script                                                                 '
'        InFile:    The complete path to the spreadsheet with the site specific data                        '
'        template:  The complete path to the CIQ template spreadsheet                                       '
'        CIQ:       The complete path of where you want the generated CIQ's saved                           '
'        log:       Flag to indicate detailed log, summary is default                                       '
'        file:      Flag to log everything to a log file, default is no log file                            '
'        hide:      Keep Excel hidden during processing, only applies to the instanse the script starts     '
'        csv:       Generate CSV files for Cable testing as well as IPControl upload of Subnet and devices  '
'        overwrite: signals to automatically overwrite existing files                                       '
'                                                                                                           '
'  All arguments are optional and the order does not matter                                                 '
'  The nessisary data you don't provide will be prompted for                                                '
'  if you provide "help" or "?" as an argument help message will be printed and script will exit            '
'                                                                                                           '
'-----------------------------------------------------------------------------------------------------------'

'User definable Constants

const DefCIQFolderName = "CIQ\"
const DefTemplateCIQ   = "C:\LTE\2015Q1\Template_CIQ.xlsx"
const DefInFile        = "C:\LTE\2015Q1\MasterAssignments.xlsx"
Const CIQSuffix        = "_MME_CIQ.xlsx"
const SubnetIPCSuffix  = "_ChildBlocks.csv"
const HostIPCSuffix    = "_Device.csv"
const CableTestSuffix  = "_CableTest.csv"
const HPNASuffix       = "_HPNA.csv"
const DefCSVFolder     = "UploadCSVs\"


'Nothing below here is user configurable proceed at your own risk.

'Variable declaration
Dim strCIQFolder, strScriptFullName, strLogFileName, strInFile, strCSVFolder
Dim strTemplate, iArg, strParts, strScriptName, strSafeDate, strFullFileName
Dim strArgParts, strDefCIQFolder, strInput, strTmp, strSafeTime, strHelpMsg, strParams
Dim dictSubnets, dictCommon, bLog, bFile, bShowExcel, fso, x, y, bCSV, bOverWrite, bChanged
Dim objLogOut, app, wbin, wsin, wbout, wsout, wsLog, wsTNE, wbtmp, wstmp, wsCSV, objShell, objExecObj

bLog       = false
bFile      = false
bShowExcel = True
bCSV       = False
bOverWrite = False
bChanged   = False

strScriptFullName = wscript.ScriptFullName
strParts          = split (strScriptFullName, "\")
strScriptName     = strParts(ubound(strParts))

strSafeDate = Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2) & DatePart("yyyy",Date)
strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)

Const xlPasteValues = -4163
Const xlCSV = 6

' Creating a File System Object to interact with the File System
Set fso = CreateObject("Scripting.FileSystemObject")
set objShell = createobject("wscript.shell")

'Process command line arguments
' If command line copy the arguments into an array.
 If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
	redim strArgParts(Wscript.Arguments.count -1)
	for iArg = 0 to Wscript.Arguments.Count - 1
		strArgParts(iArg) = WScript.Arguments(iArg)
	next
 Else
 	' if not advice user of preferred way of runnging this
		strHelpMsg =              " THIS IS NOT THE RECOMENDED WAY TO RUN THIS SCRIPT!!!!!" & vbcrlf
		strHelpMsg = strHelpMsg & " it will work but it's a sub optimal way of working, you'll get no progress, less help and less options"
		strHelpMsg = strHelpMsg & " It is recomended that you open up a command window and run it according to the example below" & vbcrlf & vbcrlf
		strHelpMsg = strHelpMsg & "  Example: cscript " & strScriptFullName & " help" & vbcrlf
		strHelpMsg = strHelpMsg & "  Example: cscript " & strScriptFullName & " infile=C:\HPNATest\testfile.txt log file" & vbcrlf & vbcrlf
		strHelpMsg = strHelpMsg & "  Type abort at the next popup to exit, help for options"
		msgbox strHelpMsg
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
		case "csv"
			bCSV = true
			wscript.echo "generating all CSV"
		case "overwrite"
			bOverWrite = true
			wscript.echo "overwriting existing files as nessisary"
		case "abort","exit","quit"
			wscript.quit
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
	strTemplate = UserInput("No template file was specified. Please provide template file " & vbcrlf & _
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
strDefCIQFolder = Mid(DefInFile, 1, InStrRev(DefInFile, "\")) & DefCIQFolderName

'Prompt if CIQ folder was not provided
if strCIQFolder = "" then
	strCIQFolder = UserInput("No CIQ folder provided. Please provide comlete path where you want the CIQ's stored " & vbcrlf & _
								"or leave blank to use the default of " & strDefCIQFolder & vbcrlf & "Your Input:")
	if strCIQFolder = "" then strCIQFolder = strDefCIQFolder
end if

if right(strCIQFolder,1)<> "\" then
	strCIQFolder = strCIQFolder & "\"
'else
'	strCIQFolder = left(strCIQFolder,len(strCIQFolder)-1) & strSafeDate & "-" & strSafeTime & "\"
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
Set wbin = app.Workbooks.Open (strInFile,0,false)
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
Set wsTNE = wbout.Worksheets("Traffica IP")
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
	if bLog then WriteLog "Parsing " & CleanStr(wsin.Cells(x,1))
	dictCommon.RemoveAll
	dictCommon.add "DevName",     CleanStr(wsin.Cells(x,1))
	dictCommon.add "Dev2Name",    CleanStr(wsin.Cells(x,2))
	dictCommon.add "SiteName",    CleanStr(wsin.Cells(x,3))
	dictCommon.add "TestIP",      CleanStr(wsin.Cells(x,6))
	dictCommon.add "SiteID",      CleanStr(wsin.Cells(x,7))
	dictCommon.add "SBType",      CleanStr(wsin.Cells(x,8))
	dictCommon.add "SB1",         CleanStr(wsin.Cells(x,9))
	dictCommon.add "SB2",         CleanStr(wsin.Cells(x,10))
	dictCommon.add "SBPort1",     CleanStr(wsin.Cells(x,11))
	dictCommon.add "SBPort2",     CleanStr(wsin.Cells(x,12))
	dictCommon.add "TNEPort",     CleanStr(wsin.Cells(x,13))
	dictCommon.add "WO",          CleanStr(wsin.Cells(x,14))
	dictCommon.add "TestWO",      CleanStr(wsin.Cells(x,15))
	dictCommon.add "ASATask",     CleanStr(wsin.Cells(x,16))
	dictCommon.add "SBTask",      CleanStr(wsin.Cells(x,17))
	dictCommon.add "IPCPriv",     CleanStr(wsin.Cells(x,18))
	dictCommon.add "IPCPub",      CleanStr(wsin.Cells(x,19))
	dictCommon.add "MAPName",     CleanStr(wsin.Cells(x,20))
	dictCommon.add "MAPTab",      CleanStr(wsin.Cells(x,21))
	dictCommon.add "TNEVlan",     CleanStr(wsin.Cells(x,22))
	dictCommon.add "DevRR",       CleanStr(wsin.Cells(x,23))
	dictCommon.add "Dev2RR",      CleanStr(wsin.Cells(x,24))
	dictCommon.add "SB1RR",       CleanStr(wsin.Cells(x,25))
	dictCommon.add "SB2RR",       CleanStr(wsin.Cells(x,26))
	dictCommon.add "DevMgmt",     CleanStr(wsin.Cells(x,27))
	dictCommon.add "OMU0",        CleanStr(wsin.Cells(x,28))
	dictCommon.add "OMU1",        CleanStr(wsin.Cells(x,29))
	dictCommon.add "Dev2Mgmt",    CleanStr(wsin.Cells(x,30))
	dictCommon.add "Dev2iLo",     CleanStr(wsin.Cells(x,31))
	dictCommon.add "NMnet",       CleanStr(wsin.Cells(x,32))
	dictCommon.add "NMnetMask",   CleanStr(wsin.Cells(x,33))
	dictCommon.add "Vlan",        CleanStr(wsin.Cells(x,34))
	dictCommon.add "ASA1",        CleanStr(wsin.Cells(x,35))
	dictCommon.add "ASA1RR",      CleanStr(wsin.Cells(x,36))
	dictCommon.add "ASA2",        CleanStr(wsin.Cells(x,37))
	dictCommon.add "ASA2RR",      CleanStr(wsin.Cells(x,38))
	dictCommon.add "ASAType",     CleanStr(wsin.Cells(x,39))
	dictCommon.add "ASAP11",      CleanStr(wsin.Cells(x,40))
	dictCommon.add "ASAP12",      CleanStr(wsin.Cells(x,41))
	dictCommon.add "TNEMgmt",     CleanStr(wsin.Cells(x,42))
	dictCommon.add "ASAP21",      CleanStr(wsin.Cells(x,43))
	dictCommon.add "ASAP22",      CleanStr(wsin.Cells(x,44))
	dictCommon.add "TNEiLo",      CleanStr(wsin.Cells(x,45))
	dictCommon.add "CSM01",       CleanStr(wsin.Cells(x,46))
	dictCommon.add "CSM02",       CleanStr(wsin.Cells(x,47))
	dictCommon.add "CSM01RR",     CleanStr(wsin.Cells(x,48))
	dictCommon.add "CSM02RR",     CleanStr(wsin.Cells(x,49))
	dictCommon.add "CSMPort1",    CleanStr(wsin.Cells(x,50))
	dictCommon.add "CSMPort2",    CleanStr(wsin.Cells(x,51))
	if CleanStr(wsin.Cells(x,52)) = "" then
		strTmp = GetIP (CleanStr(wsin.Cells(x,46)))
		if strTmp <> "" then
			wsin.Cells(x,52).value = strTmp
			bChanged = True
		end if
	end if
	dictCommon.add "CSMMgt1",     CleanStr(wsin.Cells(x,52))
	if CleanStr(wsin.Cells(x,53)) = "" then
		strTmp = GetIP (CleanStr(wsin.Cells(x,47)))
		if strTmp <> "" then
			wsin.Cells(x,52).value = strTmp
			bChanged = True
		end if
	end if
	dictCommon.add "CSMMgt2",     CleanStr(wsin.Cells(x,52))

	strTmp = CleanStr(wsin.Cells(x,4)) ' start by grabing the site private subnet and start dividing it
	strParts = split(strTmp,".")
	if ValidateIP(strTmp) then
		dictCommon.add "S6AS13PRI",   strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 0
		dictCommon.add "S6AS13SEC",   strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 4
		dictCommon.add "S1",          strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 8
		dictCommon.add "X11",         strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 16
		dictCommon.add "TrafficaA",   strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 24
		dictCommon.add "X2",          strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 32
		dictCommon.add "TrafficaB",   strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 48
		dictCommon.add "SGSPri",      strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 64
	    dictCommon.add "SGSSEC",      strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 68
	    dictCommon.add "Sv",          strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 72
	    dictCommon.add "SBc",         strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 80
	    dictCommon.add "SLsPri",      strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 88
	    dictCommon.add "SLsSec",      strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 92
	    dictCommon.add "SLgPri",      strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 96
	    dictCommon.add "SLgSec",      strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 100
	    dictCommon.add "UL11",        strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 104
	    dictCommon.add "UL12",        strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 108
	    dictCommon.add "UL21",        strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 112
	    dictCommon.add "UL22",        strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 116
	else
		writelog "Found invalid IP '" & strTmp & "' so I can't do any IP assignments for that range"
	end if

	strTmp = CleanStr(wsin.Cells(x,5)) ' Now grab the public subnet
	strParts = split(strTmp,".")

	if ValidateIP(strTmp) then
	    dictCommon.add "S11S10",      strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 0
    	dictCommon.add "Gn",          strParts(0) & "." & strParts(1) & "." & strParts(2) & "." & strParts(3) + 8
	else
		writelog "Found invalid IP '" & strTmp & "' so I can't do any IP assignments for that range"
    end if


	if bLog then WriteLog "Updating IP Tab"
	wsout.cells(1,3).value  = dictCommon.item("SiteName")
	wsout.cells(1,8).value  = dictCommon.item("SiteID")
	wsout.cells(4,3).value  = dictCommon.item("DevName")
	wsout.cells(5,3).value  = dictCommon.item("Dev2Name")
	wsout.cells(76,8).value = dictCommon.item("SB1")
	wsout.cells(78,8).value = dictCommon.item("SB2")
	wsout.cells(76,9).value = dictCommon.item("SBPort1")
	wsout.cells(77,9).value = dictCommon.item("SBPort2")
	wsout.cells(78,9).value = dictCommon.item("SBPort1")
	wsout.cells(79,9).value = dictCommon.item("SBPort2")
	wsout.cells(90,4).value = dictCommon.item("IPCPriv")
	wsout.cells(91,4).value = dictCommon.item("IPCPub")
	wsout.cells(92,4).value = dictCommon.item("MAPName")
	wsout.cells(93,4).value = dictCommon.item("MAPTab")
	wsout.cells(11,2).value = dictCommon.item("S6AS13PRI")
	wsout.cells(12,2).value = dictCommon.item("S6AS13SEC")
	wsout.cells(13,2).value = dictCommon.item("S1")
	wsout.cells(14,2).value = dictCommon.item("X11")
	wsout.cells(15,2).value = dictCommon.item("X2")
	wsout.cells(16,2).value = dictCommon.item("TrafficaA")
	wsout.cells(16,5).value = dictCommon.item("TNEVlan")
	wsout.cells(17,2).value = dictCommon.item("TrafficaB")
	wsout.cells(18,2).value = dictCommon.item("SGSPri")
	wsout.cells(19,2).value = dictCommon.item("SGSSEC")
	wsout.cells(20,2).value = dictCommon.item("S11S10")
	wsout.cells(21,2).value = dictCommon.item("Gn")
	wsout.cells(22,2).value = dictCommon.item("Sv")
	wsout.cells(23,2).value = dictCommon.item("SBc")
	wsout.cells(24,2).value = dictCommon.item("SLsPri")
	wsout.cells(25,2).value = dictCommon.item("SLsSec")
	wsout.cells(26,2).value = dictCommon.item("SLgPri")
	wsout.cells(27,2).value = dictCommon.item("SLgSec")
	wsout.cells(28,2).value = dictCommon.item("UL11")
	wsout.cells(29,2).value = dictCommon.item("UL12")
	wsout.cells(30,2).value = dictCommon.item("UL21")
	wsout.cells(31,2).value = dictCommon.item("UL22")
	wsout.cells(85,4).value = dictCommon.item("WO")
	wsout.cells(4,4).value  = dictCommon.item("DevRR")
	wsout.cells(5,4).value  = dictCommon.item("Dev2RR")
	wsout.cells(76,7).value = dictCommon.item("SB1RR")
	wsout.cells(78,7).value = dictCommon.item("SB2RR")
	wsout.cells(4,5).value  = dictCommon.item("DevMgmt")
	wsout.cells(4,7).value  = dictCommon.item("OMU0")
	wsout.cells(4,8).value  = dictCommon.item("OMU1")
	wsout.cells(5,5).value  = dictCommon.item("Dev2Mgmt")
	wsout.cells(5,6).value  = dictCommon.item("Dev2iLo")
	wsout.cells(8,3).value  = dictCommon.item("TestIP")
	wsout.cells(32,2).value = dictCommon.item("NMnet")
	wsout.cells(32,7).value = dictCommon.item("NMnetMask")
	wsout.cells(32,5).value = dictCommon.item("Vlan")
	wsout.cells(74,8).value = dictCommon.item("ASA1")
	wsout.cells(74,7).value = dictCommon.item("ASA1RR")
	wsout.cells(75,8).value = dictCommon.item("ASA2")
	wsout.cells(75,7).value = dictCommon.item("ASA2RR")
	wsout.cells(74,9).value = dictCommon.item("ASAP11")
	wsout.cells(75,9).value = dictCommon.item("ASAP21")
	wsout.cells(80,9).value = dictCommon.item("ASAP12")
	wsout.cells(81,9).value = dictCommon.item("ASAP22")
	wsout.cells(82,8).value = dictCommon.item("CSM01")
	wsout.cells(83,8).value = dictCommon.item("CSM02")
	wsout.cells(82,7).value = dictCommon.item("CSM01RR")
	wsout.cells(83,7).value = dictCommon.item("CSM02RR")
	wsout.cells(82,9).value = dictCommon.item("CSMPort1")
	wsout.cells(83,9).value = dictCommon.item("CSMPort2")
	wsout.cells(6,5).value  = dictCommon.item("CSMMgt1")
	wsout.cells(7,5).value  = dictCommon.item("CSMMgt2")
	wsout.cells(63,9).value = dictCommon.item("TestWO")
	wsout.cells(65,9).value = dictCommon.item("ASATask")
	wsout.cells(66,9).value = dictCommon.item("ASATask")
	wsout.cells(67,9).value = dictCommon.item("SBTask")
	wsout.cells(68,9).value = dictCommon.item("SBTask")
	wsout.cells(65,10).value = dictCommon.item("ASAType")
	wsout.cells(66,10).value = dictCommon.item("ASAType")
	wsout.cells(67,10).value = dictCommon.item("SBType")
	wsout.cells(68,10).value = dictCommon.item("SBType")

	if bLog then WriteLog "Updating Traffica Tab"
	wsTNE.cells(3,5).value  = dictCommon.item("TNEMgmt")
	wsTNE.cells(4,5).value  = dictCommon.item("TNEPort")
	wsTNE.cells(5,5).value  = dictCommon.item("TNEPort")
	wsTNE.cells(6,5).value  = dictCommon.item("TNEiLo")

	if bLog then WriteLog "Updating change log"

	wsLog.cells(y,1).value = now
	wsLog.cells(y,3).value = "New CIQ for " & dictCommon.item("DevName") & " auto generated by " & strScriptName

	if bLog then WriteLog "Saving "

	strFullFileName = strCIQFolder & dictCommon.item("DevName") & CIQSuffix

    If fso.FileExists(strFullFileName) and bOverWrite Then
        fso.DeleteFile strFullFileName,true
    End If

	wbout.SaveAs strFullFileName

	' Saving CSV files if so directed
	if bCSV then
		if bLog then WriteLog "Create CSV folder if needed"
		strCSVFolder = wbout.Path & "\" & DefCSVFolder
	    If Not fso.FolderExists(strCSVFolder) Then
	        fso.CreateFolder (strCSVFolder)
	        WriteLog strCSVFolder & " did not exists so I created it"
	    End If

		if bLog then WriteLog "Saving Cable testing CSV file"
	    strFullFileName = strCSVFolder & wsout.Range("C4").Value & CableTestSuffix
	    If fso.FileExists(strFullFileName) and bOverWrite Then
	        fso.DeleteFile strFullFileName,true
	    End If
	    set wsCSV = wbout.Sheets("Port Testing CSV")
	    Set wbtmp = app.Workbooks.Add
	    set wstmp = wbtmp.sheets(1)
		wstmp.Range("A1:I13").value = wsCSV.Range("A1:I13").value
	    wstmp.SaveAs strFullFileName, xlCSV
	    wbtmp.close False
	    set wstmp = nothing
	    Set wbtmp = nothing

		if bLog then WriteLog "Saving HPNA Automation CSV file"
	    strFullFileName = strCSVFolder & wsout.Range("C4").Value & HPNASuffix
	    If fso.FileExists(strFullFileName) and bOverWrite Then
	        fso.DeleteFile strFullFileName,true
	    End If
	    set wsCSV = wbout.Sheets("HPNA Access port")
	    Set wbtmp = app.Workbooks.Add
	    set wstmp = wbtmp.sheets(1)
		wstmp.Range("A1:G7").value = wsCSV.Range("A1:G7").value
	    wstmp.SaveAs strFullFileName, xlCSV
	    wbtmp.close False
	    set wstmp = nothing
	    Set wbtmp = nothing

		if bLog then WriteLog "Saving IP Control Subnets CSV file"
	    strFullFileName = strCSVFolder & wsout.Range("C4").Value & SubnetIPCSuffix
	    If fso.FileExists(strFullFileName) and bOverWrite Then
	        fso.DeleteFile strFullFileName,true
	    End If
	    set wsCSV = wbout.Sheets("Subnet IPC")
	    Set wbtmp = app.Workbooks.Add
	    set wstmp = wbtmp.sheets(1)
		wstmp.Range("A1:G21").value = wsCSV.Range("A2:G22").value
	    wstmp.SaveAs strFullFileName, xlCSV
	    wbtmp.close False
	    set wstmp = nothing
	    Set wbtmp = nothing

		if bLog then WriteLog "Saving IP Control Host IP CSV file"
	    strFullFileName = strCSVFolder & wsout.Range("C4").Value & HostIPCSuffix
	    If fso.FileExists(strFullFileName) and bOverWrite Then
	        fso.DeleteFile strFullFileName,true
	    End If
	    set wsCSV = wbout.Sheets("Device IPControl")
	    Set wbtmp = app.Workbooks.Add
	    set wstmp = wbtmp.sheets(1)
		wstmp.Range("A1:N96").value = wsCSV.Range("A2:N97").value
	    wstmp.SaveAs strFullFileName, xlCSV
	    wbtmp.close False
	    set wstmp = nothing
	    Set wbtmp = nothing
	end if

	if bLog then WriteLog "Done Next one ... "
	x=x+1

Loop Until CleanStr(wsin.Cells(x,1))=""

if bChanged then
	wbin.save
end if

wbin.close False
wbout.close False
app.quit ' Close Excel

WriteLog "Done. CIQ's can be found in " & strCIQFolder

if bLog then WriteLog "Creating a combined CSV Files"
strFullFileName = strCSVFolder & "All" & SubnetIPCSuffix
If fso.FileExists(strFullFileName) Then
    fso.DeleteFile strFullFileName,true
End If

strParams = "%comspec% /c copy "& strCSVFolder & "*" &SubnetIPCSuffix & " " & strFullFileName
Set objExecObj = objShell.exec(strParams)
Do While Not objExecObj.StdOut.AtEndOfStream
	if bLog then WriteLog objExecObj.StdOut.Readline()
Loop
set objExecObj = Nothing

strFullFileName = strCSVFolder & "All" & HostIPCSuffix
If fso.FileExists(strFullFileName) Then
    fso.DeleteFile strFullFileName,true
End If

strParams = "%comspec% /c copy " & strCSVFolder & "*" & HostIPCSuffix & " " & strFullFileName
Set objExecObj = objShell.exec(strParams)
Do While Not objExecObj.StdOut.AtEndOfStream
	if bLog then WriteLog objExecObj.StdOut.Readline()
Loop
set objExecObj = Nothing

strFullFileName = strCSVFolder & "All" & CableTestSuffix
If fso.FileExists(strFullFileName) Then
    fso.DeleteFile strFullFileName,true
End If

strParams = "%comspec% /c copy " & strCSVFolder & "*" & CableTestSuffix & " " & strFullFileName
Set objExecObj = objShell.exec(strParams)
Do While Not objExecObj.StdOut.AtEndOfStream
	if bLog then WriteLog objExecObj.StdOut.Readline()
Loop
set objExecObj = Nothing

'Cleanup, close out files, release resources, etc.
if bLog then WriteLog "Done with combining CSV's, nothing left but cleanup"
if bFile then
	objLogOut.close
	set objLogOut = nothing
end if

set objShell    = nothing
Set fso         = nothing
set wsCSV       = nothing
Set wsin        = Nothing
Set wbin        = Nothing
Set wsout       = Nothing
Set wsTNE       = Nothing
Set wsLog       = Nothing
Set wbout       = Nothing
Set app         = Nothing
set dictSubnets = Nothing
set dictCommon  = Nothing

wscript.echo now & vbtab & "All Done, Cleanup complete"

function ValidateIP (strIP)
'-------------------------------------------------------------------------------------------------'
' Function ValidateIP (strIP)                                                                     '
'                                                                                                 '
' This function accepts one input parameter, an IPv4 IP address, and validates passes basic       '
' formating requirements                                                                          '
'-------------------------------------------------------------------------------------------------'
	dim IPQuads, IP

	ValidateIP = True
	IPQuads = split(strIP,".")
	if ubound(IPQuads) = 3 then
		for each IP in IPQuads
			if isnumeric(IP) then
				if IP > 255 or IP < 0 then
					ValidateIP = false
				end if
			else
				ValidateIP = false
			end if
		Next
	else
		ValidateIP = false
	end if

end function

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

	 If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
		strHelpMsg =              " Usage: cscript " & strScriptName & " infile=inpath log file " & vbcrlf
		strHelpMsg = strHelpMsg & "        " & strScriptName & ": The name of this script" & vbcrlf
	else
		strHelpMsg = strHelpMsg & " Usage: " & strScriptName & " infile=inpath log file " & vbcrlf
		strHelpMsg = strHelpMsg & "        " & strScriptName & ": The name of this script" & vbcrlf
	end if
	strHelpMsg = strHelpMsg & "    InFile: The complete path to the spreadsheet with the site specific data" & vbcrlf
	strHelpMsg = strHelpMsg & "    template: The complete path to the CIQ template spreadsheet" & vbcrlf
	strHelpMsg = strHelpMsg & "    CIQ: The complete path of where you want the generated CIQ's saved" & vbcrlf
	strHelpMsg = strHelpMsg & "    Options include:" & vbcrlf
	strHelpMsg = strHelpMsg & "        log: Flag to indicate detailed log, summary is default" & vbcrlf
	strHelpMsg = strHelpMsg & "        file: Flag to log everything to a log file, default is no log file" & vbcrlf
	strHelpMsg = strHelpMsg & "        hide: Keep Excel hidden during processing, only applies to the instanse" & vbcrlf
	strHelpMsg = strHelpMsg & "             the script starts. Default is to show Excel." & vbcrlf
	strHelpMsg = strHelpMsg & "        csv: Generate CSV files for Cable testing as well as IPControl upload" & vbcrlf
	strHelpMsg = strHelpMsg & "             of Subnet and devices" & vbcrlf
	strHelpMsg = strHelpMsg & "        overwrite: Signals to automatically overwrite existing files" &vbcrlf
	strHelpMsg = strHelpMsg & vbcrlf
	strHelpMsg = strHelpMsg & "  All arguments are optional and the order does not matter." & vbcrlf
	strHelpMsg = strHelpMsg & "  The nessisary data you don't provide will be prompted for" & vbcrlf
	strHelpMsg = strHelpMsg & "  if you provide ""help"" or ""?"" as an argument this message will be printed" & vbcrlf
	strHelpMsg = strHelpMsg & "   and script will exit" & vbcrlf
	strHelpMsg = strHelpMsg & vbcrlf
	strHelpMsg = strHelpMsg & "  Example: " & vbcrlf
	strHelpMsg = strHelpMsg & "cscript " & strScriptName & " InFile=C:\CIQGen\2015GGSNSubnets.xlsx template=C:\CIQ\ASR5500CIQ.xlsx log file" & vbcrlf

	 ' Check if the script runs in CSCRIPT.EXE
	 If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
	 ' If so, use StdIn and StdOut
	 WScript.StdOut.Writeline strHelpMsg
	Else
	 Msgbox strHelpMsg', "Help message"
	end if

end sub

function GetIP (strHostName)
'-------------------------------------------------------------------------------------------------'
' GetIP (strHostName)                                                                             '
'                                                                                                 '
' This function takes one string inpput and does a nslookup on it, returns the result             '
'-------------------------------------------------------------------------------------------------'
	dim strParams, objExecObj, strText, strhost, strAddr
	strParams = "%comspec% /c NSlookup " & strHostName
	Set objExecObj = objShell.exec(strParams)

	Do While Not objExecObj.StdOut.AtEndOfStream
		strText = objExecObj.StdOut.Readline()
		if instr (strText, "Name") Then
			strhost = trim(replace(strText,"Name:",""))
		End if
		if instr (strText, "Address") Then
			strAddr = trim(replace(strText,"Address:",""))
		End if
	Loop
	if strhost <> "" then
		GetIP = strAddr
	else
		GetIP = ""
	end if
	Set objExecObj = nothing
end function