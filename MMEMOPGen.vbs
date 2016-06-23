Option Explicit
'----------------------------------------------------------------------------------------------------'
' This script will loop through all the files in the provided input directory of NSN MME             '
' CIQ's, parse out the appropriate variables generate a configuration file in IOS syntax             '
' for turning up those MME's                                                                         '
'                                                                                                    '
' Author: Siggi Bjarnason                                                                            '
' Date: 04/15/2015                                                                                   '
' Usage: parser folder=inpath log file hide                                                          '
'        parser: The name of this script                                                             '
'        folder: The complete path of the folder where everything is stored                          '
'        log:    flag to indicate detailed log, summary is default                                   '
'        file:   flag to log everything to a log file, default is no log file                        '
'        hide:   Keep Apps hidden during processing, only applies to the instanse the script starts  '
'        NoMOP:  Do not auto generate word based MOPs, just text files.                              '
'                                                                                                    '
'  All arguments are optional and the order does not matter                                          '
'  if you do not provide path you will be prompted for one                                           '
'  default folder configured below will be suggested while prompting for the path                    '
'  NOTE: CIQ's should be in a sub folder inside folder identified by path                            '
'  Configurations will be saved in their own sub folder                                              '
'  Template file should be in the root of this folder                                                '
'  if you provide "help" or "?" as an argument help message will be printed and script will exit     '
'                                                                                                    '
'  Example: cscript MMEMOPGen.vbs folder=C:\LTE\ log file                                            '
'                                                                                                    '
'----------------------------------------------------------------------------------------------------'

'User definable Constants

const DefFolder           = "C:\LTE\2015Q1\"
const DefCIQFolderName    = "CIQ\"
const DefConfigFolderName = "Configurations\"
const DefTemplates        = "Templates"
Const DefWordTemplate     = "TemplateMMECorenet.docx"
Const DefMOPFolder        = "MOPs\"

const VarDelim            = "$"

'Nothing below here is user configurable proceed at your own risk.

'Variable declaration
Dim strCIQFolder, strConfFolder, strTemplates, strScriptFullName, strLogFileName, strConfigFileName, strMOPName
Dim strLine, strVar, strConfig, strTemplate, iArg, strParts, strScriptName, strInput, strArgParts, strDate, strFind
Dim dictSubnets, dictCommon, dictSB1, dictSB2, bLog, bFile, bShowApps, wsMTS, dictTemplates, dictWord, bCreateMOP
Dim objLogOut, objTemplate, objConfig, objExcel, wb, ws, fso, f, fc, f1, FolderSpec, doc, sel, objWord
Dim strWordTemplate, strMOPFolder

const ForReading     = 1
const ForWriting     = 2
const ForAppending   = 8
const wdReplaceAll   = 2
const wdFindContinue = 1

bLog  = false
bFile = false
bShowApps = True
bCreateMOP = true
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
		case "folder"
			FolderSpec = strParts(1)
			wscript.echo "found an argument for folder " & FolderSpec
		case "log"
			bLog = true
			wscript.echo "detailed logging enabled"
		case "file"
			bFile = true
			wscript.echo "log file enabled"
		case "hide"
			bShowApps = false
			wscript.echo "not showing apps such as Word and Excel"
		case "nomop"
			bCreateMOP = false
			wscript.echo "not creating word based IP MOPs, just text file configurations"
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
	Set objLogOut = fso.OpenTextFile(strLogFileName, ForAppending, True)
end if
wscript.echo "starting " & strScriptName &" at " & now

'Validate input
if FolderSpec = "" then
	FolderSpec = UserInput("No folder was specified. Please provide root " & vbcrlf & _
								"folder or leave blank to use the default of " & DefFolder & vbcrlf & "Your Input:")
	if FolderSpec = "" then FolderSpec = DefFolder
end if

if right(FolderSpec,1)<> "\" then
	FolderSpec = FolderSpec & "\"
end if


' Defining full path names for all the relevant files and folders
strCIQFolder     = FolderSpec & DefCIQFolderName
strConfFolder    = FolderSpec & DefConfigFolderName
strTemplates     = FolderSpec & DefTemplates
strWordTemplate  = FolderSpec & DefWordTemplate
strMOPFolder     = FolderSpec & DefMOPFolder

' Log the full path names for all the relevant files and folders
WriteLog "Main directory " & FolderSpec
WriteLog "Reading CIQ's from " & strCIQFolder
WriteLog "Using template directory " & strTemplates
WriteLog "Saving Configurations to " & strConfFolder
WriteLog "Saving MOPs to " & strMOPFolder
WriteLog "Word template is " & strWordTemplate


if bLog then WriteLog "Validating input"

'Validating existance of the main folder, prompt if not valid
while not fso.FolderExists(FolderSpec)
	FolderSpec = UserInput("Folder " & FolderSpec & " is not valid. Please provide new one or leave blank to abort:")
	if FolderSpec = "" then
		wscript.echo "No input provided, aborting"
		wscript.quit
	end if
wend

'Validating the existance of the CIQ folder, prompt if not valid
while not fso.FolderExists(strCIQFolder)
	strCIQFolder = UserInput("Folder " & strCIQFolder & " does not seem to exists. Please provide new one or leave blank to abort:")
	if strCIQFolder = "" then
		wscript.echo "No input provided, aborting"
		wscript.quit
	end if
wend

'Validating the existance of the template folder, prompt if not valid
while not fso.FolderExists(strTemplates)
	strTemplates = UserInput("Folder " & strTemplates & " does not seem to exists. Please provide new one or leave blank to abort:")
	if strTemplates = "" then
		wscript.echo "No input provided, aborting"
		wscript.quit
	end if
wend

'Validating the existance of the configuration folder, create if not valid
if not fso.FolderExists(strConfFolder) then
	CreatePath strConfFolder
	WriteLog strConfFolder & " did not exists so I created it"
end if

if bLog then WriteLog "Starting Excel"
'Open Excel and get hook into it
on error resume next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
	WriteLog "Unable to start Excel, probably not installed correctly. Unable to read the CIQ's without it"
	wscript.quit
end if
on error goto 0

if bCreateMOP then
	'Validating the existance of the MOP folder, create if not valid
	if not fso.FolderExists(strMOPFolder) then
		CreatePath strMOPFolder
		WriteLog strMOPFolder & " did not exists so I created it"
	end if

	'Validating the existance of the Word template file, prompt if not valid
	while not fso.FileExists(strWordTemplate)
		strWordTemplate = UserInput("Word Template " & strWordTemplate & " does not seem to exists. Please provide new one or leave blank to abort:")
		if strWordTemplate = "" then
			wscript.echo "No input provided, aborting"
			wscript.quit
		end if
	wend
	if bLog then WriteLog "Starting Word"
	'Open Word and get hook into it
	on error resume next
	Set objWord = CreateObject("Word.Application")
	If Err.Number <> 0 Then
		WriteLog "Unable to start Word, probably not installed correctly. Unable to create Word formated MOPs without it"
		wscript.quit
	end if
	Set doc = objWord.Documents.Open (strWordTemplate,0,true)
	set sel = objWord.selection
	If Err.Number <> 0 Then
		WriteLog "Unable to open " & strWordTemplate & ". Aborting"
		wscript.quit
	end if
	on error goto 0
	objWord.visible = bShowApps
end if
'Set the visibility of Excel based on input parameters.
objExcel.visible = bShowApps

if bLog then WriteLog "Initializing Dictionaries"
'Initializing Dictionaries, aka ordered arrays
set dictSubnets   = CreateObject("Scripting.Dictionary")
set dictCommon    = CreateObject("Scripting.Dictionary")
set dictSB1       = CreateObject("Scripting.Dictionary")
set dictSB2       = CreateObject("Scripting.Dictionary")
set dictTemplates = CreateObject("Scripting.Dictionary")
set dictWord      = CreateObject("Scripting.Dictionary")

'Load initial values into dictionaries as applicable
InitializeDicts

dictWord.add "Baseline.txt", "Baseline"
dictWord.add "Implementation.txt", "Implementation"
dictWord.add "Rollback.txt", "BackOut"
dictWord.add "Verifications.txt", "Verification"


'Create an array of all the file names in the template folder
Set f = fso.GetFolder(strTemplates)
Set fc = f.Files

'Loop through the array of template's and stick them in a dictionary.
For Each f1 in fc
	strParts = split (f1, "\")
	strConfig = strParts(ubound(strParts))

	if bLog then WriteLog "processing " & strConfig

	'open up the template file and read it all into single variable, then close the file
	set objTemplate = fso.OpenTextFile(f1, ForReading, False)
	dictTemplates.add strConfig, objTemplate.readall
	objTemplate.close
	set objTemplate = Nothing
next

'Create an array of all the file names in the CIQ folder
Set f = fso.GetFolder(strCIQFolder)
Set fc = f.Files

'Loop through the array of CIQ's and process them.
For Each f1 in fc
	if bLog then WriteLog "Opening up CIQ " & f1
	on error resume next
	Set wb = objExcel.Workbooks.Open (f1,0,true)
	If Err.Number <> 0 Then
		WriteLog "Unable to open " & f1 & " continuing to the next one"
	else
		on error goto 0
		Set ws    = wb.Worksheets(1)
		Set wsMTS = wb.Worksheets(2)
		strLine = ""

		if bLog then WriteLog "Starting with Common elements"
		dictCommon.RemoveAll
		dictCommon.add "SiteName",  CleanStr(ws.Cells(1,3))
		dictCommon.add "S1uVRF", CleanStr(ws.Cells(89,4))
		dictCommon.add "MTS-VlanID", CleanStr(ws.Cells(16,5))
		dictCommon.add "SvMask", CleanStr(ws.Cells(22,3))
		dictCommon.add "SBcMask", CleanStr(ws.Cells(23,3))
		dictCommon.add "SBcSubnet", CleanStr(ws.Cells(23,2))
		dictCommon.add "X2Subnet", CleanStr(ws.Cells(15,2))
		dictCommon.add "MTS-Subnet", CleanStr(ws.Cells(16,2))
		dictCommon.add "MTSSubnet", CleanStr(ws.Cells(17,2))
		dictCommon.add "S1uSubnet", CleanStr(ws.Cells(13,2))
		dictCommon.add "X1-1Subnet", CleanStr(ws.Cells(14,2))
		dictCommon.add "MTS-Bits", CleanStr(ws.Cells(16,7))
		dictCommon.add "S11Subnet", CleanStr(ws.Cells(20,2))
		dictCommon.add "S1uMask", CleanStr(ws.Cells(13,3))
		dictCommon.add "S1uName", CleanStr(ws.Cells(13,8))
		dictCommon.add "X1-1Name", CleanStr(ws.Cells(14,8))
		dictCommon.add "X2Name", CleanStr(ws.Cells(15,8))
		dictCommon.add "MTS-VlanName", CleanStr(ws.Cells(16,8))
		dictCommon.add "MTSVlanName", CleanStr(ws.Cells(17,8))
		dictCommon.add "MME-Name", CleanStr(ws.Cells(4,3))
		dictCommon.add "MTS-Name", CleanStr(ws.Cells(5,3))
		dictCommon.add "S11Name", CleanStr(ws.Cells(20,8))
		dictCommon.add "GnName", CleanStr(ws.Cells(21,8))
		dictCommon.add "SvName", CleanStr(ws.Cells(22,8))
		dictCommon.add "SBcName", CleanStr(ws.Cells(23,8))
		dictCommon.add "X1-1Mask", CleanStr(ws.Cells(14,3))
		dictCommon.add "X2Mask", CleanStr(ws.Cells(15,3))
		dictCommon.add "MME-RR", FormatRR(ws.Cells(4,4))
		dictCommon.add "MTS-RR", FormatRR(ws.Cells(5,4))
		dictCommon.add "MTS-HSRP", CleanStr(ws.Cells(56,9))
		dictCommon.add "MTS-Mask", CleanStr(ws.Cells(16,3))
		dictCommon.add "MTSMask", CleanStr(ws.Cells(17,3))
		dictCommon.add "GnSubnet", CleanStr(ws.Cells(21,2))
		dictCommon.add "SvSubnet", CleanStr(ws.Cells(22,2))
		dictCommon.add "S11Mask", CleanStr(ws.Cells(20,3))
		dictCommon.add "GnMask", CleanStr(ws.Cells(21,3))
		dictCommon.add "WONum", CleanStr(ws.Cells(85,4))
		dictCommon.add "xCon", CleanStr(ws.Cells(86,4))
		dictCommon.add "TrackLo", CleanStr(ws.Cells(87,4))
		dictCommon.add "Uplink", CleanStr(ws.Cells(88,4))
		dictCommon.add "MAPName", CleanStr(ws.Cells(92,4))
		dictCommon.add "MAPTab", CleanStr(ws.Cells(93,4))

		if bLog then WriteLog "Now working on SB1 specifics"

		dictSB1.RemoveAll
		dictSB1.add "SGSMask", CleanStr(ws.Cells(18,3))
		dictSB1.add "MTS-IP", CleanStr(ws.Cells(57,9))
		dictSB1.add "SLsMask", CleanStr(ws.Cells(24,3))
		dictSB1.add "SLgMask", CleanStr(ws.Cells(26,3))
		dictSB1.add "SGSName", CleanStr(ws.Cells(18,8))
		dictSB1.add "SBPort1Mask", CleanStr(ws.Cells(28,3))
		dictSB1.add "SBPort2Mask", CleanStr(ws.Cells(29,3))
		dictSB1.add "MMEPort2IP", CleanStr(ws.Cells(67,4))
		dictSB1.add "Pri", CleanStr(ws.Cells(94,4))
		dictSB1.add "SBPort1IP", CleanStr(ws.Cells(65,3))
		dictSB1.add "MMEPort1IP", CleanStr(ws.Cells(66,3))
		dictSB1.add "SBPort2IP", CleanStr(ws.Cells(65,4))
		dictSB1.add "S6aS13Subnet", CleanStr(ws.Cells(11,2))
		dictSB1.add "MMESlot1", CleanStr(ws.Cells(76,4))
		dictSB1.add "MMESlot2", CleanStr(ws.Cells(77,4))
		dictSB1.add "SBPort1", CleanStr(ws.Cells(76,9))
		dictSB1.add "SGSSubnet", CleanStr(ws.Cells(18,2))
		dictSB1.add "SBPort2", CleanStr(ws.Cells(77,9))
		dictSB1.add "SLsSubnet", CleanStr(ws.Cells(24,2))
		dictSB1.add "MMEPort1", CleanStr(ws.Cells(76,5))
		dictSB1.add "MMEPort2", CleanStr(ws.Cells(77,5))
		dictSB1.add "SLgName", CleanStr(ws.Cells(26,8))
		dictSB1.add "SLgSubnet", CleanStr(ws.Cells(26,2))
		dictSB1.add "SLsName", CleanStr(ws.Cells(24,8))
		dictSB1.add "S6aS13Mask", CleanStr(ws.Cells(11,3))
		dictSB1.add "S6aS13Name", CleanStr(ws.Cells(11,8))
		dictSB1.add "MTS-NIC", CleanStr(wsMTS.Cells(4,1))
		dictSB1.add "MTSInt", CleanStr(wsMTS.Cells(4,5))
		dictSB1.add "SBName", CleanStr(ws.Cells(76,8))

		if bLog then WriteLog "Now working on SB2 specifics"

		dictSB2.RemoveAll
		dictSB2.add "SGSMask", CleanStr(ws.Cells(19,3))
		dictSB2.add "SGSName", CleanStr(ws.Cells(19,8))
		dictSB2.add "SLsName", CleanStr(ws.Cells(25,8))
		dictSB2.add "SLgName", CleanStr(ws.Cells(27,8))
		dictSB2.add "MTS-IP", CleanStr(ws.Cells(58,9))
		dictSB2.add "SBPort1", CleanStr(ws.Cells(78,9))
		dictSB2.add "SBPort2", CleanStr(ws.Cells(79,9))
		dictSB2.add "S6aS13Subnet", CleanStr(ws.Cells(12,2))
		dictSB2.add "SGSSubnet", CleanStr(ws.Cells(19,2))
		dictSB2.add "SLsSubnet", CleanStr(ws.Cells(25,2))
		dictSB2.add "SLgSubnet", CleanStr(ws.Cells(27,2))
		dictSB2.add "S6aS13Mask", CleanStr(ws.Cells(12,3))
		dictSB2.add "S6aS13Name", CleanStr(ws.Cells(12,8))
		dictSB2.add "SLsMask", CleanStr(ws.Cells(25,3))
		dictSB2.add "SLgMask", CleanStr(ws.Cells(27,3))
		dictSB2.add "SBPort1Mask", CleanStr(ws.Cells(30,3))
		dictSB2.add "SBPort2Mask", CleanStr(ws.Cells(31,3))
		dictSB2.add "Pri", CleanStr(ws.Cells(95,4))
		dictSB2.add "MMESlot1", CleanStr(ws.Cells(78,4))
		dictSB2.add "MMESlot2", CleanStr(ws.Cells(79,4))
		dictSB2.add "SBPort1IP", CleanStr(ws.Cells(68,5))
		dictSB2.add "MMEPort1IP", CleanStr(ws.Cells(69,5))
		dictSB2.add "MMEPort1", CleanStr(ws.Cells(78,5))
		dictSB2.add "MMEPort2", CleanStr(ws.Cells(79,5))
		dictSB2.add "SBPort2IP", CleanStr(ws.Cells(68,6))
		dictSB2.add "MMEPort2IP", CleanStr(ws.Cells(70,6))
		dictSB2.add "MTS-NIC", CleanStr(wsMTS.Cells(5,1))
		dictSB2.add "MTSInt", CleanStr(wsMTS.Cells(5,5))
		dictSB2.add "SBName", CleanStr(ws.Cells(78,8))

		if bLog then WriteLog "Now generating configuration for SB1"
		for each strTemplate in dictTemplates
			strConfig = dictTemplates.item(strTemplate)
			if bLog then WriteLog "Now working on template " & strTemplate
			for each strVar in dictCommon
				strConfig = replace(strConfig, VarDelim & strVar & VarDelim,dictCommon.item(strVar),vbTextCompare)
			next

			for each strVar in dictSB1
				strConfig = replace(strConfig, VarDelim & strVar & VarDelim,dictSB1.item(strVar),vbTextCompare)
			next

			strConfigFileName = strConfFolder & "WO" & dictCommon.item("WONum") & "-" & dictCommon.item("MME-Name") & "-" & dictSB1.item("SBName") & "-" & strTemplate
			if bLog then WriteLog "Saving configuration to " & strConfigFileName
			set objConfig = fso.createtextfile(strConfigFileName, true)
			objConfig.write strConfig
			objConfig.close
			set objConfig = nothing

			if bCreateMOP then
				strFind = VarDelim & dictWord.item(strTemplate) & "1" & VarDelim
				if bLog then WriteLog "Looking for " & strFind
				sel.find.wrap = wdFindContinue
				sel.Find.text = strFind
				while sel.Find.Execute
					if bLog then WriteLog "found " & strFind
					sel.InsertFile strConfigFileName
					if bLog then WriteLog "Inserted  " & strConfigFileName
				wend
			end if
		next

		if bLog then WriteLog "Now generating configuration for SB2"
		for each strTemplate in dictTemplates
			strConfig = dictTemplates.item(strTemplate)
			if bLog then WriteLog "Now working on template " & strTemplate
			for each strVar in dictCommon
				strConfig = replace(strConfig, VarDelim & strVar & VarDelim,dictCommon.item(strVar),vbTextCompare)
			next

			for each strVar in dictSB2
				strConfig = replace(strConfig, VarDelim & strVar & VarDelim,dictSB2.item(strVar),vbTextCompare)
			next

			strConfigFileName = strConfFolder & "WO" & dictCommon.item("WONum") & "-" & dictCommon.item("MME-Name") & "-" & dictSB2.item("SBName") & "-" & strTemplate
			if bLog then WriteLog "Saving configuration to " & strConfigFileName
			set objConfig = fso.createtextfile(strConfigFileName, true)
			objConfig.write strConfig
			objConfig.close
			set objConfig = nothing

			if bCreateMOP then
				strFind = VarDelim & dictWord.item(strTemplate) & "2" & VarDelim
				if bLog then WriteLog "Looking for " & strFind
				sel.Find.text = strFind
				sel.find.wrap = wdFindContinue
				while sel.Find.Execute
					if bLog then WriteLog "found " & strFind
					sel.InsertFile strConfigFileName
					if bLog then WriteLog "Inserted  " & strConfigFileName
				wend
			end if
		next

		if bCreateMOP then
			strDate = DatePart("m",Date) & "/" & DatePart("d",Date) & "/" & Right(DatePart("yyyy",Date),2)
			if bLog then WriteLog "Changing general MOP vars. Date:   " & strDate

			sel.find.wrap = wdFindContinue
			sel.Find.text = VarDelim & "Date" & VarDelim
			sel.Find.Replacement.Text = strDate
			sel.Find.Execute ,,,,,,,,,,wdReplaceAll

			if bLog then WriteLog "MapName:   " & dictCommon.item("MAPName")
			sel.find.wrap = wdFindContinue
			sel.Find.text = VarDelim & "MAPName" & VarDelim
			sel.Find.Replacement.Text = dictCommon.item("MAPName")
			sel.Find.Execute ,,,,,,,,,,wdReplaceAll

			if bLog then WriteLog "MAPTab:   " & dictCommon.item("MAPTab")
			sel.find.wrap = wdFindContinue
			sel.Find.text = VarDelim & "MAPTab" & VarDelim
			sel.Find.Replacement.Text = dictCommon.item("MAPTab")
			sel.Find.Execute ,,,,,,,,,,wdReplaceAll

			if bLog then WriteLog "SB1:   " & dictSB1.item("SBName")
			sel.find.wrap = wdFindContinue
			sel.Find.text = VarDelim & "SB1" & VarDelim
			sel.Find.Replacement.Text = dictSB1.item("SBName")
			sel.Find.Execute ,,,,,,,,,,wdReplaceAll

			if bLog then WriteLog "SB2:   " & dictSB2.item("SBName")
			sel.find.wrap = wdFindContinue
			sel.Find.text = VarDelim & "SB2" & VarDelim
			sel.Find.Replacement.Text = dictSB2.item("SBName")
			sel.Find.Execute ,,,,,,,,,,wdReplaceAll

			strMOPName = strMOPFolder & "WO " & dictCommon.item("WONum") & " " & dictCommon.item("MME-Name") & " " & dictCommon.item("SiteName") & " MME install.docx"
			doc.SaveAs strMOPName
			doc.close
			if bLog then WriteLog "MOP Save to " & strMOPName

			if bLog then WriteLog "Opening Template again " & strMOPName
			Set doc = objWord.Documents.Open (strWordTemplate,0,true)
			set sel = objWord.selection
		end if

		if bLog then WriteLog "Closing this CIQ and moving on to the next one"
		wb.close False
	end if
Next

if bCreateMOP then
	doc.close
	objWord.quit
	set sel         = Nothing
	set doc         = Nothing
	set objWord     = Nothing
end if



WriteLog "Done."
WriteLog "Configuration files saved to " & strConfFolder

'Cleanup, close out files, release resources, etc.
if bFile then
	objLogOut.close
	set objLogOut = nothing
end if


objExcel.quit
Set fc          = nothing
Set f           = nothing
Set fso         = nothing
Set ws          = Nothing
Set wb          = Nothing
Set objExcel    = Nothing
set dictSubnets = Nothing
set dictCommon  = Nothing
set dictSB1     = Nothing
set dictSB2     = Nothing
set dictWord    = Nothing

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
	strMsg = replace(strMsg," ", "_")
	CleanStr = strMsg
end function

function FormatRR (strMsg)
'-------------------------------------------------------------------------------------------------'
' Function CleanStr (strMsg)                                                                      '
'                                                                                                 '
' This function accepts one input parameter, a string, trims any leading or trailing spaces       '
' as well as any Carrige return or line feed characters from that string then returns it          '
'-------------------------------------------------------------------------------------------------'
dim strparts, RR1, RR2

	strMsg = CleanStr (strMsg)
	strparts = split(strMsg,".")
	RR1 = right("0000" & strparts(0),4)
	if ubound(strparts) = 0 then
		RR2="00"
	else
		RR2 = left(strparts(1)&"00",2)
	end if
	FormatRR = RR1 & "." & RR2
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
	
	strHelpMsg =              " Usage: " & strScriptName & " folder=inpath log file hide" & vbcrlf
	strHelpMsg = strHelpMsg & "        " & strScriptName & ": The name of this script" & vbcrlf
	strHelpMsg = strHelpMsg & "        folder: The complete path of the folder where everything is stored" & vbcrlf
	strHelpMsg = strHelpMsg & "        log: flag to indicate detailed log, summary is default" & vbcrlf
	strHelpMsg = strHelpMsg & "        file: flag to log everything to a log file, default is no log file" & vbcrlf
	strHelpMsg = strHelpMsg & "        hide: Keep apps hidden during processing, only applies to the instanse the script starts. Default is to show apps." & vbcrlf
	strHelpMsg = strHelpMsg & vbcrlf
	strHelpMsg = strHelpMsg & "  All arguments are optional and the order does not matter." & vbcrlf
	strHelpMsg = strHelpMsg & "  If you do not provide path you will be prompted for one" & vbcrlf
	strHelpMsg = strHelpMsg & "  default folder will be suggested while prompting for the path" & vbcrlf
	strHelpMsg = strHelpMsg & vbcrlf
	strHelpMsg = strHelpMsg & "  NOTE: CIQ's should be in a sub folder of the folder provided" & vbcrlf
	strHelpMsg = strHelpMsg & "  Configurations will be saved in their own sub folder of the folder provided" & vbcrlf
	strHelpMsg = strHelpMsg & "  Template file should be in the root of the folder provided" & vbcrlf
	strHelpMsg = strHelpMsg & "  if you provide ""help"" or ""?"" as an argument this message will be printed and script will exit" & vbcrlf
	strHelpMsg = strHelpMsg & vbcrlf
	strHelpMsg = strHelpMsg & "  Example: cscript " & strScriptName & " folder=C:\LTE\ log file" & vbcrlf
	
	 ' Check if the script runs in CSCRIPT.EXE
	 If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
	 ' If so, use StdIn and StdOut
	 WScript.StdOut.Writeline strHelpMsg
	Else
	 Msgbox strHelpMsg', "Help message"
	end if
end sub

Function CreatePath (strFullPath)
'-------------------------------------------------------------------------------------------------'
' Function CreatePath (strFullPath)                                                               '
'                                                                                                 '
' This function takes a complete path as input and builds that path out as nessisary.             '
'-------------------------------------------------------------------------------------------------'
dim pathparts, buildpath, part
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
