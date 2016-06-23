Option Explicit
'-------------------------------------------------------------------------------------------------'
' This script will loop through all the files in the provided input directory of Cisco ASR 5500   '
' CIQ's, parse out the appropriate variables and generate a single CSV file of those variables    '
' as well as generate a configuration file in IOS-XR syntax for turning up those ASR5500          '
'                                                                                                 '
' Author: Siggi Bjarnason                                                                         '
' Date: 02/04/2015                                                                                '
' Usage: parser folder=inpath log file hide                                                       '
'        parser: The name of this script                                                          '
'        folder: The complete path of the folder where everything is stored                       '
'        log: flag to indicate detailed log, summary is default                                   '
'        file: flag to log everything to a log file, default is no log file                       '
'        hide: Keep Excel hidden during processing, only applies to the instanse the script starts'
'                                                                                                 '
'  All arguments are optional and the order does not matter                                       '
'  if you do not provide path you will be prompted for one                                        '
'  default folder configured below will be suggested while prompting for the path                 '
'  NOTE: CIQ's should be in a sub folder inside folder identified by path                         '
'  Configurations will be saved in their own sub folder                                           '
'  Template file should be in the root of this folder                                             '
'  CSV file will be saved in the root of this folder                                              '
'  if you provide "help" or "?" as an argument help message will be printed and script will exit  '
'                                                                                                 '
'  Example: cscript ASR5500MOPGen.vbs folder=C:\2015ASR5500Deploy log file                        '
'                                                                                                 '
'-------------------------------------------------------------------------------------------------'

'User definable Constants

const DefFolder           = "C:\2015PCF\"
const DefCIQFolderName    = "CIQ\"
const DefConfigFolderName = "configurations\"
const CSVFolderName       = "CSVs\"
const DefTemplateName     = "ASR5500ConfDollar.txt"
const DefCSVFileName      = "ASR5500Allvars.csv"
const ACLVarCSVName       = "S1ASR55KACL.csv"
const BVIVarCSVName       = "S2ASR55KBVI.csv"
const BEDVarCSVName       = "S3ASR55KBED.csv"
const BEXVarCSVName       = "S4ASR55KBEX.csv"
const IntVarCSVName       = "S5ASR55KInt.csv"
const VPNVarCSVName       = "S6ASR55KVPN.csv"
const BGPVarCSVName       = "S7ASR55KBGP.csv"
const VarDelim            = "$"


'Nothing below here is user configurable proceed at your own risk.

'Variable declaration
Dim strCIQFolder, strConfFolder, strTemplateName, strScriptFullName, strLogFileName, strConfigFileName, strVPNVarCSVName
Dim strLine, strOutFileName, strVar, strConfig, strTemplate, iArg, strParts, strScriptName, strInput, strArgParts
Dim strCSVFolder, strACLVarCSVName, strBVIVarCSVName, strBEDVarCSVName, strBEXVarCSVName, strIntVarCSVName, strBGPVarCSVName
Dim dictSubnets, dictCommon, dictARG21, dictARG22, dictPCF, bLog, bFile, bShowExcel
Dim objFileOut, objLogOut, objTemplate, objConfig, app, wb, ws, fso, f, fc, f1, FolderSpec
Dim objACLVar, objBVIVar, objBEDVar, objBEXVar, objIntVar, objVPNVar, objBGPVar

const ForReading   = 1
const ForWriting   = 2
const ForAppending = 8

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
			bShowExcel = false
			wscript.echo "not showing excel"
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
strOutFileName   = FolderSpec & CSVFolderName & DefCSVFileName
strCIQFolder     = FolderSpec & DefCIQFolderName
strConfFolder    = FolderSpec & DefConfigFolderName
strTemplateName  = FolderSpec & DefTemplateName
strCSVFolder     = FolderSpec & CSVFolderName
strACLVarCSVName = strCSVFolder & ACLVarCSVName
strBVIVarCSVName = strCSVFolder & BVIVarCSVName
strBEDVarCSVName = strCSVFolder & BEDVarCSVName
strBEXVarCSVName = strCSVFolder & BEXVarCSVName
strIntVarCSVName = strCSVFolder & IntVarCSVName
strVPNVarCSVName = strCSVFolder & VPNVarCSVName
strBGPVarCSVName = strCSVFolder & BGPVarCSVName

' Log the full path names for all the relevant files and folders
WriteLog "Main directory " & FolderSpec
WriteLog "Reading CIQ's from " & strCIQFolder
WriteLog "Using template file " & strTemplateName
WriteLog "Saving Configurations to " & strConfFolder
WriteLog "Saving the allvar CSV file to " & strOutFileName
WriteLog "Saving the ACLVarCSVName CSV file to " & strACLVarCSVName
WriteLog "Saving the BVIVarCSVName CSV file to " & strBVIVarCSVName
WriteLog "Saving the BEDVarCSVName CSV file to " & strBEDVarCSVName
WriteLog "Saving the BEXVarCSVName CSV file to " & strBEXVarCSVName
WriteLog "Saving the IntVarCSVName CSV file to " & strIntVarCSVName
WriteLog "Saving the VPNVarCSVName CSV file to " & strVPNVarCSVName
WriteLog "Saving the BGPVarCSVName CSV file to " & strBGPVarCSVName

if bLog then WriteLog "Validating input"

'Validating existance of the main folder, prompt if not valid
while not fso.FolderExists(FolderSpec)
	FolderSpec = UserInput("Folder " & FolderSpec & " is not valid. Please provide new one or leave blank to abort:")
	if FolderSpec = "" then
		wscript.echo "No input provided, aborting"
		wscript.quit
	end if
wend

'Validating the existance of the CIQ folder, abort if not valid
if not fso.FolderExists(strCIQFolder) then
	WriteLog strCIQFolder & " Folder does not seem to exists, please rectify and re-run the script"
	wscript.quit
end if

'Validating the existance of the configuration folder, create if not valid
if not fso.FolderExists(strConfFolder) then
	fso.CreateFolder(strConfFolder)
	WriteLog strConfFolder & " did not exists so I created it"
end if

'Validating the existance of the CSV folder, create if not valid
if not fso.FolderExists(strCSVFolder) then
	fso.CreateFolder(strCSVFolder)
	WriteLog strCSVFolder & " did not exists so I created it"
end if

if bLog then WriteLog "Starting Excel"
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

if bLog then WriteLog "Initializing Dictionaries"
'Initializing Dictionaries, aka ordered arrays
set dictSubnets = CreateObject("Scripting.Dictionary")
set dictCommon  = CreateObject("Scripting.Dictionary")
set dictARG21   = CreateObject("Scripting.Dictionary")
set dictARG22   = CreateObject("Scripting.Dictionary")
set dictPCF     = CreateObject("Scripting.Dictionary")

'Load initial values into dictionaries as applicable
InitializeDicts

if bLog then WriteLog "Creating CSV files"

' Create the CSV file, overwrite any existing files.
set objFileOut = fso.createtextfile(strOutFileName, true)
if bLog then WriteLog "AllVar CSV Created"
set objACLVar  = fso.createtextfile(strACLVarCSVName, true)
if bLog then WriteLog "ACLVar CSV Created"
set objBVIVar  = fso.createtextfile(strBVIVarCSVName, true)
if bLog then WriteLog "BVIVar Created"
set objBEDVar  = fso.createtextfile(strBEDVarCSVName, true)
if bLog then WriteLog "BEDVar CSV Created"
set objBEXVar  = fso.createtextfile(strBEXVarCSVName, true)
if bLog then WriteLog "BEXVar CSV Created"
set objIntVar  = fso.createtextfile(strIntVarCSVName, true)
if bLog then WriteLog "IntVar CSV Created"
set objVPNVar  = fso.createtextfile(strVPNVarCSVName, true)
if bLog then WriteLog "VPNVar CSV Created"
set objBGPVar  = fso.createtextfile(strBGPVarCSVName, true)
if bLog then WriteLog "BGPVar CSV Created"

if bLog then WriteLog "writing headers to CSV files"


' Write header row to the AllVar CSV file
strLine = "GWName,GWRR,PCFAS,ARGAS,GaLabel,GaSubnet,GaNetMask,GaBitMask,GaID,GiLabel,GiSubnet,GiNetMask,GiBitMask,"
strline = strLine & "GiInvMask,GiID,GnLabel,GnSubnet,GnNetMask,GnBitMask,GnID,LiLabel,LiSubnet,LiNetMask,LiBitMask,"
strline = strLine & "LiID,S1uLabel,S1uSubnet,S1uNetMask,S1uBitMask,S1uID,BPLabel,BPNet,BPNetMask,BPBitMask,BPInvMask,"
strline = strLine & "BPID,GiV6Label,GiV6Subnet,GiV6MaskBit,GiV6ID,GiLoIP,GiLoV6IP,TPNet,TPInvMask,BundleID,XConBEID,SPTName,"
strline = strLine & "OMWABFName,OMWABF1,OMWABF2,OMWABF3,S1uVRF,GiVRF,GiACLName,GiACL1,GiACL2,GiV6ACLName,GiV6ACL1,GiV6ACL2,"
strline = strLine & "ABFACLName,ABFACL1,WONum,WOTask,ABFNextHop,BGPGroup,BPBGPGroup,v6BGPGroup,S1uBGPGroup,S1Label,S2Label,"
strline = strLine & "S3Label,S4Label,S5Label,S6Label,S7Label,ARG21-GaIntIP,ARG21-GiIntIP,ARG21-GnIntIP,ARG21-LiIntIP,"
strline = strLine & "ARG21-S1uIntIP,ARG21-BPIntIP,ARG21-GiV6IntIP,ARG21-GWSlot1,ARG21-GWPort1,ARGName,ARG21-ARGPort1,ARG21-GWSlot2,ARG21-GWPort2,"
strline = strLine & "ARG21-ARGPort2,ARG21-GWSlot3,ARG21-GWPort3,ARG21-ARGPort3,ARG21-GWSlot4,ARG21-GWPort4,ARG21-ARGPort4,ARG22-GaIntIP,"
strline = strLine & "ARG22-GiIntIP,ARG22-GnIntIP,ARG22-LiIntIP,ARG22-S1uIntIP,ARG22-BPIntIP,ARG22-GiV6IntIP,ARG22-GWSlot1,ARG22-GWPort1,"
strline = strLine & "ARGName,ARG22-ARGPort1,ARG22-GWSlot2,ARG22-GWPort2,ARG22-ARGPort2,ARG22-GWSlot3,ARG22-GWPort3,ARG22-ARGPort3,ARG22-GWSlot4,"
strline = strLine & "ARG22-GWPort4,ARG22-ARGPort4,PCF-GaPCFIntIP,PCF-GiPCFIntIP,PCF-GnPCFIntIP,PCF-LiPCFIntIP,PCF-S1uPCFIntIP,PCF-BPPCFIntIP,"
strline = strLine & "PCF-GiV6PCFIntIP"
objFileOut.writeline strLine

strLine = "primaryIPAddress,hostName,GiACLName,GiLoIP,GiACL1,GiACL2,GiSubnet,GiInvMask,GiV6ACLName,GiV6ACL1,GiLoV6IP,GiV6ACL2,GiV6Subnet,GiV6MaskBit,"
strLine = strLine & "ABFACLName,ABFACL1,TPNet,TPInvMask,ABFNextHop,OMWABFName,BPNet,BPInvMask,OMWABF1,OMWABF2,OMWABF3,WONum,WOTask,S1Label"
objACLVar.writeline strLine

strLine = "primaryIPAddress,hostName,GaID,GWName,GaLabel,GaSubnet,GaBitMask,GaIntIP,GaNetMask,GiID,GiLabel,GiSubnet,GiBitMask,GiVRF,GiIntIP,GiNetMask,GiACLName,"
strLine = strLine & "GnID,GnLabel,GnSubnet,GnBitMask,GnIntIP,GnNetMask,LiID,LiLabel,LiSubnet,LiBitMask,LiIntIP,LiNetMask,S1uID,S1uLabel,S1uSubnet,S1uBitMask,S1uVRF,S1uIntIP,S1uNetMask,"
strLine = strLine & "BPID,BPLabel,BPNet,BPBitMask,BPIntIP,BPNetMask,OMWABFName,GiV6ID,GiV6Label,GiV6Subnet,GiV6MaskBit,GiVRF,GiV6IntIP,GiV6MaskBit,GiV6ACLName,WONum,WOTask,S2Label"
objBVIVar.writeline strLine

strLine = "primaryIPAddress,hostName,BundleID,GWRR,GWName,BundleID,GaID,GaLabel,GaSubnet,GaBitMask,GiID,GiLabel,GiSubnet,GiBitMask,GnID,GnLabel,GnSubnet,GnBitMask,LiID,LiLabel,LiSubnet,LiBitMask,"
strLine = strLine & "S1uID,S1uLabel,S1uSubnet,S1uBitMask,BPID,BPLabel,BPNet,BPBitMask,GiV6ID,GiV6Label,GiV6Subnet,GiV6MaskBit,WONum,WOTask,S3Label"
objBEDVar.writeline strLine

strLine = "primaryIPAddress,hostName,XConBEID,GWName,GaID,GaLabel,GaSubnet,GaBitMask,GiID,GiLabel,GiSubnet,GiBitMask,GnID,GnLabel,GnSubnet,GnBitMask,LiID,LiLabel,LiSubnet,LiBitMask,"
strLine = strLine & "S1uID,S1uLabel,S1uSubnet,S1uBitMask,BPID,BPLabel,BPNet,BPBitMask,GiV6ID,GiV6Label,GiV6Subnet,GiV6MaskBit,WONum,WOTask,S4Label"
objBEXVar.writeline strLine

strLine = "primaryIPAddress,hostName,GWRR,GWName,BundleID,ARGPort1,GWSlot1,GWPort1,ARGPort2,GWSlot2,GWPort2,ARGPort3,GWSlot3,GWPort3,ARGPort4,GWSlot4,GWPort4,WONum,WOTask,S5Label"
objIntVar.writeline strLine

strLine = "primaryIPAddress,hostName,GWName,GaID,GaLabel,BundleID,XConBEID,GiID,GiLabel,GnID,GnLabel,LiID,LiLabel,S1uID,S1uLabel,BPID,BPLabel,GiV6ID,GiV6Label,WONum,WOTask,S6Label"
objVPNVar.writeline strLine

strLine = "primaryIPAddress,hostName,GWName,ARGAS,GaPCFIntIP,PCFAS,BGPGroup,GaLabel,GnPCFIntIP,GnLabel,LiPCFIntIP,LiLabel,BPPCFIntIP,BPLabel,GiVRF,GiPCFIntIP,BPBGPGroup,"
strLine = strLine & "GiLabel,GiV6PCFIntIP,v6BGPGroup,GiV6Label,S1uVRF,S1uPCFIntIP,S1uBGPGroup,S1uLabel,SPTName,BundleID,WONum,WOTask,S7Label"
objBGPVar.writeline strLine

'open up the template file and read it all into single variable, then close the file
set objTemplate = fso.OpenTextFile(strTemplateName, ForReading, False)
strTemplate = objTemplate.readall
objTemplate.close
set objTemplate = Nothing

'Create an array of all the file names in the CIQ folder
Set f = fso.GetFolder(strCIQFolder)
Set fc = f.Files

'Loop through the array of CIQ's and process them.
For Each f1 in fc
	if bLog then WriteLog "Opening up CIQ " & f1
	on error resume next
	Set wb = app.Workbooks.Open (f1,0,true)
	If Err.Number <> 0 Then
		WriteLog "Unable to open " & f1 & " continuing to the next one"
	else
		on error goto 0
		Set ws = wb.Worksheets(1)
		strLine = ""

		if bLog then WriteLog "Starting with Common elements"
		dictCommon.RemoveAll
		dictCommon.add "GWName",      CleanStr(ws.Cells(4,3))
		dictCommon.add "GWRR",        CleanStr(ws.Cells(4,4))
		dictCommon.add "PCFAS",       CleanStr(ws.Cells(4,7))
		dictCommon.add "ARGAS",       CleanStr(ws.Cells(4,8))
		dictCommon.add "GaLabel",     replace(CleanStr(ws.Cells(9,2))," ", "_")
		dictCommon.add "GaSubnet",    CleanStr(ws.Cells(9,3))
		dictCommon.add "GaNetMask",   CleanStr(ws.Cells(9,4))
		dictCommon.add "GaBitMask",   CleanStr(ws.Cells(9,5))
		dictCommon.add "GaID",        CleanStr(ws.Cells(9,6))
		dictCommon.add "GiLabel",     replace(CleanStr(ws.Cells(10,2))," ", "_")
		dictCommon.add "GiSubnet",    CleanStr(ws.Cells(10,3))
		dictCommon.add "GiNetMask",   CleanStr(ws.Cells(10,4))
		dictCommon.add "GiBitMask",   CleanStr(ws.Cells(10,5))
		dictCommon.add "GiInvMask",   CleanStr(dictSubnets.Item(dictCommon.item("GiBitMask")))
		dictCommon.add "GiID",        CleanStr(ws.Cells(10,6))
		dictCommon.add "GnLabel",     replace(CleanStr(ws.Cells(11,2))," ", "_")
		dictCommon.add "GnSubnet",    CleanStr(ws.Cells(11,3))
		dictCommon.add "GnNetMask",   CleanStr(ws.Cells(11,4))
		dictCommon.add "GnBitMask",   CleanStr(ws.Cells(11,5))
		dictCommon.add "GnID",        CleanStr(ws.Cells(11,6))
		dictCommon.add "LiLabel",     replace(CleanStr(ws.Cells(12,2))," ", "_")
		dictCommon.add "LiSubnet",    CleanStr(ws.Cells(12,3))
		dictCommon.add "LiNetMask",   CleanStr(ws.Cells(12,4))
		dictCommon.add "LiBitMask",   CleanStr(ws.Cells(12,5))
		dictCommon.add "LiID",        CleanStr(ws.Cells(12,6))
		dictCommon.add "S1uLabel",    replace(CleanStr(ws.Cells(13,2))," ", "_")
		dictCommon.add "S1uSubnet",   CleanStr(ws.Cells(13,3))
		dictCommon.add "S1uNetMask",  CleanStr(ws.Cells(13,4))
		dictCommon.add "S1uBitMask",  CleanStr(ws.Cells(13,5))
		dictCommon.add "S1uID",       CleanStr(ws.Cells(13,6))
		dictCommon.add "BPLabel",     replace(CleanStr(ws.Cells(14,2))," ", "_")
		dictCommon.add "BPNet",       CleanStr(ws.Cells(14,3))
		dictCommon.add "BPNetMask",   CleanStr(ws.Cells(14,4))
		dictCommon.add "BPBitMask",   CleanStr(ws.Cells(14,5))
		dictCommon.add "BPInvMask",   CleanStr(dictSubnets.Item(dictCommon.item("BPBitMask")))
		dictCommon.add "BPID",        CleanStr(ws.Cells(14,6))
		dictCommon.add "GiV6Label",   replace(CleanStr(ws.Cells(15,2))," ", "_")
		dictCommon.add "GiV6Subnet",  CleanStr(ws.Cells(15,3))
		dictCommon.add "GiV6MaskBit", CleanStr(ws.Cells(15,5))
		dictCommon.add "GiV6ID",      CleanStr(ws.Cells(15,6))
		dictCommon.add "GiLoIP",      CleanStr(ws.Cells(19,3))
		dictCommon.add "GiLoV6IP",    CleanStr(ws.Cells(34,3))
		dictCommon.add "TPNet",       CleanStr(ws.Cells(35,3))
		dictCommon.add "TPInvMask",   CleanStr(dictSubnets.Item(CleanStr(ws.Cells(35,5))))
		dictCommon.add "BundleID",    CleanStr(ws.Cells(65,3))
		dictCommon.add "XConBEID",    CleanStr(ws.Cells(66,3))
		dictCommon.add "SPTName",     CleanStr(ws.Cells(67,3))
		dictCommon.add "OMWABFName",  CleanStr(ws.Cells(68,3))
		dictCommon.add "OMWABF1",     CleanStr(ws.Cells(69,3))
		dictCommon.add "OMWABF2",     CleanStr(ws.Cells(70,3))
		dictCommon.add "OMWABF3",     CleanStr(ws.Cells(71,3))
		dictCommon.add "S1uVRF",      CleanStr(ws.Cells(72,3))
		dictCommon.add "GiVRF",       CleanStr(ws.Cells(73,3))
		dictCommon.add "GiACLName",   CleanStr(ws.Cells(75,3))
		dictCommon.add "GiACL1",      CleanStr(ws.Cells(76,3))
		dictCommon.add "GiACL2",      CleanStr(ws.Cells(77,3))
		dictCommon.add "GiV6ACLName", CleanStr(ws.Cells(78,3))
		dictCommon.add "GiV6ACL1",    CleanStr(ws.Cells(79,3))
		dictCommon.add "GiV6ACL2",    CleanStr(ws.Cells(80,3))
		dictCommon.add "ABFACLName",  CleanStr(ws.Cells(81,3))
		dictCommon.add "ABFACL1",     CleanStr(ws.Cells(82,3))
		dictCommon.add "WONum",       CleanStr(ws.Cells(83,3))
		dictCommon.add "WOTask",      CleanStr(ws.Cells(84,3))
		dictCommon.add "ABFNextHop",  CleanStr(ws.Cells(87,3))
		dictCommon.add "BGPGroup",    CleanStr(ws.Cells(88,3))
		dictCommon.add "BPBGPGroup",  CleanStr(ws.Cells(89,3))
		dictCommon.add "v6BGPGroup",  CleanStr(ws.Cells(90,3))
		dictCommon.add "S1uBGPGroup", CleanStr(ws.Cells(91,3))
		dictCommon.add "S1Label",     "ACL" & right(dictCommon.item("GWName"),1)
		dictCommon.add "S2Label",     "BVI" & right(dictCommon.item("GWName"),1)
		dictCommon.add "S3Label",     "BE" & dictCommon.item("BundleID") & right(dictCommon.item("GWName"),1)
		dictCommon.add "S4Label",     "BE" & dictCommon.item("XConBEID") & right(dictCommon.item("GWName"),1)
		dictCommon.add "S5Label",     "Ints" & right(dictCommon.item("GWName"),1)
		dictCommon.add "S6Label",     "L2VPN" & right(dictCommon.item("GWName"),1)
		dictCommon.add "S7Label",     "BGP" & right(dictCommon.item("GWName"),1)

		if bLog then WriteLog "Now working on ARG21 specifics"

		dictARG21.RemoveAll
		dictARG21.add "GaIntIP",   CleanStr(ws.Cells(42,3))
		dictARG21.add "GiIntIP",   CleanStr(ws.Cells(43,3))
		dictARG21.add "GnIntIP",   CleanStr(ws.Cells(44,3))
		dictARG21.add "LiIntIP",   CleanStr(ws.Cells(45,3))
		dictARG21.add "S1uIntIP",  CleanStr(ws.Cells(46,3))
		dictARG21.add "BPIntIP",   CleanStr(ws.Cells(47,3))
		dictARG21.add "GiV6IntIP", CleanStr(ws.Cells(48,3))
		dictARG21.add "GWSlot1",   CleanStr(ws.Cells(53,4))
		dictARG21.add "GWPort1",   CleanStr(ws.Cells(53,5))
		dictARG21.add "ARGName",   CleanStr(ws.Cells(53,7))
		dictARG21.add "ARGPort1",  CleanStr(ws.Cells(53,8))
		dictARG21.add "GWSlot2",   CleanStr(ws.Cells(54,4))
		dictARG21.add "GWPort2",   CleanStr(ws.Cells(54,5))
		dictARG21.add "ARGPort2",  CleanStr(ws.Cells(54,8))
		dictARG21.add "GWSlot3",   CleanStr(ws.Cells(55,4))
		dictARG21.add "GWPort3",   CleanStr(ws.Cells(55,5))
		dictARG21.add "ARGPort3",  CleanStr(ws.Cells(55,8))
		dictARG21.add "GWSlot4",   CleanStr(ws.Cells(56,4))
		dictARG21.add "GWPort4",   CleanStr(ws.Cells(56,5))
		dictARG21.add "ARGPort4",  CleanStr(ws.Cells(56,8))

		if bLog then WriteLog "Now working on ARG22 specifics"

		dictARG22.RemoveAll
		dictARG22.add "GaIntIP",   CleanStr(ws.Cells(42,4))
		dictARG22.add "GiIntIP",   CleanStr(ws.Cells(43,4))
		dictARG22.add "GnIntIP",   CleanStr(ws.Cells(44,4))
		dictARG22.add "LiIntIP",   CleanStr(ws.Cells(45,4))
		dictARG22.add "S1uIntIP",  CleanStr(ws.Cells(46,4))
		dictARG22.add "BPIntIP",   CleanStr(ws.Cells(47,4))
		dictARG22.add "GiV6IntIP", CleanStr(ws.Cells(48,4))
		dictARG22.add "GWSlot1",   CleanStr(ws.Cells(57,4))
		dictARG22.add "GWPort1",   CleanStr(ws.Cells(57,5))
		dictARG22.add "ARGName",   CleanStr(ws.Cells(57,7))
		dictARG22.add "ARGPort1",  CleanStr(ws.Cells(57,8))
		dictARG22.add "GWSlot2",   CleanStr(ws.Cells(58,4))
		dictARG22.add "GWPort2",   CleanStr(ws.Cells(58,5))
		dictARG22.add "ARGPort2",  CleanStr(ws.Cells(58,8))
		dictARG22.add "GWSlot3",   CleanStr(ws.Cells(59,4))
		dictARG22.add "GWPort3",   CleanStr(ws.Cells(59,5))
		dictARG22.add "ARGPort3",  CleanStr(ws.Cells(59,8))
		dictARG22.add "GWSlot4",   CleanStr(ws.Cells(60,4))
		dictARG22.add "GWPort4",   CleanStr(ws.Cells(60,5))
		dictARG22.add "ARGPort4",  CleanStr(ws.Cells(60,8))

		if bLog then WriteLog "Now working on PCF specifics"

		dictPCF.RemoveAll
		dictPCF.add "GaPCFIntIP",   CleanStr(ws.Cells(42,5))
		dictPCF.add "GiPCFIntIP",   CleanStr(ws.Cells(43,5))
		dictPCF.add "GnPCFIntIP",   CleanStr(ws.Cells(44,5))
		dictPCF.add "LiPCFIntIP",   CleanStr(ws.Cells(45,5))
		dictPCF.add "S1uPCFIntIP",  CleanStr(ws.Cells(46,5))
		dictPCF.add "BPPCFIntIP",   CleanStr(ws.Cells(47,5))
		dictPCF.add "GiV6PCFIntIP", CleanStr(ws.Cells(48,5))

		if bLog then WriteLog "Now adding a line to the AllVar CSV"
		for each strVar in dictCommon
			strLine = strLine & dictCommon.item(strVar) & ","
		next

		for each strVar in dictARG21
			strLine = strLine & dictARG21.item(strVar) & ","
		next

		for each strVar in dictARG22
			strLine = strLine & dictARG22.item(strVar) & ","
		next

		for each strVar in dictPCF
			strLine = strLine & dictPCF.item(strVar) & ","
		next

		strLine = left(strLine,len(strLine)-1)
		objFileOut.writeline strLine

		if bLog then WriteLog "Now generating configuration for ARG21"
		strConfig = strTemplate
		for each strVar in dictCommon
			strConfig = replace(strConfig, VarDelim & strVar & VarDelim,dictCommon.item(strVar),vbTextCompare)
		next

		for each strVar in dictARG21
			strConfig = replace(strConfig, VarDelim & strVar & VarDelim,dictARG21.item(strVar),vbTextCompare)
		next

		for each strVar in dictPCF
			strConfig = replace(strConfig, VarDelim & strVar & VarDelim,dictPCF.item(strVar),vbTextCompare)
		next

		strConfigFileName = strConfFolder & dictCommon.item("GWName") & "ARG21.txt"
		if bLog then WriteLog "Saving configuration to " & strConfigFileName
		set objConfig = fso.createtextfile(strConfigFileName, true)
		objConfig.write strConfig
		objConfig.close
		set objConfig = nothing

		if bLog then WriteLog "Now adding a line to the various section CSV files"

		strLine = "," & dictARG21.item("ARGName") & "," & dictCommon.item("GiACLName") & "," & dictCommon.item("GiLoIP") & "," & dictCommon.item("GiACL1") & ","
		strLine = strLine & dictCommon.item("GiACL2") & "," & dictCommon.item("GiSubnet") & "," & dictCommon.item("GiInvMask") & "," & dictCommon.item("GiV6ACLName") & ","
		strLine = strLine & dictCommon.item("GiV6ACL1") & "," & dictCommon.item("GiLoV6IP") & "," & dictCommon.item("GiV6ACL2") & "," & dictCommon.item("GiV6Subnet") & ","
		strLine = strLine & dictCommon.item("GiV6MaskBit") & "," & dictCommon.item("ABFACLName") & "," & dictCommon.item("ABFACL1") & "," & dictCommon.item("TPNet") & ","
		strLine = strLine & dictCommon.item("TPInvMask") & "," & dictCommon.item("ABFNextHop") & "," & dictCommon.item("OMWABFName") & "," & dictCommon.item("BPNet") & ","
		strLine = strLine & dictCommon.item("BPInvMask") & "," & dictCommon.item("OMWABF1") & "," & dictCommon.item("OMWABF2") & "," & dictCommon.item("OMWABF3") & ","
		strLine = strLine & dictCommon.item("WONum") & "," & dictCommon.item("WOTask") & "," & dictCommon.item("S1Label")
		objACLVar.writeline strLine

		strLine =  "," & dictARG21.item("ARGName") & "," & dictCommon.item("GaID") & "," & dictCommon.item("GWName") & "," & dictCommon.item("GaLabel")
		strLine = strLine & "," & dictCommon.item("GaSubnet") & "," & dictCommon.item("GaBitMask") & "," & dictARG21.item("GaIntIP") & "," & dictCommon.item("GaNetMask")
		strLine = strLine & "," & dictCommon.item("GiID") & "," & dictCommon.item("GiLabel") & "," & dictCommon.item("GiSubnet") & "," & dictCommon.item("GiBitMask")
		strLine = strLine & "," & dictCommon.item("GiVRF") & "," & dictARG21.item("GiIntIP") & "," & dictCommon.item("GiNetMask") & "," & dictCommon.item("GiACLName")
		strLine = strLine & "," & dictCommon.item("GnID") & "," & dictCommon.item("GnLabel") & "," & dictCommon.item("GnSubnet") & "," & dictCommon.item("GnBitMask")
		strLine = strLine & "," & dictARG21.item("GnIntIP") & "," & dictCommon.item("GnNetMask")
		strLine = strLine & "," & dictCommon.item("LiID") & "," & dictCommon.item("LiLabel") & "," & dictCommon.item("LiSubnet") & "," & dictCommon.item("LiBitMask")
		strLine = strLine & "," & dictARG21.item("LiIntIP") & "," & dictCommon.item("LiNetMask")
		strLine = strLine & "," & dictCommon.item("S1uID") & "," & dictCommon.item("S1uLabel") & "," & dictCommon.item("S1uSubnet") & "," & dictCommon.item("S1uBitMask")
		strLine = strLine & "," & dictCommon.item("S1uVRF") & "," & dictARG21.item("S1uIntIP") & "," & dictCommon.item("S1uNetMask")
		strLine = strLine & "," & dictCommon.item("BPID") & "," & dictCommon.item("BPLabel") & "," & dictCommon.item("BPNet") & "," & dictCommon.item("BPBitMask")
		strLine = strLine & "," & dictARG21.item("BPIntIP") & "," & dictCommon.item("BPNetMask") & "," & dictCommon.item("OMWABFName")
		strLine = strLine & "," & dictCommon.item("GiV6ID") & "," & dictCommon.item("GiV6Label") & "," & dictCommon.item("GiV6Subnet")
		strLine = strLine& "," & dictCommon.item("GiV6MaskBit") & "," & dictCommon.item("GiVRF") & "," & dictARG21.item("GiV6IntIP") & "," & dictCommon.item("GiV6MaskBit")
		strLine = strLine & "," & dictCommon.item("GiV6ACLName") & "," & dictCommon.item("WONum") & "," & dictCommon.item("WOTask") & "," & dictCommon.item("S2Label")
		objBVIVar.writeline strLine

		strLine = "," & dictARG21.item("ARGName") & "," & dictCommon.item("BundleID") & "," & dictCommon.item("GWRR") & "," & dictCommon.item("GWName") & "," & dictCommon.item("BundleID")
		strLine = strLine & "," & dictCommon.item("GaID") & "," & dictCommon.item("GaLabel") & "," & dictCommon.item("GaSubnet") & "," & dictCommon.item("GaBitMask") & "," & dictCommon.item("GiID")
		strLine = strLine & "," & dictCommon.item("GiLabel") & "," & dictCommon.item("GiSubnet") & "," & dictCommon.item("GiBitMask") & "," & dictCommon.item("GnID") & "," & dictCommon.item("GnLabel")
		strLine = strLine & "," & dictCommon.item("GnSubnet") & "," & dictCommon.item("GnBitMask") & "," & dictCommon.item("LiID") & "," & dictCommon.item("LiLabel") & "," & dictCommon.item("LiSubnet")
		strLine = strLine & "," & dictCommon.item("LiBitMask") & "," & dictCommon.item("S1uID") & "," & dictCommon.item("S1uLabel") & "," & dictCommon.item("S1uSubnet") & "," & dictCommon.item("S1uBitMask")
		strLine = strLine & "," & dictCommon.item("BPID") & "," & dictCommon.item("BPLabel") & "," & dictCommon.item("BPNet") & "," & dictCommon.item("BPBitMask") & "," & dictCommon.item("GiV6ID") & "," & dictCommon.item("GiV6Label")
		strLine = strLine & "," & dictCommon.item("GiV6Subnet") & "," & dictCommon.item("GiV6MaskBit") & "," & dictCommon.item("WONum") & "," & dictCommon.item("WOTask") & "," & dictCommon.item("S3Label")
		objBEDVar.writeline strLine

		strLine = "," & dictARG21.item("ARGName") & "," & dictCommon.item("XConBEID") & "," & dictCommon.item("GWName")
		strLine = strLine & "," & dictCommon.item("GaID") & "," & dictCommon.item("GaLabel") & "," & dictCommon.item("GaSubnet") & "," & dictCommon.item("GaBitMask") & "," & dictCommon.item("GiID")
		strLine = strLine & "," & dictCommon.item("GiLabel") & "," & dictCommon.item("GiSubnet") & "," & dictCommon.item("GiBitMask") & "," & dictCommon.item("GnID") & "," & dictCommon.item("GnLabel")
		strLine = strLine & "," & dictCommon.item("GnSubnet") & "," & dictCommon.item("GnBitMask") & "," & dictCommon.item("LiID") & "," & dictCommon.item("LiLabel") & "," & dictCommon.item("LiSubnet")
		strLine = strLine & "," & dictCommon.item("LiBitMask") & "," & dictCommon.item("S1uID") & "," & dictCommon.item("S1uLabel") & "," & dictCommon.item("S1uSubnet") & "," & dictCommon.item("S1uBitMask")
		strLine = strLine & "," & dictCommon.item("BPID") & "," & dictCommon.item("BPLabel") & "," & dictCommon.item("BPNet") & "," & dictCommon.item("BPBitMask") & "," & dictCommon.item("GiV6ID") & "," & dictCommon.item("GiV6Label")
		strLine = strLine & "," & dictCommon.item("GiV6Subnet") & "," & dictCommon.item("GiV6MaskBit") & "," & dictCommon.item("WONum") & "," & dictCommon.item("WOTask") & "," & dictCommon.item("S4Label")
		objBEXVar.writeline strLine

		strLine = "," & dictARG21.item("ARGName") & "," & dictCommon.item("GWRR") & "," & dictCommon.item("GWName") & "," & dictCommon.item("BundleID") & "," & dictARG21.item("ARGPort1")
		strLine = strLine & "," & dictARG21.item("GWSlot1") & "," & dictARG21.item("GWPort1") & "," & dictARG21.item("ARGPort2") & "," & dictARG21.item("GWSlot2")
		strLine = strLine & "," & dictARG21.item("GWPort2") & "," & dictARG21.item("ARGPort3") & "," & dictARG21.item("GWSlot3") & "," & dictARG21.item("GWPort3")
		strLine = strLine & "," & dictARG21.item("ARGPort4") & "," & dictARG21.item("GWSlot4") & "," & dictARG21.item("GWPort4") & "," & dictCommon.item("WONum") & "," & dictCommon.item("WOTask") & "," & dictCommon.item("S5Label")
		objIntVar.writeline strLine

		strLine = "," & dictARG21.item("ARGName") & "," & dictCommon.item("GWName") & "," & dictCommon.item("GaID") & "," & dictCommon.item("GaLabel") & "," & dictCommon.item("BundleID")
		strLine = strLine & "," & dictCommon.item("XConBEID") & "," & dictCommon.item("GiID") & "," & dictCommon.item("GiLabel") & "," & dictCommon.item("GnID") & "," & dictCommon.item("GnLabel")
		strLine = strLine & "," & dictCommon.item("LiID") & "," & dictCommon.item("LiLabel") & "," & dictCommon.item("S1uID") & "," & dictCommon.item("S1uLabel") & "," & dictCommon.item("BPID")
		strLine = strLine & "," & dictCommon.item("BPLabel") & "," & dictCommon.item("GiV6ID") & "," & dictCommon.item("GiV6Label") & "," & dictCommon.item("WONum") & "," & dictCommon.item("WOTask") & "," & dictCommon.item("S6Label")
		objVPNVar.writeline strLine

		strLine = "," & dictARG21.item("ARGName") & "," & dictCommon.item("GWName") & "," & dictCommon.item("ARGAS") & "," & dictPCF.item("GaPCFIntIP") & "," & dictCommon.item("PCFAS")
		strLine = strLine & "," & dictCommon.item("BGPGroup") & "," & dictCommon.item("GaLabel") & "," & dictPCF.item("GnPCFIntIP") & "," & dictCommon.item("GnLabel")
		strLine = strLine & "," & dictPCF.item("LiPCFIntIP") & "," & dictCommon.item("LiLabel") & "," & dictPCF.item("BPPCFIntIP") & "," & dictCommon.item("BPLabel")
		strLine = strLine & "," & dictCommon.item("GiVRF") & "," & dictPCF.item("GiPCFIntIP") & "," & dictCommon.item("BPBGPGroup") & "," & dictCommon.item("GiLabel")
		strLine = strLine & "," & dictPCF.item("GiV6PCFIntIP") & "," & dictCommon.item("v6BGPGroup") & "," & dictCommon.item("GiV6Label") & "," & dictCommon.item("S1uVRF")
		strLine = strLine & "," & dictPCF.item("S1uPCFIntIP") & "," & dictCommon.item("S1uBGPGroup") & "," & dictCommon.item("S1uLabel") & "," & dictCommon.item("SPTName")
		strLine = strLine & "," & dictCommon.item("BundleID") & "," & dictCommon.item("WONum") & "," & dictCommon.item("WOTask") & "," & dictCommon.item("S7Label")
		objBGPVar.writeline strLine

		if bLog then WriteLog "Now generating configuration for ARG22"
		strConfig = strTemplate
		for each strVar in dictCommon
			strConfig = replace(strConfig, VarDelim & strVar & VarDelim,dictCommon.item(strVar),vbTextCompare)
		next

		for each strVar in dictARG22
			strConfig = replace(strConfig, VarDelim & strVar & VarDelim,dictARG22.item(strVar),vbTextCompare)
		next

		for each strVar in dictPCF
			strConfig = replace(strConfig, VarDelim & strVar & VarDelim,dictPCF.item(strVar),vbTextCompare)
		next

		strConfigFileName = strConfFolder & dictCommon.item("GWName") & "ARG22.txt"
		if bLog then WriteLog "Saving configuration to " & strConfigFileName
		set objConfig = fso.createtextfile(strConfigFileName, true)
		objConfig.write strConfig
		objConfig.close
		set objConfig = nothing

		if bLog then WriteLog "Now adding a line to the various section CSV files"
		strLine = "," & dictARG22.item("ARGName") & "," & dictCommon.item("GiACLName") & "," & dictCommon.item("GiLoIP") & "," & dictCommon.item("GiACL1") & ","
		strLine = strLine & dictCommon.item("GiACL2") & "," & dictCommon.item("GiSubnet") & "," & dictCommon.item("GiInvMask") & "," & dictCommon.item("GiV6ACLName") & ","
		strLine = strLine & dictCommon.item("GiV6ACL1") & "," & dictCommon.item("GiLoV6IP") & "," & dictCommon.item("GiV6ACL2") & "," & dictCommon.item("GiV6Subnet") & ","
		strLine = strLine & dictCommon.item("GiV6MaskBit") & "," & dictCommon.item("ABFACLName") & "," & dictCommon.item("ABFACL1") & "," & dictCommon.item("TPNet") & ","
		strLine = strLine & dictCommon.item("TPInvMask") & "," & dictCommon.item("ABFNextHop") & "," & dictCommon.item("OMWABFName") & "," & dictCommon.item("BPNet") & ","
		strLine = strLine & dictCommon.item("BPInvMask") & "," & dictCommon.item("OMWABF1") & "," & dictCommon.item("OMWABF2") & "," & dictCommon.item("OMWABF3") & ","
		strLine = strLine & dictCommon.item("WONum") & "," & dictCommon.item("WOTask") & "," & dictCommon.item("S1Label")
		objACLVar.writeline strLine

		strLine =  "," & dictARG22.item("ARGName") & "," & dictCommon.item("GaID") & "," & dictCommon.item("GWName") & "," & dictCommon.item("GaLabel")
		strLine = strLine & "," & dictCommon.item("GaSubnet") & "," & dictCommon.item("GaBitMask") & "," & dictARG22.item("GaIntIP") & "," & dictCommon.item("GaNetMask")
		strLine = strLine & "," & dictCommon.item("GiID") & "," & dictCommon.item("GiLabel") & "," & dictCommon.item("GiSubnet") & "," & dictCommon.item("GiBitMask")
		strLine = strLine & "," & dictCommon.item("GiVRF") & "," & dictARG22.item("GiIntIP") & "," & dictCommon.item("GiNetMask") & "," & dictCommon.item("GiACLName")
		strLine = strLine & "," & dictCommon.item("GnID") & "," & dictCommon.item("GnLabel") & "," & dictCommon.item("GnSubnet") & "," & dictCommon.item("GnBitMask")
		strLine = strLine & "," & dictARG22.item("GnIntIP") & "," & dictCommon.item("GnNetMask")
		strLine = strLine & "," & dictCommon.item("LiID") & "," & dictCommon.item("LiLabel") & "," & dictCommon.item("LiSubnet") & "," & dictCommon.item("LiBitMask")
		strLine = strLine & "," & dictARG22.item("LiIntIP") & "," & dictCommon.item("LiNetMask")
		strLine = strLine & "," & dictCommon.item("S1uID") & "," & dictCommon.item("S1uLabel") & "," & dictCommon.item("S1uSubnet") & "," & dictCommon.item("S1uBitMask")
		strLine = strLine & "," & dictCommon.item("S1uVRF") & "," & dictARG22.item("S1uIntIP") & "," & dictCommon.item("S1uNetMask")
		strLine = strLine & "," & dictCommon.item("BPID") & "," & dictCommon.item("BPLabel") & "," & dictCommon.item("BPNet") & "," & dictCommon.item("BPBitMask")
		strLine = strLine & "," & dictARG22.item("BPIntIP") & "," & dictCommon.item("BPNetMask") & "," & dictCommon.item("OMWABFName")
		strLine = strLine & "," & dictCommon.item("GiV6ID") & "," & dictCommon.item("GiV6Label") & "," & dictCommon.item("GiV6Subnet")
		strLine = strLine& "," & dictCommon.item("GiV6MaskBit") & "," & dictCommon.item("GiVRF") & "," & dictARG22.item("GiV6IntIP") & "," & dictCommon.item("GiV6MaskBit")
		strLine = strLine & "," & dictCommon.item("GiV6ACLName") & "," & dictCommon.item("WONum") & "," & dictCommon.item("WOTask") & "," & dictCommon.item("S2Label")
		objBVIVar.writeline strLine

		strLine = "," & dictARG22.item("ARGName") & "," & dictCommon.item("BundleID") & "," & dictCommon.item("GWRR") & "," & dictCommon.item("GWName") & "," & dictCommon.item("BundleID")
		strLine = strLine & "," & dictCommon.item("GaID") & "," & dictCommon.item("GaLabel") & "," & dictCommon.item("GaSubnet") & "," & dictCommon.item("GaBitMask") & "," & dictCommon.item("GiID")
		strLine = strLine & "," & dictCommon.item("GiLabel") & "," & dictCommon.item("GiSubnet") & "," & dictCommon.item("GiBitMask") & "," & dictCommon.item("GnID") & "," & dictCommon.item("GnLabel")
		strLine = strLine & "," & dictCommon.item("GnSubnet") & "," & dictCommon.item("GnBitMask") & "," & dictCommon.item("LiID") & "," & dictCommon.item("LiLabel") & "," & dictCommon.item("LiSubnet")
		strLine = strLine & "," & dictCommon.item("LiBitMask") & "," & dictCommon.item("S1uID") & "," & dictCommon.item("S1uLabel") & "," & dictCommon.item("S1uSubnet") & "," & dictCommon.item("S1uBitMask")
		strLine = strLine & "," & dictCommon.item("BPID") & "," & dictCommon.item("BPLabel") & "," & dictCommon.item("BPNet") & "," & dictCommon.item("BPBitMask") & "," & dictCommon.item("GiV6ID") & "," & dictCommon.item("GiV6Label")
		strLine = strLine & "," & dictCommon.item("GiV6Subnet") & "," & dictCommon.item("GiV6MaskBit") & "," & dictCommon.item("WONum") & "," & dictCommon.item("WOTask") & "," & dictCommon.item("S3Label")
		objBEDVar.writeline strLine

		strLine = "," & dictARG22.item("ARGName") & "," & dictCommon.item("XConBEID") & "," & dictCommon.item("GWName")
		strLine = strLine & "," & dictCommon.item("GaID") & "," & dictCommon.item("GaLabel") & "," & dictCommon.item("GaSubnet") & "," & dictCommon.item("GaBitMask") & "," & dictCommon.item("GiID")
		strLine = strLine & "," & dictCommon.item("GiLabel") & "," & dictCommon.item("GiSubnet") & "," & dictCommon.item("GiBitMask") & "," & dictCommon.item("GnID") & "," & dictCommon.item("GnLabel")
		strLine = strLine & "," & dictCommon.item("GnSubnet") & "," & dictCommon.item("GnBitMask") & "," & dictCommon.item("LiID") & "," & dictCommon.item("LiLabel") & "," & dictCommon.item("LiSubnet")
		strLine = strLine & "," & dictCommon.item("LiBitMask") & "," & dictCommon.item("S1uID") & "," & dictCommon.item("S1uLabel") & "," & dictCommon.item("S1uSubnet") & "," & dictCommon.item("S1uBitMask")
		strLine = strLine & "," & dictCommon.item("BPID") & "," & dictCommon.item("BPLabel") & "," & dictCommon.item("BPNet") & "," & dictCommon.item("BPBitMask") & "," & dictCommon.item("GiV6ID") & "," & dictCommon.item("GiV6Label")
		strLine = strLine & "," & dictCommon.item("GiV6Subnet") & "," & dictCommon.item("GiV6MaskBit") & "," & dictCommon.item("WONum") & "," & dictCommon.item("WOTask") & "," & dictCommon.item("S4Label")
		objBEXVar.writeline strLine

		strLine = "," & dictARG22.item("ARGName") & "," & dictCommon.item("GWRR") & "," & dictCommon.item("GWName") & "," & dictCommon.item("BundleID") & "," & dictARG22.item("ARGPort1")
		strLine = strLine & "," & dictARG22.item("GWSlot1") & "," & dictARG22.item("GWPort1") & "," & dictARG22.item("ARGPort2") & "," & dictARG22.item("GWSlot2")
		strLine = strLine & "," & dictARG22.item("GWPort2") & "," & dictARG22.item("ARGPort3") & "," & dictARG22.item("GWSlot3") & "," & dictARG22.item("GWPort3")
		strLine = strLine & "," & dictARG22.item("ARGPort4") & "," & dictARG22.item("GWSlot4") & "," & dictARG22.item("GWPort4") & "," & dictCommon.item("WONum") & "," & dictCommon.item("WOTask") & "," & dictCommon.item("S5Label")
		objIntVar.writeline strLine

		strLine = "," & dictARG22.item("ARGName") & "," & dictCommon.item("GWName") & "," & dictCommon.item("GaID") & "," & dictCommon.item("GaLabel") & "," & dictCommon.item("BundleID")
		strLine = strLine & "," & dictCommon.item("XConBEID") & "," & dictCommon.item("GiID") & "," & dictCommon.item("GiLabel") & "," & dictCommon.item("GnID") & "," & dictCommon.item("GnLabel")
		strLine = strLine & "," & dictCommon.item("LiID") & "," & dictCommon.item("LiLabel") & "," & dictCommon.item("S1uID") & "," & dictCommon.item("S1uLabel") & "," & dictCommon.item("BPID")
		strLine = strLine & "," & dictCommon.item("BPLabel") & "," & dictCommon.item("GiV6ID") & "," & dictCommon.item("GiV6Label") & "," & dictCommon.item("WONum") & "," & dictCommon.item("WOTask") & "," & dictCommon.item("S6Label")
		objVPNVar.writeline strLine

		strLine = "," & dictARG22.item("ARGName") & "," & dictCommon.item("GWName") & "," & dictCommon.item("ARGAS") & "," & dictPCF.item("GaPCFIntIP") & "," & dictCommon.item("PCFAS")
		strLine = strLine & "," & dictCommon.item("BGPGroup") & "," & dictCommon.item("GaLabel") & "," & dictPCF.item("GnPCFIntIP") & "," & dictCommon.item("GnLabel")
		strLine = strLine & "," & dictPCF.item("LiPCFIntIP") & "," & dictCommon.item("LiLabel") & "," & dictPCF.item("BPPCFIntIP") & "," & dictCommon.item("BPLabel")
		strLine = strLine & "," & dictCommon.item("GiVRF") & "," & dictPCF.item("GiPCFIntIP") & "," & dictCommon.item("BPBGPGroup") & "," & dictCommon.item("GiLabel")
		strLine = strLine & "," & dictPCF.item("GiV6PCFIntIP") & "," & dictCommon.item("v6BGPGroup") & "," & dictCommon.item("GiV6Label") & "," & dictCommon.item("S1uVRF")
		strLine = strLine & "," & dictPCF.item("S1uPCFIntIP") & "," & dictCommon.item("S1uBGPGroup") & "," & dictCommon.item("S1uLabel") & "," & dictCommon.item("SPTName")
		strLine = strLine & "," & dictCommon.item("BundleID") & "," & dictCommon.item("WONum") & "," & dictCommon.item("WOTask") & "," & dictCommon.item("S7Label")
		objBGPVar.writeline strLine

		if bLog then WriteLog "Closing this CIQ and moving on to the next one"
		wb.close False
	end if
Next

app.quit ' Close Excel

WriteLog "Done. CSV saved to " & strOutFileName
WriteLog "Configuration files saved to " & strConfFolder

'Cleanup, close out files, release resources, etc.
if bFile then
	objLogOut.close
	set objLogOut = nothing
end if
objFileOut.close
objFileOut.close
objACLVar.close
objBVIVar.close
objBEDVar.close
objBEXVar.close
objIntVar.close
objVPNVar.close
objBGPVar.close

set objFileOut  = nothing
Set fc          = nothing
Set f           = nothing
Set fso         = nothing
Set ws          = Nothing
Set wb          = Nothing
Set app         = Nothing
set dictSubnets = Nothing
set dictCommon  = Nothing
set dictARG21   = Nothing
set dictARG22   = Nothing
set dictPCF     = Nothing
set objFileOut  = Nothing
set objACLVar   = Nothing
set objBVIVar   = Nothing
set objBEDVar   = Nothing
set objBEXVar   = Nothing
set objIntVar   = Nothing
set objVPNVar   = Nothing
set objBGPVar   = Nothing


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

strHelpMsg =              " Usage: " & strScriptName & " folder=inpath log file hide" & vbcrlf
strHelpMsg = strHelpMsg & "        " & strScriptName & ": The name of this script" & vbcrlf
strHelpMsg = strHelpMsg & "        folder: The complete path of the folder where everything is stored" & vbcrlf
strHelpMsg = strHelpMsg & "        log: flag to indicate detailed log, summary is default" & vbcrlf
strHelpMsg = strHelpMsg & "        file: flag to log everything to a log file, default is no log file" & vbcrlf
strHelpMsg = strHelpMsg & "        hide: Keep Excel hidden during processing, only applies to the instanse the script starts. Default is to show Excel." & vbcrlf
strHelpMsg = strHelpMsg & vbcrlf
strHelpMsg = strHelpMsg & "  All arguments are optional and the order does not matter." & vbcrlf
strHelpMsg = strHelpMsg & "  If you do not provide path you will be prompted for one" & vbcrlf
strHelpMsg = strHelpMsg & "  default folder will be suggested while prompting for the path" & vbcrlf
strHelpMsg = strHelpMsg & vbcrlf
strHelpMsg = strHelpMsg & "  NOTE: CIQ's should be in a sub folder of the folder provided" & vbcrlf
strHelpMsg = strHelpMsg & "  Configurations will be saved in their own sub folder of the folder provided" & vbcrlf
strHelpMsg = strHelpMsg & "  Template file should be in the root of the folder provided" & vbcrlf
strHelpMsg = strHelpMsg & "  CSV file containing all the variables used will be saved in the root of the folder provided" & vbcrlf
strHelpMsg = strHelpMsg & "  if you provide ""help"" or ""?"" as an argument this message will be printed and script will exit" & vbcrlf
strHelpMsg = strHelpMsg & vbcrlf
strHelpMsg = strHelpMsg & "  Example: cscript " & strScriptName & " folder=C:\2015ASR5500Deploy log file" & vbcrlf

 ' Check if the script runs in CSCRIPT.EXE
 If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
 ' If so, use StdIn and StdOut
 WScript.StdOut.Writeline strHelpMsg
Else
 Msgbox strHelpMsg', "Help message"
end if


end sub
