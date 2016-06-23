Option Explicit
'-------------------------------------------------------------------------------------------------'
' This script takes two input files, a text template file and csv with variable values            '
' it will replace variables in in the template file with values from the csv file                 '
' and it will report any discrpency it find                                                       '
'                                                                                                 '
' Author: Siggi Bjarnason                                                                         '
' Date: 02/23/2015                                                                                '
' Usage: parser infile=inpath log file                                                            '
'        parser: The name of this script                                                          '
'        infile: The complete path of one of the input file, or folder                            '
'        log: flag to indicate detailed log, summary log is default                               '
'        file: flag to log everything to a log file, default is no log file                       '
'                                                                                                 '
'  All arguments are optional and the order does not matter                                       '
'  if you do not provide infile you will be prompted for one                                      '
'  the assumption is that the text file and csv file is named the same                            '
'  one with csv extension the other with txt extension.                                           '
'  If they aren't named like that you will be prompted for the full path of the other file        '
'  If you provide just a folder, all files in that folder will be processed                       '
'  if a matching set can be found. 
'  Configurations will be saved in their own sub folder                                           '
'                                                                                                 '
'  if you provide "help" or "?" as an argument help message will be printed and script will exit  '
'                                                                                                 '
'  Example: cscript HPNAValidate.vbs infile=C:\HPNATest\testfile.txt log file                     '
'                                                                                                 '
'-------------------------------------------------------------------------------------------------'

'User definable Constants

' Sets the character that deliminates the variables in the configuration file, HPNA requires $
Const VarDelim            = "$"
Const VarDelimReEx        = "\$"

' sets various file name defaults.
const DefConfigFolderName = "config\"


'Nothing below here is user configurable proceed at your own risk.

'Variable declaration
  Dim strConfFolder, strTemplateName, strScriptFullName, strLogFileName, strConfigFileName, strInFileName, bFolder, strHostNameA
  Dim strLine, strConfig, strTemplate, iArg, strParts, strScriptName, strInput, strArgParts, strHeaderParts, bCont, iHostNo
  Dim strHostIP, strHostName, strRegPatern, strLineParts, strInFile, strVariable, strPartName, iExtLoc, iPatEndLoc, iloc, iCount
  Dim objLogOut, objTemplate, objConfig, objFileIn, dictHostNames, bLog, bFile, re, Matches, Match, x, fso, iLineNo, strVar
  dim dictNotInTemplate, dictNotInCSV, f, fc, f1, dictFiles, strHelpMsg
  
  const ForReading    = 1
  const ForWriting    = 2
  const ForAppending  = 8

	bLog              = false
	bFile             = false
	bFolder           = false
	strScriptFullName = wscript.ScriptFullName
	strParts          = split (strScriptFullName, "\")
	strScriptName     = strParts(ubound(strParts))

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
			strHelpMsg = strHelpMsg & " it will work but it's a sub optimal way or working, you'll get no progress and less options"
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
				wscript.echo "found an argument for infile: " & strInFile
			case "log"
				bLog = true
				wscript.echo "detailed logging enabled"
			case "file"
				bFile = true
				wscript.echo "log file enabled"
			case "help","?"
				PrintHelp
				wscript.quit
			case "abort","exit"
				wscript.quit
			case else
				wscript.echo "Invalid argument " & strArgParts(iArg) & ". try help for valid options"
				wscript.quit
			end select
	next

	' Creating a File System Object to interact with the File System
	Set fso = CreateObject("Scripting.FileSystemObject")

	'if log file is enabled create the log file
	strLogFileName = Mid(strScriptFullName, 1, InStrRev(strScriptFullName, ".")) & "log"
	if bFile then
		wscript.echo "Log file " & strLogFileName
		Set objLogOut = fso.OpenTextFile(strLogFileName, ForAppending, True)
	end if
	
	writelog "starting " & strScriptName &" at " & now

	if bLog then WriteLog "Validating input"

	'Validate input
	if strInFile = "" then
		strInFile = UserInput("No infile was specified. Please provide complete path to the input file or folder" & vbcrlf & "Your Input:")
		if strInFile = "" then
			writelog "No input received, aborting"
			wscript.quit
		end if
	end if

	'Attempt to add a .txt extension if infile is not valid
	'If that's not valid check to see if it is a valid folder
	if not fso.FileExists(strInFile) then
		strInFileName = strInFile & ".txt"
		if not fso.FileExists(strInFileName) then
			if fso.FolderExists(strInFile) then
				bFolder = true
				writelog "Found a valid folder path " & strInFile
				if right(strInFile,1)<> "\" then
					strInFile = strInFile & "\"
				end if
			end if
		else
				bFolder = false
				strInFile = strInFileName
		end if
	end if

	if not bFolder then
		while not fso.FileExists(strInFile)
			strInFile = UserInput("input file " & strInFile & " is not valid. Please provide new one or leave blank to abort:")
			if strInFile = "" then
				writelog "No input provided, aborting"
				wscript.quit
			end if
		wend

		select case right(strInFile,4)
			case ".txt"
				if bLog then WriteLog "Template file is " & strInFile
				strTemplateName   = strInFile
				strInFileName = Mid(strInFile, 1, InStrRev(strInFile, ".")) & "csv"
				while not fso.FileExists(strInFileName)
					strInFileName = UserInput("CSV file " & strInFileName & " is not valid. Please provide new one or leave blank to abort:")
					if strInFileName = "" then
						writelog "No input provided, aborting"
						wscript.quit
					end if
				wend
				if bLog then WriteLog "CSV file is " & strInFileName
			case ".csv"
				if bLog then WriteLog "CSV file is " & strInFile
				strInFileName   = strInFile
				strTemplateName = Mid(strInFile, 1, InStrRev(strInFile, ".")) & "txt"
				while not fso.FileExists(strTemplateName)
					strTemplateName = UserInput("template file " & strTemplateName & " is not valid. Please provide new one or leave blank to abort:")
					if strTemplateName = "" then
						writelog "No input provided, aborting"
						wscript.quit
					end if
				wend
				if bLog then WriteLog "template file is " & strTemplateName
		end select
		strConfFolder = Mid(strTemplateName, 1, InStrRev(strTemplateName, "\")) & DefConfigFolderName
	else
		strConfFolder = strInFile & DefConfigFolderName
	end if


	' Log the full path names for all the relevant files and folders
	WriteLog "saving configuations to " & strConfFolder

	'if Configuration folder doesn't exists create it
	if not fso.FolderExists(strConfFolder) then
		fso.CreateFolder(strConfFolder)
		if bLog then WriteLog strConfFolder & " did not exists so I created it"
	end if

	'Initializing Dictionaries, aka ordered arrays
	set dictHostNames     = CreateObject("Scripting.Dictionary")
	set dictNotInTemplate = CreateObject("Scripting.Dictionary")
	set dictNotInCSV      = CreateObject("Scripting.Dictionary")
	set dictFiles         = CreateObject("Scripting.Dictionary")

	if bFolder then
		if bLog then WriteLog "starting folder operations"
		'Create an array of all the file names in the provided folder
		Set f = fso.GetFolder(strInFile)
		Set fc = f.Files

		'Loop through the array of files and process them.
		For Each f1 in fc
			iExtLoc = InStrRev(f1, ".")
			iPatEndLoc = InStrRev(f1, "\") + 1
			strPartName = Mid(f1, iPatEndLoc, iExtLoc-iPatEndLoc)
			if bLog then WriteLog "evaluating " & f1
			select case right(f1,4)
				case ".txt"
					if not dictFiles.exists(strPartName & "txt") then
						dictFiles.add strPartName & "txt", "t-template"
						strInFileName = Mid(f1, 1, InStrRev(f1, ".")) & "csv"
						if fso.FileExists(strInFileName) then
							dictFiles.add strPartName & "csv", "t-csv"
							ProcessFiles strInFileName, f1
						else
							if bLog then writelog "Unable to find a match for " & f1
						end if
					end if
				case ".csv"
					if not dictFiles.exists(strPartName & "csv") then
						dictFiles.add strPartName & "csv", "c-csv"
						strTemplateName = Mid(f1, 1, InStrRev(f1, ".")) & "txt"
						if fso.FileExists(strTemplateName) then
							dictFiles.add strPartName & "txt", "c-template"
							ProcessFiles f1, strTemplateName
						else
							if bLog then writelog "Unable to find a match for " & f1
						end if
					end if
				case else
					if bLog then WriteLog "neither txt nor csv file"
				end select
		next
		if bLog then
			writelog "processed the following files in folder mode:"
			for each strVar in dictFiles
				writelog strVar & "::" & dictFiles.item(strVar) & ";"
			next
		end if
	else
		ProcessFiles strInFileName, strTemplateName
	end if

	WriteLog "configuations saved to " & strConfFolder
	WriteLog "Done. "

	iCount = dictNotInCSV.count
	if iCount > 0 then
		writelog iCount & " variables in the template were not replaced"
		writelog "as they didn't have a matching variable in the CSV"
		for each strVar in dictNotInCSV
			writelog strVar & " in " & dictNotInCSV.item(strVar)
		next
	end if

	iCount = dictNotInTemplate.count
	if iCount > 0 then
		writelog iCount & " variables in the CSV file weren't used "
		writelog "as they didn't have a matching variable in the template"
		for each strVar in dictNotInTemplate
			writelog strVar & " in " & dictNotInTemplate.item(strVar)
		next
	end if

	'Cleanup, close out files, release resources, etc.
	if bFile then
		objLogOut.close
		set objLogOut = Nothing
	end if

	Set fso = Nothing
	set dictHostNames  = Nothing
	set Matches = Nothing


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


sub PrintHelp
'-------------------------------------------------------------------------------------------------'
' sub PrintHelp                                                                                   '
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
	strHelpMsg = strHelpMsg & "        infile: The complete path of the folder where everything is stored" & vbcrlf
	strHelpMsg = strHelpMsg & "                or the complete path of one of the files you want to process" & vbcrlf
	strHelpMsg = strHelpMsg & "        log: flag to indicate detailed log, summary is default" & vbcrlf
	strHelpMsg = strHelpMsg & "        file: flag to copy all output to a log file, default is no log file" & vbcrlf
	strHelpMsg = strHelpMsg & vbcrlf
	strHelpMsg = strHelpMsg & "  All arguments are optional and the order does not matter." & vbcrlf
	strHelpMsg = strHelpMsg & "  If you do not provide a value for infile you will be prompted for one" & vbcrlf
	strHelpMsg = strHelpMsg & "  the assumption is that the text file and csv file is named the same" & vbcrlf
	strHelpMsg = strHelpMsg & "  one with csv extension the other with txt extension." & vbcrlf
	If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
		strHelpMsg = strHelpMsg & "  If they aren't named like that you will be prompted for the full path" & vbcrlf
		strHelpMsg = strHelpMsg & "  of the other file" & vbcrlf
	end if
	strHelpMsg = strHelpMsg & "  You can also just provide a folder path and the script will search that " & vbcrlf
	strHelpMsg = strHelpMsg & "  folder matchin CSV and TXT files to process." & vbcrlf
	strHelpMsg = strHelpMsg & "  Configurations will be saved in their own sub folder of the folder provided " & vbcrlf
	strHelpMsg = strHelpMsg & "  or subfolder of the folder with the template file." & vbcrlf
	strHelpMsg = strHelpMsg & vbcrlf
	strHelpMsg = strHelpMsg & "  If you provide ""help"" or ""?"" as an argument this message will be printed " & vbcrlf
	strHelpMsg = strHelpMsg & "  and script will exit" & vbcrlf
	strHelpMsg = strHelpMsg & vbcrlf
	strHelpMsg = strHelpMsg & "  Example: cscript " & strScriptName & " infile=C:\HPNATest\testfile.txt log file" & vbcrlf
	strHelpMsg = strHelpMsg & "           cscript " & strScriptName & " infile=C:\HPNATest\ log file" & vbcrlf
	
	 ' Check if the script runs in CSCRIPT.EXE
	 If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
	 ' If so, use StdIn and StdOut
	 WScript.StdOut.Writeline strHelpMsg
	Else
	 Msgbox strHelpMsg', "Help message"
	end if

end sub

function ProcessFiles (strCSVname, strTName)
'-------------------------------------------------------------------------------------------------'
' sub ProcessFiles (strCSVname, strTName)                                                         '
'                                                                                                 '
' This sub takes two file names as import and processes them by substituting variables in a       '
' text file template file with values in a csv file and reporting out any discrepenancies         '
'-------------------------------------------------------------------------------------------------'
	iExtLoc = InStrRev(strTName, ".")
	iPatEndLoc = InStrRev(strTName, "\") + 1
	strPartName = Mid(strTName, iPatEndLoc, iExtLoc-iPatEndLoc)
	dictHostNames.RemoveAll
	writelog "Processing CVS file " & strCSVName & " and template " & strTname
	if bLog then writelog "Part Name : " & strPartName
	if bLog then writelog "Reading in the template file to memory"
	'open up the template file and read it all into single variable, then close the file
	set objTemplate = fso.OpenTextFile(strTName, ForReading, False)
	strTemplate = objTemplate.readall
	objTemplate.close
	set objTemplate = Nothing

	'open up the csv file
	if bLog then writelog "Opening up the CSV and starting processing"
	Set objFileIn = fso.opentextfile(strCSVname, ForReading, False)

	if bLog then writelog "reading in header"
	'read in the first header line and split it into an array
	strLine = objFileIn.readline
	strHeaderParts = split (Trim(strLine), ",")
	bCont = True
	iLineNo = 1

	if bLog then writelog "starting the loop"
 ' loop through each line in the CSV
	While not objFileIn.atendofstream
		strHostName = ""
		strLine = objFileIn.readline
		strLineParts = split(Trim(strLine), ",")
		strConfig = strTemplate
		for x=0 to ubound(strHeaderParts)
			select case lcase (strHeaderParts(x))
				case "primaryipaddress"
					strHostIP = strLineParts(x)
				case "hostname"
					strHostName = strLineParts(x)
					strhostNameA = strHostName
					if dictHostNames.exists(strHostName) then
						do
							iHostNo = dictHostNames.item(strHostNameA)
							strhostNameA = strHostName & "-" & iHostNo
						loop while dictHostNames.exists(strHostNameA)
						dictHostNames.add strHostNameA, iHostNo+1
						strHostName = strHostNameA
						writelog "processing " & strHostName
					else
						bCont = True
						dictHostNames.add strHostName,"2"
						writelog "processing " & strHostName
					end if
				case else
					strVariable =  VarDelim & strHeaderParts(x) & VarDelim
					iloc = instr(1,strTemplate,strVariable,vbTextCompare)
					if iloc = 0 then
						if not dictNotInTemplate.exists(strHeaderParts(x)) then
							dictNotInTemplate.add strHeaderParts(x), strCSVname
						end if
						'writelog strHeaderParts(x) & " not found in template file"
					end if
					strConfig = replace(strConfig, strVariable,strLineParts(x),vbTextCompare)
			end select
		next

		if bCont then
			strRegPatern = VarDelimReEx & ".+" & VarDelimReEx
			set re = new regexp
			re.pattern = strRegPatern
			re.IgnoreCase = true
	    	re.Global = True
	    	Set Matches = re.Execute(strConfig)

		    if Matches.count > 0 then
		    	For Each Match in Matches
'		    		writelog Match.value
						if not dictNotInCSV.exists(Match.value) then
							dictNotInCSV.add Match.value, strTName
						end if
		    	next
		    end if

		    if strHostName = "" then
		    	strHostName = "Line" & iLineNo
		    end if

			strConfigFileName = strConfFolder & strHostName & "-" & strPartName & ".txt"
			if bLog then WriteLog "Saving configuration to "
			if bLog then WriteLog strConfigFileName
			set objConfig = fso.createtextfile(strConfigFileName, true)
			objConfig.write strConfig
			objConfig.close
			set objConfig = nothing
			iLineNo=iLineNo + 1
		end if
	wend

	objFileIn.close
	set objFileIn = Nothing

end function