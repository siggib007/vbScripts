Option Explicit

'|----------------------------------------------------------------------------------------------------------|
'|  This script will log into each ARG in a specified spreadsheet and confirm ACLs are up to standards.     |
'|  If deviations are found, will generate configuration files for HPNA or manual MOP.                      |
'|                                                                                                          |
'|  Author: Siggi Bjarnason                                                                                 |
'|  Authored: 04/10/2017                                                                                    |
'|  Copyright: Siggi Bjarnason 2017                                                                         |
'|----------------------------------------------------------------------------------------------------------|

 ' User Spefified values, specify values here per your needs

  Const AutoCloseResults = True
  Const ExcelVisible     = True
  const Timeout    = 5 ' Timeout in seconds for each command, if expected results aren't received withing this time, the script moves on.
  const MaxError   = 5 ' If connection error occurs how often to retry

 ' Nothing below here is user values, proceed with caution and at your own risk.
 ' File handling constants
  const ForReading    = 1
  const ForWriting    = 2
  const ForAppending  = 8


  ' button parameter options
  Const ICON_STOP = 16                 ' display the ERROR/STOP icon.
  Const ICON_QUESTION = 32             ' display the '?' icon
  Const ICON_WARN = 48                 ' display a '!' icon.
  Const ICON_INFO= 64                  ' displays "info" icon.
  Const BUTTON_OK = 0                  ' OK button only
  Const BUTTON_CANCEL = 1              ' OK and Cancel buttons
  Const BUTTON_ABORTRETRYIGNORE = 2    ' Abort, Retry, and Ignore buttons
  Const BUTTON_YESNOCANCEL = 3         ' Yes, No, and Cancel buttons
  Const BUTTON_YESNO = 4               ' Yes and No buttons
  Const BUTTON_RETRYCANCEL = 5         ' Retry and Cancel buttons

  Const DEFBUTTON1 = 0        ' First button is default
  Const DEFBUTTON2 = 256      ' Second button is default
  Const DEFBUTTON3 = 512      ' Third button is default

  ' Possible MessageBox() return values
  Const IDOK = 1              ' OK button clicked
  Const IDCANCEL = 2          ' Cancel button clicked
  Const IDABORT = 3           ' Abort button clicked
  Const IDRETRY = 4           ' Retry button clicked
  Const IDIGNORE = 5          ' Ignore button clicked
  Const IDYES = 6             ' Yes button clicked
  Const IDNO = 7              ' No button clicked

  Dim app, objShell, dictNames, dictVars, dictACLs, dictACLVarNames, dictChange, dictDevAffected
  Dim strOutPath, strOutFile, strLogFile, strAsIsOutPath, wsNames, wsVars, wsACL, wbin, wbNameIn
  Dim fso, objFileOut, objLogOut, objACLGen, objChangeOut, strMOPPath

Sub main
  Dim bComp, bRange, bNewChange, dkey, dkeys, errcode, errmsg
  Dim iNameRow, iVarRow, iACLRow, iNameCol, iACLCol, iVarCol, iStartPos, iStopPos, iHostCol
  Dim iGSeq, iASeq, iLastLine, iError, iResult, iIPCol, iChangeID
  Dim strACLVar, strTempOut, strACLName, strACLID, strACLNameVar
  Dim strHostname, strIPAddr, strResultParts, strGenOutPath
  Dim strNotes, strNoMatch, strMissing, strErr, strIPVer, strChange

  ' comparison constants
  Const vbBinaryCompare = 0 'Perform a binary comparison, i.e case sensitive
  Const vbTextCompare = 1 'Perform a textual comparison, i.e. case insensitive

  ' Display File Open dialog so that the user can selelct the input workbook.
  wbNameIn = crt.Dialog.FileOpenDialog("Please select OMW ACL Standard Spreadsheet", "Open", "", "Excel Files (*.xlsx)|*.xlsx||")
  if wbNameIn = "" then
    msgbox "No file provided, exiting"
    exit Sub
  end if
  ' Hook to the filesystem
  Set fso = CreateObject("Scripting.FileSystemObject")

  ' Doublecheck the input workbook actually does exists
  if not fso.FileExists(wbNameIn) Then
    msgbox "Input file " & wbNameIn & " not found, exiting"
    set fso = nothing
    exit sub
  end if

  ' Parse out the path only from the workbook file name, make sure the path ends in a \
  strOutPath = left (wbNameIn, InStrRev (wbNameIn,"\"))
  if right(strOutPath,1)<>"\" then
    strOutPath = strOutPath & "\"
  end if

  ' Create file names for log file and results file based on the input file
  strOutFile = left (wbNameIn, InStrRev (wbNameIn,".")-1)&"-Results.csv"
  strlogFile = left (wbNameIn, InStrRev (wbNameIn,".")-1)&"-log.txt"

  ' Creating some dictionaries
  Set dictNames       = CreateObject("Scripting.Dictionary")
  Set dictVars        = CreateObject("Scripting.Dictionary")
  Set dictACLs        = CreateObject("Scripting.Dictionary")
  Set dictACLVarNames = CreateObject("Scripting.Dictionary")
  Set dictChange      = CreateObject("Scripting.Dictionary")
  Set dictDevAffected = CreateObject("Scripting.Dictionary")

  ' Get a direct hook into specific sheets in the workbook, as well a command line hook.
  Set objShell = CreateObject("WScript.Shell")
  Set app = CreateObject("Excel.Application")
  Set wbin = app.Workbooks.Open (wbNameIn,0,true)
  Set wsNames = wbin.Worksheets("ACL Names")
  Set wsVars = wbin.Worksheets("OMW-Vars")
  Set wsACL = wbin.Worksheets("ACL Lines")

  ' Create output files and catch any errros in the process
  set objLogOut = CreateFile(strlogFile)
  set objFileOut = CreateFile(strOutFile)

  ' Now start the real work
  objLogOut.writeline "Starting at " & now()
  objFileOut.writeline "primaryIPAddress,hostName,ACL Name, comment"
  app.visible = ExcelVisible

  ' Grab all the ACL names and stick them in a dictionary, store the Abriviated ID, the Actual ACL and the variable name if applcible.
  iNameRow = 4
  dictNames.removeall
  dictACLVarNames.removeall
  Do
  	If wsNames.Cells(iNameRow,1).Value = "" Then Exit Do
  	If not dictNames.Exists(wsNames.Cells(iNameRow,1).value) then
  		dictNames.Add wsNames.Cells(iNameRow,1).value, wsNames.Cells(iNameRow,2).value
  	End If
    If not dictACLVarNames.Exists(wsNames.Cells(iNameRow,1).value) then
      dictACLVarNames.Add wsNames.Cells(iNameRow,1).value, wsNames.Cells(iNameRow,3).value
    End If
  	iNameRow = iNameRow + 1
  loop


  ' Grab the column headers of the variable sheet and stick them into a dictionary.
  iVarCol=1
  dictVars.removeall
  Do
    If wsVars.Cells(1,iVarCol).Value = "" Then Exit Do
    If not dictVars.Exists(wsVars.Cells(1,iVarCol).value) then
      dictVars.Add wsVars.Cells(1,iVarCol).value, iVarCol
    End If
    iVarCol = iVarCol + 1
  loop

  ' Get the column headers for the ACL sheet and sticke them into a dicationary.
  iACLCol=1
  dictACLs.removeall
  Do
    If wsACL.Cells(1,iACLCol).Value = "" Then Exit Do
    If not dictACLs.Exists(wsACL.Cells(1,iACLCol).value) then
      dictACLs.Add wsACL.Cells(1,iACLCol).value, iACLCol
    End If
    iACLCol = iACLCol + 1
  loop

  strChange  = ""
  strMOPPath = left (wbNameIn, InStrRev (wbNameIn,".")-1)& "-Changes\"
  if not fso.FolderExists(strMOPPath) then
    CreatePath (strMOPPath)
    objLogOut.writeline """" & strMOPPath & """ did not exists so I created it"
  end if

  ' For testing and dev purpose, focus on a single ACL from ACL Name sheet. Looping throught them all comes later.
  iNameRow = 4
  strACLID = wsNames.Cells(iNameRow,1).value
  strACLName = wsNames.Cells(iNameRow,2).value
  strACLNameVar = wsNames.Cells(iNameRow,3).value
  strIPVer = "ipv4"
  dictDevAffected.RemoveAll

  ' Setup the output paths to be ACL specific
  ' First Folder for the Generated ACL's
  strGenOutPath = strOutPath & strACLName & "-Gen\"
  if not fso.FolderExists(strGenOutPath) then
    CreatePath (strGenOutPath)
    objLogOut.writeline """" & strGenOutPath & """ did not exists so I created it"
  end if

  ' Then a folder for the ACL's we grab from the router
  strAsIsOutPath = strOutPath & strACLName & "-AsIs\"
  if not fso.FolderExists(strAsIsOutPath) then
    CreatePath (strAsIsOutPath)
    objLogOut.writeline """" & strAsIsOutPath & """ did not exists so I created it"
  end if

  ' There should be column in the ACL sheet whose header is the same as the ACLID we're working on. Get that column number or report an issue and exit.
  if dictACLs.Exists(strACLID) then
    iACLCol = dictACLs(strACLID)
  else
    objLogOut.writeline "couldn't find " & strACLID & " in dictACLs :-("
    msgbox "couldn't find " & strACLID & " in dictACLs, exiting :-("
    CleanUp
    exit sub
  end if

  ' Get the column number for the Router IP address in the Variable sheet, report an error and exit if can't be deteremined.
  if dictVars.Exists("primaryIPAddress") then
    iIPCol = dictVars("primaryIPAddress")
  else
    objLogOut.writeline "couldn't find primaryIPAddress in dictVars :-("
    msgbox "couldn't find primaryIPAddress in dictACLs, exiting :-("
    CleanUp
    exit sub
  end if

  ' Get the column number for the Router hostname in the Variable sheet, report an error and exit if can't be deteremined.
  if dictVars.Exists("hostName") then
    iHostCol = dictVars("hostName")
  else
    objLogOut.writeline "couldn't find hostName in dictVars :-("
    msgbox "couldn't find hostName in dictACLs, exiting :-("
    CleanUp
    exit sub
  end if


  iVarRow=2
  iError = 1
  do ' Now start looping throught the variable sheet.
    strIPAddr = wsVars.Cells(iVarRow,iIPCol).value
    if strHostname <> wsVars.Cells(iVarRow,iHostCol).value then
      strHostname = wsVars.Cells(iVarRow,iHostCol).value
      iError = 1 ' Ensure the error counter is set back to 1, which is default and means no error.
    end if

    ' If this ACL has variability in the name find the column number in the variable sheet.
    if strACLNameVar <> "" then
      if dictVars.Exists(strACLNameVar) then
        strACLName = wsVars.Cells(iVarRow,dictVars(strACLNameVar)).value
      end if
      strChange  = strIPVer & " access-list $" & strACLNameVar & "$" &vbcrlf
    else
      strChange  = strIPVer & " access-list " & strACLName & vbcrlf
    end if
    objLogOut.writeline "Starting on router " & strHostname & " with ACL " & strACLName ' Log that we are about to log into a router.

    strNotes   = ""
    strNoMatch = ""
    strMissing = ""
    bRange     = False
    iLastLine  = 0

    strResultParts = GetAsIsACL(strIPVer, strACLName, strHostname, iError)
    if isArray(strResultParts) then
      iError = 1
    else
      strNotes = "Failed to connect "
      if IsNumeric(strResultParts) then
        iError = strResultParts
      else
        objLogOut.writeline "GetAsIsACL returned neither array nor a number, this shouldn't happen so I'm quitting. I got back: " & strResultParts
        exit Sub
      end if
    end if
    strTempOut = ""
    iACLRow=2
    iResult=1
    if iError = 1 then ' If there has been no connection errors, analyse the results.
      set objACLGen = CreateFile(strGenOutPath & strHostname & "-" & strACLName & ".txt")
      do
        bComp = False
        if wsACL.Cells(iACLRow,iACLCol).value <> "" then ' Is current ACL standard line non-blank.
          iStartPos = instr (1,wsACL.Cells(iACLRow,1).value,"$",vbTextCompare) ' Look for $ which indicates a start of a variable in the ACL standard.
          if iStartPos > 0 then ' If the current line has a variable parse out the variable, and substitute it with the proper value.
            iStopPos = instr (iStartPos+1,wsACL.Cells(iACLRow,1).value,"$",vbTextCompare) ' Locate the end of the variable name.
            strACLVar = mid(wsACL.Cells(iACLRow,1).value,iStartPos+1,iStopPos-iStartPos-1) ' Store the name of the variable.
            if strACLVar = "ACLName" then ' If the variable is "ACLName" then substitute it with the actual ACL Name
              strTempOut = replace(wsACL.Cells(iACLRow,1).value,"$ACLName$",strACLName)
              objACLGen.writeline strTempOut
              bComp = True
            end if ' End if the Variable is ACLName.

            if dictVars.Exists(strACLVar) then ' If the variable we found exists in the variable sheet.
              iVarCol = dictVars(strACLVar)
              if wsVars.Cells(iVarRow, iVarCol) <> "" then ' If the varible value is not an empty string do the substitution and write the generate ACL line to file.
                strTempOut = replace(wsACL.Cells(iACLRow,1).value,"$"&strACLVar&"$",wsVars.Cells(iVarRow, iVarCol))
                objACLGen.writeline strTempOut
                bComp = True
              end if ' End if Variable is not empty string.
            end if ' end if variable exists
          else ' If the current line has no variables, just write it to the file.
            strTempOut = wsACL.Cells(iACLRow,1).value
            objACLGen.writeline strTempOut
            bComp = True
          end if ' End if analyzing the current line of the ACL standard.
          if bComp then ' If the ACL line was found applicable, compare the generated line with the same line in the ACL capture.
            iGSeq = trim(left(strTempOut,instr(1,trim(strTempOut)," ",vbTextCompare))) ' Grab the sequence number of the generated ACL line we're looking at
            iASeq = trim(left(strResultParts(iResult),instr(1,trim(strResultParts(iResult))," ",vbTextCompare))) ' Grab the sequence number of the router ACL line we're looking at
            if strTempOut <> trim(strResultParts(iResult)) Then ' If generated and AsIs lines aren't identical, note it.
              if iGSeq > iASeq Then
                objLogOut.writeline "Line " & iResult & ": Extra line on router not in standard: " & trim(strResultParts(iResult))
                strMissing = strMissing & iResult & "(only on router) "
                bRange = False
                strChange = strChange & "no " & iASeq & vbcrlf
              end if
              if iGSeq < iASeq Then
                objLogOut.writeline "Line " & iResult & ": This standard line missing from router: " & strTempOut
                strMissing = strMissing & iResult & "(missing from router) "
                bRange = False
                strChange = strChange & wsACL.Cells(iACLRow,1).value & vbcrlf
              end if
              if iGSeq = iASeq then ' If seqences match report the lines don't match
                if iLastLine > 0 and iLastLine + 1 = iResult Then ' If last line didn't match
                  bRange = True
                else
                  if bRange = True Then
                    strNoMatch = trim(strNoMatch) & "-" & iLastLine & " " & iResult & " "
                  else
                    strNoMatch = strNoMatch & iResult & " "
                  end if ' end if in a range
                  bRange = False
                end if ' end if last line didn't match.
                objLogOut.writeline strHostname & " " & strACLName & " no matchy on line " & iResult
                objLogOut.writeline " Gen: " & strTempOut
                objLogOut.writeline "AsIs: " & trim(strResultParts(iResult))
                objLogOut.writeline "--------------------------------"
                strChange = strChange & wsACL.Cells(iACLRow,1).value & vbcrlf
              end if ' end if seqences match
              iLastLine = iResult
            end if ' End if generated and AsIs are different.
            ' If there are lines left in the captured ACL and router ACL sequence is lower or equal move on to the next line.
            if iResult < ubound(strResultParts) and iGSeq >= iASeq then iResult = iResult + 1
          end if ' End If ACL is applicable
        end if ' end Is current ACL standard line non-blank.
        if iGSeq <= iASeq then iACLRow = iACLRow + 1 ' Move down line in the ACL sheet if we are in sync or we are too low
      loop until wsACL.Cells(iACLRow,1).Value = "" ' Unless the new line is blank, loop back and repeat.
      objACLGen.Close
    end if ' End of checking for error prior to analysis
    if bRange = True then
      strNoMatch = trim(strNoMatch) & "-" & iLastLine
    end if

    if iError > MaxError then objLogOut.writeline "No connection after " & MaxError & " attempts, giving up and moving on."
    if iError = 1 or iError > MaxError then
      iVarRow = iVarRow + 1
      if strNoMatch = "" and strMissing = "" and strNotes = "" then
        strNotes = "Good"
      else
        if strNotes <> "Failed to connect " then
          ' if strNoMatch <> "" then strNotes = trim(strNotes) & " " & "Lines " & trim(strNoMatch) & " Don't match; "
          ' if strMissing <> "" then strNotes = trim(strNotes) & " " & "These lines are only on one side: " & trim(strMissing) & ";"
          if strNoMatch <> "" or strMissing <> "" then strNotes = strNotes & "Updates required"
          bNewChange = True
          if dictChange.Exists(strChange) then
            bNewChange = False
            if not dictDevAffected.Exists(strHostname) then
              dictChange.item(strChange) = dictChange.item(strChange) & vbcrlf & strHostname
              dictDevAffected.add strHostname,""
            end if
          else
            dictChange.add strChange,strHostname
          end if
        end if
      end if
      objFileOut.writeline strIPAddr & "," & strHostname & "," & strACLName & "," & strNotes
    end if
  loop until wsVars.Cells(iVarRow,1).Value = "" ' This is the end of the loop to go through the Variable sheet
  dkeys = dictChange.keys
  iChangeID = 1
  for each dkey in dkeys
    set objChangeOut = CreateFile(strMOPPath & "HPNAScript-Change" & iChangeID & ".txt")
    objChangeOut.writeline "************ Devices Affected ************" & vbcrlf & dictChange.item(dkey) & vbcrlf & "******************************************"
    objChangeOut.writeline dkey
    CreateCSVs dictChange.item(dkey),dkey,iChangeID
    iChangeID = iChangeID + 1
    objChangeOut.close
  next

  objLogOut.writeline "All done at " & now()
  objLogOut.close
  objFileOut.close
  CleanUp
End Sub ' End of the Main sub

Sub CreateCSVs (strDevlist, strChange, iChangeID)
'-------------------------------------------------------------------------------------------------'
' Sub CreateCSVs (strDevlist, strChange, iChangeID)                                               '
'                                                                                                 '
' This sub takes a Devicelist (CRLF seperated), configuration script and changeID and generates   '
' all the CSV files to be used by both HPNA and PIER to push out those changes.                   '
'-------------------------------------------------------------------------------------------------'
dim strDevListParts, strChangeLines, x, y, strVarCol, dictDevices, iRow, iVarColList, objHPNAout
dim iStartPos, iStopPos, strACLVar, strColHead, iCol

  set objHPNAout = CreateFile(strMOPPath & "HPNAVars-Change" & iChangeID & ".csv")
  Set dictDevices = CreateObject("Scripting.Dictionary")
  strDevListParts = split(strDevlist,vbcrlf)
  strChangeLines  = split(strChange,vbcrlf)
  strVarCol = ""
  objLogOut.writeline "in CreateCSVs, checking change ID " & iChangeID
  for x = 0 to ubound(strChangeLines)
    objLogOut.writeline "checking for variable in " & strChangeLines(x)
    iStartPos = instr (1,strChangeLines(x),"$",vbTextCompare) ' Look for $ which indicates a start of a variable in the ACL standard.
    if iStartPos > 0 then ' If the current line has a variable parse out the variable, and substitute it with the proper value.
      iStopPos = instr (iStartPos+1,strChangeLines(x),"$",vbTextCompare) ' Locate the end of the variable name.
      strACLVar = mid(strChangeLines(x),iStartPos+1,iStopPos-iStartPos-1) ' Store the name of the variable.
      objLogOut.writeline "found variable: " & strACLVar
      if dictVars.Exists(strACLVar) then ' If the variable we found exists in the variable sheet.
        strVarCol = strVarCol & dictVars(strACLVar) & ","
        objLogOut.writeline "it's in column " & dictVars(strACLVar)
      end if ' end if variable exists
    end if
  next
  ' Grab the device name column of the variable sheet and stick them into a dictionary.
  iRow=2
  dictDevices.removeall
  Do until wsVars.Cells(iRow,2).Value = ""
    If not dictDevices.Exists(wsVars.Cells(iRow,2).value) then
      dictDevices.Add wsVars.Cells(iRow,2).value, iRow
    End If
    iRow = iRow + 1
  loop
  objLogOut.writeline "strVarCol:" & strVarCol
  iVarColList = split(strVarCol,",")
  objLogOut.writeline "contains " & ubound(iVarColList) & " elements"
  objHPNAout.write wsVars.Cells(1,1).value
  objHPNAout.write "," & wsVars.Cells(1,2).value
  for x=0 to ubound(iVarColList)
    if IsNumeric(iVarColList(x)) then
      iCol = cint(iVarColList(x))
      objHPNAout.write "," & wsVars.Cells(1,iCol).value
    end if
  next
  objHPNAout.writeline
  for x=0 to ubound(strDevListParts)
    if dictDevices.Exists(strDevListParts(x)) then
      iRow = dictDevices(strDevListParts(x))
    else
      objLogOut "something weird just happened, can't find " & strDevListParts(x) & " in the spreadsheet!"
    end if
    objHPNAout.write wsVars.Cells(iRow,1).value
    objHPNAout.write "," & wsVars.Cells(iRow,2).value
    for y=0 to ubound(iVarColList)
      if IsNumeric(iVarColList(y)) then
        iCol = cint(iVarColList(y))
        objHPNAout.write "," & wsVars.Cells(iRow,iCol).value
      end if
    next
    objHPNAout.writeline
  next
End Sub

Sub CleanUp()
'-------------------------------------------------------------------------------------------------'
' Sub CleanUp()                                                               '
'                                                                                                 '
' This is a cleanup function.             '
'-------------------------------------------------------------------------------------------------'
  crt.Session.Disconnect
  if AutoCloseResults = True then
    if IsObject(wbin) then wbin.Close
    ' app.Quit
  end if

  if IsObject(app) then Set wbin = app.Workbooks.Open (strOutFile,0,False)
  if IsObject(objShell) then objShell.run ("notepad " & strlogFile)

  if IsObject(wbin) then Set wbin = Nothing
  if IsObject(wsNames) then Set wsNames = Nothing
  if IsObject(wsVars) then Set wsVars = Nothing
  if IsObject(wsACL) then Set wsACL = Nothing
  if IsObject(app) then Set app = Nothing
  if IsObject(objShell) then set objShell = Nothing
  if IsObject(objACLGen) then
    objACLGen.close
    set objACLGen = Nothing
  end if
  if IsObject(objChangeOut) then
    objChangeOut.Close
    set objChangeOut = Nothing
  end if
  if IsObject(objLogOut) then
    objLogOut.Close
    set objLogOut = Nothing
  end if
  if IsObject(objFileOut) then
    objFileOut.Close
    set objFileOut = Nothing
  end if
end Sub ' End of CleanUp sub

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
end Function ' End of CreatePath Function

Function CreateFile (strFilePath)
'-----------------------------------------------------------------------------------------------------'
' Function CreateFile (strFilePath)                                                                    '
'                                                                                                     '
' This function takes a filepath and returns a file handle, while doing all nessisary error handling. '
'-----------------------------------------------------------------------------------------------------'
dim objFileOut, iInt, strOrigional, iPos, iLen

  iInt = 1
  strOrigional = strFilePath
  iPos = InStrRev (strFilePath,".") - 1
  iLen = len(strFilePath)
  on error resume next
  set objFileOut = fso.OpenTextFile(strFilePath, ForWriting, True)
  if err.number <> 0 Then
    if err.number = 70 then
      while err.number = 70
        strFilePath = left(strFilePath, iPos) & "-" & iInt & right(strFilePath, iLen-iPos)
        err.clear
        set objFileOut = fso.OpenTextFile(strFilePath, ForWriting, True)
        iInt = iInt + 1
        objLogOut.writeline "trying " & strFilePath
      wend
      objLogOut.writeline "Permission denied error when attempting to create file " & strOrigional & ". Created " & strFilePath & " instead."
    else
      MsgBox ("Create file Error # " & CStr(Err.Number) & " " & Err.Description)
      Exit Function
    end if
  end if
  on error goto 0
  set CreateFile = objFileOut
end Function ' End of CreateFile Function

Function GetAsIsACL(strIPVer, strACLName, strHostname, iError)
'-----------------------------------------------------------------------------------------------------'
' Function GetAsIsACL()                                                                               '
'                                                                                                     '
' This function logs into a router and feteches the current AsIs ACL.                                 '
'-----------------------------------------------------------------------------------------------------'
dim strVerifyCmd, strConnection, strResult, objACLAsIs, szHostName, objTab, objConfig

    strVerifyCmd = "show run " & strIPVer & " access-list " & strACLName ' construct the verification command to run.

    ' If session is connected, disconnect it.
    If crt.Session.Connected Then
      Set objTab = crt.GetScriptTab ' Hook into the current SecureCRT Tab
      Set objConfig = objTab.Session.Config ' Grab the session configuration for the current tab
      szHostName = objConfig.GetOption("Hostname") ' Get the currently connected hostname from the current session configuration
      if szHostName <> strHostname then crt.Session.Disconnect ' Unless the current connection is to the router we need disconnect
    end if

    ' Make new connection unless already connected.
    If not crt.Session.Connected Then
      strConnection = "/SSH2 /ACCEPTHOSTKEYS "  & strHostname ' connect string
      on error resume next
      crt.Session.Connect strConnection
      on error goto 0
    end if

    If crt.Session.Connected Then ' If we have a successful connection, run the verification command, write the ACL to a file and keep it in a variable.
      iError = 1 ' Ensure the error counter is set back to 1, which is default and means no error.
      crt.Screen.Synchronous = True
      crt.Screen.WaitForString "#",Timeout
      crt.Screen.Send("term len 0" & vbcr)
      crt.Screen.WaitForString "#",Timeout
      crt.Screen.Send(strVerifyCmd & vbcr)
      crt.Screen.WaitForString vbcrlf,Timeout
      strResult=trim(crt.Screen.Readstring (vbcrlf&"RP/",Timeout))
      ' crt.Session.Disconnect
      set objACLAsIs = CreateFile(strAsIsOutPath & strHostname & "-" & strACLName & ".txt")
      objACLAsIs.write strResult
      objACLAsIs.close
      set objACLAsIs = Nothing
      GetAsIsACL = split (strResult,vbcrlf)
    else ' If no connection, increase an error counter and note the failure.
      objLogOut.writeline "No connection to " & strHostname & " " & strACLName & ". Attempt #" & iError
      GetAsIsACL = iError + 1
      ' strNotes = "Failed to connect "
    end if ' End of connection verification

end Function ' End GetAsIsACL function