Option Explicit

Sub main
dim AutoCloseResults, AutoSaveResults, wbNameIn, DateInFileName, ExcelVisible
'|----------------------------------------------------------------------------------------------------------|
'|  This script will log into each ARG in a specified spreadsheet and confirm ACLs are up to standards.     |
'|  If deviations are found, will generate configuration files for HPNA or manual MOP.                      |
'|                                                                                                          |
'|  Author: Siggi Bjarnason                                                                                 |
'|  Authored: 04/10/2017                                                                                    |
'|  Copyright: Siggi Bjarnason 2017                                                                         |
'|----------------------------------------------------------------------------------------------------------|

' User Spefified values, specify values here per your needs

	DateInFileName = True
	AutoSaveResults = True
	AutoCloseResults = True
  ExcelVisible = True
  const Timeout    = 5    ' Timeout in seconds for each command, if expected results aren't received withing this time, the script moves on.

' Non user section, changes to this section can have undesired results
  Dim app, objShell, dictNames, dictVars, dictACLs, dictACLVarNames, bComp
  Dim wsNames, wsVars, wsACL, wbin, fso, objFileOut, objLogOut, objACLGen, objACLAsIs
  Dim iNameRow, iVarRow, iACLRow, iNameCol, iACLCol, iVarCol, iStartPos, iStopPos, iHostCol, iIPCol, iError, iResult
  Dim strOutPath, strOutFile, strlogFile, strACLVar, strTempOut, strACLName, strACLID, strACLNameVar, strErr, strIPVer
  Dim strHostname, strIPAddr, strResult, strResultParts, strConnection, strGenOutPath, strAsIsOutPath, strVerifyCmd


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

  ' comparison constants
  Const vbBinaryCompare = 0 'Perform a binary comparison, i.e case sensitive
  Const vbTextCompare = 1 'Perform a textual comparison, i.e. case insensitive

  wbNameIn = crt.Dialog.FileOpenDialog("Please select OMW ACL Standard Spreadsheet", "Open", "", "Excel Files (*.xlsx)|*.xlsx||")

  Set fso = CreateObject("Scripting.FileSystemObject")

  if not fso.FileExists(wbNameIn) Then
    msgbox "Input file " & wbNameIn & " not found, exiting"
    exit sub
  end if

  strOutPath = left (wbNameIn, InStrRev (wbNameIn,"\"))
  if right(strOutPath,1)<>"\" then
    strOutPath = strOutPath & "\"
  end if

  strOutFile = left (wbNameIn, InStrRev (wbNameIn,".")-1)&"-Results.csv"
  strlogFile = left (wbNameIn, InStrRev (wbNameIn,".")-1)&"-log.txt"

  Set dictNames = CreateObject("Scripting.Dictionary")
  Set dictVars = CreateObject("Scripting.Dictionary")
  Set dictACLs = CreateObject("Scripting.Dictionary")
  Set dictACLVarNames = CreateObject("Scripting.Dictionary")

  Set objShell = CreateObject("WScript.Shell")
  Set app = CreateObject("Excel.Application")
  Set wbin = app.Workbooks.Open (wbNameIn,0,true)
  Set wsNames = wbin.Worksheets("ACL Names")
  Set wsVars = wbin.Worksheets("OMW-Vars")
  Set wsACL = wbin.Worksheets("ACL Lines")

  set objLogOut = fso.OpenTextFile(strlogFile, ForWriting, True)
  set objFileOut  = fso.OpenTextFile(strOutFile, ForWriting, True)

  objLogOut.writeline "Starting at " & now()
  objFileOut.writeline "primaryIPAddress,hostName,comment"
  app.visible = ExcelVisible

  iNameRow = 4
  dictNames.removeall
  dictACLVarNames.removeall
  Do
  	If wsNames.Cells(iNameRow,1).Value = "" Then Exit Do
    ' objLogOut.writeline wsNames.Cells(iNameRow,1).value & "=" & wsNames.Cells(iNameRow,2).value & "/" & wsNames.Cells(iNameRow,3).value
  	If not dictNames.Exists(wsNames.Cells(iNameRow,1).value) then
  		dictNames.Add wsNames.Cells(iNameRow,1).value, wsNames.Cells(iNameRow,2).value
  	End If
    If not dictACLVarNames.Exists(wsNames.Cells(iNameRow,1).value) then
      dictACLVarNames.Add wsNames.Cells(iNameRow,1).value, wsNames.Cells(iNameRow,3).value
    End If
  	iNameRow = iNameRow + 1
  loop

  iVarCol=1
  dictVars.removeall
  Do
    If wsVars.Cells(1,iVarCol).Value = "" Then Exit Do
    ' objLogOut.writeline wsVars.Cells(1,iVarCol).value & "=" & iVarCol
    If not dictVars.Exists(wsVars.Cells(1,iVarCol).value) then
      dictVars.Add wsVars.Cells(1,iVarCol).value, iVarCol
    End If
    iVarCol = iVarCol + 1
  loop

  iACLCol=1
  dictACLs.removeall
  Do
    If wsACL.Cells(1,iACLCol).Value = "" Then Exit Do
    ' objLogOut.writeline wsACL.Cells(1,iACLCol).value & "=" & iACLCol
    If not dictACLs.Exists(wsACL.Cells(1,iACLCol).value) then
      dictACLs.Add wsACL.Cells(1,iACLCol).value, iACLCol
    End If
    iACLCol = iACLCol + 1
  loop

  iNameRow = 4
  strACLID = wsNames.Cells(iNameRow,1).value
  strACLName = wsNames.Cells(iNameRow,2).value
  strACLNameVar = wsNames.Cells(iNameRow,3).value
  strIPVer = "ipv4"

  strGenOutPath = strOutPath & strACLName & "-Gen\"
  if not fso.FolderExists(strGenOutPath) then
    CreatePath (strGenOutPath)
    objLogOut.writeline """" & strGenOutPath & """ did not exists so I created it"
  end if

  strAsIsOutPath = strOutPath & strACLName & "-AsIs\"
  if not fso.FolderExists(strAsIsOutPath) then
    CreatePath (strAsIsOutPath)
    objLogOut.writeline """" & strAsIsOutPath & """ did not exists so I created it"
  end if

  if dictACLs.Exists(strACLID) then
    iACLCol = dictACLs(strACLID)
    objLogOut.writeline strACLID & " is column " & iACLCol
  else
    objLogOut.writeline "couldn't find " & strACLID & " in dictACLs :-("
    msgbox "couldn't find " & strACLID & " in dictACLs, exiting :-("
    exit sub
  end if

  if dictVars.Exists("primaryIPAddress") then
    iIPCol = dictVars("primaryIPAddress")
    objLogOut.writeline "primaryIPAddress is column " & iIPCol
  else
    objLogOut.writeline "couldn't find primaryIPAddress in dictVars :-("
    msgbox "couldn't find primaryIPAddress in dictACLs, exiting :-("
    exit sub
  end if

  if dictVars.Exists("hostName") then
    iHostCol = dictVars("hostName")
    objLogOut.writeline "hostName is column " & iHostCol
  else
    objLogOut.writeline "couldn't find hostName in dictVars :-("
    msgbox "couldn't find hostName in dictACLs, exiting :-("
    exit sub
  end if

  iVarRow=2
  do
    strIPAddr = wsVars.Cells(iVarRow,iIPCol).value
    strHostname = wsVars.Cells(iVarRow,iHostCol).value
    if strACLNameVar <> "" then
      if dictVars.Exists(strACLNameVar) then
        strACLName = wsVars.Cells(iVarRow,dictVars(strACLNameVar)).value
      end if
    end if
    strVerifyCmd = "show run " & strIPVer & " access-list " & strACLName
    objLogOut.writeline "Starting on router " & strHostname & " with ACL " & strACLName
    If crt.Session.Connected Then
      crt.Session.Disconnect
    end if

    strConnection = "/SSH2 /ACCEPTHOSTKEYS "  & strHostname
    on error resume next
    crt.Session.Connect strConnection
    on error goto 0

    If crt.Session.Connected Then
      crt.Screen.Synchronous = True
      crt.Screen.WaitForString "#",Timeout
      iError = Err.Number
      strErr = Err.Description
      If iError <> 0 Then
        result = "Error " & iError & ": " & strErr
      end if
      crt.Screen.Send("term len 0" & vbcr)
      crt.Screen.WaitForString "#",Timeout
      crt.Screen.Send(strVerifyCmd & vbcr)
      crt.Screen.WaitForString vbcrlf,Timeout
      strResult=trim(crt.Screen.Readstring (vbcrlf&"RP/",Timeout))
      crt.Session.Disconnect
      set objACLAsIs = fso.OpenTextFile(strAsIsOutPath & strHostname & "-" & strACLName & ".txt", ForWriting, True)
      objACLAsIs.write strResult
      objACLAsIs.close
      strResultParts = split (strResult,vbcrlf)
      objLogOut.writeline strACLName & " contains " & ubound(strResultParts) & " on " & strHostname
    else
      objLogOut.writeline "No connection to " & strHostname & " will generate what the ACL should be, no as-is"
    end if
    set objACLGen = fso.OpenTextFile(strGenOutPath & strHostname & "-" & strACLName & ".txt", ForWriting, True)
    strTempOut = ""
    iACLRow=2
    iResult=1
    do
      if wsACL.Cells(iACLRow,iACLCol).value <> "" then
        iStartPos = instr (1,wsACL.Cells(iACLRow,1).value,"$",vbTextCompare)
        if iStartPos > 0 then
          iStopPos = instr (iStartPos+1,wsACL.Cells(iACLRow,1).value,"$",vbTextCompare)
          strACLVar = mid(wsACL.Cells(iACLRow,1).value,iStartPos+1,iStopPos-iStartPos-1)
          if strACLVar = "ACLName" then
            strTempOut = replace(wsACL.Cells(iACLRow,1).value,"$ACLName$",strACLName)
            objACLGen.writeline strTempOut
          end if
          if dictVars.Exists(strACLVar) then
            iVarCol = dictVars(strACLVar)
            if wsVars.Cells(iVarRow, iVarCol) <> "" then
              strTempOut = replace(wsACL.Cells(iACLRow,1).value,"$"&strACLVar&"$",wsVars.Cells(iVarRow, iVarCol))
              objACLGen.writeline strTempOut
            end if
          end if
        else
          strTempOut = wsACL.Cells(iACLRow,1).value
          objACLGen.writeline strTempOut
        end if
        if strTempOut = trim(strResultParts(iResult)) Then
          objLogOut.writeline "Line " & iACLRow & " matches"
          objLogOut.writeline "strTempOut: " & strTempOut
        else
          objLogOut.writeline "--------------"&vbcrlf&"start Line mismatch"&vbcrlf&"---------------"
          objLogOut.writeline "strTempOut: " & strTempOut
          objLogOut.writeline "strResultParts(" & iResult & "): " & strResultParts(iResult)
          objLogOut.writeline "--------------"&vbcrlf&"end Line mismatch"&vbcrlf&"---------------"
        end if
        if iResult < ubound(strResultParts) then iResult = iResult + 1
      end if
      iACLRow = iACLRow + 1
    loop until wsACL.Cells(iACLRow,1).Value = ""
    iVarRow = iVarRow + 1
    objACLGen.Close
    objFileOut.writeline strIPAddr & "," & strHostname & ", Asis contains" & ubound(strResultParts)
  loop until wsVars.Cells(iVarRow,1).Value = "10.250.80.98"
  objLogOut.writeline "All done at " & now()

  if AutoCloseResults = True then
    wbin.Close
    app.Quit
  end if

  Set wbin = Nothing
  Set wsNames = Nothing
  Set wsVars = Nothing
  Set wsACL = Nothing
  Set app = Nothing
  objShell.run ("notepad " & strlogFile)
  ' msgbox "Done"

End Sub


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
