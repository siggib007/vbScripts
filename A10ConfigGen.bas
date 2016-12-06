Attribute VB_Name = "A10ConfigGen"
Option Explicit
Dim fso, dictSG, dictSNAT

Sub ConfigGen()
Dim strWBFullName, strLBName1, strLBName2, iPathLen, strWBPath, strSavePath, strScriptName, strCRNum, strProject, ObjScript, ObjLog, x, strCurSG, strSPName, strSGName, iExtPos, strLogName, strPersistType
Dim strSGProt, strSGMethod, strSGHealth, iSGRow, bSGGood, strServerName, strServerIP, strServerPort, strVIPName, strVIPip, strVIPPort, strVIPProt, strSNAT, strHA, strPersist, strAflex, strSSL, strSNATPart2
Dim strBaselineName, strValidateName, strRollbackName, objBaseline, objValidate, objRollback, strBaseTemp, strValtemp, strRBTemp, strPriority, strProfile, strSNAT1, strSNAT2, strSNATMask, strSNATParts
Dim strOutput, dictServer, strBaseline, strVerify, strServiceName

Set fso = CreateObject("Scripting.FileSystemObject")
Set dictSG = CreateObject("Scripting.Dictionary")
Set dictSNAT = CreateObject("Scripting.Dictionary")
Set dictServer = CreateObject("Scripting.Dictionary")


strWBFullName = ThisWorkbook.FullName
strCRNum = Worksheets("Overview").Cells(4, 4).Value
strProject = Worksheets("Overview").Cells(5, 4).Value
strLBName1 = Worksheets("Overview").Cells(8, 4).Value
strLBName2 = Worksheets("Overview").Cells(9, 4).Value
strSavePath = Worksheets("Overview").Cells(12, 4).Value
iPathLen = InStrRev(strWBFullName, "\")
iExtPos = InStrRev(strWBFullName, ".")
strLogName = Left(strWBFullName, iExtPos) & "log"
Set ObjLog = fso.createtextfile(strLogName, True)

ObjLog.writeline "starting at " & Now()
strWBPath = Left(strWBFullName, iPathLen)
If strSavePath = "" Then
  strSavePath = strWBPath
  ObjLog.writeline "saving configuration to " & strSavePath
Else
  If InStr(strSavePath, "\") > 0 Then
    If Not fso.FolderExists(strSavePath) Then
      CreatePath (strSavePath)
      ObjLog.writeline "Path " & strSavePath & " didn't exists so I created it for you"
    End If
  Else
    strSavePath = "C:\" & strSavePath
    If Not fso.FolderExists(strSavePath) Then
      CreatePath (strSavePath)
      ObjLog.writeline "Path " & strSavePath & " didn't exists so I created it for you"
    End If
  End If
End If
If Right(strSavePath, 1) <> "\" Then
  strSavePath = strSavePath & "\"
End If
strScriptName = strSavePath & strCRNum & "-" & strProject & "-" & strLBName1 & "-" & strLBName2 & "-implement.txt"
strBaselineName = strSavePath & strCRNum & "-" & strProject & "-" & strLBName1 & "-" & strLBName2 & "-baseline.txt"
strValidateName = strSavePath & strCRNum & "-" & strProject & "-" & strLBName1 & "-" & strLBName2 & "-validate.txt"
strRollbackName = strSavePath & strCRNum & "-" & strProject & "-" & strLBName1 & "-" & strLBName2 & "-rollback.txt"

ObjLog.writeline "Implementation script saved to " & strScriptName
ObjLog.writeline "Baseline script saved to " & strBaselineName
ObjLog.writeline "Validation script saved to " & strValidateName
ObjLog.writeline "Rollback script saved to " & strRollbackName
Set ObjScript = fso.createtextfile(strScriptName, True)
Set objBaseline = fso.createtextfile(strBaselineName, True)
Set objValidate = fso.createtextfile(strValidateName, True)
Set objRollback = fso.createtextfile(strRollbackName, True)

ObjScript.writeline " !!! #### IMPLEMENTATION FOR " & strLBName1 & " and " & strLBName2 & " #### !!!"
objBaseline.writeline " !!! #### Baseline Verifications FOR " & strLBName1 & " and " & strLBName2 & " #### !!!"
objValidate.writeline " !!! #### Post Implement Validation FOR " & strLBName1 & " and " & strLBName2 & " #### !!!"
objRollback.writeline " !!! #### Rollback FOR " & strLBName1 & " and " & strLBName2 & " #### !!!"

x = 2
dictSG.RemoveAll

Do
  dictSG.Add Worksheets("Service Group Details").Cells(x, 1).Value, x
  x = x + 1
Loop Until Worksheets("Service Group Details").Cells(x, 1).Value = ""

strCurSG = ""
x = 2
Do
  If strCurSG <> Worksheets("Service Member").Cells(x, 1).Value Then
    If strCurSG <> "" Then
      ObjScript.writeline "exit"
      objRollback.writeline "no slb service-group " & strCurSG & " " & strSGProt
    End If
    strCurSG = Worksheets("Service Member").Cells(x, 1).Value
    If dictSG.exists(strCurSG) Then
      iSGRow = dictSG.Item(strCurSG)
      strSGProt = Worksheets("Service Group Details").Cells(iSGRow, 2).Value
      strSGMethod = Worksheets("Service Group Details").Cells(iSGRow, 3).Value
      strSGHealth = Worksheets("Service Group Details").Cells(iSGRow, 4).Value
      strBaseline = strBaseline & "show slb service-group " & strCurSG & " config" & vbCrLf
      strVerify = strVerify & "show slb service-group  " & strCurSG & " | include State" & vbCrLf
      ObjScript.writeline "slb service-group " & strCurSG & " " & strSGProt
      ObjScript.writeline " method " & strSGMethod
      If strSGHealth <> "" And strSGHealth <> "none" Then ObjScript.writeline " health-check " & strSGHealth
      strServerName = Worksheets("Service Member").Cells(x, 2).Value
      strServerIP = Worksheets("Service Member").Cells(x, 3).Value
      strServerPort = Worksheets("Service Member").Cells(x, 4).Value
      strPriority = Worksheets("Service Member").Cells(x, 5).Value
      If strPriority > 0 Then
        ObjScript.writeline " member " & strServerName & "_" & strServerIP & ":" & strServerPort & " " & strServerIP & " priority " & strPriority
      Else
        ObjScript.writeline " member " & strServerName & "_" & strServerIP & ":" & strServerPort & " " & strServerIP
      End If
      If Not dictServer.exists(strServerIP) Then
        objBaseline.writeline "show slb server all-partitions | include  " & strServerIP
        objValidate.writeline "show slb server all-partitions | include  " & strServerIP
        'strOutput = "no slb server " & strServerName & "_" & strServerIP & " " & strServerIP
        objRollback.writeline "no slb server " & strServerName & "_" & strServerIP & " " & strServerIP
        dictServer.Add strServerIP, strServerName
      End If
      bSGGood = True
    Else
      ObjLog.writeline "Service Group " & strCurSG & " not define"
      bSGGood = False
    End If
  Else
    If bSGGood Then
      strServerName = Worksheets("Service Member").Cells(x, 2).Value
      strServerIP = Worksheets("Service Member").Cells(x, 3).Value
      strServerPort = Worksheets("Service Member").Cells(x, 4).Value
      strPriority = Worksheets("Service Member").Cells(x, 5).Value
      If strPriority > 0 Then
        ObjScript.writeline " member " & strServerName & "_" & strServerIP & ":" & strServerPort & " " & strServerIP & " priority " & strPriority
      Else
        ObjScript.writeline " member " & strServerName & "_" & strServerIP & ":" & strServerPort & " " & strServerIP
      End If
      If Not dictServer.exists(strServerIP) Then
        objBaseline.writeline "show slb server all-partitions | include  " & strServerIP
        objValidate.writeline "show slb server all-partitions | include  " & strServerIP
        'strOutput = "no slb server " & strServerName & "_" & strServerIP & " " & strServerIP
        objRollback.writeline "no slb server " & strServerName & "_" & strServerIP & " " & strServerIP
        dictServer.Add strServerIP, strServerName
      End If
    End If
  End If
  x = x + 1
Loop Until Worksheets("Service Member").Cells(x, 1).Value = ""
ObjScript.writeline "exit"
objRollback.writeline "no slb service-group " & strCurSG & " " & strSGProt
objBaseline.writeline "!"
objBaseline.writeline "! *** Verify above servers do not exist ***"
objBaseline.writeline "!"
objValidate.writeline "!"
objValidate.writeline "! *** Verify above servers exist ***"
objValidate.writeline "!"
objBaseline.write strBaseline
objValidate.write strVerify
objBaseline.writeline "!"
objBaseline.writeline "! *** No such service group ***"
objBaseline.writeline "!"
objValidate.writeline "!"
objValidate.writeline "! *** Verify above service groups are up ***"
objValidate.writeline "!"

strBaseTemp = ""
strValtemp = ""
strRBTemp = ""
x = 2
dictSNAT.RemoveAll
Do
  strVIPName = Worksheets("VIP details").Cells(x, 1).Value
  strVIPip = Worksheets("VIP details").Cells(x, 2).Value
  strVIPPort = Worksheets("VIP details").Cells(x, 3).Value
  If strVIPPort = "any" Then strVIPPort = 0
  strVIPProt = Worksheets("VIP details").Cells(x, 4).Value
  strServiceName = Worksheets("VIP details").Cells(x, 13).Value
  If strVIPProt = "any" Then strVIPProt = "others"
  strSNAT = Worksheets("VIP details").Cells(x, 5).Value
  strSPName = ""
  strHA = Worksheets("VIP details").Cells(x, 6).Value
  strPersistType = Worksheets("VIP details").Cells(x, 7).Value
  strPersist = Worksheets("VIP details").Cells(x, 8).Value
  strAflex = Worksheets("VIP details").Cells(x, 9).Value
  strSSL = Worksheets("VIP details").Cells(x, 10).Value
  strSGName = Worksheets("VIP details").Cells(x, 11).Value
  strProfile = Worksheets("VIP details").Cells(x, 12).Value
  If strProfile = "tcp" Then strProfile = ""
  If Not dictSG.exists(strSGName) Then
    ObjLog.writeline "Warning!!! Service Group " & strSGName & " not defined by this script, make sure it exists on the load balancer. "
  End If
  If strSNAT = "none" Then strSNAT = ""
  If strSNAT <> "" Then
    If strSNAT = "automap" Then strSNAT = "auto"
    If strSNAT = "auto" Then
      strSPName = "auto"
    Else
        strSPName = strVIPName & "-SP"
        If dictSNAT.exists(strSNAT) Then
          strSPName = dictSNAT.Item(strSNAT)
        Else
          dictSNAT.Add strSNAT, strSPName
          If InStr(strSNAT, "-") > 0 Then
            strSNATParts = Split(strSNAT, "-")
            strSNAT1 = strSNATParts(0)
            strSNATPart2 = Split(strSNATParts(1), "/")
            strSNAT2 = strSNATPart2(0)
            If UBound(strSNATPart2) > 0 Then
              strSNATMask = strSNATPart2(1)
            Else
              strSNATMask = ""
            End If
          Else
            strSNAT1 = strSNAT
            strSNAT2 = strSNAT
            strSNATMask = 32
          End If
          ObjScript.writeline "!" & vbCrLf & "ip nat pool " & strSPName & " " & strSNAT1 & " " & strSNAT2 & " netmask /" & strSNATMask & " ha-group-id " & strHA & vbCrLf & "!"
          strBaseTemp = strBaseTemp & "show ip nat pool " & strSPName & vbCrLf
          strValtemp = strValtemp & "show ip nat pool " & strSPName & vbCrLf
          strRBTemp = strRBTemp & "no ip nat pool " & strSPName & " " & strSNAT1 & " " & strSNAT2 & " netmask /" & strSNATMask & " ha-group-id " & strHA & vbCrLf
        End If
    End If
  End If
  ObjScript.writeline "slb virtual-server " & strVIPName & " " & strVIPip
  ObjScript.writeline " ha-group " & strHA
  ObjScript.writeline " port " & strVIPPort & " " & strVIPProt
  ObjScript.writeline "  name " & strServiceName
  If strSPName <> "" Then
    If strSPName = "auto" Then
      ObjScript.writeline "  source-nat auto "
    Else
      ObjScript.writeline "  source-nat pool " & strSPName
    End If
  End If
  ObjScript.writeline "  service-group " & strSGName
  If strSSL <> "" Then
    ObjScript.writeline "  template client-ssl " & strSSL
  End If
  If strProfile <> "" Then
    ObjScript.writeline "  template " & strProfile
  End If
  If strPersistType <> "" Then
    ObjScript.writeline "  template persist " & strPersistType & " " & strPersist
  End If
  If strAflex <> "" Then
    ObjScript.writeline "  aflex " & strAflex
  End If
  ObjScript.writeline " exit"
  ObjScript.writeline "exit"
  objBaseline.writeline "show slb virtual-server " & strVIPName & " all-partitions"
  objValidate.writeline "show slb virtual-server " & strVIPName & " all-partitions"
  objRollback.writeline "no slb virtual-server " & strVIPName
  x = x + 1
Loop Until Worksheets("VIP details").Cells(x, 1).Value = ""
  objBaseline.writeline "!"
  objBaseline.writeline "! *** Verify above virtual server does not exist ***"
  objBaseline.writeline "!"
  objValidate.writeline "!"
  objValidate.writeline "! *** Verify above virtual server exist ***"
  objValidate.writeline "!"

objBaseline.write strBaseTemp
objValidate.write strValtemp
objRollback.write strRBTemp
      
      objBaseline.writeline "!"
      objBaseline.writeline "! *** Verify above NAT pool do not exist ***"
      objBaseline.writeline "!"
      objValidate.writeline "!"
      objValidate.writeline "! *** Verify above NAT pool exist ***"
      objValidate.writeline "!"

ObjLog.writeline "cleaning up..."
ObjScript.Close
Set ObjScript = Nothing
Set dictSG = Nothing
Set dictSNAT = Nothing
ObjLog.writeline "All Done at " & Now()
ObjLog.Close
Set ObjLog = Nothing
Set fso = Nothing

'Workbooks.OpenText Filename:=strLogName, Origin:=437, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True

Dim Shex As Object
Set Shex = CreateObject("Shell.Application")
Shex.Open (strLogName)
Set Shex = Nothing
   
End Sub

Function CreatePath(strFullPath)
'-------------------------------------------------------------------------------------------------'
' Function CreatePath (strFullPath)                                                               '
'                                                                                                 '
' This function takes a complete path as input and builds that path out as nessisary.             '
'-------------------------------------------------------------------------------------------------'
Dim pathparts, buildpath, part
    pathparts = Split(strFullPath, "\")
    buildpath = ""
    For Each part In pathparts
        If buildpath <> "" Then
            If buildpath = "\" Then
                buildpath = buildpath & part
            Else
                buildpath = buildpath & "\" & part
            End If
            If Not fso.FolderExists(buildpath) Then
                fso.CreateFolder (buildpath)
            End If
        Else
            If part = "" Then
                buildpath = "\"
            Else
                buildpath = part
            End If
        End If
    Next
End Function


