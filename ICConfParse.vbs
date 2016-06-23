Option Explicit
Dim inFileObj, strLine, fso, strOut, strParts, strInFileName, strOutPath, objLogFileOut
Dim Priority, SubInt, Broadcast, Collision, Discard, iError, ErrorTraffic, QueDrop, util
Dim strSection, strType, objIntFileOut, objSysFileOut, objMatchFileOut
Dim UseBridging, GenericOI, MaxUptime, NumBridgeVia, NumBridges, RestartTrapNum, RestartTrapWindow, TestingNotification
Dim BackPlane, FreeMemory, BuffMiss, BuffUtil, MemFrag, ProcUtil, FanSpeed, RelTemp, RelVoltage, HighTemp, MaxUtil, MinAvail

Const IntFileName = "\ICIntThresholds.csv"
Const SysFileName = "\ICSysThresholds.csv"
Const LogFileName = "\ICThresholdLog.csv"
Const MatchFileName = "\ICMatchItem.csv"

If WScript.Arguments.Count <> 2 Then
  WScript.Echo "Usage: " & wscript.scriptname & " infilename, outpath"
  WScript.Quit
End If

strInFileName = WScript.Arguments(0)
strOutPath = WScript.Arguments(1)
strSection = ""

Set fso = CreateObject("Scripting.FileSystemObject")
Set objLogFileOut = fso.createtextfile(strOutPath & LogFileName)
logout "Group,Type,Object,Value" '& vbcrlf
Set inFileObj = fso.opentextfile(strInFileName)
Set objSysFileOut = fso.createtextfile(strOutPath & SysFileName)
Set objIntFileOut = fso.createtextfile(strOutPath & IntFileName)
Set objMatchFileOut = fso.createtextfile(strOutPath & MatchFileName)
objIntFileOut.writeline "Group,Type,Priority,SubInt,Broadcast,Collision,Discard,Error,ErrorTraffic,QueDrop,Util"
objsysfileout.writeline "Type,Priority,UseBridging,GenericOI,MaxUptime,NumBridgeVia,NumBridges,RestartTrapNum,RestartTrapWindow,TestingNotification,BackPlane,FreeMemory,BuffMiss,BuffUtil,MemFrag,ProcUtil,FanSpeed,RelTemp,RelVoltage,HighTemp,MaxUtil,MinAvail"
objmatchfileout.writeline "Group,Type,MatchCriteria"
While not infileobj.atendofstream
	strLine = Trim(inFileObj.readline)
	strLine = replace(strLine, vbtab," ")
	'logout "-- " & strLine
		If Left(strline,33) = "# Start of Configuration Group - " Then
			strSection = Mid(strline,34)
			If Len(strSection) > 20 Then
				If Left(strSection,14) = "Port Groups - " Then strsection = Mid(strSection,15)
			End If
			'logout "** strsection: " & strsection
		End If
		If strSection <> "" And strSection <> "Polling Groups" Then
			If Left(strline,6) = "config" Then
        If strType = "" Then strType = Trim(Mid(strline,7))
				'logout "** section, type: " & strSection & "," & strType
				If IsArray(strparts) Then 
					If Left(strsection,6) <> "System" Then 
						objIntFileOut.writeline strSection & "," & strType & "," & Priority & "," & SubInt & "," &  Broadcast & "," & Collision & "," & Discard & "," & iError & "," & ErrorTraffic & "," & QueDrop & "," & util
						'logout "++ writing to intfile: " & strSection & "," & strType & "," & Priority & "," & SubInt & "," &  Broadcast & "," & Collision & "," & Discard & "," & iError & "," & ErrorTraffic & "," & QueDrop & "," & util
					End If
					If strSection = "System Resource Groups" Then 
						objSysFileOut.write strType & "," & Priority & "," & UseBridging & "," & GenericOI & "," & MaxUptime & "," & NumBridgeVia & "," & NumBridges & "," & RestartTrapNum & "," & RestartTrapWindow & "," & TestingNotification & ","
						objsysfileout.writeline BackPlane & "," & FreeMemory & "," & BuffMiss & "," & BuffUtil & "," & MemFrag & "," & ProcUtil & "," & FanSpeed & "," & RelTemp & "," & RelVoltage & "," & HighTemp & "," & MaxUtil & "," & MinAvail
					End If 
				  strType = Trim(Mid(strline,7))
          Priority = ""
          SubInt = ""
          Broadcast = ""
          Collision = ""
          Discard = ""
          iError = ""
          ErrorTraffic = ""
          QueDrop = ""
          util = ""
          Priority = ""
          UseBridging = ""
          GenericOI = ""
          MaxUptime = ""
          NumBridgeVia = ""
          NumBridges = ""
          RestartTrapNum = ""
          RestartTrapWindow = ""
          TestingNotification = ""
          BackPlane = ""
          FreeMemory = ""
          BuffMiss = ""
          BuffUtil = ""
          MemFrag = ""
          ProcUtil = ""
          FanSpeed = ""
          RelTemp = ""
          RelVoltage = ""
          HighTemp = ""
          MaxUtil = ""
          MinAvail = ""
				End If
			End If
			If Left(strline,5) = "match" Then
				objmatchfileout.writeline strSection & "," & strType & "," & Trim(Mid(strline,6))
			End If 
			If Left(strline,5) = "param" Then
				strParts = split(strline, " ")
				logout strSection & "," & strType & "," & strparts(1)  & "," & strparts(2)
				Select Case strparts(1)
					case "Priority"
						Priority=strparts(2)
					case "AnalysisModeOfSubInterfacePerformance"
						SubInt = strparts(2)
					case "BroadcastThreshold"
						Broadcast = strparts(2)
					case "CollisionThreshold"
						Collision = strparts(2)
					case "DiscardThreshold"
						Discard = strparts(2)
					case "ErrorThreshold"
						iError = strparts(2)
					case "ErrorTrafficThreshold"
						ErrorTraffic = strparts(2)
					case "QueueDropThreshold"
						QueDrop = strparts(2)
					case "UtilizationThreshold"
						util = strparts(2)
					case "CorrelationUseBridgingMode"
						UseBridging = strparts(2)
					case "EnableGenericOIEvent"
						GenericOI = strparts(2)
					case "MaxUpTimeThreshold"
						MaxUptime = strparts(2)
					case "NumberOfBridgedViaThreshold"
						NumBridgeVia = strparts(2)
					case "NumberOfBridgesThreshold"
						NumBridges = strparts(2)
					case "RestartTrapThreshold"
						RestartTrapNum = strparts(2)
					case "RestartTrapWindow"
						RestartTrapWindow = strparts(2)
					case "TestingNotificationMode"
						TestingNotification = strparts(2)
					case "BackplaneUtilizationThreshold"
						BackPlane = strparts(2)
					case "FreeMemoryThreshold"
						FreeMemory = strparts(2)
					case "MemoryBufferMissThreshold"
						BuffMiss = strparts(2)
					case "MemoryBufferUtilizationThreshold"
						BuffUtil = strparts(2)
					case "MemoryFragmentationThreshold"
						MemFrag = strparts(2)
					case "ProcessorUtilizationThreshold"
						ProcUtil = strparts(2)
					case "FanSpeedThreshold"
						FanSpeed = strparts(2)
					case "RelativeTemperatureThreshold"
						RelTemp = strparts(2)
					case "RelativeVoltageThreshold"
						RelVoltage = strparts(2)
					case "HighTemperatureThreshold"
						HighTemp = strparts(2)
					case "MaxUtilizationPct"
						MaxUtil = strparts(2)
					case "MinAvailableSpace"
						MinAvail = strparts(2)						
				End Select 
			End If
		End If
Wend
If strOut <> "" Then
	'wscript.echo strOut
	logout strOut
End If
inFileObj.close
strOut = ""

objLogFileOut.close
objsysfileout.close
objintfileout.close
objmatchfileout.close

Set inFileObj = Nothing
Set objSysFileOut = Nothing
Set objIntFileOut = Nothing
Set objMatchFileOut = Nothing

Set inFileObj = nothing
Set objLogFileOut = nothing
Set fso = nothing

Sub logout(strText)
	wscript.echo strText
	objLogFileOut.writeline strText
End Sub
