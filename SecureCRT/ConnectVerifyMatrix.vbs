Option Explicit
public TestMode
	TestMode = False

Sub main
dim SaveFileName, AutoCloseResults, AutoSaveResults, wbNameIn, SavePath, DateInFileName, user
'|----------------------------------------------------------------------------------------------------------|
'|  This script will take the specified input file and generate a matrix with the site and interface type   |
'|  Then if it is not in test mode it will attempt to connect to specified device and check the status      |
'|  of the specified interface, then color the cell based on the specified colors for each connect state    |
'|  Author: Siggi Bjarnason                                                                                 |
'|  Authored: 8/16/2012                                                                                     |
'|  Copyright: Siggi Bjarnason 2012                                                                         |
'|----------------------------------------------------------------------------------------------------------|

' User Spefified values, specify values here per your needs

  wbNameIn = "C:\Scripts\VBScript\SecureCRT\MMEInts.xlsx" 'The Excel spreadsheet with all the values
	DateInFileName = True
	AutoSaveResults = True
	AutoCloseResults = False
  SavePath = "C:\Users\sbjarna\Documents\Projects\LTE\Port Audits\" ' Just the path where you want the results stored
  SaveFileName = "MME Port and Console Audit" ' First part of the name you want to call the results file. A timestamp (if specified) plus .xlsx will be appended
  user = "sbjarna" ' The device login username

' Connection state color, standard decimal values for specified color. Also can use RGB(x,x,x) function

  Const NotConfColor = 16777215
  Const ConnectColor = 9498256
  Const NotConnectColor = 255
  Const AdminDownColor = 13882323
  Const ConfIssueColor = 65535
  Const ProblemColor = 16436871

' Non user section, changes to this section can have undesired results
  Dim app, wb, wbin, wsInts, wsDevs, wsCapability, wsSites, wsOut(), wsLog
  Dim objShell, dictSites, dictCapability, dictDevices, dictTabs, dictTypes(), dictOutLines()
  Dim row, x, TabCount, curTabNum, curOutCol, curOutRow, SiteCode, StartTime, ElapseTime
  Dim wbOutName, CleanToday, LogRow, CurColor, passwd, cNum, DevType

  ReDim Preserve dictTypes(1)
  ReDim Preserve dictOutLines(1)

  Const xlHAlignCenter  = -4108
  Const xlHAlignGeneral = 1
  Const xlHAlignJustify = -4130
  Const xlHAlignLeft = -4131
  Const xlHAlignRight = -4152
  Const xlVAlignBottom  = -4107
  Const xlVAlignCenter = -4108
  Const xlVAlignTop = -4160

  CleanToday = replace(Now, "/", "-")
  CleanToday = replace(CleanToday, ":", "-")
  if DateInFileName = True then
  	wbOutName = SavePath & SaveFileName & " " & CleanToday & ".xlsx"
  else
  	wbOutName = SavePath & SaveFileName & ".xlsx"
  end if

  if TestMode <> True then
  	' Prompt for a password instead of embedding it in a script...
  	passwd = crt.Dialog.Prompt("Enter password for " & user , "Login", "", True)
  end if


  Set dictSites = CreateObject("Scripting.Dictionary")
  Set dictCapability = CreateObject("Scripting.Dictionary")
  Set dictDevices = CreateObject("Scripting.Dictionary")
  Set dictTabs = CreateObject("Scripting.Dictionary")
  Set dictTypes(1) = CreateObject("Scripting.Dictionary")
  Set dictOutLines(1) = CreateObject("Scripting.Dictionary")

  Set objShell = CreateObject("WScript.Shell")
  Set app = CreateObject("Excel.Application")
  Set wbin = app.Workbooks.Open (wbNameIn,0,true)
  Set wb = app.Workbooks.Add
  Set wsInts = wbin.Worksheets(1)
  Set wsDevs = wbin.Worksheets(2)
  Set wsCapability = wbin.Worksheets(3)
  Set wsSites = wbin.Worksheets(4)
  Set wsLog = wb.Worksheets(1)

  app.visible = true
  wsLog.name = "Log"
  LogRow = 1

	StartTime = now
  wsLog.Cells(LogRow,1) =  "Start Time: " & now
  LogRow = LogRow + 1

  wsLog.Cells(LogRow,1) =  "Input file: " & wbNameIn
  LogRow = LogRow + 1

  While wb.worksheets.count > 1
  	wb.Worksheets(2).delete
  Wend
  wsLog.Cells(LogRow,1) =  "initializtion complete, deleted all but one sheet in the new workbook"
  LogRow = LogRow + 1

  row = 2
  Do
  	If wsSites.Cells(row,1).Value = "" Then Exit Do
  	If not dictSites.Exists(wsSites.Cells(row,1).value) then
  		dictSites.Add wsSites.Cells(row,1).value, wsSites.Cells(row,2).value
  	End If
  	row = row + 1
  loop

  wsLog.Cells(LogRow,1) = "Imported Sites into dictionary object"
  LogRow = LogRow + 1

  row = 2
  Do
  	If wsDevs.Cells(row,1).Value = "" Then Exit Do
  	If not dictDevices.Exists(wsDevs.Cells(row,1).value) then
  		dictDevices.Add wsDevs.Cells(row,1).value, wsDevs.Cells(row,2).value
  	End If
  	row = row + 1
  loop

  wsLog.Cells(LogRow,1) =  "Imported Devices into dictionary object"
  LogRow = LogRow + 1

  row = 2
  Do
  	If wsCapability.Cells(row,1).Value = "" Then Exit Do
  	If not dictCapability.Exists(wsCapability.Cells(row,1).value) then
  		dictCapability.Add wsCapability.Cells(row,1).Value, row
  	End If
  	row = row + 1
  loop

  wsLog.Cells(LogRow,1) =  "Imported Capabilities into dictionary object"
  LogRow = LogRow + 1

  row = 2
  curTabNum = 0
  curOutCol = 0
  curOutRow = 0

  Do
  	If wsInts.Cells(row,1).Value = "" Then Exit Do
  	TabCount = dictTabs.count
  	SiteCode = Mid(wsInts.Cells(row,2).Value,4,3)
	  wsLog.Cells(LogRow,1) =  "Working on " & wsInts.Cells(row,2).Value & " " & wsInts.Cells(row,3).Value
  	LogRow = LogRow + 1
  	If not dictTabs.Exists(wsInts.Cells(row,1).value) Then
  		curTabNum = TabCount + 2
  		dictTabs.Add wsInts.Cells(row,1).value, curTabNum
  		wb.Sheets.Add ,wb.Worksheets(wb.Worksheets.Count)
  		ReDim Preserve wsOut(curTabNum)
			Set wsOut(curTabNum) = wb.Worksheets(curTabNum)
			wsOut(curTabNum).name = wsInts.Cells(row,1).Value
			With wsOut(curTabNum).Cells(1,1)
			    .value = "Site Name"
  			  With .Font
        		.Name = "Calibri"
        		.Size = 12
        		.Bold = True
        		.Strikethrough = False
        		.Superscript = False
        		.Subscript = False
        		.OutlineFont = False
        		.Shadow = False
        		.Color = 0
        	End With
     		  .Interior.Color = RGB(255,255,255)
	    	  .HorizontalAlignment = xlHAlignCenter
    		  .VerticalAlignment = xlVAlignCenter
      	  .WrapText = False
  		End With
			With wsOut(curTabNum).Cells(1,2)
				.value = "Code"
  				With .Font
    	    		.Name = "Calibri"
    	    		.Size = 12
    	    		.Bold = True
    	    		.Strikethrough = False
    	    		.Superscript = False
    	    		.Subscript = False
    	    		.OutlineFont = False
    	    		.Shadow = False
    	    		.Color = 0
    	    End With
    	 		.Interior.Color = RGB(255,255,255)
	  	  	.HorizontalAlignment = xlHAlignCenter
    			.VerticalAlignment = xlVAlignCenter
    	  	.WrapText = False
  		End With
			ReDim Preserve dictTypes(curTabNum)
			ReDim Preserve dictOutLines(curTabNum)
			Set dictTypes(curTabNum) = CreateObject("Scripting.Dictionary")
			Set dictOutLines(curTabNum) = CreateObject("Scripting.Dictionary")
			wsLog.Cells(LogRow,1) =  "created tab " & wsInts.Cells(row,1).Value
			LogRow = LogRow + 1
		Else
			curTabNum=dictTabs.Item(wsInts.Cells(row,1).value)
			wsOut(curTabNum).activate
			wsLog.Cells(LogRow,1) =  "Working on tab " & wsInts.Cells(row,1).Value
			LogRow = LogRow + 1
  	End If
  	If not dictTypes(curTabNum).Exists(wsInts.Cells(row,4).value) Then
  		curOutCol = 3 + dictTypes(curTabNum).count
  		dictTypes(curTabNum).Add wsInts.Cells(row,4).value, curOutCol
		With wsOut(curTabNum).Cells(1,curOutCol)
			.value = wsInts.Cells(row,4).value
  			With .Font
        		.Name = "Calibri"
        		.Size = 12
        		.Bold = True
        		.Strikethrough = False
        		.Superscript = False
        		.Subscript = False
        		.OutlineFont = False
        		.Shadow = False
        		.Color = 0
        	End With
 	    	.HorizontalAlignment = xlHAlignCenter
    		.VerticalAlignment = xlVAlignCenter
      		.WrapText = False
  		End With
  		wsLog.Cells(LogRow,1) =  "Started colum " & curOutCol & " for type " & wsInts.Cells(row,4).value
  		LogRow = LogRow + 1
  	Else
  		curOutCol = dictTypes(curTabNum).Item(wsInts.Cells(row,4).value)
  		wsLog.Cells(LogRow,1) =  "Looked up colum number for " & wsInts.Cells(row,4).value & " and found it to be " & curOutCol
  		LogRow = LogRow + 1
  	End If
  	If not dictOutLines(curTabNum).Exists(SiteCode) Then
  		curOutRow = 2 + dictOutLines(curTabNum).count
  		dictOutLines(curTabNum).Add SiteCode, curOutRow
		With wsOut(curTabNum).Cells(curOutRow,1)
			.value = dictSites.Item(SiteCode)
  			With .Font
        		.Name = "Calibri"
        		.Size = 11
        		.Bold = True
        		.Strikethrough = False
        		.Superscript = False
        		.Subscript = False
        		.OutlineFont = False
        		.Shadow = False
        		.Color = 0
        	End With
			.HorizontalAlignment = xlHAlignRight
    		.VerticalAlignment = xlVAlignCenter
      		.WrapText = False
  		End With
		With wsOut(curTabNum).Cells(curOutRow,2)
			.value = SiteCode
  			With .Font
        		.Name = "Calibri"
        		.Size = 11
        		.Bold = True
        		.Strikethrough = False
        		.Superscript = False
        		.Subscript = False
        		.OutlineFont = False
        		.Shadow = False
        		.Color = 0
        	End With
			.HorizontalAlignment = xlHAlignCenter
    		.VerticalAlignment = xlVAlignCenter
      		.WrapText = False
   		End With
  		wsLog.Cells(LogRow,1) = "Started row " & curOutRow & " for Site " & dictSites.Item(SiteCode) & " and Code " & SiteCode
  		LogRow = LogRow + 1
  	Else
  		curOutRow = dictOutLines(curTabNum).Item(SiteCode)
  		wsLog.Cells(LogRow,1) = "Looked up colum number for " & SiteCode & " and found it to be " & curOutRow
  		LogRow = LogRow + 1
  	End If
  	DevType = dictDevices.item(wsInts.Cells(row,2).value)
  	cNum = dictCapability.item(DevType)
  	wsLog.Cells(LogRow,1) = "CNum = " & cNum
  	LogRow = LogRow + 1
  	Select Case wsInts.Cells(row,5).value
  		Case "Yes"
  			Select case CheckIntState (wsInts.Cells(row,2).value, wsInts.Cells(row,3).value, user, passwd, cNum, wsCapability, wsLog, LogRow)
  				case "Connected"
		  			curColor = ConnectColor
					case "NotConnect"
		  			curColor = NotConnectColor
					case "AdminDown"
		  			curColor = AdminDownColor
					case else
		  			curColor = ProblemColor
				end select
  		Case "No"
  			curColor = NotConfColor
  		Case "Partial", "Issues"
  			curColor = ConfIssueColor
  		Case Else
  			curColor = ProblemColor
		  	wsLog.Cells(LogRow,1) = "Unknown Configuration Confirmed value " & wsInts.Cells(row,5).value
		  	LogRow = LogRow + 1
  	End Select

  	With wsOut(curTabNum).Cells(curOutRow,curOutCol)
  		.value = wsInts.Cells(row,2).value & " " & wsInts.Cells(row,3).value
  		With .Font
        	.Name = "Calibri"
        	.Size = 10
        	.Strikethrough = False
        	.Superscript = False
        	.Subscript = False
        	.OutlineFont = False
        	.Shadow = False
        	.Color = 0
        End With
    .Interior.Color = curColor
  	End With
  	wsLog.Cells(LogRow,1) =  "Inserted " & wsInts.Cells(row,2).value & " for site " & sitecode & " and type " & wsInts.Cells(row,4).value & " row & col " & curOutRow & "," & curOutCol
  	LogRow = LogRow + 1
  	row = row + 1
  loop

  For x=2 To wb.worksheets.count
    wsLog.Cells(LogRow,1) =  "adjusted col width and adding ledgent to tab #" & x
    LogRow = LogRow + 1
  	wsOut(x).Cells(dictOutLines(x).count + 3,1).value = "Not Configured"
  	wsOut(x).Cells(dictOutLines(x).count + 3,1).Interior.Color = NotConfColor
		wsOut(x).Cells(dictOutLines(x).count + 4,1).value = "connect"
		wsOut(x).Cells(dictOutLines(x).count + 4,1).Interior.Color = ConnectColor
		wsOut(x).Cells(dictOutLines(x).count + 5,1).value = "not connect"
		wsOut(x).Cells(dictOutLines(x).count + 5,1).Interior.Color = 255
		wsOut(x).Cells(dictOutLines(x).count + 6,1).value = "admin disabled"
		wsOut(x).Cells(dictOutLines(x).count + 6,1).Interior.Color = AdminDownColor
		wsOut(x).Cells(dictOutLines(x).count + 7,1).value = "Configured with issues"
		wsOut(x).Cells(dictOutLines(x).count + 7,1).Interior.Color = ConfIssueColor
		wsOut(x).Cells(dictOutLines(x).count + 8,1).value = "Unknown Problem"
		wsOut(x).Cells(dictOutLines(x).count + 8,1).Interior.Color = ProblemColor
 		wsOut(x).Range(wsOut(x).Cells(1, 1), wsOut(x).Cells(1, 5 + dictTypes(x).count)).EntireColumn.AutoFit
  Next
  wsOut(2).activate

  wsLog.Cells(LogRow,1) =  "Completion Time: " & now
  LogRow = LogRow + 1
  ElapseTime = datediff("n", StartTime, now)
  wsLog.Cells(LogRow,1) =  "Elapse Time: " & ElapseTime & " minutes."
  LogRow = LogRow + 1

  wsLog.visible = False
  wsLog.Cells(1,1).EntireColumn.AutoFit

  if AutoSaveResults = True then
    msgbox "Done. Now saving to " & wbOutName
  	wb.SaveAs(wbOutName)
  else
  	  msgbox "Done. Remeber to save "
  end if
  wbin.Close

  if AutoCloseResults = True then
  	wb.Close
  	app.Quit
  end if

  Set wb = Nothing
  Set wsInts = Nothing
  Set wsDevs = Nothing
  Set wsCapability = Nothing
  Set wsSites = Nothing
  Set wsLog = Nothing
  Set app = Nothing

end sub

function CheckIntState (host, intname, user, pass, cLineNum, wsCap, wsLog, LogRow)
dim ConCmd, IntCmd, ConType, Connected, NotConnect, AdminDown, x, screenrow, tmp
dim readline, Prompt, LogOut, nError, strErr, IntShort, AscValue, ConOpen

	ConType    = wsCap.Cells(cLineNum,2).Value
	IntCmd     = wsCap.Cells(cLineNum,3).Value
	Connected  = wsCap.Cells(cLineNum,4).Value
	NotConnect = wsCap.Cells(cLineNum,5).Value
	AdminDown  = wsCap.Cells(cLineNum,6).Value
	Prompt     = wsCap.Cells(cLineNum,7).Value
	LogOut     = wsCap.Cells(cLineNum,8).Value

	CheckIntState = "Unknown"
	crt.screen.synchronous = true

 	wsLog.Cells(LogRow,1) = "In CheckIntState, ConType = " & ConType
 	LogRow = LogRow + 1

	x = 1
	while (not IsNumeric(mid(intname,x,1))) and (x < len(intname))
		x = x + 1
	wend
	IntShort = mid(intname,x)

	if wsCap.Cells(cLineNum,1).Value = "CatOS" then
		if not IsNumeric(left(intname,1)) then
		 	wsLog.Cells(LogRow,1) = "Device is Catos and intname doesn't start with a number. Intname: " & intname
 			LogRow = LogRow + 1
			intname = IntShort
		 	wsLog.Cells(LogRow,1) = "Intname corrected to : " & intname
 			LogRow = LogRow + 1
		end if
	end if

	if TestMode = True then
		CheckIntState = "Just Testing"
		exit function
	end if
on error resume next
	If crt.Session.Connected Then crt.Session.Disconnect
	select case ConType
		case "SSH2"
			ConCmd = "/SSH2 /L " & user & " /PASSWORD " & pass & " " & host
		  crt.Session.Connect ConCmd
		case "SSH1"
			ConCmd = "/SSH1 /L " & user & " /PASSWORD " & pass &  " " & host
		  crt.Session.Connect ConCmd
		Case "Telnet"
		  crt.Session.Connect "/TELNET " & host & " 23"
  		crt.Screen.WaitForString "name:"
	  	crt.Screen.Send user & vbCr
  		crt.Screen.WaitForString "assword:"
  		crt.Screen.Send pass & vbCr
		Case Else
			CheckIntState = "Unknown connection protocol"
			exit function
	end select
	nError = Err.Number
	strErr = Err.Description
  If nError <> 0 Then
 		wsLog.Cells(LogRow,1) = "Error  " & nError & " occured: " & strErr
 		LogRow = LogRow + 1
 	end if
  If crt.Session.Connected Then
  	crt.Screen.WaitForString Prompt
		If wsCap.Cells(cLineNum,1).Value = "LTConsole" Then
			if not IsNumeric(left(intname,1)) then
			 	wsLog.Cells(LogRow,1) = "Device is a console and intname doesn't start with a number. Intname: " & intname
 				LogRow = LogRow + 1
				intname = IntShort
			 	wsLog.Cells(LogRow,1) = "Intname corrected to : " & intname
 				LogRow = LogRow + 1
			end if
  		crt.Screen.Send(IntCmd & " " & intname & vbcr )
  		ConOpen = false
  		if crt.Screen.WaitForString("Escape sequence is ESC A" & vbcr, 15) then
  			ConOpen = True
  			crt.Screen.Send(vbcr)
 				crt.Screen.WaitForString vbcr, 2
  			while crt.Screen.WaitForString (vbcr, 2)
  			wend
 				screenrow = crt.screen.CurrentRow
 				readline = trim(crt.Screen.Get(screenrow, 1, screenrow, crt.Screen.Columns ))
 			end if
 			if readline <> "" then
			 	wsLog.Cells(LogRow,1) = "readline : " & readline
 				LogRow = LogRow + 1
 				AscValue = asc(mid(readline,1,1))
			 	wsLog.Cells(LogRow,1) = "ASCII value of first character : " & AscValue
 				LogRow = LogRow + 1
 				if AscValue > 32 and AscValue < 127 then
 					CheckIntState="Connected"
 				else
 					CheckIntState = "NotConnect"
 				end if
			else
			 	wsLog.Cells(LogRow,1) = "readline is blank "
 				LogRow = LogRow + 1
				CheckIntState = "NotConnect"
			end if
			If ConOpen then 
				crt.Screen.Send chr(27) & "A"
			else
				CheckIntState = "Problem"
			end if
	 		crt.Screen.WaitForString Prompt, 10
		else
  		crt.Screen.Send(IntCmd & " " & intname & vbcr )
  		crt.Screen.WaitForString (vbcr)
  		crt.Screen.WaitForString (IntShort)
  		crt.Screen.WaitForString (vbcr)
   		screenrow = crt.screen.CurrentRow
  		readline = trim(crt.Screen.Get(screenrow, 1, screenrow, crt.Screen.Columns ))
 			wsLog.Cells(LogRow,1) = "Read line:  " & readline
 			LogRow = LogRow + 1
			If InStr(readline,Connected) > 0 Then CheckIntState  = "Connected"
			If InStr(readline,NotConnect) > 0 Then CheckIntState = "NotConnect"
			If InStr(readline,AdminDown) > 0 Then CheckIntState  = "AdminDown"
 		end if
 	end if
	If crt.Session.Connected Then crt.Screen.Send LogOut & vbcr

 	wsLog.Cells(LogRow,1) = "Return value:  " & CheckIntState
 	LogRow = LogRow + 1
 	crt.screen.synchronous = false
	on Error goto 0

end function

if TestMode = True then main