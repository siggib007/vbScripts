Option Explicit

Dim FileObj, strLine, fso, f, fc, f1, strParts, FolderSpec, strOutFileName, objFileOut
Dim app, wb, ws4, ws6, wsAPN, strPCFName, strContext, strPoolGroup, bExcelMode, strAns
Dim strAPN_Map, x, dictSites, dictSubnets, strAPN, row4, row6, strTemp, strSiteParts
Dim oRange, dictLocation, dictAG, dictGA, agRow, gaRow, iLoc, strPCFContext, bLog

'-------------------------------------------------------------------------------------------------'
'  This script will loop through all the files in the provided input directory of Cisco ASR 5000  '
'  GGSN configurations and parse out all the important details about handset pool configurations  '
'                                                                                                 '
'  Author: Siggi Bjarnason                                                                        '
'  Date: 8/27/2012                                                                                '
'  Usage:  parser inpath outfilename                                                              '
'          parser: The name of this script                                                        '
'          inpath: The complete path of the folder where the ASR 5000 configurations are stored   '
'          outfilename: The name of the file you want the results saved to, with complete path    '
'  Example: cscript parser.vbs C:\Cisco5KCfg c:\temp\ASR5KPools.xlsx                                  '
'-------------------------------------------------------------------------------------------------'

Const strLink = "http://docs.eng.t-mobile.com/InfoRouter/docs/~D3405019"
const DefFolder = "C:\HPNAConfigs\GGSN"
const DefFile = "C:\Users\sbjarna\Documents\IP Projects\ASR5KPools.xlsx"
'const DefFile = "C:\temp\ASR5KPools.xlsx"

Const xlHAlignCenter  = -4108
Const xlHAlignGeneral = 1
Const xlHAlignJustify = -4130
Const xlHAlignLeft = -4131
Const xlHAlignRight = -4152
Const xlVAlignBottom  = -4107
Const xlVAlignCenter = -4108
Const xlVAlignTop = -4160
Const xlExcel8 = 56

bLog = false
if Wscript.Arguments.Count > 0 then
	FolderSpec = WScript.Arguments(0)
end if
if Wscript.Arguments.Count > 1 then
	strOutFileName = WScript.Arguments(1)
end if
if Wscript.Arguments.Count > 2 then
	if lcase(WScript.Arguments(2)) = "log" then bLog = true
end if

if FolderSpec = "" then
	FolderSpec = UserInput("No input folder was specified. Please provide input" & vbcrlf & _
								"folder or leave blank to use the default of " & DefFolder & "." & vbcrlf & "Your Input:")
	if FolderSpec = "" then FolderSpec = DefFolder
end if
if strOutFileName = "" then
	strOutFileName = UserInput("Filename to save the results in was not specified. " & vbcrlf & _
										"Please provide file name with full path to save " & vbcrlf & _
										"the results in or leave blank to use the default of " & _
										DefFile & "." & vbcrlf & "Your Input:")
	if strOutFileName = "" then strOutFileName = DefFile
end if

wscript.echo "Reading from directory " & FolderSpec & vbcrlf & "Saving to " & strOutFileName

Set fso = CreateObject("Scripting.FileSystemObject")

while not fso.FolderExists(FolderSpec)
	FolderSpec = UserInput("Folder " & FolderSpec & " is not valid. Please provide new one or leave blank to abort:")
	if FolderSpec = "" then
		wscript.echo "No input provided, aborting"
		wscript.quit (1)
	end if
wend


strTemp = Mid(strOutFileName, 1, InStrRev(strOutFileName, "\"))
while not fso.FolderExists(strTemp) 
	strAns = AskYesNo ("Folder " & strTemp & " doesn't exists, should I create it?")
	if strAns = "Yes" then
		fso.CreateFolder(strTemp)
		wscript.echo "Folder " & strTemp & " created. "
	else
		strOutFileName = UserInput("Path in " & strOutFileName & " is not valid. Please provide new one or leave blank to abort:")
		if strOutFileName = "" then
			wscript.echo "No input provided, aborting"
			wscript.quit (1)
		end if
		strTemp = Mid(strOutFileName, 1, InStrRev(strOutFileName, "\"))
	end if
wend


set dictSubnets = CreateObject("Scripting.Dictionary")
set dictSites = CreateObject("Scripting.Dictionary")
set dictLocation = CreateObject("Scripting.Dictionary")
set dictAG = CreateObject("Scripting.Dictionary")
set dictGA = CreateObject("Scripting.Dictionary")

strAPN_Map = "GGSN, Context, APN, Pool Group" & vbcrlf
row4 = 3
row6 = 3
agRow = 3
gaRow = 3

InitializeDicts

strParts = split(strOutFileName,".")
if left(strParts(ubound(strParts)),3) = "xls" then
	on error resume next
	Set app = CreateObject("Excel.Application")
	If Err.Number <> 0 Then
		wscript.echo "Unable to start Excel, probably not installed correctly. Unable to create a xls or xlsx file without it"
		wscript.echo "Please re-start the script using a txt or csv extension on your file. Aborting"
		wscript.quit (1)
	end if
	on error goto 0
	Set wb = app.Workbooks.Add
	Set ws4 = wb.Worksheets(1)
	app.visible = True
	ws4.name = "IPv4"
	ws4.Cells(1,1).value = "For a GGSN to ARG and LBG mapping see:"
	ws4.Range("A1:C1").Merge
	ws4.Cells(1,1).HorizontalAlignment = xlHAlignRight
	ws4.Cells(1,4).value = strLink
	set oRange = ws4.Range(ws4.Cells(1, 4), ws4.Cells(1, 7))
	oRange.merge
	ws4.hyperlinks.add oRange,strLink
	ws4.Cells(2, 1).value = "Site Name"
	ws4.Cells(2, 2).value = "Site Code"
	ws4.Cells(2, 3).value = "GGSN"
	ws4.Cells(2, 4).value = "Context"
	ws4.Cells(2, 5).value = "Name"
	ws4.Cells(2, 6).value = "Address"
	ws4.Cells(2, 7).value = "Type"
	ws4.Cells(2, 8).value = "Group"
	ws4.Cells(2, 9).value = "APN"
	With ws4.Range(ws4.Cells(2, 1), ws4.Cells(2, 9))
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

	wb.Sheets.Add ,wb.Worksheets(wb.Worksheets.Count)
	Set ws6 = wb.Worksheets(2)
	ws6.name = "IPv6"
	ws6.Cells(1, 1).value = "For a GGSN to ARG and LBG mapping see:"
	ws6.Range("A1:C1").Merge
	ws6.Cells(1,1).HorizontalAlignment = xlHAlignRight
	ws6.Cells(1, 4).value = strLink
	set oRange = ws6.Range(ws6.Cells(1, 4), ws6.Cells(1, 7))
	oRange.merge
	ws6.hyperlinks.add oRange,strLink
	ws6.Cells(2, 1).value = "Site Name"
	ws6.Cells(2, 2).value = "Site Code"
	ws6.Cells(2, 3).value = "GGSN"
	ws6.Cells(2, 4).value = "Context"
	ws6.Cells(2, 5).value = "Name"
	ws6.Cells(2, 6).value = "Address"
	ws6.Cells(2, 7).value = "Type"
	ws6.Cells(2, 8).value = "Group"
	ws6.Cells(2, 9).value = "APN"
	With ws6.Range(ws6.Cells(2, 1), ws6.Cells(2, 9))
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

	wb.Sheets.Add ,wb.Worksheets(wb.Worksheets.Count)
	Set wsAPN = wb.Worksheets(3)
	wsAPN.name = "APN"
	wsAPN.Cells(1,1).value = "Pool to APN map"
  wsAPN.Range(wsAPN.Cells(1, 1), wsAPN.Cells(1, 4)).Merge
	wsAPN.Columns("E:E").ColumnWidth = 3
	wsAPN.Cells(1,6).value = "APN to Pool map"
  wsAPN.Range(wsAPN.Cells(1, 6), wsAPN.Cells(1, 9)).Merge	
	wsAPN.Cells(2, 1).value = "GGSN"
	wsAPN.Cells(2, 2).value = "Context"
	wsAPN.Cells(2, 3).value = "Pool Group"
	wsAPN.Cells(2, 4).value = "APN"
	wsAPN.Cells(2, 6).value = "GGSN"
	wsAPN.Cells(2, 7).value = "Context"
	wsAPN.Cells(2, 8).value = "APN"
	wsAPN.Cells(2, 9).value = "Pool Group"
	With wsAPN.Range(wsAPN.Cells(1, 1), wsAPN.Cells(2, 9))
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
	bExcelMode = True
else
on error resume next
	Set objFileOut = fso.createtextfile(strOutFileName)
	If Err.Number <> 0 Then		
    WScript.Echo "Unable to save to " & strOutFileName
    on error goto 0
    strParts = split(strOutFileName,".")
    strTemp = strParts(ubound(strParts)-1)
    x = 1
    strParts(ubound(strParts)-1) = strTemp & x
    strOutFileName = join(strParts,".")
    while fso.fileexists(strOutFileName)
    	x = x + 1
    	strParts(ubound(strParts)-1) = strTemp & x
    	strOutFileName = join(strParts,".")
    wend
    Err.Clear
    wscript.echo "saving to " & strOutFileName & " instead"
		Set objFileOut = fso.createtextfile(strOutFileName)
	End If
	objFileOut.writeline "For a GGSN to ARG and LBG mapping see: http://docs.eng.t-mobile.com/InfoRouter/docs/~D3405019"
	objFileOut.writeline "IP Version, Site Name, Site Code, GGSN, Context,Name, Address, Type, Group"
end if

Set f = fso.GetFolder(folderspec)
Set fc = f.Files

For Each f1 in fc
	If f1.name <> strOutFileName Then
		Set FileObj = fso.opentextfile(folderspec & "\" & f1.name)
		While not fileobj.atendofstream
			strLine = Trim(FileObj.readline)
			strParts = split(strline," ")
			if left(strline,15) = "system hostname" then
				WriteMapOut
				strPCFName = mid(strline,17)
				if dictSites.exists(strPCFName) then
					strSiteParts=split(dictSites.Item(strPCFName),",")
				else
					strSiteParts=split("n/a,unknown",",")
				end if
				if bLog then wscript.echo "Found Hostname: " & strPCFName
			end if
			if left(strline,7) = "context" then
				if strParts(0) = "context" and strParts(1) <> "schema" then
					WriteMapOut
					strContext = ucase(strParts(1))
					if bLog then wscript.echo "in context " & strContext
				end if
			end if
			if left(strline, 7) = "ip pool" then
				if strParts(5) = "static" then
					strPoolGroup = strParts(7)
				else
					strPoolGroup = strParts(8)
				end if
				if strParts(3) = "range" then
					if bExcelMode then
						ws4.activate
						ws4.Cells(row4, 1).value = strSiteParts(1)
						ws4.Cells(row4, 2).value = strSiteParts(0)
						ws4.Cells(row4, 3).value = strPCFName
						ws4.Cells(row4, 4).value = strContext
						ws4.Cells(row4, 5).value = strParts(2)
						ws4.Cells(row4, 6).value = strParts(4) & "-" & strParts(5)
						ws4.Cells(row4, 7).value = strParts(6)
						ws4.Cells(row4, 8).value = strParts(8)
						ws4.Cells(row4, 9).value = "=VLOOKUP(H" & row4 & ",APN!C:D,2,FALSE)"
						row4 = row4 + 1
					else
						objFileOut.writeline "IPV4," & strSiteParts(1) & "," & strSiteParts(0) & "," & strPCFName & "," & strContext & "," & strParts(2) & "," & strParts(4) & "-" & strParts(5) & "," & strParts(6) & "," & strParts(8)
					end if
				else
					if bExcelMode then
						ws4.activate
						ws4.Cells(row4, 1).value = strSiteParts(1)
						ws4.Cells(row4, 2).value = strSiteParts(0)
						ws4.Cells(row4, 3).value = strPCFName
						ws4.Cells(row4, 4).value = strContext
						ws4.Cells(row4, 5).value = strParts(2)
						ws4.Cells(row4, 6).value = strParts(3) & dictSubnets.Item(strParts(4))
						ws4.Cells(row4, 7).value = strParts(5)
						ws4.Cells(row4, 8).value = strPoolGroup
						ws4.Cells(row4, 9).value = "=VLOOKUP(H" & row4 & ",APN!C:D,2,FALSE)"
						row4 = row4 + 1
					else
						objFileOut.writeline "IPV4," & strSiteParts(1) & "," & strSiteParts(0) & "," & strPCFName & "," & strContext & "," & strParts(2) & "," & strParts(3) & dictSubnets.Item(strParts(4)) & "," & strParts(5) & ","  & strPoolGroup
					end if
				end if
			end if
			if left(strline,9)="ipv6 pool" then
				strParts = split(strline," ")
				if bExcelMode then
					ws6.activate
					ws6.Cells(row6, 1).value = strSiteParts(1)
					ws6.Cells(row6, 2).value = strSiteParts(0)
					ws6.Cells(row6, 3).value = strPCFName
					ws6.Cells(row6, 4).value = strContext
					ws6.Cells(row6, 5).value = strParts(2)
					ws6.Cells(row6, 6).value = strParts(4)
					ws6.Cells(row6, 7).value = strParts(5)
					ws6.Cells(row6, 8).value = strParts(8)
					ws6.Cells(row6, 9).value = "=VLOOKUP(H" & row6 & ",APN!C:D,2,FALSE)"
					row6 = row6 + 1
				else
					objFileOut.writeline "IPV6," & strSiteParts(1) & "," & strSiteParts(0) & "," & strPCFName & "," & strContext & "," & strParts(2) & "," & strParts(4) & "," & strParts(5) & "," & strParts(8)
				end if
			end if
			if left(strline,4)="apn " then
				strAPN = mid(strline,5)
			end if
			if left(strline,20)="ip address pool name" then
				strPoolGroup = mid(strline,22)
					if dictGA.exists(strPoolGroup) then
						dictGA.item(strPoolGroup) = dictGA.item(strPoolGroup) & ", " & strAPN
					else
						dictGA.add strPoolGroup, strAPN
					end if
					if dictAG.exists(strAPN) then
						dictAG.item(strAPN) = dictAG.item(strAPN) & ", " & strPoolGroup
					else
						dictAG.add strAPN, strPoolGroup
					end if
			end if
			if left(strline,24) = "ipv6 address prefix-pool" then
				strPoolGroup = mid(strline,26)
					if dictGA.exists(strPoolGroup) then
						dictGA.item(strPoolGroup) = dictGA.item(strPoolGroup) & ", " & strAPN
					else
						dictGA.add strPoolGroup, strAPN
					end if
					if dictAG.exists(strAPN) then
						dictAG.item(strAPN) = dictAG.item(strAPN) & ", " & strPoolGroup
					else
						dictAG.add strAPN, strPoolGroup
					end if
			end if				
		Wend
		FileObj.close
	End If
Next

WriteMapOut
if bExcelMode = false then objFileOut.write vbcrlf & strAPN_Map

if bExcelMode then
	row4 = 3
	while ws4.cells(row4,3).value <> ""
		strPCFContext = ws4.cells(row4,3).value & "," & ws4.cells(row4,4).value
		if dictLocation.exists(strPCFContext) then
			iLoc = split(dictLocation.item(strPCFContext),",")
			ws4.Cells(row4, 9).value = "=VLOOKUP(H" & row4 & ",APN!C" & iLoc(0) & ":D" & iLoc(1) & ",2,FALSE)"
		else
			if bLog then wscript.echo "no "  & strPCFContext & " in location dict"
		end if
		row4 = row4 + 1
	wend
	row6 = 3
	while ws6.cells(row6,3).value <> ""
		strPCFContext = ws6.cells(row6,3).value & "," & ws6.cells(row6,4).value
		if dictLocation.exists(strPCFContext) then
			iLoc = split(dictLocation.item(strPCFContext),",")
			ws6.Cells(row6, 9).value = "=VLOOKUP(H" & row6 & ",APN!C" & iLoc(0) & ":D" & iLoc(1) & ",2,FALSE)"
		else
			if bLog then wscript.echo "IPV6: no "  & strPCFContext & " in location dict"
		end if
		row6 = row6 + 1
	wend
	
	'ws6.Range(ws6.Cells(1, 1), ws6.Cells(1, 9)).EntireColumn.AutoFit
	with ws6
	   .Range(ws6.Cells(1, 1), ws6.Cells(1, 9)).EntireColumn.AutoFit
	   .activate
	   .cells(1,1).activate
     .Application.ActiveWindow.SplitRow = 2
     .Application.ActiveWindow.FreezePanes = True
	end with 
	'wsAPN.Range(wsAPN.Cells(1, 1), wsAPN.Cells(1, 9)).EntireColumn.AutoFit
	with wsAPN
	   .Range(wsAPN.Cells(1, 1), wsAPN.Cells(1, 9)).EntireColumn.AutoFit
	   .activate
	   .cells(1,1).activate
     .Application.ActiveWindow.SplitRow = 2
     .Application.ActiveWindow.FreezePanes = True
	end with 
	with ws4
	   .Range(ws4.Cells(1, 1), ws4.Cells(1, 9)).EntireColumn.AutoFit
	   .activate
	   .cells(1,1).activate
     .Application.ActiveWindow.SplitRow = 2
     .Application.ActiveWindow.FreezePanes = True
	end with 

	on error resume next
	if right(strOutFileName,5) = ".xlsx" then wb.SaveAs strOutFileName
	if right(strOutFileName,4) = ".xls" then wb.SaveAs strOutFileName, xlExcel8
	If Err.Number <> 0 Then		
    WScript.Echo "Unable to save to " & strOutFileName
    on error goto 0
    Err.Clear
    strParts = split(strOutFileName,".")
    strTemp = strParts(ubound(strParts)-1)
    x = 1
    strParts(ubound(strParts)-1) = strTemp & x
    strOutFileName = join(strParts,".")
    while fso.fileexists(strOutFileName)
    	x = x + 1
    	strParts(ubound(strParts)-1) = strTemp & x
    	strOutFileName = join(strParts,".")
    wend
    wscript.echo "saving to " & strOutFileName & " instead"
		if right(strOutFileName,5) = ".xlsx" then wb.SaveAs strOutFileName
		if right(strOutFileName,4) = ".xls" then wb.SaveAs strOutFileName, xlExcel8
	End If
'	wb.Close
'	app.Quit
	Set ws4 = Nothing
	Set wb = Nothing
	Set app = Nothing
else
	objFileOut.close
	Set objFileOut = nothing
end if

Set FileObj = nothing
Set fc = nothing
Set f = nothing
Set fso = nothing

wscript.echo "Done. Results saved to " & strOutFileName


sub WriteMapOut
	dim strGroup, strAPN, iStart, iEnd, strPCFContext
		if dictGA.count > 0 then 
			strPCFContext = strPCFName & "," & strContext
			iStart = gaRow
			
			if bExcelMode then
				wsAPN.activate
				for each strGroup in dictGA			
					wsAPN.Cells(gaRow, 1).value = strPCFName
					wsAPN.Cells(gaRow, 2).value = strContext
					wsAPN.Cells(gaRow, 3).value = strGroup
					wsAPN.Cells(gaRow, 4).value = dictGA.item(strGroup)
					gaRow = gaRow + 1
				next
			end if
			iEnd = gaRow
			if not dictLocation.exists(strPCFContext) then
				dictLocation.add strPCFContext, iStart & "," & iEnd
			end if 
			dictGA.RemoveAll
		end if
		if dictAG.count > 0 then
			for each strAPN in dictAG	
				if bExcelMode then
					wsAPN.Cells(agRow, 6).value = strPCFName
					wsAPN.Cells(agRow, 7).value = strContext
					wsAPN.Cells(agRow, 8).value = strAPN
					wsAPN.Cells(agRow, 9).value = dictAG.item(strAPN)
					agRow = agRow + 1
				end if
				strAPN_Map = strAPN_Map & strPCFName & "," & strContext & "," & strAPN  & "," &  replace(dictAG.item(strAPN),", ",";") & vbcrlf
			next
			dictAG.RemoveAll
		end if
end sub

sub InitializeDicts
	dictSubnets.add "255.255.255.255", "/32"
	dictSubnets.add "255.255.255.254", "/31"
	dictSubnets.add "255.255.255.252", "/30"
	dictSubnets.add "255.255.255.248", "/29"
	dictSubnets.add "255.255.255.240", "/28"
	dictSubnets.add "255.255.255.224", "/27"
	dictSubnets.add "255.255.255.192", "/26"
	dictSubnets.add "255.255.255.128", "/25"
	dictSubnets.add "255.255.255.0", "/24"
	dictSubnets.add "255.255.254.0", "/23"
	dictSubnets.add "255.255.252.0", "/22"
	dictSubnets.add "255.255.248.0", "/21"
	dictSubnets.add "255.255.240.0", "/20"
	dictSubnets.add "255.255.224.0", "/19"
	dictSubnets.add "255.255.192.0", "/18"
	dictSubnets.add "255.255.128.0", "/17"
	dictSubnets.add "255.255.0.0", "/16"
	dictSubnets.add "255.254.0.0", "/15"
	dictSubnets.add "255.252.0.0", "/14"
	dictSubnets.add "255.248.0.0", "/13"
	dictSubnets.add "255.240.0.0", "/12"
	dictSubnets.add "255.224.0.0", "/11"
	dictSubnets.add "255.192.0.0", "/10"
	dictSubnets.add "255.128.0.0", "/9"
	dictSubnets.add "255.0.0.0", "/8"
	dictSubnets.add "254.0.0.0", "/7"
	dictSubnets.add "252.0.0.0", "/6"
	dictSubnets.add "248.0.0.0", "/5"
	dictSubnets.add "240.0.0.0", "/4"
	dictSubnets.add "224.0.0.0", "/3"
	dictSubnets.add "192.0.0.0", "/2"
	dictSubnets.add "128.0.0.0", "/1"

	dictSites.add "ATPCF000", "ATL,ATLANTA RNOC"
	dictSites.add "ATPCF001", "ATL,ATLANTA RNOC"
	dictSites.add "CHPCF000", "ELG,ELGIN MSO"
	dictSites.add "CHPCF001", "ELG,ELGIN MSO"
	dictSites.add "CRPCF000", "CHR,CHARLOTTE SERVICE ST MSO"
	dictSites.add "DAPCF000", "DAL,DALLAS SWITCH"
	dictSites.add "DAPCF001", "DAL,DALLAS SWITCH"
	dictSites.add "DCPCF000", "BLT,BELTSVILLE MD SWITCH"
	dictSites.add "DEPCF000", "DET,LIVONIA (Detroit) SWITCH"
	dictSites.add "DNPCF000", "DEN,DENVER SWITCH"
	dictSites.add "HNPCF000", "HOU,HOUSTON SWITCH"
	dictSites.add "LAPCF000", "IRV,IRVINE MSC"
	dictSites.add "LAPCF001", "RVS,RIVERSIDE CA SWITCH"
	dictSites.add "NEPCF000", "NRT,NORTON SWITCH"
	dictSites.add "NYPCF000", "WAY,WAYNE SWITCH"
	dictSites.add "NYPCF001", "SYO,SYOSSET, NY MSO"
	dictSites.add "NYPCF003", "MAN,MANHATTAN SWITCH"
	dictSites.add "ORPCF000", "ORL,ORLANDO SWITCH"
	dictSites.add "PHPCF000", "PHI,PHILADELPHIA SWITCH"
	dictSites.add "PHPCF001", "PHI,PHILADELPHIA SWITCH"
	dictSites.add "PXPCF000", "O11,ORION DATACENTER"
	dictSites.add "SCPCF000", "WSC,WEST SACRAMENTO MSC"
	dictSites.add "SEPCF000", "SNQ,SNOQUALMIE Datacenter"
	dictSites.add "SEPCF001", "SNQ,SNOQUALMIE Datacenter"
	dictSites.add "SEPCF002", "SNQ,SNOQUALMIE Datacenter"
	dictSites.add "NVPCF000", "NVL,NASHVILLE 2 MSC"
	dictSites.add "ATPCF002","ATL,ATLANTA RNOC"
	dictSites.add "CHPCF002","ELG,ELGIN MSO"
	dictSites.add "CRPCF001","CHR,CHARLOTTE SERVICE ST MSO"
	dictSites.add "DAPCF002","DAL,DALLAS SWITCH"
	dictSites.add "DCPCF001","BLT,BELTSVILLE MD SWITCH"
	dictSites.add "DEPCF001","DET,LIVONIA (Detroit) SWITCH"
	dictSites.add "DNPCF001","DEN,DENVER SWITCH"
	dictSites.add "HNPCF001","HOU,HOUSTON SWITCH"
	dictSites.add "LAPCF002","RVS,RIVERSIDE CA SWITCH"
	dictSites.add "LAPCF003","IRV,IRVINE MSC"
	dictSites.add "NEPCF001","NRT,NORTON SWITCH"
	dictSites.add "NVPCF001","NVL,NASHVILLE 2 MSC"
	dictSites.add "NYPCF002","SYO,SYOSSET, NY MSO"
	dictSites.add "NYPCF004","WAY,WAYNE SWITCH"
	dictSites.add "NYPCF005","MAN,MANHATTAN SWITCH"
	dictSites.add "ORPCF001","ORL,ORLANDO SWITCH"
	dictSites.add "PHPCF002","PHI,PHILADELPHIA SWITCH"
	dictSites.add "SCPCF001","WSC,WEST SACRAMENTO MSC"
	dictSites.add "SEPCF003","SNQ,SNOQUALMIE Datacenter"
	dictSites.add "PHPCF001(Engg)","PHI,PHILADELPHIA SWITCH"
	dictSites.add "CHPCF003","ELG,ELGIN MSO"
	dictSites.add "CHPCF004","ELG,ELGIN MSO"
	dictSites.add "CHPCF005","ELG,ELGIN MSO"
	dictSites.add "CHPCF006","ELG,ELGIN MSO"
	dictSites.add "CHPCF007","ELG,ELGIN MSO"
	dictSites.add "CHPCF009","ELG,ELGIN MSO"
	dictSites.add "CRPCF002","CHR,CHARLOTTE SERVICE ST MSO"
	dictSites.add "CRPCF003","CHR,CHARLOTTE SERVICE ST MSO"
	dictSites.add "CRPCF004","CHR,CHARLOTTE SERVICE ST MSO"
	dictSites.add "CRPCF005","CHR,CHARLOTTE SERVICE ST MSO"
	dictSites.add "CRPCF006","CHR,CHARLOTTE SERVICE ST MSO"
	dictSites.add "CRPCF007","CHR,CHARLOTTE SERVICE ST MSO"
	dictSites.add "CRPCF008","CHR,CHARLOTTE SERVICE ST MSO"
	dictSites.add "DAPCF003","DAL,DALLAS SWITCH"
	dictSites.add "DAPCF004","DAL,DALLAS SWITCH"
	dictSites.add "DAPCF005","DAL,DALLAS SWITCH"
	dictSites.add "DAPCF006","DAL,DALLAS SWITCH"
	dictSites.add "DAPCF007","DAL,DALLAS SWITCH"
	dictSites.add "DAPCF008","DAL,DALLAS SWITCH"
	dictSites.add "DAPCF011","DAL,DALLAS SWITCH"
	dictSites.add "DCPCF002","BLT,BELTSVILLE MD SWITCH"
	dictSites.add "DCPCF003","BLT,BELTSVILLE MD SWITCH"
	dictSites.add "DCPCF004","BLT,BELTSVILLE MD SWITCH"
	dictSites.add "DCPCF005","BLT,BELTSVILLE MD SWITCH"
	dictSites.add "DCPCF006","BLT,BELTSVILLE MD SWITCH"
	dictSites.add "DEPCF002","DET,LIVONIA (Detroit) SWITCH"
	dictSites.add "DEPCF003","DET,LIVONIA (Detroit) SWITCH"
	dictSites.add "DEPCF004","DET,LIVONIA (Detroit) SWITCH"
	dictSites.add "DNPCF002","DEN,DENVER SWITCH"
	dictSites.add "DNPCF004","DEN,DENVER SWITCH"
	dictSites.add "DNPCF005","DEN,DENVER SWITCH"
	dictSites.add "HIPCF000","HON,KOAPAKA HI SWITCH"
	dictSites.add "HNPCF002","HOU,HOUSTON SWITCH"
	dictSites.add "HNPCF003","HOU,HOUSTON SWITCH"
	dictSites.add "HNPCF004","HOU,HOUSTON SWITCH"
	dictSites.add "HNPCF005","HOU,HOUSTON SWITCH"
	dictSites.add "HNPCF006","HOU,HOUSTON SWITCH"
	dictSites.add "LAPCF004","IRV,IRVINE MSC"
	dictSites.add "LAPCF005","RVS,RIVERSIDE CA SWITCH"
	dictSites.add "LAPCF006","RVS,RIVERSIDE CA SWITCH"
	dictSites.add "LAPCF007","RVS,RIVERSIDE CA SWITCH"
	dictSites.add "LAPCF008","IRV,IRVINE MSC"
	dictSites.add "LAPCF009","IRV,IRVINE MSC"
	dictSites.add "LAPCF010","IRV,IRVINE MSC"
	dictSites.add "LAPCF011","IRV,IRVINE MSC"
	dictSites.add "LAPCF012","IRV,IRVINE MSC"
	dictSites.add "LAPCF013","RVS,RIVERSIDE CA SWITCH"
	dictSites.add "LAPCF014","RVS,RIVERSIDE CA SWITCH"
	dictSites.add "LAPCF015","IRV,IRVINE MSC"
	dictSites.add "LAPCF016","IRV,IRVINE MSC"
	dictSites.add "LAPCF017","IRV,IRVINE MSC"
	dictSites.add "LAPCF018","RVS,RIVERSIDE CA SWITCH"
	dictSites.add "LAPCF019","RVS,RIVERSIDE CA SWITCH"
	dictSites.add "LAPCF020","RVS,RIVERSIDE CA SWITCH"
	dictSites.add "NEPCF002","NRT,NORTON SWITCH"
	dictSites.add "NEPCF003","NRT,NORTON SWITCH"
	dictSites.add "NEPCF004","NRT,NORTON SWITCH"
	dictSites.add "NEPCF005","NRT,NORTON SWITCH"
	dictSites.add "NVPCF002","NVL,NASHVILLE 2 MSC"
	dictSites.add "NVPCF003","NVL,NASHVILLE 2 MSC"
	dictSites.add "NVPCF004","NVL,NASHVILLE 2 MSC"
	dictSites.add "NVPCF005","NVL,NASHVILLE 2 MSC"
	dictSites.add "NVPCF006","NVL,NASHVILLE 2 MSC"
	dictSites.add "NYPCF006","SYO,SYOSSET, NY MSO"
	dictSites.add "NYPCF007","SYO,SYOSSET, NY MSO"
	dictSites.add "NYPCF008","SYO,SYOSSET, NY MSO"
	dictSites.add "NYPCF009","WAY,WAYNE SWITCH"
	dictSites.add "NYPCF010","WAY,WAYNE SWITCH"
	dictSites.add "NYPCF011","SYO,SYOSSET, NY MSO"
	dictSites.add "NYPCF012","SYO,SYOSSET, NY MSO"
	dictSites.add "NYPCF013","WAY,WAYNE SWITCH"
	dictSites.add "NYPCF014","WAY,WAYNE SWITCH"
	dictSites.add "NYPCF015","WAY,WAYNE SWITCH"
	dictSites.add "ORPCF002","ORL,ORLANDO SWITCH"
	dictSites.add "ORPCF003","ORL,ORLANDO SWITCH"
	dictSites.add "ORPCF004","ORL,ORLANDO SWITCH"
	dictSites.add "ORPCF005","ORL,ORLANDO SWITCH"
	dictSites.add "ORPCF006","ORL,ORLANDO SWITCH"
	dictSites.add "ORPCF007","ORL,ORLANDO SWITCH"
	dictSites.add "ORPCF008","ORL,ORLANDO SWITCH"
	dictSites.add "ORPCF009","ORL,ORLANDO SWITCH"
	dictSites.add "PHPCF003","PHI,PHILADELPHIA SWITCH"
	dictSites.add "PHPCF004","PHI,PHILADELPHIA SWITCH"
	dictSites.add "PHPCF005","PHI,PHILADELPHIA SWITCH"
	dictSites.add "PHPCF006","PHI,PHILADELPHIA SWITCH"
	dictSites.add "PHPCF007","PHI,PHILADELPHIA SWITCH"
	dictSites.add "PRPCF000","BYM,BAYAMON PR  MSO"
	dictSites.add "SCPCF002","WSC,WEST SACRAMENTO MSC"
	dictSites.add "SCPCF003","WSC,WEST SACRAMENTO MSC"
	dictSites.add "SCPCF004","WSC,WEST SACRAMENTO MSC"
	dictSites.add "SCPCF005","WSC,WEST SACRAMENTO MSC"
	dictSites.add "SCPCF006","WSC,WEST SACRAMENTO MSC"
	dictSites.add "SCPCF007","WSC,WEST SACRAMENTO MSC"
	dictSites.add "SCPCF008","WSC,WEST SACRAMENTO MSC"
	dictSites.add "SCPCF009","WSC,WEST SACRAMENTO MSC"
	dictSites.add "SCPCF010","WSC,WEST SACRAMENTO MSC"
	dictSites.add "SEPCF005","PLR, Polaris MSO"
	dictSites.add "SEPCF006","PLR, Polaris MSO"
	dictSites.add "SEPCF007","PLR, Polaris MSO"
	dictSites.add "SEPCF008","PLR, Polaris MSO"
	dictSites.add "SEPCF009","PLR, Polaris MSO"
	dictSites.add "SEPCF011","PLR, Polaris MSO"
	dictSites.add "SEPCF012","PLR, Polaris MSO"
	dictSites.add "SEPCF013","PLR, Polaris MSO"
	dictSites.add "VGPCF000","NLV,LAS VEGAS NV MSC"
	dictSites.add "VGPCF001","NLV,LAS VEGAS NV MSC"
	dictSites.add "VGPCF002","NLV,LAS VEGAS NV MSC"
	dictSites.add "VGPCF003","NLV,LAS VEGAS NV MSC"
	dictSites.add "VGPCF004","NLV,LAS VEGAS NV MSC"
	dictSites.add "VGPCF005","NLV,LAS VEGAS NV MSC"
	dictSites.add "VGPCF006","NLV,LAS VEGAS NV MSC"
	dictSites.add "VGPCF007","NLV,LAS VEGAS NV MSC"
end sub

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
	Dim strInput 
    If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
        WScript.StdOut.Write myPrompt & " "
        strInput = WScript.StdIn.ReadLine
        if left(ucase(strInput),1)="Y" then
        	AskYesNo = "Yes"
        else
        	AskYesNo = "No"
        end if
    Else
        strInput = Msgbox(myPrompt, vbYesNo, "Question for you")
				If strInput = vbYes Then
       		AskYesNo = "Yes"
        else
        	AskYesNo = "No"
        end if
    End If
End Function
