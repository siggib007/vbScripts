'* ********** ********** ********** ********** ********** ********** **
' *** If the provided string is less than the length
'					given, the returned string will be padded with
'					zeros on it's left side
'  strPad		- the string that needs to be padded
'  lngLen		- the minimum length that the string should be 
'					when it's returned
'* ********** ********** ********** ********** ********** ********** */
Function FillLeft(strPad, lngLen)
	If lngLen > Len(strPad) Then
		FillLeft = String(lngLen - Len(strPad), Asc("0")) & strPad
	Else
		FillLeft = strPad
	End If
End Function

'* ********** ********** ********** ********** ********** ********** **
' *** Takes a date object and returns a file name based on
'					on the following equation:
'		Year Month Day '.txt'
'  datetime		- The date that the file name will be based on
'* ********** ********** ********** ********** ********** ********** */
Function FormatDate(datetime)
	FormatDate = FillLeft(DatePart("yyyy", datetime), 4) & FillLeft(DatePart("m", datetime), 2) & FillLeft(DatePart("d", datetime), 2) & ".txt"
End Function

'* ********** ********** ********** ********** ********** ********** **
' *** Create ADO parameters for TRNDdpEnumNetwork stored procedure
'  Cmd	- The Command object that the parameters are to be added to.
'  Con	- The Connection object that will be used.
'* ********** ********** ********** ********** ********** ********** **
Function CreateEnumRegionGroupParams(Cmd, Con)
	On Error Resume Next

	CreateEnumRegionGroupParams = False

	Cmd.Parameters.Append Cmd.CreateParameter("RETURN_VALUE",  adInteger, adParamReturnValue, 4)
	If Not ADO_CheckErr(Con, "Cmd.Parameters.Append") Then Exit Function

	Cmd.Parameters.Append Cmd.CreateParameter("@vcNetworkName", adVarChar, adParamInput, 50)
	If Not ADO_CheckErr(Con, "Cmd.Parameters.Append") Then Exit Function

	Cmd.Parameters.Append Cmd.CreateParameter("@cNetworkTypeCode", adChar, adParamInput, 1)
	If Not ADO_CheckErr(Con, "Cmd.Parameters.Append") Then Exit Function

	Cmd.Parameters.Append Cmd.CreateParameter("@iOrderByName", adInteger, adParamInput, 4)
	If Not ADO_CheckErr(Con, "Cmd.Parameters.Append") Then Exit Function

	Cmd.Parameters.Append Cmd.CreateParameter("@vcReturnDetailList", adVarChar, adParamInput, 3)
	If Not ADO_CheckErr(Con, "Cmd.Parameters.Append") Then Exit Function

	CreateEnumRegionGroupParams = True
End Function

'* ********** ********** ********** ********** ********** ********** **
' *** Calls the TRNDdpEnumRegionGroup stored procedure
'* ********** ********** ********** ********** ********** ********** **
Sub CallEnumRegionGroup(Cmd)

	Dim blnEmptySet, lngItem, rgItems(), strItem, rstFields, iRetVal

	On Error Resume Next

	' execute TRNDdpEnumRegionGroup @vcRegionGroupName varchar(50) = NULL,
	'                               @iOrderByName      int         = NULL

	Cmd.CommandText		= "TRNDdpEnumNetwork"
	Cmd.CommandType		= adCmdStoredProc
	Cmd.CommandTimeout	= Application("cCommandTimeout")

	' Execute the stored procedure
	Set Rst = Cmd.Execute

	If IsEmpty(Rst) Then
		ADO_MsgBox Con.Errors
	Else
		' Is the record set empty?
		blnEmptySet = Rst.BOF And Rst.EOF

		If blnEmptySet Then
			Response.Write "No Regions" & vbCrLf
		Else
			Set rstFields = Rst.Fields

			lngItem = 0
			DIM NetworkID, NetworkName
			Set NetworkID	= Rst.Fields("NetworkID")
			Set NetworkName	= Rst.Fields("NetworkName")

			Do While Not Rst.EOF
				Redim Preserve rgItems(lngItem)

				rgItems(lngItem) = NetworkID & "," & Trim(NetworkName)

				Rst.MoveNext
				ADO_CheckErr Con, "Rst.MoveNext"

				lngItem = lngItem + 1
			Loop

			Application.Lock
			Application("rgRegionGroupNames") = rgItems
			Application.Unlock
		End If

		'Close recordset
		Rst.close
		Set Rst = Nothing

		'Get stored procedure return value
		iRetVal = Cmd("RETURN_VALUE")
		If iRetVal <> 0 Then
%>
			RETURN_VALUE=<%= iRetVal %>
<%
		End If
	End If
End Sub

'* ********** ********** ********** ********** ********** ********** **
' *** Calls the TRNDdpEnumRegionGroup stored procedure
'* ********** ********** ********** ********** ********** ********** **
Sub GetRegionGroups()
	DIM Cmd, Con
	If Not ADO_Connect(Con, Cmd, Application("iConfigDB")) Then
		'Database connection error
		Set Con = Nothing
		Set Cmd = Nothing
		
		Exit Sub			
	End If
	
	If Cmd is Nothing or IsEmpty(Cmd) Then
		Exit Sub
	End If
	
	If CreateEnumRegionGroupParams(Cmd, Con) Then
		CallEnumRegionGroup Cmd
	End If
	
	Set Cmd = nothing
End Sub


'* ********** ********** ********** ********** ********** **
' *** Returns a list of the most current files in
'  the \\ServerName\cache\cu30day directory
'  return - an array of valid files paths
'* ********** ********** ********** ********** ********** */
Function GetGoodFileName(fso, szFolder)
	GetGoodFileName = ""
	
	Dim today, yesterday
	'The files are named the previous day because that's the data that they contain
	today     = szFolder & "\" & FormatDate(DateAdd("d", -1, now))
	yesterday = szFolder & "\" & FormatDate(DateAdd("d", -2, now))
	Minus3    = szFolder & "\" & FormatDate(DateAdd("d", -3, now))
	Minus4    = szFolder & "\" & FormatDate(DateAdd("d", -4, now))
	Minus5    = szFolder & "\" & FormatDate(DateAdd("d", -5, now))
	Minus6    = szFolder & "\" & FormatDate(DateAdd("d", -6, now))
	Minus7    = szFolder & "\" & FormatDate(DateAdd("d", -7, now))

	' Let's work with the most recent version of the file within the last week	
	If fso.FileExists(today) Then
		GetGoodFileName = today
	ElseIf fso.FileExists(yesterday) Then
		GetGoodFileName = yesterday
	ElseIf fso.FileExists(Minus3) Then
		GetGoodFileName = Minus3
	ElseIf fso.FileExists(Minus4) Then
		GetGoodFileName = Minus4
	ElseIf fso.FileExists(Minus5) Then
		GetGoodFileName = Minus5
	ElseIf fso.FileExists(Minus6) Then
		GetGoodFileName = Minus6
	ElseIf fso.FileExists(Minus7) Then
		GetGoodFileName = Minus7
	Else
		GetGoodFileName = ""
	End If
	
End Function

'* ********** ********** ********** ********** ********** **
' *** Returns a list of the most current files in
'  the \\ServerName\cache\cu30day directory
'  return - an array of valid files paths
'* ********** ********** ********** ********** ********** */
function GetFileList(rgRegionGroupNames)
	Dim Files()
	Dim szBase, fso, curSize
	
	szBase = Server.MapPath("/cache/cu30day")

	Set fso = CreateObject("Scripting.FileSystemObject")
	curSize = 0
	If IsObject(fso) Then
		
		For each rg in rgRegionGroupNames
			rgValues = Split(rg, ",")
			
			'Make sure that we have good data
			If UBound(rgValues) >= 0 and rgValues(1) <> "" Then
				Dim rgGroupId, rgGroupName
				rgGroupId  = rgValues(0)
				rgGroupName = rgValues(1)
			
				'Check to see if the sub-folder is there
				If fso.FolderExists(szBase & "\" & rgGroupName) Then
				
					'Check to see if there's a recent file there
					Dim GoodFileName
					rgGroupName = GetGoodFileName(fso, szBase & "\" & rgGroupName)
						
					'If there was one there, append it to the list
					If rgGroupName <> "" Then
						Dim rgTuple
						ReDim rgTuple(1)
						rgTuple(0) = rgGroupId
						rgTuple(1) = rgGroupName
						
						ReDim Preserve Files(curSize)
						Files(curSize) = rgTuple
						curSize = curSize + 1
					End If
				End If
			End  If
		Next

	End If

	Set fso = nothing
	
	GetFileList = Files
End Function


'* ********** ********** ********** ********** ********** **
' *** Opens the file and outputs any line with str in it to 
'Response.Write
'* ********** ********** ********** ********** ********** */
Sub InsertFile(ts, Strs, appendText)
	Set OutFileObj = CreateObject("Scripting.FileSystemObject")
	Set Outfile=OutfileObj.CreateTextFile ("e:\globalnet\cache\SNMP.txt",true)

	If Not IsObject(ts) Then
		Exit Sub
	End If
	If Not IsArray(Strs) Then
		Exit Sub
	End If
	
	Dim line
	Do while Not ts.AtEndOfStream
		line = ts.ReadLine()
		
		'Check to see if the str is in the current line.
		for each str in Strs
			Dim result
			result = InStr(1, line, str, 1)
			
			If result and result <> 0 Then
				OutFile.Write line + appendText + vbCrLf
				line = ""
			End If
		Next
	Loop
End Sub


'* ********** ********** ********** ********** ********** ********** **
' *** Creates a filed that is a grep from the SNMP files 
'  The words that it greps on are the Request.QueryString("Find")
'  Able to grep multiple finds
'* ********** ********** ********** ********** ********** ********** **

Dim iCacheMinutes

Set OutFileObj = CreateObject("Scripting.FileSystemObject")
Set Outfile=OutfileObj.CreateTextFile ("e:\globalnet\cache\SNMP.txt",true)

iCacheMinutes = CInt(Application("iCacheMinutes"))
If CacheExpired("dtRegionGroupsCached", iCacheMinutes) _
	or Not IsArray(Application("rgRegionGroupNames")) Then
	GetRegionGroups
End If

Dim iIndex, rgRegionGroupNames, rgValues

' Get the application-level array
rgRegionGroupNames = Application("rgRegionGroupNames")
If IsArray(rgRegionGroupNames) Then
		
	DIM fso, f
	Set fso = CreateObject("Scripting.FileSystemObject")
		
	'Make sure that we've successfully created the object
	If IsObject(fso) Then
		Dim FileList
		FileList = GetFileList(rgRegionGroupNames)
			
		For each rgTuple in FileList
			' Response.Write rgTuple(1) & ", " & rgTuple(0) & vbCrLf
			Set f = fso.OpenTextFile(rgTuple(1))
			Do While not f.AtEndOfStream
				OutFile.Writeline f.ReadLine() 
			Loop
		Next
	End If
End If

'* ********** ********** ********** ********** ********** ********** **
'* ********** ********** ********** ********** ********** ********** **