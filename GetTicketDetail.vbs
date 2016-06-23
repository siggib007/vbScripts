Option Explicit 

Dim strURL, sDatasourceName, sSavedReportUserAlias, sSavedReportName, sMaxNumRows, sArguments, sDSN, sXML, fso, fFile
Dim DBName, OpsDBServer, UserName, Password, oPoster, objDBConn, objBL, objDocument, cn, cmd, StartDate, rs, WshNetwork
Dim XMLFile, XSDFile, XMLOk, BulkLoadErrLogPath, sTableName, RootDir, cmdText, NetDBrs, NetDBcn, enddate, NetDBcmd, NetDBServer

' Initialize global variables
Set NetDBcn = CreateObject("ADODB.Connection")
Set NetDBrs = CreateObject("ADODB.Recordset")
Set cn      = CreateObject("ADODB.Connection")
Set cmd     = CreateObject("ADODB.Command")
Set rs      = CreateObject("ADODB.Recordset")
Set NetDBcmd     = CreateObject("ADODB.Command")

OpsDBServer = "satnetengfs01"     ' -- Replace with db server name
NetDBServer = "TK2STGDWSQL05"
DBName      = "smarts"     ' -- Replace with db name
UserName    = "ScriptRW"     ' -- Replace with user name
Password    = "thisbites2."     ' -- Replace with pwd

main

Sub main()

	wscript.echo "Starting main process at " & Now 
	Set WshNetwork = WScript.CreateObject("WScript.Network")
	WScript.Echo "Domain = " & WshNetwork.UserDomain
	WScript.Echo "Computer Name = " & WshNetwork.ComputerName
	WScript.Echo "User Name = " & WshNetwork.UserName
	
	wscript.echo "Connecting to the Netpro DB"
	
	cn.Provider = "sqloledb"
	cn.Properties("Data Source").Value = OpsDBServer
	cn.Properties("Initial Catalog").Value = DBName
	cn.Properties("User ID").Value = UserName
	cn.Properties("Password").Value = Password
	'cn.Properties("Integrated Security").Value = "SSPI"
	cn.Open

	Cmd.ActiveConnection = cn
	cmd.CommandTimeout = 60
		
	cmdtext = "select isnull(max(CREATEDAT),'') from smarts.dbo.EventTicket"
	rs.open cmdtext, cn
	StartDate = rs.fields(0).value
	If startdate > "2006-01-01" Then
		StartDate = "8/1/06"
	End If 
	rs.close
	wscript.echo "Start date set to " & startdate
	
	wscript.echo "Connecting to NetDB for a ticketlist"
	'Set ADO connection properties.
	NetDBcn.Provider = "sqloledb"
	NetDBcn.Properties("Data Source").Value = NetDBServer
	NetDBcn.Properties("Initial Catalog").Value = "smarts"
	NetDBcn.Properties("Integrated Security").Value = "SSPI"
	NetDBcn.Open
	
	wscript.echo "Connected. Now fetching list. " & Now 
	cmdtext = "select TROUBLETICKETID, min(CREATEDAT) FROM [Smarts].[dbo].[IC_T_NOTIFICATION_OCCURRENCES] WHERE TROUBLETICKETID <>'' and CREATEDAT > '" & StartDate & "' group by TROUBLETICKETID order by min(CREATEDAT) "
	'cmdtext = "select distinct TROUBLETICKETID FROM [Smarts].[dbo].[IC_T_NOTIFICATION_OCCURRENCES] WHERE TROUBLETICKETID <>'' and CREATEDAT > '" & StartDate & "' group by TROUBLETICKETID"
	netdbcmd.commandtext = cmdtext
	netdbCmd.ActiveConnection = NetDBcn
	netdbcmd.CommandTimeout = 60
	'wscript.echo cmdtext
	'Exit Sub
	Set NetDBrs = netdbcmd.Execute
	'NetDBrs.Open cmdText, NetDBcn
	
	wscript.echo "Got list, calling xmlinterface on each ticket. " & Now 

	Do until NetDBrs.eof
		'wscript.echo NetDBrs.fields(0).value
		wscript.echo "processing " & NetDBrs.fields(1).value & " " & NetDBrs.fields(0).value & " at " & now
		mainxmlget NetDBrs.fields(0).value
		Cmd.CommandText = "insert into smarts.dbo.EventTicket (TROUBLETICKETID, CREATEDAT) values (" & NetDBrs.fields(0).value & ",'" & NetDBrs.fields(1).value & "')"
		Cmd.Execute
		wscript.echo "Inserted IC data into DB. " & Now
		
		Cmd.CommandText = "insert smarts.dbo.tblUTSTickets select * from dbo.xmlUTSTickets"
		Cmd.Execute
		'wscript.echo "Moved UTS data into perm table"
		
		NetDBrs.movenext
	loop  

	'StartDate = "5/1/06"
	'enddate = DateAdd("d",-1,DateAdd("m",1,DatePart("m",startdate) & "/1/" & DatePart("yyyy",startdate)))
	enddate = Now

	While enddate < Now
		wscript.echo "Daterange from " & startdate & " to " & enddate
		cmd.commandtext = "smarts.dbo.spCalcTicketCreateStat '" & startdate & "','" & enddate & "'"
		cmd.execute
		startdate = DateAdd("d",1,enddate)
		enddate = DateAdd("d",-1,DateAdd("m",1,startdate))
	Wend
	
	NetDBrs.close
	NetDBcn.close
	Set NetDBrs = nothing
	Set NetDBcn = nothing	
	Set rs = nothing
	Set cmd = nothing
	Set cn = nothing
	
	wscript.echo "All done at " & Now 
end Sub 'End of main sub


'=================================================================
'Sub: MainXMLGet(Argument)                                      ==
'    This is the main procedure that calls all the other ones.  ==
'=================================================================
Sub MainXMLGet(Argument)

Set objDocument = CreateObject("msxml2.DOMDocument")
Set objBL = CreateObject("SQLXMLBulkLoad.SQLXMLBulkLoad.3.0")
Set fso = CreateObject("Scripting.FileSystemObject")
Set objDBConn = CreateObject("ADODB.Connection")

     '--- Initialize Script variables
     RootDir		   		= "C:\~"     
     
     sDatasourceName       = "UTS - Ticket"
     sSavedReportUserAlias = "siggib"
     sSavedReportName      = "TicketInfo"
     sMaxNumRows           = 0
     sArguments            = Argument '"1572647"	
     sTableName            = "xmlUTSTickets"
     
     '--- Make sure you are posting the data to the correct URL
     strURL                = "http://XMLInterface/XMLPullRS.asp"
 
     sDSN                  = "provider=SQLOLEDB.1;data source=" & OpsDBServer & ";database=" & DBName & ";uid=" & UserName & ";pwd=" & Password

     BulkLoadErrLogPath    = RootDir & "XMLPullError.xml"
     XMLFile               = RootDir & sTableName & ".xml"
     XSDFile               = RootDir & sTableName & "Schema.xsd"
     sXML                  = ""


'================================================================
'Script Main Processing    --  Begins here                  =====
'================================================================

'wscript.echo "processing " & sArguments & " at " & now
'wscript.echo "sDatasourceName=" & sDatasourceName
'wscript.echo "sSavedReportUserAlias=" & sSavedReportUserAlias
'wscript.echo "sSavedReportName=" & sSavedReportName
'wscript.echo "sArguments=" & sArguments
'wscript.echo "sMaxNumRows=" & sMaxNumRows
'wscript.echo "sTableName=" & sTableName
'wscript.echo "XMLFile=" & XMLFile

'--- Request the XSD Schema
Set fFile = fso.CreateTextFile(XMLFile, True, True)
Set oPoster = CreateObject("Microsoft.XMLHTTP")
GetPosterData "GetElementBasedXSDSchema2", URLEncode(sDatasourceName), URLEncode(sSavedReportUserAlias), URLEncode(sSavedReportName), "[dbo]." & URLEncode(sTableName), "", "", "", ""
CheckPosterStatus
sXML = oPoster.responseXML.xml

'--- oPoster.responseXML is an XMLDocument object containing the query results
If oPoster.responseXML.xml <> "" Then
     '--- Write the schema to the output file
     sXML = Replace(sXML, "unsignedbyte", "unsignedByte") 
     Set fFile = fso.CreateTextFile(XSDFile, True, True)
     fFile.Write (sXML)
     fFile.Close
     Set fFile = Nothing
Else
     ExitWithErrors("Unable to output Schema to file.  XML results to follow." & Chr(13) & Chr(10) & oPoster.responseXML.xml & Chr(13) & Chr(10) & oPoster.responseText)
End If

'--- Get the data
GetPosterData "ScriptedPull2", URLEncode(sDatasourceName), URLEncode(sSavedReportUserAlias), URLEncode(sSavedReportName), "[dbo]." & URLEncode(sTableName), CInt(sMaxNumRows), URLEncode(sArguments), True, False
CheckPosterStatus

'--- Write the XML to the output file
Set fFile = fso.CreateTextFile(XMLFile, True, True)
sXML = oPoster.responseXML.xml
sXML = Replace(sXML, "<?xml<version>1.0</version>?>", "<?xml version=""1.0""?>")
sXML = Replace(sXML, "<DATASET>", "<DATASET xmlns:od=""urn:schemas-microsoft-com:officedata"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:noNamespaceSchemaLocation=""~XMLPull.xsd"">")
fFile.Write(Replace(sXML, "<?xml version=""1.0""?>", "<?xml version=""1.0"" encoding=""UTF-16""?>"))
fFile.Close()

'--- Check for errors
If Err.Number > 0 Then
     ExitWithErrors("Unable to output XML to file.  XML results to follow." & Chr(13) & Chr(10) & oPoster.responseXML.xml & Chr(13) & Chr(10) & oPoster.responseText)
Else
     ValidateXMLDocument XMLFile
     BulkLoadXML XMLFile, XSDFile

     If Err.Number <> 0 Then
          ExitWithErrors("There were errors while processing the job. Please try running the job later or contact OSSG for more info.")
     Else
          WScript.Echo "Results successfully written to database."
     End If
End If

CleanUpObjects

End Sub
'=================================================================
'End of Script      ==============================================
'=================================================================



'=================================================================
'Sub: ValidateXMLDocument(XMLFile)                          ==
'    This subroutine loads and parse the XML document             ==
'=================================================================
Sub ValidateXMLDocument(XMLFile)
     On Error Resume Next

     objDocument.async = False
     objDocument.validateOnParse = True
     XMLOk = objDocument.load(XMLFile)

     If Not XMLOk Then
          XMLErrorMsgDesc = "XML Document is invalid: " & objDocument.parseError.reason
          iXMLTagBegPos = InStrRev(objDocument.parseError.srctext, "<",(objDocument.parseError.linepos - 1))
          If iXMLTagBegPos = 0 Then
               If objDocument.parseError.line > 1 Then
                    XMLErrorMsgDesc = XMLErrorMsgDesc & "XML syntax on line " & (objDocument.parseError.line - 1)
               Else
                    XMLErrorMsgDesc = XMLErrorMsgDesc & "Cannot determine source text in error."
               End If
          Else
               XMLErrorMsgDesc = XMLErrorMsgDesc & " XML:" & Mid(objDocument.parseError.srctext, iXMLTagBegPos, (objDocument.parseError.linepos - iXMLTagBegPos))
          End If

          ExitWithErrors(XMLErrorMsgDesc)
     End If
End Sub

'=================================================================
'Sub: BulkLoadXML (XMLFile, XSDFile)                            ==
'    This subroutine sets values for the SQLXML bulkload object ==
'    and uses the content in XSD & XML documents to populate    ==
'    a SQL database table.                                      ==
'=================================================================
Sub BulkLoadXML(XMLFile, XSDFile)
     On Error Resume Next
     objDBConn.Open sDSN
     If Err.Number <> 0 Then
          ExitWithErrors("Database connection error: Cannot connect to destination database [" & DBName & "] on server [" & OpsDBServer & "]." & Chr(13) & Chr(10) & Err.Description)
     End If

     objBL.SGDropTables = True
     objBL.SchemaGen = True
     objBL.ForceTableLock = True
     objBL.ConnectionString = sDSN
     objBL.ErrorLogFile = BulkLoadErrLogPath
     objBL.Execute XSDFile, XMLFile
     If Err.Number <> 0 Then
          ExitWithErrors("Error while updating the database." & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Check the file named ~XMLPullError.xml on the root of Drive C for more information.")
     End If
End Sub

'=================================================================
'Sub: GetPosterData(p0, p1, p2, p3, p4, p5, p6, p7, p8)          ==
'    This procedure posts request through the poster object     ==
'    to a particular subroutine in the server script             ==
'=================================================================
Sub GetPosterData (strAction, strDSN, strSavedRptUser, strSavedRptName, strTableName, strRowCount, strArguments, bElementBasedSchema, bXMLDateFormat)
     oPoster.Open "POST", strURL, 0
     oPoster.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"

     Select Case UCase(strAction)
          Case "SCRIPTEDPULL2"
               oPoster.Send "sAction=" & strAction & "&p1=" & strDSN & "&p2=" & strSavedRptUser & "&p3=" & strSavedRptName & "&p4=" & strRowCount & "&p5=" & strArguments & "&p6=" & bElementBasedSchema & "&p7=" & bXMLDateFormat 

          Case "GETELEMENTBASEDXSDSCHEMA2"
               oPoster.Send "sAction=" & strAction & "&p1=" & strDSN & "&p2=" & strSavedRptUser & "&p3=" & strSavedRptName & "&p4=" & strTableName
     End Select
End Sub


'=================================================================
'Sub: CheckPosterStatus                                   ==
'    This subroutine checks the return status of the           ==
'    poster object                               ==
'=================================================================
Sub CheckPosterStatus
     '--- Check the return status
     Select Case Left(oPoster.Status, 3)
          Case "200"
               If oPoster.responseXML.xml = "" Then
                    ExitWithErrors("No XML results were returned." & Chr(13) & Chr(10) & oPoster.responseText )
               End If

          Case "602"
               ExitWithErrors("Processing was unsuccessful. Invalid request. " & Chr(13) & Chr(10) & oPoster.responseText)

          Case "604"
               ExitWithErrors("Processing was unsuccessful. Database error. " & Chr(13) & Chr(10) & oPoster.responseText)

          Case Else
               ExitWithErrors("Processing returned an unexpected HTTP status code - " & oPoster.status)
     End Select
End Sub

'=================================================================
'Sub: ExitWithErrors(XMLErrorMsgDesc)                           ==
'    This subroutine writes an error message to the console     ==
'    and terminates the script execution                  ==
'=================================================================
Sub ExitWithErrors(XMLErrorMsgDesc)
     On Error Resume Next
     WScript.Echo XMLErrorMsgDesc
     CleanUpObjects
     WScript.Quit(1)
End Sub

'=================================================================
'Sub: CleanUpObjects                                        ==
'    This subroutine cleans up all objects and temp files       ==
'    created by this script                           ==
'=================================================================
Sub CleanUpObjects
     On Error Resume Next
     fso.DeleteFile (XMLFile)
     fso.DeleteFile (XSDFile)
     Set fFile = Nothing
     Set fso = Nothing
     Set oPoster = Nothing
     Set objDBConn = Nothing
     Set objBL = Nothing
     Set objDocument = Nothing
     wscript.echo
     'wscript.echo "completed at " & now
End Sub

'================================================================
'Function: URLEncode                                            =
'    This function simply encodes all characters that are not   =
'    valid in a URL.  It is a direct implementation of ASP's    =
'    built-in Server.URLEncode function                         =
'================================================================
Function URLEncode(tmpStr) 
     Dim temp, onechar, j
     Const URLComponent_SET_OF_VALID_UNESCAPED_CHARACTERS = "abcdefghijklmnopqrstuvwxyz1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ;/:@=$-_.!*'(), "
     For j = 1 To Len(tmpStr) 
          onechar = Mid(tmpStr, j, 1) 
          If InStr(URLComponent_SET_OF_VALID_UNESCAPED_CHARACTERS, onechar) = 0 Then 
               ' Encode this character 
               temp = temp + "%" + Hex(AscB(onechar)) 
          Else 
               ' Good character 
               temp = temp + onechar 
          End If 
     Next 
     URLEncode = Replace(temp, " ", "+")
End Function 