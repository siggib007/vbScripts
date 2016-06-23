'..................................................................................................................................
'...This wsf script can be invoked from a command prompt using:                                                           ...
'...    cscript scriptfile.wsf "outputfile.xml" "DatasourceName" "SavedReportUserAlias" "SavedReportName" MaxNumRows "Arguments"...
'...Where:                                                                                                                ...
'...    scriptfile.wsf          - name of this script                                                                     ...
'...     outputfile.xml               - name of the xml file to create to hold the data                                                  ...
'...    DatasourceName          - name of the datasource to be queried                                                    ...
'...    SavedReportUserAlias    - user alias of the person owning the saved report                                        ...
'...    SavedReportName         - name of the saved report to be loaded                                                   ...
'...    MaxNumRows              - maximum number of rows to return (0 for all)                                            ...
'...    Arguments               - [optional] comma-separated list of search parameter values                              ...
'...    bElementBased           - [optional] True returns Element based XML. Valid values: True/False (Default:False)           ...
'...    bXMLDateFormat          - [optional] True returns dates in XMLDate format. Valid values: True/False (Default:False)     ...
'..................................................................................................................................
Dim strURL, oPoster, sOutputFileName, sDatasourceName, sSavedReportUserAlias, sSavedReportName, sMaxNumRows, sArguments, sElementBased, sXMLDateFormat, fso, fFile
'--- Make sure you are posting the data to the correct URL
strURL = "http://XMLInterface/XMLPullRS.asp"     
If WScript.Arguments.Count < 5 Then
   WScript.Echo "Missing argument. Please check documentation."
   usage
   WScript.Quit 
End if
If WScript.Arguments.Count > 8 Then
   WScript.Echo "Extra arguments. Please check documentation."
   usage
   WScript.Quit 
End If
'--- Read Script Arguments
sOutputFileName = WScript.Arguments(0)
sDatasourceName = WScript.Arguments(1)
sSavedReportUserAlias = WScript.Arguments(2)
sSavedReportName = WScript.Arguments(3)
sMaxNumRows = WScript.Arguments(4)
sArguments = ""
sElementBased  = "False"
sXMLDateFormat = "False"
If WScript.Arguments.Count > 5 Then
     If WScript.Arguments(5) = "" Then
          sArguments = ""
     Else
          sArguments = WScript.Arguments(5)
     End If
End If
If WScript.Arguments.Count > 6 Then
     If WScript.Arguments(6) = "" Then
          sElementBased = "False"
     Else
          sElementBased = WScript.Arguments(6)
     End If
End If
If WScript.Arguments.Count > 7 Then
     If WScript.Arguments(7) = "" Then
          sXMLDateFormat = "False"
     Else
          sXMLDateFormat = WScript.Arguments(7)
     End If
End If
'--- Send the data
Set oPoster = CreateObject("Microsoft.XMLHTTP")
oPoster.Open "POST", strURL, 0
oPoster.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
oPoster.Send "sAction=ScriptedPull2&p1=" & URLEncode(sDatasourceName) & "&p2=" & URLEncode(sSavedReportUserAlias) & "&p3=" & URLEncode(sSavedReportName) & "&p4=" & CInt(sMaxNumRows) & "&p5=" & URLEncode(sArguments) & "&p6=" & URLEncode(sElementBased) & "&p7=" & URLEncode(sXMLDateFormat)
'--- Check the return status
Select Case Left(oPoster.status, 3)
     Case "200"
          WScript.Echo "Processing was successful!"
     Case "602"
          WScript.Echo "Processing was unsuccessful. Invalid request. Error information to follow."
     Case "604"
          WScript.Echo "Processing was unsuccessful. Database error. Error information to follow."
     Case "610"
          WScript.Echo "Processing was unsuccessful. XmlInterface is currently not available. Please try later."
     Case Else
          WScript.Echo  "Processing returned an unexpected HTTP status code - " & oPoster.status
End Select
'--- See if we got valid XML back
If oPoster.responseXML.xml <> "" Then          'oPoster.responseXML is an XMLDocument object containing the query results
		 'wscript.echo oPoster.responseXML.xml   
     'On Error Resume Next
     '--- Write the XML to the output file
     '--- The XML could be parsed here instead
     Set fso = CreateObject("Scripting.FileSystemObject")
     Set fFile = fso.CreateTextFile(sOutputFileName, True, True)
     fFile.write(Replace(oPoster.responseXML.xml, "<?xml version=""1.0""?>", "<?xml version=""1.0"" encoding=""UTF-16""?>"))
	   'wscript.echo oPoster.responseXML.xml   
     fFile.write oPoster.responseXML.xml 
     If Err.Number > 0 Then
          WScript.Echo "Unable to output XML to file.  XML results to follow."
          WScript.Echo oPoster.responseXML.xml
     Else
          WScript.Echo "Results successfully written to file."     
     End If
     fFile.close()
     Set fFile = Nothing
     Set fso = Nothing
     '--- Check for errors
     '--- More sophisticated handling could also be done here
Else
     '--- Write error information, if any
     '--- More sophisticated handling could also be done here
     If Left(oPoster.status, 3) = "200" Then
          WScript.Echo "No XML results were returned.  Text results to follow."
     End If
     WScript.Echo oPoster.responseText     'Error information would be here
End If
'================================================================
'End of Script===================================================
'================================================================
'================================================================
'Function: URLEncode                                            =
'    This function simply encodes all characters that are not   =
'    valid in a URL.  It is a direct implementation of ASP's    =
'    built-in Server.URLEncode function                         =
'================================================================
Function URLEncode(tmpStr) 
     Dim temp, onechar
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

Sub usage
	wscript.echo "This wsf script can be invoked from a command prompt using:"
	wscript.echo "cscript scriptfile.vbs ""outputfile.xml"" ""DatasourceName"" ""SavedReportUserAlias"" ""SavedReportName"" ""MaxNumRows"" ""Arguments"" ""bElementBased"" ""bXMLDateFormat"""
	wscript.echo "Where:"
	wscript.echo "scriptfile.vbs       - name of this script"
	wscript.echo "outputfile.xml       - name of the xml file to create to hold the data"
	wscript.echo "DatasourceName       - name of the datasource to be queried"
	wscript.echo "SavedReportUserAlias - user alias of the person owning the saved report"
	wscript.echo "SavedReportName      - name of the saved report to be loaded"
	wscript.echo "MaxNumRows           - maximum number of rows to return (0 for all)"
	wscript.echo "Arguments            - [optional] comma-separated list of search parameters"
	wscript.echo "bElementBased        - [optional] True returns Element based XML. Valid values: True/False (Default:False)"
	wscript.echo "bXMLDateFormat       - [optional] True returns dates in XMLDate format. Valid values: True/False (Default:False)"
	wscript.echo "Check XML API on http://xmlinterface/ for more details. "
End Sub