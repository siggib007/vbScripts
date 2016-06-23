Option Explicit 
Dim cn, fso, strOutFile
Dim rs, ToAddress, ResultInMail, ResultToFile, strOutFilePath
Dim fld, fldValue, StrOut, strTD,OverWrite

Const MailServerName = "smarthost.dns.microsoft.com" 
Const SMTPTimeout = 10
Const FromAddress = "ghtools@microsoft.com"
Const CCAddress = ""
Const cdoSendUsingPort = 2
Const TristateTrue = -1
Const TristateFalse = 0
Const TristateUseDefault = -2
Const Indent = "    "
Const DBServer = "ineteng"
Const DefaultDB = "network"
Const UserID = ""
Const PWD = ""
Const cmdText1 = "select * from Network.dbo.vwCompareInetEngToGNOC"
Const ReportTitle = "Devices Comparison between InetEng and GNOC"
Const Table1Title = "Devices Unknown to GNOC"
Const cmdText2 = "select * from Network.dbo.vwCompareGNOCToInetEng"
Const Table2Title = "Devices Unknown to Ineteng"

ToAddress = ""
strOutFilePath =""
ResultInMail = false
ResultToFile = false
OverWrite = false

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")
Set fso = CreateObject("Scripting.FileSystemObject")

Sub Main()
Dim arg1, arg2, arg3, ians

	arg1=""
	arg2=""
	arg2=""

	If wscript.arguments.count > 0 Then
		arg1 = wscript.arguments(0)
	End If
	
	If wscript.arguments.count > 1 Then
		arg2 = wscript.arguments(1)
	End If

	If wscript.arguments.count > 2 Then
		arg3 = wscript.arguments(2)
	End If	

	If arg1 <> "" Then
		If InStr(1,arg1,"@")=0 Then 
			If Left(arg1,2) = "/y" Then 
				OverWrite = true
			Else
					wscript.echo "Found File Name in arg1"
					strOutFilePath = arg1
			End If
		Else
			wscript.echo "Found Email Address arg1"
			ToAddress = arg1
		End If
	End If
	
	If arg2 <> "" Then
		If InStr(1,arg2,"@")=0 Then 
			If Left(arg2,2) = "/y" Then 
				OverWrite = true
			Else
				wscript.echo "Found File Name in arg2"
				strOutFilePath = arg2
			End If
		Else
			wscript.echo "Found Email Address arg2"
			ToAddress = arg2
		End If
	End If

	If arg3 <> "" Then
		If InStr(1,arg3,"@")=0 Then 
			If Left(arg3,2) = "/y" Then 
				OverWrite = true
			Else
				wscript.echo "Found File Name in arg3"
				strOutFilePath = arg3
			End If
		Else
			wscript.echo "Found Email Address arg3"
			ToAddress = arg3
		End If
	End If
	
	If strOutFilePath = "" and ToAddress="" Then
		wscript.echo "Neither email address or file path provided." 
		wscript.echo "One or both needs to be provided so results are saved."
		wscript.echo "Proper syntax: "
		wscript.echo " "
		wscript.echo "cscript DeviceCompare.vbs [email | filePath] [/y]"
		wscript.echo " "
		wscript.echo "Note that if providing both email and filepath they can appear in any order"
		wscript.echo "Remeber to put quotes around paths with spaces"
		wscript.echo "To force a file to be overwritten, include the /y"
		wscript.echo "Exiting..."
		wscript.quit
	End If

	wscript.echo "File Path = " & strOutFilePath
	wscript.echo "ToAddress = " & ToAddress
	If ToAddress <> "" Then ResultInMail=True
	If strOutFilePath <> "" Then ResultToFile=True
	wscript.echo "ResultInMail = " & ResultInMail
	wscript.echo "ResultToFile = " & ResultToFile
	wscript.echo "Overwrite = " & OverWrite
	If fso.fileexists(strOutFilePath) and not overwrite Then
		ians = MsgBox (strOutFilePath & " already exists. Would you like to overwrite?",vbyesno)
		If ians = vbno Then 
			wscript.echo "File exits that can't be overwriten. Exiting."
			Exit Sub
		End If
	End If
	
	'Set ADO connection properties.
    cn.Provider = "sqloledb"
    cn.Properties("Data Source").Value = DBServer
    cn.Properties("Initial Catalog").Value = DefaultDB
	If UserID = "" Then 
		cn.Properties("Integrated Security").Value = "SSPI"
	Else
		cn.Properties("User ID").Value = UserID
		cn.Properties("Password").Value = PWD
	End If
    
    cn.Open

	wscript.echo "Creating Report headers"

    'Create HTML headers
	strOut = "<html>" & vbcrlf
	strout = strout & "<head>" & vbcrlf
	strout = strout & "<title>" & ReportTitle & "</title>" & vbcrlf
	strout = strout & "</head>" & vbcrlf & vbcrlf
	strout = strout & "<body>" & vbcrlf & vbcrlf
	strout = strout & "<center> " & vbcrlf
	strout = strout & "<h1>" & ReportTitle & "</h1>" & vbcrlf
	strout = strout & "Report Generated on " & Now()  & vbcrlf

	wscript.echo "Creating Table1"
	'Call the DB for recordset
    rs.Open cmdText1, cn
        
    'Create HTML Table from the results
	ConvertRS2HTMLTable Table1Title
	rs.close
	strout = strout & "<h1><br></h1>" & vbcrlf

	wscript.echo "Creating Table2"
	'Call the DB for recordset
    rs.Open cmdText2, cn
        
    'Create HTML Table from the results
	ConvertRS2HTMLTable Table2Title
	rs.close
	strout = strout & "<h1><br></h1>" & vbcrlf

	'Create HTML Footer stuff
	strout = strout & "</Center>" & vbcrlf & vbcrlf
	strout = strout & "</body>" & vbcrlf
	strout = strout & "</html>" & vbcrlf
	
	'Cleanup prior to exit
	cn.close
	
End Sub

Sub ConvertRS2HTMLTable(TableHeading)
Dim X
	strout = strout & "<table border=1 cellpadding=5 >" & vbcrlf
	strout = strout & "<caption align=center><h3>" & TableHeading & "</h3></caption>" & vbcrlf
  	strout = strout & indent & "<tr>" & vbcrlf
	For Each fld In rs.Fields
		strout = strout & indent & indent & "<td align=center><b>" & fld.name & "</b></td>" & vbcrlf
	Next
  	strout = strout & indent & "</tr>" & vbcrlf
  	strout = strout & indent & "<tr>" & vbcrlf
	While not rs.eof
	  For x=0 to rs.fields.count - 1
	  	Select Case rs.fields(x).type
	  		Case 3 'Int
	  			'wscript.echo "Field: " & rs.fields(x).name & " is an int. Value=" & rs.fields(x).value
	  			If not IsNull(rs.fields(x).value) Then 
	  				fldValue = FormatNumber (rs.fields(x).value,0,TristateTrue,TristateFalse,TristateTrue)
	  				'wscript.echo "Field: " & rs.fields(x).name & " is an int. Value=" & rs.fields(x).value
	  			Else
	  				fldvalue = "<BR>"
	  			End If 
	  			'wscript.echo "fldvalue = " & fldvalue
	  			strTD= "<td align=right>"
	  		Case 4,5 ' Numeric
		  		'wscript.echo "Field: " & rs.fields(x).name & " is a float. Value=" & rs.fields(x).value
	  			If not IsNull(rs.fields(x).value) Then 
			  		fldValue = FormatPercent (rs.fields(x).value,2,TristateTrue,TristateFalse,TristateTrue)
	  				'wscript.echo "Field: " & rs.fields(x).name & " is an int. Value=" & rs.fields(x).value
	  			Else
	  				fldvalue = rs.fields(x).value
	  			End If 
		  		strTD= "<td align=right>"
		  	Case 200 'String
		  		'wscript.echo "Field: " & rs.fields(x).name & " is a string. Value=" & rs.fields(x).value
		  		fldValue = Trim(rs.fields(x).value)
		  		If fldValue="" or IsNull(fldValue) Then 
		  			'wscript.echo "empty string replaced with space"
		  			fldvalue = "<br>"
		  		End If 
		  		If IsNumeric(fldvalue) Then
		  			strTD= "<td align=right>"
		  		Else
		  			strTD= "<td>"
		  		End If 
		  	Case Else
		  		'wscript.echo "Field: " & rs.fields(x).name & " is a type#" & rs.fields(x).type & ". Value=" & rs.fields(x).value
		  		fldValue = rs.fields(x).value
		  		strTD= "<td>"
		End Select 
		If fldValue = "zTotal" Then fldValue = "Total"
		strout = strout & indent & indent & strTD & fldValue & "</td>" & vbcrlf
	  Next
  	  strout = strout & indent & "</tr>"  & vbcrlf
  	  strout = strout & indent & "<tr>"  & vbcrlf
	  rs.movenext
	Wend
	strout = strout & indent & "</tr>" & vbcrlf
	strout = strout & "</table>" & vbcrlf
end Sub

Sub MySendMail(StrSubject,msg)
	Dim iMsg,iConf,flds 
	
	Set iMsg = CreateObject("CDO.Message") 
	Set iConf = CreateObject("CDO.Configuration") 
	Set Flds = iConf.Fields 
	
	With Flds 
	  .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort 
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MailServerName 
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = SMTPTimeout
	  .Update 
	End With 
	
	With iMsg 
	  Set .Configuration = iConf 
	      .To       = ToAddress
	      .CC	= CCAddress
	      .From     = FromAddress 
	      .Subject  = StrSubject 
	      .HTMLBody = msg
	      .Send 
	End With
End Sub

main
If ResultInMail = True Then
	wscript.echo "Sending Mail..."
	MySendMail ReportTitle, strout
	wscript.echo "Mail Sent"
End If 
Set rs=nothing
Set cn=nothing

If ResultToFile = True Then
	wscript.echo "Writing to file..."
	Set strOutFile = fso.CreateTextFile(strOutFilePath, True)
	stroutfile.writeline strout
End If

Set strOutFile = nothing
Set fso=nothing