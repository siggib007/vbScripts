Option Explicit
Dim cn
Dim rs, ToAddress, ResultInMail, ResultToFile, strOutFilePath
Dim fld, fldValue
Dim cmdText, StrOut, strTD
Dim fso, strOutFile

Const MailServerName = "smarthost.dns.microsoft.com" 
Const SMTPTimeout = 10
Const FromAddress = "ghtools@microsoft.com"
Const CCAddress = ""
Const cdoSendUsingPort = 2
Const TristateTrue = -1
Const TristateFalse = 0
Const TristateUseDefault = -2
Const Indent = "    "
Const strQuery = "exec network.dbo.GNOCUnmanaged"

Const TableTitle = "Devices not managed by GNOC"
Const ReportTitle = "Un-Managed device report"

ToAddress = ""
strOutFilePath =""
ResultInMail = false
ResultToFile = false

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")
Set fso = CreateObject("Scripting.FileSystemObject")

Sub Main()
Dim arg1, arg2

	arg1=""
	arg2=""

	If wscript.arguments.count > 0 Then
		arg1 = wscript.arguments(0)
	Else
		wscript.echo "Neither email address or file path provided." 
		wscript.echo "One or both needs to be provided so results are saved."
		wscript.echo "Exiting..."
		wscript.quit
	End If
	If wscript.arguments.count > 1 Then
		arg2 = wscript.arguments(1)
	End If

	If InStr(1,arg1,"@")=0 Then 
		wscript.echo "Found File Name in arg1"
		strOutFilePath = arg1
	Else
		wscript.echo "Found Email Address arg1"
		ToAddress = arg1
	End If
	
	If arg2 <> "" Then
		If InStr(1,arg2,"@")=0 Then 
			strOutFilePath = arg2
			wscript.echo "Found File Name in arg2"
		Else
			wscript.echo "Found Email Address arg2"
			ToAddress = arg2
		End If
	End If
	
	wscript.echo "File Path = " & strOutFilePath
	wscript.echo "ToAddress = " & ToAddress
	If ToAddress <> "" Then ResultInMail=True
	If strOutFilePath <> "" Then ResultToFile=True
	wscript.echo "ResultInMail = " & ResultInMail
	wscript.echo "ResultToFile = " & ResultToFile
	
	'Set ADO connection properties.
    cn.Provider = "sqloledb"
    cn.Properties("Data Source").Value = "ineteng"
    cn.Properties("Initial Catalog").Value = "network"
	'cn.Properties("User ID").Value = ""
	'cn.Properties("Password").Value = ""
    cn.Properties("Integrated Security").Value = "SSPI"
    cn.Open

	wscript.echo "Creating Report headers"

    'Create HTML headers
	strOut = "<html>" & vbcrlf
	strout = strout & "<head>" & vbcrlf
	strout = strout & "<title>ReportTitle</title>" & vbcrlf
	strout = strout & "</head>" & vbcrlf & vbcrlf
	strout = strout & "<body>" & vbcrlf & vbcrlf
	strout = strout & "<center> " & vbcrlf
	strout = strout & "<h1>" & ReportTitle & "</h1>" & vbcrlf
	strout = strout & "Report Generated on " & Now()  & vbcrlf
	strout = strout & "<h3><br></h3>" & vbcrlf

	wscript.echo "Creating Table"
	'Calling database fetching recordset
    rs.Open strQuery, cn
        
    'Create HTML Table from the results
	ConvertRS2HTMLTable TableTitle
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
	  	Select Case rs.fields(x).type 'What type of column is this?
	  		Case 3 'Int
	  			fldValue = FormatNumber (rs.fields(x).value,0,TristateTrue,TristateFalse,TristateTrue)
	  			If IsNull(fldValue) Then fldvalue = 0
	  			strTD= "<td align=right>"
	  		Case 4,5 ' Numeric
		  		fldValue = FormatPercent (rs.fields(x).value,2,TristateTrue,TristateFalse,TristateTrue)
		  		If IsNull(fldValue) Then fldvalue = 0
		  		strTD= "<td align=right>"
		  	Case 200 'String
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
	Dim iMsg,iConf 
	
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
Set rsidc=nothing
Set cn=nothing

If ResultToFile = True Then
	wscript.echo "Writing to file..."
	Set strOutFile = fso.CreateTextFile(strOutFilePath, True)
	stroutfile.writeline strout
End If

Set strOutFile = nothing
Set fso=nothing