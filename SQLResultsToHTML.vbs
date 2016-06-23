Dim cn, dtCalcDate, dtStartPeriod, dtEndPeriod, strCycleType
Dim rs, rsIDC, ToAddress, ResultInMail, ResultToFile, strOutFilePath
Dim fld, fldValue
Dim cmdText, StrOut, strTD
Dim fso, strOutFile

Const MailServerName = "smarthost.dns.microsoft.com" 
Const SMTPTimeout = 10
Const FromAddress = "ghtools@microsoft.com"
'Const ToAddress = "idcperf@microsoft.com"
Const CCAddress = ""
'Const ResultInMail = false
'Const ResultToFile = true
Const cdoSendUsingPort = 2
Const TristateTrue = -1
Const TristateFalse = 0
Const TristateUseDefault = -2
Const Indent = "    "

ToAddress = ""
strOutFilePath =""
ResultInMail = false
ResultToFile = false

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")
Set rsIDC = CreateObject("ADODB.Recordset")
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
    cn.Properties("Data Source").Value = "b11gnmona04"
    cn.Properties("Initial Catalog").Value = "resultsdb"
	'cn.Properties("User ID").Value = "ihsuser"
	'cn.Properties("Password").Value = "ihsuser"
    cn.Properties("Integrated Security").Value = "SSPI"
    cn.Open

	wscript.echo "Creating Report headers"
	'Call the DB for last log entry
	cmdText = "select * from vwPerfLastLog"
    rs.Open cmdText, cn
    If rs.eof Then
      	Exit Sub
    Else
       	dtCalcDate=rs.fields(0).value
       	dtStartPeriod=rs.fields(1).value
       	dtEndPeriod=rs.fields(2).value
       	Select Case rs.fields(3).value
       		Case 0
       			strCycleType = "Calendar Month"
       		Case 1
       			strCycleType = "Billing Month"
       		Case Else
       			strCycleType = "unknown Month Type"
       	End Select
    End If
    rs.close

    'Create HTML headers
	strOut = "<html>" & vbcrlf
	strout = strout & "<head>" & vbcrlf
	strout = strout & "<title></title>" & vbcrlf
	strout = strout & "</head>" & vbcrlf & vbcrlf
	strout = strout & "<body>" & vbcrlf & vbcrlf
	strout = strout & "<center> " & vbcrlf
	strout = strout & "<h1>IDC Egress Performance Report</h1>" & vbcrlf
	strout = strout & "<h2>For " & strCycleType & " of " 
	strout = strout & MonthName(DatePart("m",dtEndPeriod)) & " " & DatePart("yyyy",dtEndPeriod) & "</h2>" &vbcrlf
	strout = strout & "Report Generated on " & Now() & "<br>" & vbcrlf
	strout = strout & "Calculations done on " & dtCalcDate & "<br>" & vbcrlf
	strout = strout & "Data Interval from " & dtstartperiod & " to " & dtendperiod & vbcrlf
	strout = strout & "<h1><br></h1>" & vbcrlf
	
	wscript.echo "Creating Date Center Overview Table"
	'Call the DB for DC Overview data
	cmdText = "select * from vwPerfDCOverview order by idc"
    rs.Open cmdText, cn
        
    'Create HTML Table from the results
	ConvertRS2HTMLTable "Data Center Overview"
	rs.close
	strout = strout & "<h1><br></h1>" & vbcrlf
	
	cmdText = "select * from tblPerfIDCList"
    rsIDC.Open cmdText, cn
	
	While not rsIDC.eof
		wscript.echo "Creating Egress tables for " & rsidc.fields(0).value & " IDC"
		'Call the DB
		cmdText = "select [Link Type], [Carrier], [Circuit], [95th Mbps], [Max Mbps], [95th % Used], [Max % Used], [95th % Useable], [Max % Useable] from vwPerfTopEgress where idc='" & rsidc.fields(0).value & "' order by[95th Mbps] desc"
        	rs.Open cmdText, cn
        
        	'Create HTML Table from the results
		ConvertRS2HTMLTable "Top Used Egress Links for " & rsidc.fields(0).value & " IDC"
		strout = strout & "<h1><br></h1>" & vbcrlf
		rs.close
	
		'select [Link Type], [Carrier], [Circuit], [95th Mbps], [Max Mbps], [95th % Used], [Max % Used], [95th % Useable], [Max % Useable] from vwPerfTopPercent order by idc, [95th % Used] desc
		'Call the DB
		cmdText = "select [Link Type], [Carrier], [Circuit], [95th Mbps], [Max Mbps], [95th % Used], [Max % Used], [95th % Useable], [Max % Useable] from vwPerfTopPercent where idc='" & rsidc.fields(0).value & "' order by [95th % Used] desc"
        	rs.Open cmdText, cn
        	
        	'Create HTML Table from the results
		ConvertRS2HTMLTable "Top Congested Egress Links for " & rsidc.fields(0).value & " IDC"
		strout = strout & "<h1><br></h1>" & vbcrlf
		rs.close
		rsidc.movenext
	Wend
	rsIDC.close
	
	cmdText = "select * from tblPerfIDCList"
        rsIDC.Open cmdText, cn
	
	While not rsIDC.eof
		wscript.echo "Creating Cluster tables for " & rsidc.fields(0).value & " IDC"

		'Call the DB
		cmdText = "select Cluster, [95th Mbps], [Max Mbps] from vwPerfTopClusters where idc='" & rsidc.fields(0).value & "' order by [95th Mbps] desc"
	        rs.Open cmdText, cn
	        
	        'Create HTML Table from the results
		ConvertRS2HTMLTable "Top Clusters in " & rsidc.fields(0).value & " IDC"
		strout = strout & "<h1><br></h1>" & vbcrlf
		rs.close
		rsidc.movenext
	Wend
	rsIDC.close

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
	  			fldValue = FormatNumber (rs.fields(x).value,0,TristateTrue,TristateFalse,TristateTrue)
	  			strTD= "<td align=right>"
	  		Case 4,5 ' Numeric
		  		'wscript.echo "Field: " & rs.fields(x).name & " is a float. Value=" & rs.fields(x).value
		  		fldValue = FormatPercent (rs.fields(x).value,2,TristateTrue,TristateFalse,TristateTrue)
		  		strTD= "<td align=right>"
		  	Case 200 'String
		  		'wscript.echo "Field: " & rs.fields(x).name & " is a string. Value=" & rs.fields(x).value
		  		fldValue = Trim(rs.fields(x).value)
		  		strTD= "<td>"
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
	MySendMail "IDC Network Performance Report", strout
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