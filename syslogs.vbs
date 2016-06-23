Option Explicit
Dim cn,cmd, RS, fso, spath, sfilepath, maxdate, strDate, dtline, device, deviceIP, Severity, message, tablename
Dim fileobj, strline, strparts, x, DBServerName, DBName, cmdtext, dtstart, dtstop, linecount, TotalLines, strclass

Const ADOTimeOut = 60

DBServerName = "by2netsql01"     ' -- Replace with db server name
DBName       = "syslogs"     ' -- Replace with db name
TableName    = "syslog"

If wscript.arguments.count > 0 Then
	spath = wscript.arguments(0)
Else
	wscript.echo "Please provide syslog file location"
	wscript.quit (1)
End If 

Set fso = CreateObject("Scripting.FileSystemObject")

If not fso.FolderExists(spath) Then
	wscript.echo "Folder " & spath & " is not valid. Please provide a valid syslog syslog location"
	Set fso = Nothing
	wscript.quit (1)
End If

Set cn  = CreateObject("ADODB.Connection")
Set cmd = CreateObject("ADODB.Command")
Set RS  = CreateObject("ADODB.Recordset")

	cn.Provider = "sqloledb"
	cn.Properties("Data Source").Value = DBServerName
	cn.Properties("Initial Catalog").Value = DBName
	cn.Properties("Integrated Security").Value = "SSPI"
	
	On Error Resume Next
	cn.Open
	checkerror "Error while opening connection to database on " & dbsersvername & "."
	On Error GoTo 0
	Cmd.ActiveConnection = cn
	cmd.CommandTimeout = ADOTimeOut
	wscript.echo "Starting process at " & Now()
	dtStart = Now()
	
	wscript.echo "Getting start date, querying database."

	cmdtext = "select isnull(max(dtTimeStamp),0) from " & tablename

	On Error Resume Next	
	RS.open cmdtext, cn
	checkerror "Error while fetching start date from database."
	On Error GoTo 0
	
	maxdate = rs(0)
	rs.close
	If maxdate < "1/31/07" or maxdate = "" Then 
		maxdate = "1/31/07"
		wscript.echo  "invalid date, it's now " & maxdate
	Else
		wscript.echo  "date is fine"
	End If
	maxdate = DateAdd("d", 1, maxdate)
	strdate = DatePart("YYYY",maxdate) & Right("0" & DatePart("m",maxdate),2) & Right("0" & DatePart("d",maxdate),2)
	sfilepath = spath & "\" & strdate & ".log"
	wscript.echo "MaxDate is " & maxdate & " and yesterday is " & DateAdd("d",-1,Now())
	Do until maxdate > DateAdd("d",-1,Now())
		If fso.fileexists(sfilepath) Then
			wscript.echo Now & " Processing " & sfilepath
			Set FileObj = fso.opentextfile(sfilepath)
			linecount = 0
			Do until fileobj.atendofstream
				On Error Resume Next
				strLine = Trim(FileObj.readline)
				checkerror "Error while reading next line."
				On Error GoTo 0
				strline = replace(strline,"'","''")
				strparts= split(strline," ")
				dtline = strparts(1) & " " & strparts(2) & " " & strparts(0) & " " & strparts(3) 
				device = strparts(4) 
				deviceIP = strparts(5)
				Severity = strparts(7)
				message = ""
				strclass = ""
				For x = 8 To UBound(strparts)
					Message = message & " " & strparts(x)
					If Left(strparts(x),1) = "%" Then strclass = strparts(x)
				Next 
				cmdtext = "INSERT INTO " & tablename & " (dtTimeStamp,DeviceName,DeviceIP,Severity,message,vcclass) VALUES ('"
				cmdtext = cmdtext & dtline & "','" & device & "','" & deviceip & "','" & severity & "','" 
				cmdtext = cmdtext & message & "','" & strclass & "')"
				Cmd.CommandText=cmdtext
				On Error Resume Next
				Cmd.Execute
				checkerror "Error while inserting into database table."
				On Error GoTo 0
				
				linecount = linecount + 1
				'If (linecount Mod 1000) = 0 Then wscript.echo "processed " & linecount & " lines"
			loop
			fileobj.close
			Set fileobj = nothing
			wscript.echo Now & " Processed " & linecount & " lines."
			wscript.echo Now & " Done with " & sfilepath
		Else
			wscript.echo "Can't find " & sfilepath
		End If 
		totallines = totallines + linecount
		'wscript.echo "Next Date"
		maxdate = DateAdd("d", 1, maxdate)
		strdate = DatePart("YYYY",maxdate) & Right("0" & DatePart("m",maxdate),2) & Right("0" & DatePart("d",maxdate),2)
		sfilepath = spath & "\" & strdate & ".log"
	loop 

Set cn        = nothing
Set cmd       = nothing
Set RS        = nothing
Set fso       = nothing
dtstop = Now()
wscript.echo Now & " Done."
wscript.echo "Processed " & Totallines & " lines in " & DateDiff("h",dtstart,dtstop) & " hours."
wscript.echo "Which is " & DateDiff("n",dtstart,dtstop) & " minutes."


Sub CheckError (strMsg)
	If err.number <> 0 Then
		wscript.echo strmsg & " Error number:" & err.number & " " & err.Description
		Set cn        = nothing
		Set cmd       = nothing
		Set RS        = nothing
		Set fso       = Nothing		
		wscript.quit (1)
	End If
End Sub