Option Explicit
Dim cn,cmd, RS, fso, spath, sfilepath, maxdate, strDate, dtline, device, deviceIP, Severity, message, tablename
Dim fileobj, strline, strparts, x, DBServerName, DBName, cmdtext, dtstart, dtstop, linecount, TotalLines, strclass
Dim objfileout, sOutPath, errmsg, today, dbDumpFreq, dtDay, iInst, stroutFile

Const ADOTimeOut = 300
Const MinDate = "06/30/08"

DBServerName = "by2netsql01"     ' -- Replace with db server name
DBName       = "syslogs"     ' -- Replace with db name
TableName    = "syslog"
dbDumpFreq   = 100000
iInst = 0

If wscript.arguments.count > 1 Then
	spath    = wscript.arguments(0)
	sOutPath = wscript.arguments(1)
Else
	wscript.echo "Please provide both syslog file location and output folder"
	wscript.quit (1)
End If 

Set fso = CreateObject("Scripting.FileSystemObject")

If not fso.FolderExists(spath) Then
	wscript.echo "Folder " & spath & " is not valid. Please provide a valid syslog syslog location"
	Set fso = Nothing
	wscript.quit (1)
End If

If not fso.FolderExists(soutpath) Then
	wscript.echo "Output folder " & soutpath & " does not exists, creating folder"
	On Error Resume Next
	fso.CreateFolder(soutpath)
	checkerror "Can't create folder " & soutpath & "."
	On Error GoTo 0
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

maxdate = CDate(formatdatetime(rs(0),2))
wscript.echo "Maxdate: *" & maxdate & "*"
rs.close
If maxdate < CDate(MinDate) or maxdate = "" Then 
	maxdate = CDate(MinDate)
	wscript.echo  "invalid date, it's now " & maxdate
Else
	wscript.echo  "date is fine"
End If
strdate = DatePart("YYYY",maxdate) & Right("0" & DatePart("m",maxdate),2) & Right("0" & DatePart("d",maxdate),2) & iInst
today = DatePart("YYYY",Now) & Right("0" & DatePart("m",Now),2) & Right("0" & DatePart("d",Now),2) & iInst
sfilepath = spath & "\" & strdate & Right("0" & iInst, 2) & ".log"
strOutFile = soutpath & "\" & strdate & ".txt"
wscript.echo "strdate: " & strdate
wscript.echo "strtoday: " & today
wscript.echo "sfilepath: " & sfilepath
wscript.echo "MaxDate is " & maxdate & " and today is " & Now()
Do until maxdate > Now()
	If fso.fileexists(sfilepath) Then
		wscript.echo Now & " Processing " & sfilepath
		Set FileObj = fso.opentextfile(sfilepath)
		Set objfileout = fso.createtextfile(strOutFile)	
		linecount = 0
		Do until fileobj.atendofstream
			today = DatePart("YYYY",Now) & Right("0" & DatePart("m",Now),2) & Right("0" & DatePart("d",Now),2) & iInst
			wscript.echo "strtoday: " & today
			On Error Resume Next
			strLine = Trim(FileObj.readline)
			checkerror "Error while reading next line."
			On Error GoTo 0
			strparts= split(strline," ")
			If UBound(strparts) > 3 Then
				dtline = strparts(1) & " " & strparts(2) & " " & strparts(0) & " " & strparts(3) 
				If IsDate(dtline) Then
					dtline = CDate(dtline)
					dtday = CDate(formatdatetime(dtline,2))
					If dtline > DateAdd("d", 0, maxdate) And UBound(strparts) > 7 Then
						device = strparts(4) 
						deviceIP = strparts(5)
						Severity = strparts(7)
						message = ""
						strclass = ""
						For x = 8 To UBound(strparts)
							Message = message & " " & strparts(x)
							If Left(strparts(x),1) = "%" Then strclass = strparts(x)
						Next 
						Message = replace(Message,vbtab,"    ")
						strparts = split(message,": ")
						If (UBound(strparts) > 0) and (strclass <> "") Then
							errmsg = strparts(UBound(strparts))
						Else
							strparts = split(message,":")
							If (UBound(strparts) > 0) and (strclass <> "") Then 
								errmsg = strparts(UBound(strparts))
							Else 
								errmsg = ""
							End If
						End If
						On Error Resume Next
						objfileout.writeline dtline & vbtab & dtday & vbtab & device & vbtab & deviceip & vbtab & severity & vbtab & strclass & vbtab & errmsg & vbtab & message
						checkerror "Error while writing outputline."
						On Error GoTo 0			
						linecount = linecount + 1
						If (linecount Mod dbDumpFreq) = 0 and strdate = today Then 
							wscript.echo Now() & " Processed " & linecount & " lines, dumping to DB."
							objfileout.close
							cmdtext = "bulk insert " & dbname & ".dbo." & tablename & " from '" & soutpath & "\" & strdate & ".txt'"
							wscript.echo "cmdtext: " & cmdtext
							Cmd.CommandText=cmdtext
							On Error Resume Next
							Cmd.Execute
							checkerror "Error while bulk importing " & soutpath & "\" & strdate & ".txt."
							Set objfileout = fso.createtextfile(soutpath & "\" & strdate & ".txt")	
							checkerror "Error while reopening output file"
							On Error GoTo 0
						End If
					End If
				End If 
			End If
		loop
		fileobj.close
		Set fileobj = nothing
		objfileout.close
		Set objfileout = Nothing
		cmdtext = "bulk insert " & dbname & ".dbo." & tablename & " from '" & soutpath & "\" & strdate & ".txt'"
		wscript.echo "cmdtext: " & cmdtext
		Cmd.CommandText=cmdtext
		On Error Resume Next
		Cmd.Execute
		checkerror "Error while bulk importing " & soutpath & "\" & strdate & ".txt."
		'fso.deletefile soutpath & "\" & strdate & ".txt", true
		On Error GoTo 0			
		wscript.echo Now & " Processed " & linecount & " lines."
		wscript.echo Now & " Done with " & sfilepath
	Else
		wscript.echo "Can't find " & sfilepath
	End If 
	totallines = totallines + linecount
	iInst = iInst + 1
	If iInst > 99 Then 	
		maxdate = DateAdd("d", 1, maxdate)
		iInst = 0
	End If 
	strdate = DatePart("YYYY",maxdate) & Right("0" & DatePart("m",maxdate),2) & Right("0" & DatePart("d",maxdate),2) ' & iInst
	sfilepath = spath & "\" & strdate & Right("0" & iInst, 2) & ".log"
	today = DatePart("YYYY",Now) & Right("0" & DatePart("m",Now),2) & Right("0" & DatePart("d",Now),2) '& iInst
	wscript.echo "strdate: " & strdate
	wscript.echo "strtoday: " & today
	wscript.echo "sfilepath: " & sfilepath
loop 

Set cn         = nothing
Set cmd        = nothing
Set RS         = nothing
Set FileObj    = nothing
Set objfileout = nothing
Set fso        = nothing
dtstop = Now()

'wscript.echo Now & " Done."
wscript.echo "Processed " & Totallines & " lines in " & DateDiff("n",dtstart,dtstop) & " minutes."
wscript.echo "or about " & DateDiff("h",dtstart,dtstop) & " hours."

If totallines > 0 Then
	wscript.echo "Completed successfully"
	wscript.quit (0)
Else
	wscript.echo "Did nothing, exiting abnormally"
	wscript.quit (2)
End If

Sub CheckError (strMsg)
	If err.number <> 0 Then
		On Error GoTo 0 
		wscript.echo strmsg & " Error number:" & err.number & " " & err.Description
		fileobj.close
		objfileout.close
		rs.close
		cn.close
		Set FileObj    = nothing
		Set objfileout = nothing
		Set cn         = nothing
		Set cmd        = nothing
		Set RS         = nothing
		Set fso        = Nothing		
		'fso.deletefile soutpath & "\" & strdate & ".txt", true
		'fso.deletefile soutpath & "\" & strdate & ".log", true
		wscript.quit (1)
	End If
End Sub