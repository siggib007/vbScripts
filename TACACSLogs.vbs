Option Explicit
Dim cn,cmd, ServerRS, FolderRS, DateRS, fso, server, spath, sfolder,stable,sfile, sfilepath, maxdate, strDate
Dim fileobj, strline, strparts, x, DBServerName, DBName, cmdtext, dtstart, dtstop, objfileout, linecount, TotalLines

Const ADOTimeOut = 60

DBServerName   = "by2netsql01"     ' -- Replace with db server name
DBName         = "TACACS"     ' -- Replace with db name

Set cn         = CreateObject("ADODB.Connection")
Set cmd        = CreateObject("ADODB.Command")
Set ServerRS   = CreateObject("ADODB.Recordset")
Set FolderRS   = CreateObject("ADODB.Recordset")
Set DateRS     = CreateObject("ADODB.Recordset")

Set fso        = CreateObject("Scripting.FileSystemObject")

Set objfileout = fso.createtextfile("ACSPars.log")


cn.Provider                                = "sqloledb"
cn.Properties("Data Source").Value         = DBServerName
cn.Properties("Initial Catalog").Value     = DBName
cn.Properties("Integrated Security").Value = "SSPI"
On Error Resume Next
	cn.Open
	checkerror "Error while opening connection to database on " & dbsersvername & "."
On Error GoTo 0

Cmd.ActiveConnection                       = cn
cmd.CommandTimeout                         = ADOTimeOut

wscript.echo "Starting process at " & Now()
dtStart = Now()
cmdtext = "select FolderName, TableName, BaseFileName from TACACS.dbo.Folder2Table where tablename <>''"
On Error Resume Next	
FolderRS.open cmdtext, cn
checkerror "Error while fetching folder information from database."
On Error GoTo 0

cmdtext = "select ServerName, ShareName from TACACS.dbo.LogFileLocations"
On Error Resume Next	
ServerRS.open cmdtext, cn
checkerror "Error while fetching server information from database."
On Error GoTo 0
wscript.echo "Got both server and folder info."

Do until ServerRS.eof
	Server = ServerRS.fields(0).value
	spath = ServerRS.fields(1).value
	folderrs.movefirst
	Do until FolderRS.eof
		sfolder = FolderRS.fields(0).value
		stable =  FolderRS.fields(1).value
		sFile =  FolderRS.fields(2).value
		cmdtext = "select isnull(max(timestamp),0)  from [" & stable & "]"
		objfileout.writeline cmdtext
		On Error Resume Next	
		dateRS.open cmdtext, cn
		checkerror "Error while fetching date information from database."
		On Error GoTo 0
		maxdate = CDate(formatdatetime(daters(0),2))
		daters.close
		objfileout.writeline "maxdate: " & maxdate
		If maxdate < "2/15/07" or maxdate = "" Then 
			maxdate = CDate("2/15/07")
			objfileout.writeline  "invalid date, it's now " & maxdate
		Else
			objfileout.writeline  "date is fine"
		End If
		maxdate = DateAdd("d", 1, maxdate)
		strdate = DatePart("YYYY",maxdate) & "-" & Right("0" & DatePart("m",maxdate),2) & "-" & Right("0" & DatePart("d",maxdate),2)
		sfilepath = spath & "\" & sfolder & "\" & sfile & " " & strdate & ".csv"
		objfileout.writeline sfilepath
		objfileout.writeline  Now & " Processing " & spath & "\" & sfolder & "\" & sfile & "......"
		objfileout.writeline  "MaxDate is " & maxdate & " and yesterday is " & DateAdd("d",-1,Now())
		Do until maxdate > DateAdd("d",-1,Now())
			objfileout.writeline  "Check to see if " & sfilepath & " exists."
			If fso.fileexists(sfilepath) Then
				wscript.echo Now & " Processing " & sfilepath
				Set FileObj = fso.opentextfile(sfilepath)
				strLine = Trim(FileObj.readline) ' skip first line the header line
				linecount = 0
				While not fileobj.atendofstream
					strLine = Trim(FileObj.readline)
					strline = replace(strline,"'","''")
					strparts = split(strline,",")
					cmdtext = "insert into [" & stable & "] values ('" 
					cmdtext = cmdtext & strparts(0) & " " & strparts(1) & "','"
					For x = 2 to UBound(strparts)
						cmdtext = cmdtext & strparts(x) & "','"
					Next
					cmdtext = Left(cmdtext,Len(cmdtext)-2) & ")"
					objfileout.writeline  cmdtext
					Cmd.CommandText=cmdtext
					On Error Resume Next
					Cmd.Execute
					checkerror "Error while inserting data."
					On Error GoTo 0
					linecount = linecount + 1			
				Wend
				fileobj.close
				Set fileobj = nothing
				wscript.echo Now & " Processed " & linecount & " lines."
			Else
				objfileout.writeline "Can't find " & sfilepath
			End If 
			totallines = totallines + linecount
			'wscript.echo "Next Date"
			maxdate = DateAdd("d", 1, maxdate)
			objfileout.writeline  "current: " & maxdate	
			strdate = DatePart("YYYY",maxdate) & "-" & Right("0" & DatePart("m",maxdate),2) & "-" & Right("0" & DatePart("d",maxdate),2)
			sfilepath = spath & "\" & sfolder & "\" & sfile & " " & strdate & ".csv"
			objfileout.writeline  sfilepath						
		loop
		FolderRS.movenext
	loop
	ServerRS.movenext
loop

serverrs.close
folderrs.close

Set cn        = nothing
Set cmd       = nothing
Set ServerRS  = nothing
Set FolderRS  = nothing
Set DateRS    = nothing
Set fso       = nothing
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
		wscript.echo strmsg & " Error number:" & err.number & " " & err.Description
		Set cn        = nothing
		Set cmd       = nothing
		Set ServerRS  = nothing
		Set FolderRS  = nothing
		Set DateRS    = nothing
		Set fso       = nothing		
		wscript.quit (1)
	End If
End Sub