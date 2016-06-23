Const DBSERVERNAME = "Gnettools53" 
Const DATABASE = "labinv" 
Const USERID = "trenduser" 
Const DBPASSWORD = "trenduser" 
'Const QUERYSTR = "select * from ConfigfilesView" ' For .col files
'Const QUERYSTR = "select * from InfilesView" ' For Infiles
Const FIELDNAME = "0" 
Const DEBUGLEVEL = 1

Sub Main(QUERYSTR, TextStr)
	 
	Dim NeedtoAlert, DetailMsg, FilePathName
	Dim ConfigDBCN, ConfigDBRS, NetworkFlds, Networkfld
	Dim iFieldID, strFieldID, FSO
	
	CRLF = Chr(13) + Chr(10)

	Set fso = CreateObject("Scripting.FileSystemObject")
	'Create ADO Objects
	Set ConfigDBCN = CreateObject("ADODB.Connection")
	Set ConfigDBRS = CreateObject("ADODB.Recordset")
	If DEBUGLEVEL >  0 then wscript.echo "Establishing Database connection"
	'Set ADO connection properties.
    	ConfigDBCN.Provider = "sqloledb"
    	ConfigDBCN.Properties("Data Source").Value = DBSERVERNAME
    	ConfigDBCN.Properties("Initial Catalog").Value = DATABASE
   	ConfigDBCN.Properties("User ID").Value  = USERID
	ConfigDBCN.Properties("Password").Value = DBPASSWORD
    	ConfigDBCN.Open
	If DEBUGLEVEL > 0 then wscript.echo "Database connection Established. Opening Recordset"

	'Open ADO Record set
    	ConfigDBRS.Open QUERYSTR, ConfigDBCN

    	'Set the cursor to the first filed in the first row
	Set NetworkFlds = ConfigDBRS.Fields

	If isnumeric(FIELDNAME) Then 
		iFieldID=cint(FIELDNAME)
	    Set Networkfld = NetworkFlds(iFieldID)
	Else
		strFieldID = FIELDNAME
	    Set Networkfld = NetworkFlds(strFieldID)
	End If

	If DEBUGLEVEL > 0 then wscript.echo "recordset open and filed set. Looping throught the results."

	'Loop through the recordset and create a comma seperated result string	
	Do Until  ConfigDBRS.EOF
		FilePathName = "\\" & trim(Networkfld.Value) & "\c$\boot.ini"
		If DEBUGLEVEL > 0 then wscript.echo "FileName=" & FilePathName 
		If FilePathName <> "" Then
			If Not fso.FileExists(FilePathName) Then 
				NeedtoAlert = true
				detailmsg = detailmsg & crlf & FilePathName 
			End If			
		Else
			detailmsg = "no file name entered for the check"
			wscript.echo detailmsg
		End If
		ConfigDBRS.movenext
	Loop 

	If NeedtoAlert Then
		wscript.echo crlf & TextStr &". The follow files couln't be found:" & detailmsg
	Else 
		wscript.echo crlf & TextStr & " All files successfully accessed"
	End If
	Set ConfigDBCN=Nothing
	Set ConfigDBRS=Nothing
	Set NetworkFlds=Nothing
	Set Networkfld=Nothing


End Sub

main "select machinename from servers", "Machine Check"
'main "select * from ConfigfilesView", "Config.col File Search"
