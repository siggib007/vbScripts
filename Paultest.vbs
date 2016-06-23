Dim cn, cmd, i, rs, cmdText1
Dim fld, fldValue, StrOut, strTD, strTableHeader

Const TristateTrue = -1
Const TristateFalse = 0
Const TristateUseDefault = -2
Const adUseClient = 3
Const adUseNone = 1
Const adUseServer = 2
Const adOpenStatic = 3

Const Indent = "    "
Const DBServer = "ppmsql"
Const DefaultDB = "ppmReporting"
Const UserID = "ppmMetrics"
Const PWD = "ppmMetrics"
Const ReportTitle = ""

ReDim strAttach(0)

cmdText1 = "SELECT count(*) FROM PPMReporting.dbo.vProjOrg vProjOrg WHERE (vProjOrg.ResponsibleIT='MSNetwork Engineering') AND (vProjOrg.Status Not In ('Cancelled','Deleted','Closed'))"
Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")
Set cmd = CreateObject("ADODB.Command")
strOut = "Number of open PPM projects: "

Sub Main()
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
	rs.cursorlocation = adUseClient
	rs.cursortype = adopenstatic

	cn.Open
	Cmd.ActiveConnection = cn
	'Call the DB for recordset
	Cmd.CommandText = cmdText1
	Set rs = Cmd.Execute
	
	Do Until rs Is Nothing
		'Set rs = rs.NextRecordset
		If not rs.BOF Then
			strOut = strout & rs.fields(0).value
		End If 
	Loop	
	cn.close	
End Sub

main
Set rs=nothing
Set cn=nothing
Set cmd=nothing
wscript.echo strout
