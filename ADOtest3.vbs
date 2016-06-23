Option Explicit 
Dim cn,rs, fld, outstr, cmdtext

Const DBServer = "sgbl1"
Const DefaultDB = "test"

cmdText = "select * from test.dbo.dcgroups"

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

cn.Provider = "sqloledb"
cn.Properties("Data Source").Value = DBServer
cn.Properties("Initial Catalog").Value = DefaultDB
cn.Properties("Integrated Security").Value = "SSPI"
wscript.echo "Attempting to open Connection"
cn.open
wscript.echo "attempting to execute query"
rs.Open cmdText, cn
While not rs.eof
	outstr = ""
	wscript.echo rs.fields(0).value
	rs.movenext
Wend
rs.close
Set rs=nothing
cn.close
Set cn=nothing