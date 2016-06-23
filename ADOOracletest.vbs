Option Explicit 
Dim cn,rs,cmdtext, conn

cmdText = "SELECT DISTINCT sla_id, sla_name FROM sld where end_time is null"

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

Conn = "Data Source=BrixOracle;UID=registry;PWD=registry;"
cn.connectionstring = conn
wscript.echo "Attempting to open Connection"
cn.open
wscript.echo "attempting to execute query"
rs.Open cmdText, cn
While not rs.eof
	wscript.echo rs.fields(0).value & vbtab & rs.fields(1).value
	rs.movenext
Wend
rs.close
Set rs=nothing
cn.close
Set cn=nothing