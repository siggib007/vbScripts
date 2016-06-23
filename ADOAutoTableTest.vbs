Option Explicit 
Dim cn,rs, iRA, Cmd, x, fld, outstr, cmdtext

Const DBServer = "sgbl1"
Const adUseClient = 3
Const adUseNone = 1
Const adUseServer = 2
Const adOpenStatic = 3

cmdText = "select datacenter, shortcode, DCGroup from cmdb.dbo.dclist"

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

cn.Provider = "sqloledb"
cn.Properties("Data Source").Value = DBServer
cn.Properties("Integrated Security").Value = "SSPI"
wscript.echo "Attempting to open Connection"
cn.open
wscript.echo "attempting to execute query"
rs.Open cmdText, cn
For each fld in rs.fields
	outstr = outstr & fld.name & " " & fld.type & vbtab
Next
wscript.echo outstr
While not rs.eof
	outstr = ""
	For each fld in rs.fields
		outstr = outstr & fld.value & vbtab
	Next
	wscript.echo outstr
	rs.movenext
Wend
rs.close
Set rs=nothing
cn.close
Set cn=nothing