Dim cn, rs, cmd

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")
Set cmd = CreateObject("ADODB.Command")


cn.Provider = "sqloledb"
cn.Properties("Data Source").Value = "satnetengfs01"
'cn.Properties("Initial Catalog").Value = "gnsim"
cn.Properties("Integrated Security").Value = "SSPI"
cn.Open

Cmd.ActiveConnection = cn
Cmd.CommandText = "delete from gns.dbo.dclist"

'Set rs = Cmd.Execute
cmd.execute
'wscript.echo rs.fields(0).value 
'rs.Close
cn.close

Set cn = nothing
Set rs = nothing
Set cmd = nothing