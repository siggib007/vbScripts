Option Explicit 
Dim srccn, srcrs, dstcn, dstrs, cmdtext, dsn, cmd

Const DBServer = "by2netsql01"
Const DestServer = "satnetengfs01"

cmdText = "select datacenter, shortcode, DCGroup from cmdb.dbo.dclist"

Set srccn = CreateObject("ADODB.Connection")
Set dstcn = CreateObject("ADODB.Connection")
Set srcrs = CreateObject("ADODB.Recordset")
Set dstrs = CreateObject("ADODB.Recordset")
Set cmd = CreateObject("ADODB.Command")

srccn.Provider = "sqloledb"
srccn.Properties("Data Source").Value = DBServer
srccn.Properties("Integrated Security").Value = "SSPI"
wscript.echo "Attempting to open connection to source"
srccn.open
wscript.echo "attempting to execute query"
srcrs.Open cmdText, srccn
dstcn.Provider = "sqloledb"
dstcn.Properties("Data Source").Value = destserver
dstcn.Properties("Integrated Security").Value = "SSPI"
wscript.echo "Attempting to open connection to destination"
dstcn.open
wscript.echo "Opening dest table"
dstrs.LockType = 3 'adLockOptimistic
dstrs.ActiveConnection = dstcn
dstrs.Source = "gns.dbo.dclist"
dstrs.Open
Cmd.ActiveConnection = dstcn
Cmd.CommandText = "delete from gns.dbo.dclist"
cmd.execute

While not srcrs.eof
	wscript.echo "Copying: " & srcrs.fields("datacenter").value & "   " & srcrs.fields("shortcode").value & "   " & srcrs.fields(2).value
	dstrs.addnew
		dstrs("datacenter").value = srcrs.fields("datacenter").value
		dstrs("shortcode").value = srcrs.fields("shortcode").value
		dstrs("DCGroup").value = srcrs.fields("DCGroup").value
	dstrs.update
	srcrs.movenext
Wend

srcrs.close
Set srcrs=nothing
Set dstrs = nothing
srccn.close
Set cmd = nothing
Set srccn = nothing
Set dstcn = nothing