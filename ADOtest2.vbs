Option Explicit 
Dim cn,rs, iRA, Cmd, x, fld, outstr, cmdtext

Const DBServer = "cpsssqla02"
Const DefaultDB = "reports"
Const UserID = "guest"
Const PWD = "readonly"
Const adUseClient = 3
Const adUseNone = 1
Const adUseServer = 2
Const adOpenStatic = 3

'cmdText = "select * from dbo.tblDataCenters"
'cmdtext = "sp_who"
cmdtext = "spWhereServer 'cpnettools2,cpsssqla02,cpsssqla01,cpssfsa01,sgb1,b11gnmona03,b11gnmona04,b11gnmona05,b11gnmona06'"

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

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
'cn.sp_who2 rs
wscript.echo "Attempting to open Connection"
cn.open
wscript.echo "attempting to execute query"
rs.Open cmdText, cn
'Set rs=cn.exectute(cmdtext)
'Set Cmd = CreateObject("ADODB.Command")
'Cmd.ActiveConnection = cn
'strSQLProducts = "select * from pubs.dbo.authors"
'Cmd.CommandText = "spCheckClusterCount"
'Set rs = Cmd.Execute(iRA)
'rs.movelast
x=0
'While not rs.eof
'	rs.movenext
'	x=x+1
'Wend
'wscript.echo "Records Affected: " & irA
While not rs.eof
	outstr = ""
	For each fld in rs.fields
		outstr = outstr & fld.value & vbtab
	Next
	wscript.echo outstr
	rs.movenext
	'x=x+1
Wend
'wscript.echo "BOF: " & rs.bof
'wscript.echo "EOF: " & rs.eof
'wscript.echo "State: " & rs.state
'wscript.echo "Num Records: " & rs.recordcount
'wscript.echo "X = " & x
rs.close
Set rs=nothing
cn.close
Set cn=nothing