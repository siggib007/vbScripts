Dim cn
Dim rs
Dim cmdText

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

'Set ADO connection properties.
cn.Provider = "MSDASQL"
cn.properties("DRIVER").value = "Microsoft Access Driver (*.mdb)"

cn.Properties("DefaultDir").Value = "D:\My Documents"
cn.Properties("DBQ").Value = "D:\My Documents\cdinv.mdb"
cn.Open
cmdText = "select count(*) albumcount from albums"
rs.Open cmdText, cn

wscript.echo "There are " & rs.fields(0).value & " CD's inventoried"