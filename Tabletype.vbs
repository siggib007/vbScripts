Dim cn
Dim rs
Dim fld
Dim cmdText

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")



        'Set ADO connection properties.
        cn.Provider = "sqloledb"
        cn.Properties("Data Source").Value = "b11gnmona04"
        cn.Properties("Initial Catalog").Value = "resultsdb"
	'cn.Properties("User ID").Value = "ihsuser"
	'cn.Properties("Password").Value = "ihsuser"
        cn.Properties("Integrated Security").Value = "SSPI"
        cn.Open

	'Call the DB
	'cmdText = "vwPerfTopEgress"
	cmdText = "vwPerfDCOverview"
        rs.Open cmdText, cn
        For Each fld In rs.Fields
		Wscript.echo "Field: " & fld.name & " type: " & fld.Type
	Next
	rs.close
	Set rs=nothing
	cn.close
	Set cn=nothing