Option Explicit 
Dim cn
Dim rs
Dim cmdText

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

function main()

        'Set ADO connection properties.
        cn.Provider = "sqloledb"
        cn.Properties("Data Source").Value = "olbessy"
        cn.Properties("Initial Catalog").Value = "cdinv"
	'cn.Properties("User ID").Value = "ihsuser"
	'cn.Properties("Password").Value = "ihsuser"
        cn.Properties("Integrated Security").Value = "SSPI"
        cn.Open

        'cmdText = "Select 'rowcount' = Count(*) from Customers"
	cmdText = "select 'rowcount' = count(*) from dbo.album"
        rs.Open cmdText, cn


        wscript.echo "cdinv.dbo.album contains " & rs.fields(0).value & " rows."

end function

main
