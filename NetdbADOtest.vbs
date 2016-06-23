Option Explicit 
Dim cn
Dim rs
Dim cmdText

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

function main()

	'Set ADO connection properties.
	cn.Provider = "sqloledb"
	cn.Properties("Data Source").Value = "netdb"
	cn.Properties("Initial Catalog").Value = "smarts"
	'cn.Properties("User ID").Value = "ihsuser"
	'cn.Properties("Password").Value = "ihsuser"
	cn.Properties("Integrated Security").Value = "SSPI"
	cn.Open
	
	cmdtext = "select distinct TROUBLETICKETID, SEVERITY, CREATEDAT, CLEAREDAT FROM [Smarts].[dbo].[IC_T_NOTIFICATION_OCCURRENCES] WHERE TROUBLETICKETID <>'' and CREATEDAT > dateadd(day,-1,getdate())"
	
	rs.Open cmdText, cn
	
	While not rs.eof
		wscript.echo rs.fields(0).value
		rs.movenext
	Wend 
	
	rs.close
	cn.close
	Set rs=nothing
	Set cn=nothing	

end function

main
