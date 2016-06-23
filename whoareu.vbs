'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/netdir/adsi/searching_with_activex_data_objects_ado.asp


Set objArgs = wscript.Arguments

If objArgs.Count <> 1 then
	usage
end if

strUser_to_Search = objArgs(0)

objConnector = objConn
strNC = getNC(strDomain)
Set rsComputer = getServer(strNC, objConnector, strUser_to_Search)

wscript.echo rsComputer("displayname")

function objConn()
	Set oConn = CreateObject("ADODB.Connection")
	oConn.Provider = "ADsDSOObject"
	oConn.Open "Active Directory Service Provider"
	Set objConn = oConn
end function

function getNC(domain)
	Set AdsObject = GetObject("GC://RootDSE")
	getNC = ADsObject.Get("rootDomainNamingContext")
	Set AdsObject=nothing
end function


function getServer(nc, siteConn, SamAccountName)
	strSearchBase = "GC://" & nc
	strFilter = "SamAccountName=" & SamAccountName
	strAttribs = "displayname"
	strScope = "subtree"
	
	Set siteComm = CreateObject("ADODB.Command")
	siteComm.ActiveConnection = siteConn
	strCommandText =  "<" & strSearchBase & ">;(" & strFilter & ");" & strAttribs & ";" & strScope
	siteComm.commandText = strCommandText
	Set getServer = siteComm.Execute()

end function

Sub Usage
	Wscript.Echo "Usage:  whoareu.vbs [username]"
	wscript.quit
end sub