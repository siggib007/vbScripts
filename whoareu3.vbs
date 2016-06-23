'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/netdir/adsi/searching_with_activex_data_objects_ado.asp
Option Explicit
Dim objArgs
Set objArgs = wscript.Arguments

If objArgs.Count <> 1 then
	usage
end if

wscript.echo getserver(objArgs(0))

Function getServer(SamAccountName)
Dim AdsObject, oConn, SiteComm, objServer, NC

	Set AdsObject = GetObject("GC://RootDSE")
	NC = ADsObject.Get("rootDomainNamingContext")
	Set AdsObject=Nothing
	
	Set oConn = CreateObject("ADODB.Connection")
	oConn.Provider = "ADsDSOObject"
	oConn.Open "Active Directory Service Provider"

	Set siteComm = CreateObject("ADODB.Command")
	siteComm.ActiveConnection = oConn
	siteComm.commandText = "<GC://" & NC & ">;(SamAccountName=" & SamAccountName & ");displayname;subtree"
	Set objServer = siteComm.Execute()
	getserver = objserver("displayname")
	Set objserver = Nothing
	Set sitecomm = Nothing
	Set oConn = Nothing

end Function 

Sub Usage
	Wscript.Echo "Usage:  whoareu3.vbs [username]"
	wscript.quit
end sub