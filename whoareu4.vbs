Option Explicit
Dim WshNetwork

Set WshNetwork = WScript.CreateObject("WScript.Network")
WScript.Echo "Domain = " & WshNetwork.UserDomain
WScript.Echo "Computer Name = " & WshNetwork.ComputerName
WScript.Echo "User Name = " & WshNetwork.UserName

wscript.echo "Full Name = " & getserver(WshNetwork.UserName)
Set WshNetwork = Nothing

Function getServer(SamAccountName)
Dim AdsObject, oConn, SiteComm, objServer, NC
  'wscript.echo "Looking up the full name for " & samaccountname
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
	getserver = objserver("displayname").value
	Set objserver = Nothing
	Set sitecomm = Nothing
	Set oConn = Nothing
End Function 
