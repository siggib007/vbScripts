Option Explicit
Dim WshNetwork
Dim objArgs, strUserName, strInName

Set objArgs = wscript.Arguments

Set WshNetwork = WScript.CreateObject("WScript.Network")

If objArgs.Count = 1 then
	strUserName = objArgs(0)
	strInName = strUserName
	'wscript.echo "Input accepted"
Else
	strUserName = WshNetwork.UserName
	strInName = ""
	'wscript.echo "No input detected"
end if

WScript.Echo "Domain = " & WshNetwork.UserDomain
WScript.Echo "Computer Name = " & WshNetwork.ComputerName
WScript.Echo "Your User Name = " & WshNetwork.UserName
wscript.echo "Supplied User Name = " & strInName

wscript.echo "Full Name = " & getserver(strUserName)
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
  If objserver.eof = False Then
	   getserver = objserver("displayname").value
	Else
		 getserver = "not found error"
  End If  	
	Set objserver = Nothing
	Set sitecomm = Nothing
	Set oConn = Nothing
End Function 

Sub Usage
	Wscript.Echo "Usage:  whoareu3.vbs [username]"
	wscript.quit
end sub