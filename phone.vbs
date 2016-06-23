'==========================================================================
'
'  NAME: phone.vbs
'
' AUTHOR: a-rastok
' DATE  : 4/5/2004
'
' Modified by: a-siggib
' Date  : 9/6/04
' COMMENT: retrieves current logged on user name for computer and the phone extension
'*                
'
'==========================================================================
'********************************************************************
'Option Explicit
on error resume next

Dim strComputerName
Dim WMIServices
Dim objUserSet
Dim oWshShell
Dim User
Set objArgs = wscript.Arguments

If objArgs.Count <> 1 then
	usage
	wscript.quit
end if

Set oWshShell = CreateObject("Wscript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set objTextFile = objFSO.OpenTextFile("\\tkihsfs01\public\siggib\VBScript\SvrList.txt", ForReading)
Set objOutputFile = objFSO.createTextFile(objArgs(0), true)
objOutputFile.writeline "Analyst" & space(6) & "Extension"  & space(5) & "Console#"   &  space(5) & "GROUP"
objOutputFile.writeline "................................................."

Do

  comps = objTextFile.ReadLine
  compt = split(comps)
  strComputerName = compt(0)

  'wscript.Echo comps
  'Wscript.Echo strcomputername & space(3)& compt(1)
  WMIServices = "winmgmts:{impersonationLevel=impersonate}!//"& strComputerName &""

  On Error Resume Next
  Set objUserSet = GetObject( WMIServices ).InstancesOf ("Win32_ComputerSystem")
  objConnector = objConn
  strNC = getNC()

  for each User in objUserSet

    If User.UserName <> "" Then
       namespace = 23 - len(User.UserName)
       consolenum = Right(strComputerName, 2)
       namespace2 = 15 - len(compt(1))
       username = split(User.UserName, "\")
       Set rsComputer = getServer(strNC, objConnector, UserName(1))
       objOutputFile.writeline rsComputer("displayname")
       objOutputFile.writeline  ucase(UserName(1)) & space(namespace) & compt(1) & space(10) & consolenum & space(namespace2) & compt(2)  
    Else
      wscript.Sleep 1 'Wscript.Echo "There are no users currently logged in at " & strComputerName
    End If
  Next
Loop While objTextFile.AtEndOfStream = False

objOutputFile.writeline 
objOutputFile.writeline 
objOutputFile.writeline "................................................."
objOutputFile.writeline 
objOutputFile.writeline "NB: verify your phone extension above with phone next to you"

function objConn()
	Set oConn = CreateObject("ADODB.Connection")
	oConn.Provider = "ADsDSOObject"
	oConn.Open "Active Directory Service Provider"
	Set objConn = oConn
end function

function getNC()
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
	Wscript.Echo "Usage:  phone.vbs outputfilename"
	wscript.quit
end sub

