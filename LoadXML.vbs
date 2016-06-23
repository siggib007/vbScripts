Option Explicit
Dim DBName, XMLFile, XSDFile, DBServerName, UserName, Password, sDSN, fso, objBL

     Set objBL = CreateObject("SQLXMLBulkLoad.SQLXMLBulkload.3.0")
     Set fso = CreateObject("Scripting.FileSystemObject")

     wscript.echo "Number of arguments supplied: " & wscript.arguments.count
     If wscript.arguments.count > 2 Then
	DBName   = wscript.arguments(0)
	XMLFile = wscript.arguments(1)
	XSDFile = wscript.arguments(2)
     Else
     	wscript.echo "Supply all arguments"
     	wscript.quit(1)
     End If
     
     If not fso.fileexists(xmlfile) Then 
     	wscript.echo xmlfile & " doesn't exist"
     	wscript.quite(2)
     End If 

     If not fso.fileexists(xsdfile) Then 
     	wscript.echo xsdfile & " doesn't exist"
     	wscript.quite(2)
     End If      		
     
     DBServerName          = "by2netsql01"     ' -- Replace with db server name
     UserName              = "ScriptRW"        ' -- Replace with user name
     Password              = "thisbites2."     ' -- Replace with pwd
     
     sDSN                  = "provider=SQLOLEDB.1;data source=" & DBServerName & ";database=" & DBName & ";uid=" & UserName & ";pwd=" & Password
     
     'wscript.echo sdsn
     
     objBL.ConnectionString = sDSN
     objBL.SGDropTables = True
     objBL.SchemaGen = True
     objBL.ForceTableLock = True     
     objBL.ErrorLogFile = "c:\XMLPullError.xml"
     objBL.Execute XSDFile, XMLFile
     set objBL=Nothing
     wscript.echo "Done."