Option Explicit 
Dim cn,rs, fld, outstr, cmdtext, ticketid, updatecmd

Const MailServerName = "tk2smtp.phx.gbl" 
Const SMTPTimeout = 30
Const FromAddress = """GNS Infrastructure Operation"" <gnsops@phx.gbl>"
Const ToAddress =   """GNS Problem Management"" <netpro@microsoft.com>"
'Const ToAddress =   """Siggi Bjarnason"" <siggib@microsoft.com>"
Const CCAddress = """Linn Comptom"" <linnco@microsoft.com>"
Const Subject = "Password and/or snmp string needs to be changed" 
Const cdoSendUsingPort = 2
Const cdoNTLM = 2

Const DBServer = "by2netsql01"

Const teamname = "GNS Problem Management"
Const strurl = "http://ppexmlinterface/Post/Ticket_Update.asp"

'cmdText = "select vcType, dtUpdate, imaxage from cmdb.dbo.pwdchanged where dtnotified is null"
cmdText = "select vcType, dtUpdate, imaxage from cmdb.dbo.pwdchanged"
Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

cn.Provider = "sqloledb"
cn.Properties("Data Source").Value = DBServer
cn.Properties("Integrated Security").Value = "SSPI"
wscript.echo "Attempting to open Connection"
cn.open
wscript.echo "attempting to execute query"
rs.Open cmdText, cn
wscript.echo "got recordset, analyzing..."
outstr = ""
While not rs.eof
	'wscript.echo rs.fields(0).value & vbtab & rs.fields(1).value & vbtab & rs.fields(2).value
	If DateAdd("m",rs.fields(2).value -1,rs.fields(1).value) < Now() Then
		outstr = outstr & rs.fields(0).value & " needs to be changed by " & DateAdd("m",rs.fields(2).value,rs.fields(1).value)& vbcrlf
		updatecmd = Updatecmd & "Update cmdb.dbo.pwdchanged set dtUpdate = getdate() where vcType ='" & rs.fields(0) & "'" & vbcrlf
	End If 
	rs.movenext
Wend
rs.close

If outstr = "" Then 
	wscript.echo "everything is good"
Else
	wscript.echo "cuting a UTS ticket for remediation"
	ticketid = createticket(teamname,subject,outstr,strurl)
	If IsNumeric(ticketid) Then 
		outstr = outstr & vbcrlf & "Ticket " & ticketid & " has been cut to " & teamname & " for this."
	Else
		outstr = outstr & vbcrlf & "Failed to create a ticket or get the ticket number. " & ticketnumber
	End If
	wscript.echo outstr
	mysendmail subject,outstr
	wscript.echo "Mail sent"
	rs.Open updatecmd, cn
End If 

Set rs=nothing
cn.close
Set cn=Nothing

Sub MySendMail(StrSubject,msg)
	Dim iMsg,iConf,Flds
	
	Set iMsg = CreateObject("CDO.Message") 
	Set iConf = CreateObject("CDO.Configuration") 
	Set Flds = iConf.Fields 
	
	With Flds 
	  .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort 
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MailServerName 
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = SMTPTimeout
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")= cdoNTLM
	  .Update 
	End With 
	
	With iMsg 
	  Set .Configuration = iConf 
	      .To       = ToAddress
	      .CC	= CCAddress
	      .From     = FromAddress 
	      .Subject  = StrSubject 
	      .textbody = Msg
	      .Send 
	End With
End Sub

Function CreateTicket (teamname, subject, strmsg, strurl)
Dim ticketid, oPoster, strData, xmlOK, objDocument, errind
Dim statuscode, statusdesc, errorcnt, stroutput, messagedesc, ErrorNum, rootNode, childNode

	strdata = "<?xml version=""1.0"" ?> <XMLFILE UserLogName=""gnsops"" FileId=""TicketCreate"" Action=""Add"">"
	strdata = strdata & "<TICKET AssignedToTeamName=""" & TeamName 
	strdata = strdata & """ ClientImpactInd=""No"" CloseControlInd=""No"" PropertyName=""Network"" TicketDesc=""" & strmsg
	strdata = strdata & """ TicketPriority=""2:Med"" TicketProblemType=""Request"" TicketTitle=""" &  subject
	strdata = strdata & """ TicketType=""Request""></TICKET></XMLFILE>"
	
	Set objDocument = CreateObject("msxml2.DOMDocument")
	Set oPoster = CreateObject("Microsoft.XMLHTTP")
	oPoster.Open "POST", strURL, 0
	oPoster.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oPoster.Send strData
	
	If oPoster.responseXML.xml <> "" Then
	     strData = oPoster.responseXML.xml
	     objDocument.async = False
	     xmlOK = objDocument.loadXML(strData)
	     If xmlOK Then      'XML load success
		     Set rootNode = objDocument.documentElement	
		     For Each childNode in rootNode.childNodes
		          If childNode.nodeName = "XMLLogDetail" Then
		               ErrorNum = childNode.getAttribute("ErrorNum")               
		               If Not IsNull(ErrorNum) Then 
		                    ticketid = "Return status " & oposter.status & " Error# " & errornum & " occured. " & childNode.getAttribute("ErrorDesc")
		               Else  
		               			errind = childNode.getAttribute("HasErrorInd")
		               			If errind = 1 Then 
		               				ticketid = childNode.getAttribute("MessageDesc")
		               			Else
		                    	ticketid = childnode.getattribute("TicketId")
		                    End If
		               End If     
		          End If     
		     Next
	     Else
	          ticketid =  "Failed to load XML response."
	     End If
	Else
	     ticketid = oPoster.responseText
	End If
	createticket = ticketid
End Function 