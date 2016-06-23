option explicit
Const strURL = "http://ppexmlinterface/Post/Ticket_Diary_Create.asp"

creatediary 2430966,"test title",""

Sub CreateDiary (TicketID, subject, strmsg)
Dim oPoster, strData, xmlOK, objDocument, errind
Dim statuscode, statusdesc, errorcnt, stroutput, messagedesc, ErrorNum, rootNode, childNode

	strdata = "<?xml version=""1.0"" ?> <XMLFILE UserLogName=""gnsops"" FileId=""diaryCreate""> "
	strdata = strdata & "<TicketDiary TicketId=""" & ticketid & """ TicketDiaryType=""Actions Performed""" 
	strdata = strdata & " TicketDiaryTitle=""" & subject & """ TicketDiaryDesc=""" & strmsg & """></TicketDiary></XMLFILE>"
	'wscript.echo strdata
	
	Set objDocument = CreateObject("msxml2.DOMDocument")
	Set oPoster = CreateObject("Microsoft.XMLHTTP")
	oPoster.Open "POST", strURL, 0
	oPoster.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oPoster.Send strData
	'wscript.echo oposter.responsexml.xml
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
		                    wscript.echo "Return status " & oposter.status & " Error# " & errornum & " occured. " & childNode.getAttribute("ErrorDesc")
		               Else  
		               			errind = childNode.getAttribute("HasErrorInd")
		               			If errind = 1 Then 
		               				wscript.echo childNode.getAttribute("MessageDesc")
		               			Else
		               				wscript.echo "Success"
		                    End If
		               End If     
		          End If     
		     Next
	     Else
	          wscript.echo "Failed to load XML response."
	     End If
	Else
	     wscript.echo oPoster.responseText
	End If
End Sub 