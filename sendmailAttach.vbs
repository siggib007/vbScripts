'Const MailServerName = "smtphost.redmond.corp.microsoft.com" 
Const MailServerName = "satsmtpa01"
Const SMTPTimeout = 10
Const FromAddress = """GNS Problem Management"" <netpro@microsoft.com>"
Const ToAddress =   """Siggi Bjarnason"" <siggib@microsoft.com>"
Const CCAddress = ""
Const Subject = "Attachment Test" 
Const MsgBody = "This is the body of the message. There should be an attachment with a TACACS report."
Const cdoSendUsingPort = 2
Const cdoNTLM = 2

mysendmail subject,msgbody
wscript.echo "Mail sent"

Sub MySendMail(StrSubject,msg)
	Dim iMsg,iConf,Flds
	Dim iBp 'As CDO.IBodyPart
	
	'Set iBp = creaateobject("CDO.IBodyPart")
	'Set iBp = iMsg.AddAttachment("c:\myfiles\file.doc")
	
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
	      .AddAttachment("d:\siggib\netpro\TACACSUsers.csv")
	      .AddAttachment("d:\siggib\netpro\TACACSGroup.txt")
	      .Send 
	End With
End Sub
