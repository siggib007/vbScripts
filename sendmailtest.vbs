Const MailServerName = "smtphost.redmond.corp.microsoft.com" 
Const SMTPTimeout = 10
Const FromAddress = """Siggi Bjarnason"" <siggib@microsoft.com>"
Const ToAddress =   """Siggi Bjarnason"" <siggib@microsoft.com>"
Const CCAddress = ""
Const Subject = "Hows it going?" 
Const MsgBody = "This is the body of the message"
Const cdoSendUsingPort = 2
Const cdoNTLM = 2

mysendmail subject,msgbody
wscript.echo "Mail sent"

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
