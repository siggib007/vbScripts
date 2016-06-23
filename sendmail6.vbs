'Const MailServerName = "smarthost.dns.microsoft.com" 
Const MailServerName = "mail.talnet.is" 
Const SMTPTimeout = 10
Const FromAddress = """Siggi"" <siggi@icecomputing.com>"
Const ToAddress = """Siggi Bjarnason"" <siggib@microsoft.com>,""Siggi G. Bjarnason"" <siggib@foxinternet.com>"
Const CCAddress = ""'"2062951027@mobile.att.net"
Const Subject = "Hows it going?" 
Const MsgBody = "This is the body of the message"
Const cdoSendUsingPort = 2

mysendmail subject,msgbody

Sub MySendMail(StrSubject,msg)
	Dim iMsg,iConf 
	
	Set iMsg = CreateObject("CDO.Message") 
	Set iConf = CreateObject("CDO.Configuration") 
	Set Flds = iConf.Fields 
	
	With Flds 
	  .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort 
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MailServerName 
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = SMTPTimeout
	  .Update 
	End With 
	
	With iMsg 
	  Set .Configuration = iConf 
	      .To       = ToAddress
	      .CC	= CCAddress
	      .From     = FromAddress 
	      .Subject  = StrSubject 
	      .textbody = Msg
      	      '.HTMLBody = msg
	      '.CreateMHTMLBody "http://ihsi"

	      .Send 
	End With
End Sub
