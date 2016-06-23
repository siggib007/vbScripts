Dim iMsg,iConf 

Const MailServerName = "157.54.9.106" 
Const SMTPTimeout = 10
Const FromAddress = """Siggi Bjarnason"" <siggib@microsoft.com>"
Const ToAddress = """Siggi"" <siggi@icecomputing.com>, ""Siggi G. Bjarnason"" <siggib@foxinternet.net>"
Const Subject = "Hows it going?" 
Const MsgBody = "This is the body of the message"
Const cdoSendUsingPort = 2

wscript.echo "Starting test"

Set iMsg = CreateObject("CDO.Message") 
Set iConf = CreateObject("CDO.Configuration") 
Set Flds = iConf.Fields 

wscript.echo "setting mail server, etc."

With Flds 
  .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort 
  .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = MailServerName 
  .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = SMTPTimeout
  .Update 
End With 

wscript.echo "creating the message"

With iMsg 
  Set .Configuration = iConf 
      .To       = ToAddress
      .From     = FromAddress 
      .Subject  = Subject 
      .textbody = MsgBody
	wscript.echo "Sending the message"
      .Send 
End With

wscript.echo "done."