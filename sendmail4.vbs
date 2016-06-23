"http://schemas.microsoft.com/cdo/configuration/sendusing"
"http://schemas.microsoft.com/cdo/configuration/sendusername"
"http://schemas.microsoft.com/cdo/configuration/sendpassword"

Dim iMsg,iConf

Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields

With Flds
  .Item(cdoSendUsingMethod)       = 2 ' cdoSendUsingPort
  .Item(cdoSMTPServerName)        = "smarthost"
  .Item(cdoSMTPConnectionTimeout) = 10 ' quick timeout
  .Update
End With

With iMsg
  Set .Configuration = iConf
      .To       = """Siggi Bjarnason"" <siggib@microsoft.com>"
      .From     = """Siggi"" <siggi@icecomputing.com>"
      .Subject  = "Hows it going? I've attached my web page"
      .CreateMHTMLBody "http://mypage"
      .Send
End With
