   'Body of email message
   Dim msgBody
   msgBody="A mail from the Windows Script Host!"

   'Call our function with recipient, message and subject
   MySendMail "Siggi Bjarnason","siggib;siggi@icecomputing.com","Automated Message.",msgBody

Sub MySendMail(profilename,strrecipient,subject,msg)
Dim objSession, oInbox, colMessages, oMessage, colRecipients,recipientarray,recipient

       Set objSession = CreateObject("MAPI.Session")
       objSession.Logon profilename

       Set oInbox = objSession.Inbox
       Set colMessages = oInbox.Messages
       Set oMessage = colMessages.Add()
       Set colRecipients = oMessage.Recipients

	recipientarray = split (strrecipient,";")
	For each recipient in recipientarray
		colRecipients.Add recipient
		colRecipients.Resolve
	Next

       oMessage.Subject = subject
       oMessage.Text = msg
       oMessage.Send

       objSession.Logoff
       Set objSession = nothing

   End Sub 
