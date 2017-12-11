Const olFolderInbox = 6
Const olMail = 43
Const olEmbeddeditem = 5
Const PropName = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"

Set oShell = CreateObject( "WScript.Shell" )
strTemp=oShell.ExpandEnvironmentStrings("%temp%")
strTempmsg = strTemp & "\DPHJITSQHEAFEMTTBCGF.msg"
Set fso = CreateObject("Scripting.FileSystemObject")
Set app = CreateObject("Outlook.Application")
set objNamespace = app.GetNamespace("MAPI")
set objInbox = objNameSpace.GetDefaultFolder(olFolderInbox)
set objRootFolder = objInbox.Parent
Set objDestFolder = objRootFolder.Folders("00PhisingTest")
' set objInboxItems = objInbox.items
set objInboxItems = objRootFolder.Folders("A0Test 1").Items

wscript.echo "Have your inbox open checking for fish tests or emails as attachments"
for each objItem in objInboxItems
	with objItem
		if .Class = olMail then
			if .attachments.count > 0 then
				set objAttachment = .attachments.item(1)
				if objAttachment.type = olEmbeddeditem then
					wscript.echo "Has an email Attachment. From: " & .SenderName & " at: " & .ReceivedTime & " subject: " & .Subject
					wscript.echo " - Filename: " & objAttachment.Filename
					objAttachment.SaveAsFile (strTempmsg)
					set objExtMsg = app.CreateItemFromTemplate(strTempmsg)
					strExtHeader = objExtMsg.PropertyAccessor.GetProperty(PropName)
					iLoc1 = instr(1,strExtHeader,"X-PHISHTEST",1)
					if iLoc1 > 0 then
						wscript.echo " ++ This is a phish test message"
						.Move objDestFolder
	                    Set objReplyMsg = .Reply
	                    objReplyMsg.Body = "Thanks for reporting this. This message was a phishing test"
	                    objReplyMsg.Send
	                    wscript.sleep 20
	                    wscript.echo "Message moved, reply sent"
					else
						wscript.echo " -- Just a normal email attachment"
					end if
				else
					wscript.echo "Has a file Attachment. From: " & .SenderName & " at: " & .ReceivedTime & " subject: " & .Subject
					wscript.echo " - Filename: " & objAttachment.Filename
				end if
			else
				strHeader = .PropertyAccessor.GetProperty(PropName)
				iLoc1 = instr(1,strHeader,"X-PHISHTEST",1)
				if iLoc1 > 0 then
					wscript.echo "++ Go Fish. From: " & .SenderName & " at: " & .ReceivedTime & " subject: " & .Subject
					.Move objDestFolder
					wscript.sleep 20
					wscript.echo "message moved"
				else
					wscript.echo "Normal message, no attachment. From: " & .SenderName & " at: " & .ReceivedTime & " subject: " & .Subject
				end if
			end if
		else
			wscript.echo "Class: " & .class &  " From: " & .SenderName &  " subject: " & .Subject
		end if
	end with
next
If fso.FileExists(strTempmsg) Then
    fso.DeleteFile strTempmsg,true
End If
wscript.echo "That's all folks"