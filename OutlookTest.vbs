Const olFolderInbox = 6
Const olMail = 43
Const olEmbeddeditem = 5
Const PropName = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
const TempMsg = "c:\temp\TempEmail.msg"

Set fso = CreateObject("Scripting.FileSystemObject")
Set app = CreateObject("Outlook.Application")
set objNamespace = app.GetNamespace("MAPI")
set objInboxItems = objNameSpace.GetDefaultFolder(olFolderInbox).items
wscript.echo "Have your inbox open checking for fish tests or emails as attachments"
for each objItem in objInboxItems
	if objItem.Class = olMail then
		with objItem
			strHeader = .PropertyAccessor.GetProperty(PropName)
			iLoc1 = instr(1,strHeader,"X-Testing",1)
			if iLoc1 > 0 then
				wscript.echo "mytest. From: " & .Sender & " at: " & .ReceivedTime & " subjet: " & .Subject
			end if
			iLoc1 = instr(1,strHeader,"X-PHISHTEST",1)
			if iLoc1 > 0 then
				wscript.echo "Go Fish. From: " & .Sender & " at: " & .ReceivedTime & " subjet: " & .Subject
			end if
			if .attachments.count > 0 then
				set objAttachment = .attachments.item(1)
				if objAttachment.type = olEmbeddeditem then
					wscript.echo "Has Attachment. From: " & .Sender & " at: " & .ReceivedTime & " subjet: " & .Subject
					wscript.echo " - Filename: " & objAttachment.Filename
					objAttachment.SaveAsFile (TempMsg)
					set objExtMsg = app.CreateItemFromTemplate(TempMsg)
					strExtHeader = objExtMsg.PropertyAccessor.GetProperty(PropName)
					iLoc1 = instr(1,strExtHeader,"X-Testing",1)
					if iLoc1 > 0 then wscript.echo " ++ This is a plain test message"
					iLoc1 = instr(1,strExtHeader,"X-PHISHTEST",1)
					if iLoc1 > 0 then wscript.echo " ++ This is a phish test message"
				end if
			end if
		end with
	end if
next
If fso.FileExists(TempMsg) Then
    fso.DeleteFile TempMsg,true
End If
wscript.echo "That's all folks"