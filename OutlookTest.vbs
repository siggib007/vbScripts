Const olFolderInbox = 6
PropName = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"

Set app = CreateObject("Outlook.Application")
set objNamespace = app.GetNamespace("MAPI")
set objInboxItems = objNameSpace.GetDefaultFolder(olFolderInbox).items
' iCount = objFolder.items.Count
' wscript.echo "Inbox opened, has " & iCount & " items"
for each objItem in objInboxItems
	if objItem.Class = 43 and objItem.attachments.count = 0 then
		with objItem
			if .UnRead then
				strUnread = " [not read] "
			else
				strUnread = " [read] "
			end if
			wscript.echo "From: " & .Sender & " at: " & .ReceivedTime & strUnread & " subjet: " & .Subject
			strHeader = .PropertyAccessor.GetProperty(PropName)
			iLoc1 = instr(1,strHeader,"Reply-To:",1)
			if iLoc1 > 0 then
				iLoc2 = instr(iLoc1,strHeader,vbcrlf)
				if iLoc2 > iLoc1 then
					wscript.echo " Found Reply-To at " & iLoc1 & " to " & iLoc2
					strReplyTo = mid(strHeader,iLoc1-2,iLoc2 - iLoc1)
					wscript.echo " which is " & len(strReplyTo) & " long"
					 wscript.echo " - Reply-To: " & strReplyTo
				end if
			end if
			wscript.echo "trying again from " & iLoc2
			iLoc1 = instr(iLoc2,strHeader,"Reply-To:",1)
			if iLoc1 > 0 then
				iLoc2 = instr(iLoc1,strHeader,vbcrlf)
				if iLoc2 > iLoc1 then
					wscript.echo " Found Reply-To at " & iLoc1 & " to " & iLoc2
					strReplyTo = mid(strHeader,iLoc1-2,iLoc2 - iLoc1)
					wscript.echo " which is " & len(strReplyTo) & " long"
					 wscript.echo " - Reply-To: " & strReplyTo
				end if
			end if
			if .subject = "This is a test message" then
				wscript.echo .HTMLBody
			end if
		end with
	else
		' wscript.echo "Not a mail item"
	end if
next