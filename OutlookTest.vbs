Const olFolderInbox = 6
Const olMail = 43
Const olEmbeddeditem = 5
Const PropName = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
Const strRootFolderName = "Siggi.Bjarnason@T-Mobile.com"
Const strDestFolderName = "01 Phising Test"
Const strInboxName = "A0 Test 1"
Const PhisingIndicate = "X-PHISHTEST"
Const MsgBody = "Thanks for reporting this. This message was a phishing test"

strMonthYear = monthname(month(now))&year(now)
wscript.echo "It is now " & strMonthYear
Set oShell = CreateObject( "WScript.Shell" )
set dictItems = CreateObject("Scripting.Dictionary")
Set fso = CreateObject("Scripting.FileSystemObject")
Set app = CreateObject("Outlook.Application")
set objNamespace = app.GetNamespace("MAPI")
' set objInbox = objNameSpace.GetDefaultFolder(olFolderInbox)
set objRootFolder = objNameSpace.folders(strRootFolderName)
if not FolderExists(objRootFolder,strDestFolderName) then
	objRootFolder.folders.add strDestFolderName
	wscript.echo strDestFolderName & " did not exists so I created it"
else
	wscript.echo strDestFolderName & " exists"
end if
if not FolderExists(objRootFolder,strInboxName) then
	' objRootFolder.folders.add strInboxName
	wscript.echo strInboxName & " does not exists, can't continue without a valid inbox"
	wscript.quit
else
	wscript.echo strInboxName & " exists"
end if
Set objDestRoot = objRootFolder.Folders(strDestFolderName)
if not FolderExists (objDestRoot,strMonthYear) then
	objDestRoot.folders.add strMonthYear
	wscript.echo strMonthYear & " did not exists inside " & strDestFolderName & " so I created it."
else
	wscript.echo strMonthYear & "esists inside " & strDestFolderName
end if
set objDestFolder = objDestRoot.folders(strMonthYear)
set objInboxItems = objRootFolder.Folders(strInboxName).Items

strTemp=oShell.ExpandEnvironmentStrings("%temp%")
strTempmsg = strTemp & "\DPHJITSQHEAFEMTTBCGF.msg"

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
					iLoc1 = instr(1,strExtHeader,PhisingIndicate,1)
					if iLoc1 > 0 then
						wscript.echo " ++ This is a phish test message"
						dictItems.add .entryid, objItem
	                    Set objReplyMsg = .Reply
	                    objReplyMsg.Body = MsgBody
	                    objReplyMsg.save
	                    wscript.echo "   reply sent"
					else
						wscript.echo " -- Just a normal email attachment"
					end if
				else
					wscript.echo "Has a file Attachment. From: " & .SenderName & " at: " & .ReceivedTime & " subject: " & .Subject
					wscript.echo " - Filename: " & objAttachment.Filename
				end if
			end if
			strHeader = .PropertyAccessor.GetProperty(PropName)
			if .attachments.count = 0 then
				wscript.echo "No Attachment. From: " & .SenderName & " at: " & .ReceivedTime & " subject: " & .Subject
			end if
			iLoc1 = instr(1,strHeader,PhisingIndicate,1)
			if iLoc1 > 0 then
				wscript.echo " ++ This is a phish test message"
				dictItems.add .entryid, objItem
			else
				wscript.echo " -- Normal message"
			end if
		else
			wscript.echo "Class: " & .class &  " From: " & .SenderName &  " subject: " & .Subject
		end if
	end with
next
wscript.echo vbcrlf & "Done Analysing inbox, filing the messages identified as test messages." & vbcrlf
for each ItemID in dictItems
	with dictItems(ItemID)
		wscript.echo "Moving message ID " & right(.EntryID,10)
		.move objDestFolder
	end with
next
If fso.FileExists(strTempmsg) Then
    fso.DeleteFile strTempmsg,true
End If
wscript.echo vbcrlf & "That's all folks"

Function FolderExists (objFolder, strFolderName)
	for each objSubFolder in objFolder.folders
		if lcase(objSubFolder.Name) = lcase(strFolderName) then
			FolderExists = true
			exit Function
		end if
	next
	FolderExists = false
End Function