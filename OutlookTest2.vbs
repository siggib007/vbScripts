Const olFolderInbox = 6
Const olMail = 43
Const olEmbeddeditem = 5
Const PropName = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"

Set app = CreateObject("Outlook.Application")
set objNamespace = app.GetNamespace("MAPI")
set objFolders = objNameSpace.folders
for each objFolder in objFolders
  wscript.echo objFolder.name
next
set objInbox = objNameSpace.GetDefaultFolder(olFolderInbox)
set objRootFolder = objInbox.parent
set objFolderItems = objRootFolder.folders("A0Test 1").items
wscript.echo "Have your folder open listing items"
for each objItem in objFolderItems
	with objItem
		if objItem.Class = olMail then
			wscript.echo "Mail Item subject: " & .Subject
			wscript.echo " - ID:" & right(.EntryID,10)
		else
    		wscript.echo "Class: " & objItem.Class & " subject: " & .Subject
    		wscript.echo " - ID:" & right(.EntryID,10)
		end if
	end with
next
wscript.echo "That's all folks"