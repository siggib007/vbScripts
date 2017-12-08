Const olFolderInbox = 6
Const olMail = 43
Const olEmbeddeditem = 5
Const PropName = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"

Set app = CreateObject("Outlook.Application")
set objNamespace = app.GetNamespace("MAPI")
set objFolderItems = objNameSpace.folders(1).folders("A10").items
' for each objFolder in objFolders
'   wscript.echo objFolder.name
' next
wscript.echo "Have your folder A10 open listing items"
for each objItem in objFolderItems
	if objItem.Class = olMail then
		with objItem
			wscript.echo "From: " & .Sender & " at: " & .ReceivedTime & " subjet: " & .Subject
		end with
  else
    wscript.echo "Class: " & objItem.Class
	end if
next
wscript.echo "That's all folks"