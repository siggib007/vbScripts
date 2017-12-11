Set oShell = CreateObject( "WScript.Shell" )
strTemp=oShell.ExpandEnvironmentStrings("%temp%")
wscript.echo "Listing directory for " & strTemp
wscript.echo ShowFolderList (strTemp)

Function ShowFolderList(folderspec)
   Dim fso, f, f1, s, sf
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFolder(folderspec)
   Set sf = f.SubFolders
   For Each f1 in sf
      s = s & f1.name
      s = s & vbcrlf
   Next
   ShowFolderList = s
End Function