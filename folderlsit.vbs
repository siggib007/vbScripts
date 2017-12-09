wscript.echo ShowFolderList ("%temp%")

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