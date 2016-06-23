'wscript.echo ShowFolderList ("\\tk2netdocs01\omni\Projects\10000-10999")
Option Explicit

   Dim fso, f, f1, s, sf, ssf, folderspec, f2, fc, objShell, rf, projnum, rfname
   
   folderspec = "\\tk2netdocs01\omni\Projects\"
   projnum = "10183"
   Set objShell = CreateObject("Shell.Application")
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFolder(folderspec)
   Set sf = f.SubFolders
   f2=""
   For each rf in sf
   	   rfname = split(rf.name,"-")
   	   wscript.echo CLng(projnum)
   	   wscript.echo rf.name
   	   wscript.echo CLng(rfname(0))
   	   wscript.echo CLng(rfname(1))
   	   'wscript.echo (rfname(0) < projnum) and (rfname(1) > projnum)
   	   'If (CLng(rfname(0)) < CLng(projnum)) and CLng((rfname(1)) < CLng(projnum)) Then 
	   'If  CLng((rfname(1)) > CLng(projnum)) Then 
   	   If (rfname(0) < projnum) and (projnum < rfname(1) ) Then 
		   wscript.echo "found the right root folder " & rf.name
	   	   Set ssf = rf.SubFolders
		   For Each f1 in ssf
		      s = s & f1.name & vbcrlf
		      'wscript.echo f1.name & vbcrlf
			  If InStr(f1.name,"10183") > 0 Then
			  	f2=folderspec & "\" & rf.name & "\" & f1.name
			  	Exit For
			  End If 
		   Next
	   Else
	   	  wscript.echo "Skipping " & rf.name
	   End If 
	   If f2 <> "" Then Exit For
   Next
   Set f = fso.GetFolder(f2)
   Set fc = f.files
   For each f1 in fc
       'If InStr(f1.name, "vsd") > 0 Then
       	If Right(f1.name,4) = ".vsd" Then
       		wscript.echo f2 & "\" & f1.name
    		objShell.ShellExecute (f2 & "\" & f1.name)
    		'objshell.shellexecute (f2)
       End If 
   Next   
   'wscript.echo s
   ' objShell.ShellExecute "iexplore", DMIURL & Me.cmbDevices.Value
