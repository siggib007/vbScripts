Option Explicit
Dim fso,srcFile,dstFile,f,BaseName,dtCreated,CreatedDate,strExt
Set fso = CreateObject("Scripting.FileSystemObject")
If wscript.arguments.count > 1 Then
	srcFile = wscript.arguments(0)
	dstFile = wscript.arguments(1)
Else
	wscript.echo "Need two parameters, source file path/name and destination path"
	wscript.quit
End If 
If not fso.fileexists(srcFile) Then
	wscript.echo "Source file," & srcFile & ", not found"
	wscript.quit
End If
If not fso.folderexists(dstfile) Then
	wscript.echo "Destination path does not exists, please specify a valid path"
	wscript.quit
End If 
Set f = fso.GetFile(srcFile)
CreatedDate = f.DateLastModified
BaseName = fso.GetBaseName(srcFile)
strExt = fso.GetExtensionName(srcFile)
dtCreated = DatePart("m",CreatedDate) & "-" & DatePart("d",CreatedDate) & _
			 "-" & DatePart("yyyy",createddate)
dstfile = fso.buildpath(dstfile, BaseName & dtCreated & "." & strExt)
wscript.echo "copy " & srcfile & " " & dstfile & " /y"
fso.copyfile srcfile, dstfile, true
If fso.fileexists(dstfile) Then
	wscript.echo "copy successful"
Else
	wscript.echo "copy failed"
End If



