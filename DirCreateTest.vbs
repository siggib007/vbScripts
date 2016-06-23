Set fso = CreateObject("Scripting.FileSystemObject")

strCIQFolder = "\test4\more\stuff\testing\this"
CreatePath strCIQFolder
writelog "Done"

Function CreatePath (strFullPath)
dim pathparts, buildpath, part
	pathparts = split(strFullPath,"\")
	buildpath = ""
	wscript.echo "starting with fullpath='" & strFullPath & "' and buildpath='" & buildpath & "'"
	for each part in pathparts
		wscript.echo "in loop. part='" & part & "' and buildpath='" & buildpath & "'"
		if buildpath<>"" then 
			if buildpath = "\" then
				buildpath = buildpath & part
			else
				buildpath = buildpath & "\" & part
			end if
			if not fso.FolderExists(buildpath) then
				wscript.echo "need to create '" & buildpath & "'"
				fso.CreateFolder(buildpath)
			end if		
		else
			if part="" then
				buildpath = "\"
			else
				buildpath = part
			end if
		end if
	next
end function

Function WriteLog (strMsg)
'-------------------------------------------------------------------------------------------------'
' Function WriteLog (strMsg)                                                                      '
'                                                                                                 '
' This function accepts one input parameter, a string, and writes it to the screen or a file      '
' based on command line arguments                                                                 '
'-------------------------------------------------------------------------------------------------'

' Check if the script runs in CSCRIPT.EXE, i.e. is being run from command line, if so write to screen, otherwise do nothing
' Need to avoid having the script initiate 100 popups
	If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
		wscript.echo now & vbtab & strMsg
	end if


end function
