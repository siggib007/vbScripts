
strOutFileName = "C:\Users\sbjarna\Documents\IP Projects\Automation\ARGACLs\ARGBatchCases.csv"
strListFileName = "C:\Users\sbjarna\Documents\IP Projects\Automation\ARGACLs\CaseCount.csv"
folderspec = "C:\Users\sbjarna\Documents\IP Projects\Automation\ARGACLs\HPNA Data"
strOut = ""

Set fso = CreateObject("Scripting.FileSystemObject")
Set objFileOut = fso.createtextfile(strOutFileName)
Set objListout = fso.createtextfile(strListFileName)
Set f = fso.GetFolder(folderspec)
Set fc = f.Files
objListout.writeline "File name,Case,ARG Count"
objFileOut.writeline "ARG Name,Old ACL Name,File Name,Case"

For Each f1 in fc
	x=0
	Set FileObj = fso.opentextfile(folderspec & "\" & f1.name)
	While not fileobj.atendofstream
		strLine = Trim(FileObj.readline)
		if strline<>"" then
			strParts = split(strline,",")
			if strParts(1)<>"hostName" then
				objFileOut.writeline strParts(1) & "," & strParts(2) & "," & f1.name & "," & mid(f1.name,8,len(f1.name)-11)
				x=x+1
			end if
		end If
	wend
	objListout.writeline f1.name & "," & mid(f1.name,8,len(f1.name)-11) & "," & x
next
wscript.echo "Done"