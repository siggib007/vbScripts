const ForReading    = 1
const ForWriting    = 2
const ForAppending  = 8

strInFile = "C:\Users\sbjarna\Documents\IP Projects\Automation\UltraMPeerGroup\NewGGSN.txt"
' Creating a File System Object to interact with the File System
Set fso = CreateObject("Scripting.FileSystemObject")
if not fso.FileExists(strInFile) then
	wscript.echo "Can't find the input file " & strInFile & ". Aborting"
	wscript.quit
else
	wscript.echo "Found the input file"
end if

Set objFileIn = fso.opentextfile(strInFile, ForReading, False)
wscript.echo "Opened the input file starting to process"
While not objFileIn.atendofstream
	strHostName = objFileIn.readline
	Wscript.Stdout.Write "Trying to resolve " & strHostName & " ... "
	strAddr = GetIP (strHostName)
	if strAddr <> "" then
		Wscript.echo strHostName & "=" & strAddr
	else
		wscript.echo "Failed "
	end if
wend

objFileIn.close
set objFileIn = Nothing
set fso = nothing

function GetIP (strHostNmae)
	set objShell = createobject("wscript.shell")
	strParams = "%comspec% /c NSlookup " & strHostNmae
	Set objExecObj = objShell.exec(strParams)

	iAddr = 0
	Do While Not objExecObj.StdOut.AtEndOfStream
		strText = objExecObj.StdOut.Readline()
		if instr (strText, "Name") Then
			strhost = trim(replace(strText,"Name:",""))
		End if
		if instr (strText, "Address") Then
			strAddr = trim(replace(strText,"Address:",""))
		End if
	Loop
	if strhost <> "" then
		GetIP = strAddr
	else
		GetIP = ""
	end if
end function