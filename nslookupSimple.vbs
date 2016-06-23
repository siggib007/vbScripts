strHostNmae = WScript.Arguments(0)

strAddr = GetIP (strHostNmae)

if strAddr <> "" then
	Wscript.echo strHostNmae & ": " & strAddr
else
	wscript.echo "Unable to resolve " & strHostNmae
end if

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