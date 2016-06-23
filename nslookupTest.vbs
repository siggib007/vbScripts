set objShell = createobject("wscript.shell")
strHostNmae = WScript.Arguments(0)
strParams = "%comspec% /c NSlookup " & strHostNmae
Set objExecObj = objShell.exec(strParams)

iAddr = 0
Do While Not objExecObj.StdOut.AtEndOfStream
	strText = objExecObj.StdOut.Readline()

	If instr(strText, "Server") then
		strServer = trim(replace(strText,"Server:",""))
	end if
	if instr (strText, "Name") Then
		strhost = trim(replace(strText,"Name:",""))
	End if
	if instr (strText, "Address") Then
		strAddr = trim(replace(strText,"Address:",""))
		if iAddr = 0 then
			strDNSAddr = strAddr
			iAddr = 1
		else
			strHostIP = strAddr
			iAddr = 2
		end if
	End if
Loop

if strHostIP <> "" then
	Wscript.echo "DNS Server: " & strServer & " @ " & strDNSAddr &" resolved " & strhost & " to " & strHostIP
else
	wscript.echo "Unable to resolve " & strHostNmae
end if
