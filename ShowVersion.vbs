scriptfullname = wscript.scriptfullname
scriptpath = left(scriptfullname,InStrRev(scriptfullname, "\"))
if instr(1,wscript.fullname, "wscript.exe",1) > 0 then
	strout = "This needs to be run from a command line. Please open up a command prompt, "
	strout = strout & "change to directory " & scriptpath & " and type in ""cscript.exe " & wscript.scriptname & """"
	wscript.echo strout
	wscript.quit
end if
wscript.echo wscript.fullname & " is version " & WScript.Version
