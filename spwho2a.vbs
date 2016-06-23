Option Explicit
Dim WshShell, strCmd, Quo, dtString, PathStr, tempFileName, PermFileName
Quo = Chr(34) ' Quote because of spaces
PathStr = Left (wscript.scriptFullname,InStr(wscript.scriptFullName,wscript.scriptname)-1)
tempFileName = PathStr & "spwho2out.tmp"
PermFileName = PathStr & "spwho2out.txt"
Set WshShell =  wscript.createobject("wscript.shell")
dtString = Year( Now ) & "-" & Right( "0" & Month(Now), 2 ) & "-" & Right( "0" & Day(Now), 2 ) & "_" & Right("0" & Hour (Now),2) & "-" & Right("0" & Minute(now),2) & "-" & Right("0" & Second(Now),2) 
strCmd = "osql -Sgnettools15 -E -Q " & quo & "sp_who2" & quo & " -o " & quo & tempFileName & quo & " -w255"
wscript.echo strcmd
'WshShell.run strcmd,8,true
strcmd = "echo " & dtstring & " 1>> " & PermFileName
wscript.echo strcmd
WshShell.run strcmd,8,true
strcmd = "Type " & tempFileName & " >> " & PermFileName
wscript.echo strcmd
WshShell.run strcmd,8,true

