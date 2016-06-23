Option Explicit
Dim WshShell, strCmd, Quo, dtString, PathStr, FileName
Quo = Chr(34) ' Quote because of spaces
PathStr = Left (wscript.scriptFullname,InStr(wscript.scriptFullName,wscript.scriptname)-1)
FileName = PathStr & "spwho2out.txt"
Set WshShell =  wscript.createobject("wscript.shell")
dtString = Year( Now ) & "-" & Right( "0" & Month(Now), 2 ) & "-" & Right( "0" & Day(Now), 2 ) & "_" & Right("0" & Hour (Now),2) & "-" & Right("0" & Minute(now),2) & "-" & Right("0" & Second(Now),2) 
strCmd = "osql -Sgnettools15 -E -Q " & quo & "sp_who2" & quo & " -o " & quo & pathStr & dtString & ".txt" & quo & " -w255"
'wscript.echo strcmd
WshShell.run (strcmd)