Option Explicit
Dim Quo, dtString, PathStr, SourcePath, DestinationPath, Filename, SrcExtension
Dim fso, LogFileObj, strLogFilePath, Scriptname, srcFile, dtFile, dstFile, tmpFile

Const ForAppending = 8  'vbscript FSO constant

Quo = Chr(34) ' Quote because of spaces
PathStr = Left (wscript.scriptFullname,InStr(wscript.scriptFullName,wscript.scriptname)-1)
Scriptname = Left (wscript.scriptname, Len(wscript.scriptname)-4)
strLogFilePath = PathStr & Scriptname & ".log"
dtString = Year( Now ) & Right( "0" & Month(Now), 2 ) & Right( "0" & Day(Now), 2 ) & Right("0" & Hour (Now),2) & Right("0" & Minute(now),2) 

Set fso = CreateObject("Scripting.FileSystemObject")
Set LogFileObj = fso.OpenTextFile(strLogFilePath, ForAppending, True)

Sub LogMessage (strMessage)
	wscript.echo Now & vbtab & strMessage
	LogFileObj.writeline Now & vbtab & strMessage
End Sub

Sub LogError(CodeSegment)
	logMessage "Error #" & CStr(Err.Number) & " " & quo & Err.Description & quo & " occured at " & CodeSegment & ". Aborting"
	logMessage "Script aborted"
	wscript.quit
End Sub

On Error Resume Next

LogMessage "Script starting"
If wscript.arguments.count > 3 Then
	SourcePath = wscript.arguments(0)
	DestinationPath = wscript.arguments(1)
	Filename = wscript.arguments(2)
	SrcExtension = wscript.arguments(3)
Else
	LogMessage "Required input not detected. Aborting Script"
	logMessage "Script aborted"	
	wscript.echo "Usage: " & wscript.scriptname & " SourcePath DestinationPath Filename filenameExt [dstExt]"
	wscript.quit
End If
srcFile = SourcePath & "\" & FileName & "." & SrcExtension
dtFile = SourcePath & "\" & dtString & ".txt"
dstFile = DestinationPath & "\" & dtString & ".txt"
tmpFile = SourcePath & "\" & FileName & ".tmp"
LogMessage "SourcePath=" & SourcePath
LogMessage "DestinationPath=" & DestinationPath
LogMessage "Filename=" & Filename
LogMessage "SrcExtension=" & SrcExtension
LogMessage "Full Source path filename=" & srcFile
LogMessage "dtString=" & dtString
LogMessage "tmpFile=" & tmpFile
LogMessage "dstFile=" & dstFile
If wscript.arguments.count > 4 Then
	logmessage "Destination extension detected"
	dstFile = DestinationPath & "\" & FileName & "." & wscript.arguments(4)
	logmessage "dstFile now " & dstFile
End If
If err.number <> 0 Then LogError "init"

If fso.FileExists(tmpFile) Then
	LogMessage "Temp file exists, last attempt assumed failed. Retrying"
	fso.CopyFile tmpFile, dstFile, false
	If err.number <> 0 Then LogError "Retry CopyFile"
Else
	If fso.FileExists(srcFile) Then
		fso.MoveFile srcFile, tmpFile
		If err.number <> 0 Then LogError "MoveFile"
		fso.CopyFile tmpFile, dstFile, false
		If err.number <> 0 Then LogError "CopyFile"
	Else
		LogMessage srcFile & " not found, aborting."
		logMessage "Script aborted"		
		wscript.quit
	End If 
End If

If fso.FileExists(dstFile) Then
	fso.deletefile tmpFile
	If err.number <> 0 Then LogError "deletefile"
Else
	LogMessage dtstring & ".txt copied successfully to " & dstFile & " yet it isn't found there!"
	logMessage "Script aborted"
	wscript.quit
End If 
If err.number <> 0 Then LogError

Logmessage "Successfully copied " & srcFile & " to " & dstFile
LogMessage "Script completed sucessfully"