Option Explicit

Dim FileObj, DeviceName, fso, TicketID, SrcFileName, FileSpec

If WScript.Arguments.Count <> 3 Then 
  WScript.Echo "Usage: dupfile.vbs DeviceList, TicketNum, sourcescript"
  WScript.Quit
End If

FileSpec = WScript.Arguments(0)
TicketID = WScript.Arguments(1)
SrcFileName = WScript.Arguments(2)

Set fso = CreateObject("Scripting.FileSystemObject")
Set FileObj = fso.opentextfile(filespec)

While not fileobj.atendofstream
	DeviceName = Trim(FileObj.readline)
	wscript.echo "copy /y " & srcFileName & " " & ticketID & "_" & devicename & ".txt"
	fso.copyfile srcFileName, ticketID & "_" & devicename & ".txt", true
Wend
wscript.echo "DONE!!!" 
FileObj.close

Set FileObj = nothing
Set fso = nothing
