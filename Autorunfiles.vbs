Option Explicit

Dim FileObj, DeviceName, fso, TicketID, SrcFileName, FileSpec, strline, strparts, srcFileObj
Dim Masterscript, outscript, x, outFileObj, outpath, strvar, lastdevice, lastticket

If WScript.Arguments.Count <> 3 Then 
  WScript.Echo "Usage: autorunfiles.vbs DeviceList, sourcescript, outpath"
  WScript.Quit
End If

FileSpec = WScript.Arguments(0)
SrcFileName = WScript.Arguments(1)
outpath = WScript.Arguments(2)

If Right(outpath,1) <> "\" Then outpath = outpath & "\"
wscript.echo "outpath = " & outpath

Set fso = CreateObject("Scripting.FileSystemObject")
Set FileObj = fso.opentextfile(filespec)
Set srcFileObj = fso.opentextfile(srcfilename)

masterscript = srcfileobj.readall
lastdevice = ""
lastticket = ""

While not fileobj.atendofstream
	outscript = masterscript
	strline = Trim(FileObj.readline)
	strparts = split(strline,",")
	devicename = strparts(1)
	ticketid = strparts(0)
	For x = 2 to UBound(strparts)
		strvar = "%p" & x-1 & "%"
		outscript = replace(outscript,strvar,strparts(x),1,-1,1)
	Next
	If devicename <> lastdevice or ticketid <> lastticket Then
		If lastdevice <> "" Then 	outFileObj.close
		Set outFileObj = fso.createtextfile(outpath & ticketID & "_" & devicename & ".txt")
		wscript.echo "Created " & outpath & ticketID & "_" & devicename & ".txt"
		lastdevice = devicename
		lastticket = ticketid
	End If
	outFileObj.write outscript
	wscript.echo "Added to " & outpath & ticketID & "_" & devicename & ".txt"
Wend
wscript.echo "DONE!!!" 
FileObj.close
srcfileobj.close
outFileObj.close

Set srcfileobj = nothing
Set outfileobj = nothing
Set FileObj = nothing
Set fso = nothing
