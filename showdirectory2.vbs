Dim fso, f, f1, fc, strout, folderspec, strOutFilePath, tmpPath
strOutFilePath = "\\cpnettools2\d$\ihsopeng_root\Documents\TOC.htm"
folderspec = "\\cpnettools2\d$\ihsopeng_root\Documents"
strOut = "<html>" & vbcrlf
strout = strout & "<head>" & vbcrlf
strout = strout & "<A HREF=""http://ihs"">Home</A><BR>" & vbcrlf
strout = strout & "<title>Archived Files</title>" & vbcrlf
strout = strout & "</head>" & vbcrlf & vbcrlf
strout = strout & "<body>" & vbcrlf & vbcrlf
strout = strout & "<h1> Archived IDC Performance Reports </h1>" & vbcrlf
strout = strout & "<A HREF=""http://ihs/IDC/NetworkPerformanceReport.htm"">Current Report</A><BR>" & vbcrlf
strout = strout & "<h3><br></h3>" & vbcrlf
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder(folderspec)
Set fc = f.Files
For Each f1 in fc
	tmpPath = "http://ihs/idc/archive/" & f1.name
	wscript.echo tmpPath
	strout = strout & "<A HREF=""" & tmpPath & """>" & 	f1.name & "</A><BR>" & vbcrlf
Next
strout = strout & "</body>" & vbcrlf
strout = strout & "</html>" & vbcrlf
Set strOutFile = fso.CreateTextFile(strOutFilePath, True)
stroutfile.writeline strout
