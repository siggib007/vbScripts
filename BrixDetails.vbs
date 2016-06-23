Option Explicit
Dim FileObj, strLine, fso, f, fc, f1, strOut, strParts, FolderSpec, strOutFullName
Dim objFileOut, strFileNameParts, x, cn, cmd, section, subsection

Const strFileNameCriteria = ".conf"
Const DBServerName = "satnetengfs01"
Const DBName = "Reports"
Const strOutFileName = "BrixDevices.csv"


Set cn      = CreateObject("ADODB.Connection")
Set cmd     = CreateObject("ADODB.Command")

cn.Provider = "sqloledb"
cn.Properties("Data Source").Value = DBServerName
cn.Properties("Initial Catalog").Value = DBName
cn.Properties("Integrated Security").Value = "SSPI"
'cn.Open
Cmd.ActiveConnection = cn

If WScript.Arguments.Count <> 1 Then 
  WScript.Echo "Usage: parser inpath"
  WScript.Quit
End If

FolderSpec = WScript.Arguments(0)
If Right(folderspec,1) = "\" Then folderspec = Left(folderspec,Len(folderspec)-1)
stroutfullname = folderspec & "\" & stroutfilename
wscript.echo Now & " Starting analyzing " & folderspec
wscript.echo "writing results to " & stroutfullname

Set fso = CreateObject("Scripting.FileSystemObject")
Set objfileout = fso.createtextfile(stroutfullname)
Set f = fso.GetFolder(folderspec)
Set fc = f.Files
For Each f1 in fc
	section = ""
	subsection = ""
	note = ""
	If f1.name <> strOutFileName AND InStr(f1.name,strFileNameCriteria) > 0 Then
		strFileNameParts = split(f1.name,".")
		Set FileObj = fso.opentextfile(folderspec & "\" & f1.name)
		While not fileobj.atendofstream
			strLine = Trim(FileObj.readline)
			If strline <> "" Then
				strparts = split(strline," ")
				If strline = strfilenameparts(0) & "#show run" Then section = "sh run"
				If strline = strfilenameparts(0) & "#show bench" Then section = "sh bench"
				If strparts(0) = Interface and strparts(1) = "Eth0" the subsection = "Eth0"
				If strparts(0) = Interface and strparts(1) = "Eth0" the subsection = "Eth1"
				If strparts(0) = "Device name:" Then devicename = strparts(1)
				If strparts(0) = "address:" and section = "sh run" Then
					If subsection = "Eth0" Then e0ip = strparts(1)
					If subsection = "Eth1" Then e1ip = strparts(1)
				End If 
				If strparts(0) = "default-gateway:" and section = "sh run" Then
					If subsection = "Eth0" Then e0dg = strparts(1)
					If subsection = "Eth1" Then e1dg = strparts(1)
				End If 
				If strparts(0) = "address:" and section = "sh bench" Then
					If subsection = "Eth0" and e0ip <> strparts(1) Then Note = Note & "E0 Bench IP wrong;"
					If subsection = "Eth1" and e1ip <> strparts(1) Then Note = Note & "E1 Bench IP wrong;"
				End If 
				If strparts(0) = "default-gateway:" and section = "sh bench" Then
					If subsection = "Eth0" and e0dg <> strparts(1) Then Note = Note & "E0 Bench DG wrong;"
					If subsection = "Eth1" e1dg <> strparts(1) Then Note = Note & "E1 Bench DG wrong;"
				End If 
				If strparts(0) = "IP Settings" Then subsection = "DNS"
				If strparts
				
			End If 
		Wend 
		If strOut <> "" Then 
			strparts = split(strout,",")
			Cmd.CommandText = "insert into Reports.dbo.ChassisTypes (DeviceName,ChassisType) values ('" & strparts(0) & "','" & strparts(1) &  "')"
			wscript.echo strout
			'wscript.echo cmd.commandtext
			'Cmd.Execute			
		End If 
		FileObj.close
		strOut = ""
	End If
Next
'cn.close

Set cmd = nothing
Set cn = nothing
Set FileObj = nothing
Set fc = nothing
Set f = nothing
Set fso = nothing

wscript.echo Now & " Analysis complete"

