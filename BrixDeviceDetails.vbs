Option Explicit

Dim FileObj, strLine, fso, f, fc, f1, strParts, FolderSpec, strOutFullName, strtemp
Dim objFileOut, strFileNameParts, x, cn, cmd, section, subsection, strout, errnote
Dim devicename, e0ip, e0dg, e0Label, e0State, e0Mac, e1ip, e1dg, e1Label, e1State, e1Mac
Dim Domainname, DNS, NTPState, ntpservers, udisc, ndisc, ldisc, ssh, telnet, snmp, LogLevel, DefUser, DefPWD

Const strFileNameCriteria = ".conf"
Const DBServerName = "satnetengfs01"
Const DBName = "Reports"
Const strOutFileName = "Verifiers.csv"


Set cn      = CreateObject("ADODB.Connection")
Set cmd     = CreateObject("ADODB.Command")

cn.Provider = "sqloledb"
cn.Properties("Data Source").Value = DBServerName
cn.Properties("Initial Catalog").Value = DBName
cn.Properties("Integrated Security").Value = "SSPI"
'cn.Open
'Cmd.ActiveConnection = cn

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
objfileout.writeline "devicename, e0ip, e0dg, e0Label, e0State, e0Mac, e1ip, e1dg, e1Label, e1State, e1Mac, Domainname, DNS, NTPState, ntpservers, udisc, ndisc, ldisc, ssh, telnet, snmp, LogLevel, DefUser, DefPWD, errnote"
Set f = fso.GetFolder(folderspec)
Set fc = f.Files
For Each f1 in fc
	section = ""
	subsection = ""
	errnote = ""
	devicename = ""
	e0ip = ""
	e0dg = ""
	e0Label = ""
	e0State = ""
	e0Mac = ""
	e1ip = ""
	e1dg = ""
	e1Label = ""
	e1State = ""
	e1Mac = ""
	Domainname = ""
	DNS = ""
	NTPState = ""
	ntpservers = ""
	udisc = ""
	ndisc = ""
	ldisc = ""
	ssh = ""
	telnet = ""
	snmp = ""
	LogLevel = ""
	DefUser = ""
	DefPWD = ""
	strout = ""
	If f1.name <> strOutFileName AND InStr(f1.name,strFileNameCriteria) > 0 Then
		strFileNameParts = split(f1.name,".")
		Set FileObj = fso.opentextfile(folderspec & "\" & f1.name)
		wscript.echo "analyzing " & f1.name
		While not fileobj.atendofstream
			strLine = Trim(FileObj.readline)
			If strline <> "" Then
				strparts = split(strline," ")
				If strline = strfilenameparts(0) & "#show run" Then section = "sh run"
				If strline = strfilenameparts(0) & "#show bench" Then section = "sh bench"
				If strline = strfilenameparts(0) & "#sh int" Then section = "Int"
				'wscript.echo strfilenameparts(0) & "," & section & "," & subsection & "," & strparts(0) & "*"' & strparts(1) & ":"
				Select Case strparts(0)
				Case "interface"
					If UBound(strparts) > 0 Then 
						'wscript.echo strparts(1)
						'wscript.echo strparts(15)
						If strparts(1) = "Eth0" Then subsection = "Eth0"
						If strparts(1) = "Eth1" Then subsection = "Eth1"
					End If 
				Case "Device"
					If UBound(strparts) > 1 Then
						If section = "sh run" and strparts(1) = "name:" Then devicename = strparts(2)
						If section = "sh bench" and strparts(1) = "name:" and devicename <> strparts(2) Then errnote = errnote & "name; "
					Else 
						If section = "sh bench" and devicename <> "" Then errnote = errnote & "name; "
					End If 
				Case "address:" 
					If UBound(strparts) > 0 Then
						If section = "sh run" Then
							If subsection = "Eth0" Then e0ip = strparts(1)
							If subsection = "Eth1" Then e1ip = strparts(1)
						End If 
						If section = "sh bench" Then
							If subsection = "Eth0" and e0ip <> strparts(1) Then errnote = errnote & "E0 IP; "
							If subsection = "Eth1" and e1ip <> strparts(1) Then errnote = errnote & "E1 IP; "
						End If 
					End If
				Case "default-gateway:" 
					If UBound(strparts) > 0 Then 
						If section = "sh run" Then
							If subsection = "Eth0" Then e0dg = strparts(1)
							If subsection = "Eth1" Then e1dg = strparts(1)
						End If 
						If section = "sh bench" Then
							If subsection = "Eth0" and e0dg <> strparts(1) Then errnote = errnote & "E0 DG; "
							If subsection = "Eth1" and e1dg <> strparts(1) Then errnote = errnote & "E1 DG; "
						End If 
					End If 
				Case "IP" 
					If strparts(1) = "Settings" Then subsection = "DNS"
				Case "domain-name:"
					strtemp = ""
					For x = 1 to UBound(strparts)
						strtemp = strtemp & strparts(x) & " " 
					Next 
					strtemp = Trim(strtemp)
					If section = "sh run" Then Domainname = strtemp
					If section = "sh bench" and domainname <> strtemp Then errnote = errnote & "domain; "
				Case "name-server:"
					strtemp = ""
					For x = 1 to UBound(strparts)
						strtemp = strtemp & strparts(x) & " " 
					Next 
					strtemp = Trim(strtemp)
					If section = "sh run" Then DNS = strtemp
					If section = "sh bench" and DNS <> strtemp Then errnote = errnote & "DNS; "
				Case "NTP" 
					If strparts(1) = "Information" Then subsection = "NTP" 
				Case "ntp:"
					If section = "sh run" Then NTPState = strparts(1) 
					If section = "sh bench" and NTPState <> strparts(1) Then errnote = errnote & "NTP; "	
				Case "ntp" 
					strtemp =""
					If strparts(1) = "servers:" Then 
						For x = 2 to UBound(strparts)
							strtemp = strtemp & strparts(x) & " " 
						Next 
						strtemp = Trim(strtemp)
						If section = "sh run" Then ntpservers = strtemp
						If section = "sh bench" and ntpservers <> strtemp Then errnote = errnote & "NTP Server; "
					End If
				Case "Server"
					'wscript.echo "str1: " & strparts(1)
					If strparts(1) = "Discovery" Then subsection = "disc"
				Case "Universal"
					strtemp =""
					If subsection = "disc" Then 
						For x = 3 to UBound(strparts)
							strtemp = strtemp & strparts(x) & " " 
						Next 
						strtemp = Trim(strtemp)
						If section = "sh run" Then udisc = strtemp
						If section = "sh bench" and udisc <> strtemp Then errnote = errnote & "udisc; "
					End If
				Case "Network"
					strtemp =""
					If subsection = "disc" Then 
						For x = 3 to UBound(strparts)
							strtemp = strtemp & strparts(x) & " " 
						Next 
						strtemp = Trim(strtemp)
						If section = "sh run" Then ndisc = strtemp
						If section = "sh bench" and ndisc <> strtemp Then errnote = errnote & "ndisc; "
					End If
				Case "Local"
					strtemp =""
					If subsection = "disc" Then 
						For x = 3 to UBound(strparts)
							strtemp = strtemp & strparts(x) & " " 
						Next 
						strtemp = Trim(strtemp)
						If section = "sh run" Then ldisc = strtemp
						If section = "sh bench" and ldisc <> strtemp Then errnote = errnote & "ldisc; "
					End If
				Case "Terminal"
					'wscript.echo "s1: " & strparts(1)
					If strparts(1) = "Settings" Then subsection = "term"
				Case "ssh:", "ssh"
					ssh = strparts(UBound(strparts))
				Case "telnet:", "telnet"
					telnet = strparts(UBound(strparts))
				Case "snmp:"
					snmp = strparts(1)
				Case "System"
					If strparts(1) = "Logging" Then LogLevel = strparts(3)
				Case "Default"
					If UBound(strparts) > 2 Then 
						If strparts(2) = "username:" Then DefUser = strparts(3)
						If strparts(2) = "password:" Then DefPWD = strparts(3)
					End If 
				Case "eth0"
					If section = "Int" Then
						e0Label = Right(strparts(1),Len(strparts(1))-1)	
						e0State = strparts(5)
						subsection = "e0"
					End If
				Case "eth1"
					If section = "Int" Then
						e1Label = Right(strparts(1),Len(strparts(1))-1)	
						e1State = strparts(5)
						subsection = "e1"
					End If
				Case "mac"
					If section = "Int" and subsection = "e0" Then e0Mac = strparts(2)
					If section = "Int" and subsection = "e1" Then e1Mac = strparts(2)
				End Select 
			End If 
		Wend 
		FileObj.close
		strout = devicename & "," & e0ip & "," & e0dg & "," & e0Label & "," & e0State & "," & e0Mac _
				& "," & e1ip & "," & e1dg & "," & e1Label & "," & e1State & "," & e1Mac & "," _
				& Domainname & "," & DNS & "," & NTPState & "," & ntpservers & "," _ 
				& udisc & "," & ndisc & "," & ldisc & "," & ssh & "," & telnet & "," & snmp & "," _ 
				& LogLevel & "," & DefUser & "," & DefPWD & "," & errnote
		'wscript.echo strout
		objfileout.writeline strout	
	End If
Next
'cn.close
objfileout.close
Set objfileout = nothing
Set cmd = nothing
Set cn = nothing
Set FileObj = nothing
Set fc = nothing
Set f = nothing
Set fso = nothing

wscript.echo Now & " Analysis complete"

