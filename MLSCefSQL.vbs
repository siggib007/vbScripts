Option Explicit
Dim FileObj, strLine, fso, f, fc, f1, strOut, strParts, FolderSpec
Dim strFileNameParts, strReport, cn, cmd, maxTCAM, currTCAM

Const MailServerName = "smtphost.redmond.corp.microsoft.com" 
'Const MailServerName = "Tk2smtp2.phx.gbl" 
Const SMTPTimeout = 10
Const FromAddress = """Siggi Bjarnason"" <siggib@microsoft.com>"
Const ToAddress = """Internal Netpro"" <inetpro@microsoft.com>"
Const CCAddress = ""
Const Subject = "IPC Queue report" 
Const cdoSendUsingPort = 2
Const cdoNTLM = 2

Const strFileNameCriteria = "_sh_mls_cef.txt"
Const DBServerName = "satnetengfs01"
Const DBName = "Reports"

Set cn      = CreateObject("ADODB.Connection")
Set cmd     = CreateObject("ADODB.Command")

cn.Provider = "sqloledb"
cn.Properties("Data Source").Value = DBServerName
cn.Properties("Initial Catalog").Value = DBName
'cn.Properties("User ID").Value = UserName
'cn.Properties("Password").Value = Password
cn.Properties("Integrated Security").Value = "SSPI"
cn.Open
Cmd.ActiveConnection = cn

If WScript.Arguments.Count <> 1 Then 
  WScript.Echo "Usage: parser inpath"
  WScript.Quit
End If

FolderSpec = WScript.Arguments(0)
strreport = Now & " Starting analyzing " & folderspec & vbcrlf
strreport = strreport &"DeviceName,maxTCAM,CurrentTCAM,%Used" & vbcrlf
strreport = strreport & String(65,"-") & vbcrlf
wscript.echo strreport
'wscript.echo strreport
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder(folderspec)
Set fc = f.Files
For Each f1 in fc
	If InStr(f1.name,strFileNameCriteria) > 0 Then
		strFileNameParts = split(f1.name,"_")
		Set FileObj = fso.opentextfile(folderspec & "\" & f1.name)
		maxTCAM = 0 
		currTCAM = 0 
		'wscript.echo "processing " & strFileNameParts(0)
		While not fileobj.atendofstream
			strLine = Trim(FileObj.readline)
			If strline <> "" Then
				strparts = split(strline," - ")
				If UBound(strparts) > 0 Then
					If Left(strparts(0),4) = "IPv4" Then 
						maxTCAM = Left(Trim(strparts(1)),3) * 1000
					End If 
				End If 
				If Left (strline,12) = "Total routes" Then
					strparts = split(strline,":")
					currTCAM = Trim(strparts(1))
				End If 				
			End If 
		Wend 
		strOut = strFileNameParts(0) & "," & maxTCAM & "," & currTCAM & "," & FormatPercent(currTCAM / maxTCAM)
		wscript.echo strOut	
		Cmd.CommandText = "insert into Reports.dbo.TCAMReport (DeviceName, MaxTCAM, CurrentTCAM, dttimestamp) values ('" & strFileNameParts(0) & "'," & maxTCAM & ",'" & currTCAM & "','" & Now & "')"
		Cmd.Execute	
		strreport = strreport & strout & vbcrlf
		FileObj.close
		strOut = ""
	End If
Next
cn.close

Set cmd = nothing
Set cn = nothing
Set FileObj = nothing
Set fc = nothing
Set f = nothing
Set fso = nothing

'wscript.echo strreport
'wscript.echo Now & " Analysis complete"

'mysendmail subject,strreport
'wscript.echo "Mail sent"


Sub MySendMail(StrSubject,msg)
	Dim iMsg,iConf,Flds
	
	Set iMsg = CreateObject("CDO.Message") 
	Set iConf = CreateObject("CDO.Configuration") 
	Set Flds = iConf.Fields 
	
	With Flds 
	  .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort 
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MailServerName 
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = SMTPTimeout
	  .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")= cdoNTLM
	  .Update 
	End With 
	
	With iMsg 
	  Set .Configuration = iConf 
	      .To       = ToAddress
	      .CC	= CCAddress
	      .From     = FromAddress 
	      .Subject  = StrSubject 
	      .textbody = Msg
	      .Send 
	End With
End Sub
