Option Explicit
Dim FileObj, strLine, fso, f, fc, f1, strOut, FolderSpec, bps, fileparts, device

Const MailServerName = "smtphost.redmond.corp.microsoft.com" 
Const SMTPTimeout = 10
Const FromAddress = """Siggi Bjarnason"" <siggib@microsoft.com>"
Const ToAddress =   """Siggi Bjarnason"" <siggib@microsoft.com>"
Const CCAddress = ""
Const Subject = "F5 BigIP's with PowerSupply problems" 
Const cdoSendUsingPort = 2
Const cdoNTLM = 2

If WScript.Arguments.Count <> 1 Then 
  WScript.Echo "Usage: parser inpath"
  WScript.Quit
End If

FolderSpec = WScript.Arguments(0)

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder(folderspec)
Set fc = f.Files
strout = ""
For Each f1 in fc
	bps = False
	If InStr(f1.name,"_F5_Platform_Audit.txt") > 0 Then
		'wscript.echo "working on " & f1.name
		fileparts = split(f1.name,"_")
		device = fileparts(0)
		Set FileObj = fso.opentextfile(folderspec & "\" & f1.name)
		While not fileobj.atendofstream
			strLine = Trim(FileObj.readline)
			If Left(strline,1) = "|" Then
				strline = Trim(Right(strline,Len(strline)-1))
			End If
			strline = replace(strline,"10","")
			If bps And strline <> "(1) active   (2) active" Then
				strout = strout & device & " failed:  " & strline & vbcrlf
				'wscript.echo device & " failed"
				'wscript.echo strline
			End If
			If strline = "POWER SUPPLY" Then 
				bps = True
				'wscript.echo "Found PS line"
			End If
		Wend 
		If Not bps Then 
				'wscript.echo "Didn't find powersupply for " & device
				strout = strout & "Didn't find powersupply for " & device & vbcrlf
		End If 
		FileObj.close
	End If
Next
wscript.echo strout
mysendmail subject,strout
wscript.echo "Mail sent"

Set FileObj = nothing
Set fc = nothing
Set f = nothing
Set fso = Nothing

Sub MySendMail(StrSubject,msg)
	Dim iMsg,iConf,Flds
	
	Set iMsg = CreateObject("CDO.Message") 
	Set iConf = CreateObject("CDO.Configuration") 
	Set Flds = iConf.Fields 
	
	With Flds 
	  .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort 
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MailServerName 
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = SMTPTimeout
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")= cdoNTLM
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