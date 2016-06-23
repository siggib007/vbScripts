Option Explicit
Dim inFileObj, strLine, fso, strOut, strParts, inFileName, outFileObj, userOutFileName, bSendMail, Subject, i, DictKeys, strName, outstr, perms, path, tmpstr2
Dim ProfileNum, Profile, ProfilePermDict, ProfileNameDict, ProfileSettingsDict, groupOutFileName, UserDict, SettingsArray(2), UserArray(4), tmpstr, EmailBodyFileName

'Const MailServerName = "smtphost.redmond.corp.microsoft.com" 
Const MailServerName = "satsmtpa01"
Const SMTPTimeout = 10
Const FromAddress = "harrydup@microsoft.com"
Const ToAddress = "linnco@microsoft.com;netpro@microsoft.com;gnspm@microsoft.com"
Const CCAddress = ""
Const BCCAddress = "harrydup@microsoft.com "
Const cdoSendUsingPort = 2
Const cdoNTLM = 2

Subject = "TACACS Audit report for "  & FormatDateTime(Now,vbLongDate)

Set ProfilePermDict = CreateObject("Scripting.Dictionary")
Set ProfileNameDict = CreateObject("Scripting.Dictionary")
Set ProfileSettingsDict = CreateObject("Scripting.Dictionary")
Set UserDict = CreateObject("Scripting.Dictionary")
Set fso = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count <> 2 Then 
  WScript.Echo "Usage: TACACSReport inFileName, EmailBodyFileName"
  WScript.Quit
End If

inFileName = WScript.Arguments(0)
EmailBodyFileName = WScript.Arguments(1)

If fso.FileExists(inFileName) = False Then
	wscript.echo "File " & infilename & " does not exists."
	wscript.quit
End If

If fso.FileExists(EmailBodyFileName) = False Then
	wscript.echo "File " & EmailBodyFileName & " does not exists."
	wscript.quit
End If

strparts = split(infilename,"\")
For i =0 to UBound(strparts) - 1
	path = path & strparts(i) & "\"
Next 

Useroutfilename = path & "TACACSUsers.csv"
Groupoutfilename = path & "TACACSGroup.txt"

Set inFileObj = fso.opentextfile(inFileName)


While not infileobj.atendofstream
	strLine = Trim(inFileObj.readline)
	If strline = "##--- Values End" Then
		If Left(strname,10) = "###profile" Then
			ProfileSettingsDict.add profilenum, SettingsArray
		End If 
		If strname <> "" and Left(strname,3) <> "###" and strname <> "CSMonRad" and strname <> "CSMonTac" and strname <> "system" Then
			userdict.add strname,userarray
		End If 	
		strName = ""
		ProfileNum = ""
		SettingsArray(0) = ""
		SettingsArray(1) = ""
		SettingsArray(2) = ""
		userarray(0) = ""
		userarray(1) = ""
		userarray(2) = ""
		userarray(3) = ""
		userarray(4) = ""
	End If 
	If strline <> "" Then
		strparts = split(strline,vbtab)
		If UBound(strparts) > 2 Then
			If strparts(0) = "App01" and Left(strparts(1),7) = "PROFILE" Then
				ProfileNum = CInt(Right(strparts(1),Len(strparts(1))-7))
				Profile = strparts(3)
				ProfilePermDict.add ProfileNum, Profile
			End If 
			If strparts(0) = "App03" and Left(strparts(1),10) = "PROFMAP-T-" Then
				ProfileNum = CInt(Right(strparts(1),Len(strparts(1))-10))
				Profile = strparts(3)
				ProfileNameDict.add ProfileNum, Profile
			End If 
		End If 
		If UBound(strparts) > 0 Then 
			If strparts(0) = "Name          :" Then 
				strName = strparts(1)
				If Left(strparts(1),10) = "###profile" Then ProfileNum = CInt(Right(strparts(1),Len(strparts(1))-10))
			End If	
			If strparts(0) = "App01" and strparts(1) = "Filters\NAS\records" Then
				tmpstr = ""
				For i = 3 to UBound(strparts)
					tmpstr = tmpstr & strParts(i) & vbtab
				Next 
				SettingsArray(0) = Left(tmpstr,Len(tmpstr)-1)
			End If 
			If strparts(0) = "App01" and strparts(1) = "Filters\NAS\enabled" Then
				SettingsArray(1) = strparts(3)
			End If 
			If strparts(0) = "App01" and strparts(1) = "Filters\NAS\option" Then
				SettingsArray(2) = strparts(3)
			End If 
			If strname <> "" and Left(strname,3) <> "###" Then
				If strparts(0) = "Profile       :"  Then
					profilenum = CLng(strparts(1))
					userarray(0) = profilenum
				End If 
				If strparts(1) = "USER_DEFINED_FIELD_0" and strparts(0) = "App00" Then
					userarray(1) = strparts(3)
				End If 
				If strparts(1) = "USER_DEFINED_FIELD_1" and strparts(0) = "App00" Then
					userarray(2) = strparts(3)
				End If 
				If strparts(1) = "USER_DEFINED_FIELD_2" and strparts(0) = "App00" Then
					userarray(3) = strparts(3)
				End If 								
				If strparts(1) = "USER_DEFINED_FIELD_3" and strparts(0) = "App00" Then
					userarray(4) = strparts(3)
				End If 								
			End If 
		End If 		
	End If 
Wend 
inFileObj.close

'generate users report
Set outFileObj = fso.createtextfile(userOutFileName)
DictKeys = userdict.Keys
outstr = "username, group, real name, property, UTS, approver, HWGroup"
'wscript.echo outstr
outfileobj.writeline outstr
For i = 0 to userdict.count - 1 
	tmpstr = userdict.item(DictKeys(i))
	tmpstr2 = ProfileSettingsDict.item(tmpstr(0))
	outstr = DictKeys(i) & ", " & ProfileNameDict.item(tmpstr(0)) & ", " & tmpstr(1) & ", " & tmpstr(2) & ", " & tmpstr(3) & ", " & tmpstr(4) & ", " & tmpstr2(0)
	'wscript.echo outstr
	outfileobj.writeline outstr
Next
outFileObj.close


'generate users report
Set outFileObj = fso.createtextfile(groupOutFileName)
DictKeys = ProfileSettingsDict.Keys

For i = 0 to ProfileSettingsDict.count - 1 
	tmpstr = ProfileSettingsDict.item(DictKeys(i))
	If tmpstr(1) <> "" Then
		outstr = String(60,"-") & vbcrlf & "USER GROUP NAME: " & ProfileNameDict.item(DictKeys(i)) & vbcrlf 
		outstr = outstr & "Associated Hardware Group: " & tmpstr(0) & vbcrlf
		outstr = outstr & "Enable: " & tmpstr(1)  & vbcrlf
		outstr = outstr & "Option:" & tmpstr(2) & vbcrlf
		outstr = outstr & "Commands: "
		perms = ProfilePermDict.item(dictkeys(i))
		perms = replace(perms,"{", vbcrlf & vbtab)
		perms = replace(perms,"}","")
		perms = replace(perms,"default", vbcrlf & "default")
		perms = replace(perms,"cmd", vbcrlf & "cmd")
		perms = replace(perms,"  ", vbcrlf & vbtab & vbtab)
		While InStr(perms,vbcrlf&vbtab&vbcrlf) > 0 
		perms = replace(perms,vbcrlf&vbtab&vbcrlf,vbcrlf)
		Wend
		While InStr(perms,vbcrlf&vbtab&vbtab&vbcrlf) > 0 
		perms = replace(perms,vbcrlf&vbtab&vbtab&vbcrlf,vbcrlf)
		Wend 
		'outstr = outstr & Right(perms,Len(perms)-2)
		outstr = outstr & perms
		'wscript.echo outstr
		outfileobj.writeline outstr
	End If		
Next
ouTfileobj.writeline String(60,"-")
'wscript.echo String(60,"-")
outFileObj.close

Set inFileObj = fso.opentextfile(EmailBodyFileName)
tmpstr = infileobj.readall
inFileObj.close

Set inFileObj = nothing
Set outFileObj = nothing
Set fso = nothing

wscript.echo "DONE PARSING!!!"

'mysendmail subject,tmpstr
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
	      .CC	    = CCAddress
	      .bcc      = BCCAddress
	      .From     = FromAddress 
	      .Subject  = StrSubject 
	      .textbody = Msg
	      .AddAttachment(Useroutfilename)
	      .AddAttachment(Groupoutfilename)
	      .Send          
	End With
End Sub
