Option Explicit
Dim FileObj, strLine, fso, f, fc, f1, strOut, strParts, FolderSpec, strOutFileName
Dim objFileOut, strFileNameParts, x, iLimit, strReport, cn, cmd

Const strFileNameCriteria = "_ShTemp.txt"
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
strreport = Now & " Starting analyzing " & folderspec
strreport = strreport & String(65,"-") & vbcrlf

'wscript.echo strreport
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder(folderspec)
Set fc = f.Files
For Each f1 in fc
	If f1.name <> strOutFileName AND InStr(f1.name,strFileNameCriteria) > 0 Then
		strFileNameParts = split(f1.name,"_")
		Set FileObj = fso.opentextfile(folderspec & "\" & f1.name)
		wscript.echo f1.name
		While not fileobj.atendofstream
			strLine = Trim(FileObj.readline)
			If strline <> "" Then
				strparts = split(strline,":")
				'wscript.echo "UBound(strparts): " & UBound(strparts)
				If UBound(strparts) = 1 Then
					Cmd.CommandText = "INSERT INTO [Reports].[dbo].[ModuleTemp] ([dtTimeStamp],[vcDeviceName],[vcDescription],[vcTemp]) values ('" & _ 
											Now & "','" & strFileNameParts(0) & "','" & strparts(0) & "','" & strparts(1) &  "')"
					'wscript.echo cmd.commandtext	
					Cmd.Execute		
					wscript.echo strFileNameParts(0) & ": " & strline					
				End If 
			End If 
		Wend 
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
wscript.echo Now & " Analysis complete"