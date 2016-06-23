Option Explicit
Dim fso, logObj

Const ForReading = 1
Const strLogFileName = "testlog.txt"
Const strTmpFileName = "Checklog.tmp"
Const strExcludePath = "exclude.txt"
Const strOutFileName = "wmicollectionOut.txt"
Const strINIFileName = "wmicollect.ini"
Const strJobName = "test"
Const bUseIntegrated = true
Const txtServer = "sgb"
Const JobStartTimeout = 120
Const strComplete = "Job Complete."
Const Subject = "WMI Collector is hung"
Const MsgBody = "Please fix"
Const SMTPTimeout = 10
Const FromAddress = "ghtools@microsoft.com"
Const ToAddress = "siggib@microsoft.com"
Const CCAddress = ""
Const cdoSendUsingPort = 2
Const MailServerName = "smarthost.dns.microsoft.com" 
Const ForAppending = 8 

Set fso = CreateObject("Scripting.FileSystemObject")
Set logObj = fso.opentextfile("checklog.log",ForAppending,true)

Main

Sub Main()
	Dim LogFileObj, strLine, Strlastfile, array1, array2,tmpFileObj, Success
	
	If fso.fileexists(strtmpfilename) Then
		Set tmpFileObj = fso.opentextfile(strtmpfilename)
		If not tmpfileobj.atendofstream Then
			strlastFile = tmpfileobj.readline
		Else
			strlastfile = "blank" & vbtab & "file"
		End If 
		tmpfileobj.close
	Else
		strlastfile = "blank" & vbtab & "file"
	End If
	'writelog strlastfile
	array2 = split(strlastfile,vbtab)
	If Trim(array2(1)) = strComplete Then
		writelog "Done"
		wscript.quit
	End If
	
	If fso.FileExists(strlogfilename) Then
		Set LogFileObj = fso.OpenTextFile(strlogFileName, ForReading)
	Else
		writelog "Logfile " & strlogfilename & " not found"
		wscript.quit
	End If 
	
	While not LogFileObj.atendofstream
		strLine = LogFileObj.readline
	Wend
	
	array1=split(strline,vbtab)
	
	If (Trim(array1(0)) = Trim(array2(0))) and (Trim(array1(1)) <> strComplete) Then
		writelog "Collector hung, attempting to restart..."
		writelog strline
		UpdateFiles(Trim(array1(1)))
		Success = RestartJob()
		Select Case Success
			Case 0
				writelog "Successfully restarted Collector. Sending notification mail."
				mysendmail "Collector hung, automatically restarted","Hung while attempting:" & vbcrlf & strline
				Set tmpfileobj=fso.createtextfile(strtmpfilename,true) 
				tmpfileobj.close
			Case 1
				writelog "Collector is hung, unable to restart it. Hung while attempting to stop."
				mysendmail subject,"Unable to automatically restart it. Job hung while attempting to stop" & vbcrlf & strline
				Set tmpfileobj=fso.createtextfile(strtmpfilename,true) 
				tmpfileobj.writeline "Collector Hung" & vbtab & strComplete
				tmpfileobj.close
			Case 2
				writelog "Collector is hung, unable to restart it. Didn't respond to start command."
				mysendmail subject,"Unable to automatically restart it. Didn't respond to start command." & vbcrlf & strline
				Set tmpfileobj=fso.createtextfile(strtmpfilename,true) 
				tmpfileobj.writeline "Collector Hung" & vbtab & strComplete
				tmpfileobj.close
			Case 3
				writelog "Collector is hung, unable to restart it. Hung while attempting to start."
				mysendmail subject,"Unable to automatically restart it. Hung while attempting to start." & vbcrlf & strline
				Set tmpfileobj=fso.createtextfile(strtmpfilename,true) 
				tmpfileobj.writeline "Collector Hung" & vbtab & strComplete
				tmpfileobj.close
			Case 4
				writelog "Collector is hung, unable to restart it. Failed to start."
				mysendmail subject,"Unable to automatically restart it. Failed to start." & vbcrlf & strline
				Set tmpfileobj=fso.createtextfile(strtmpfilename,true) 
				tmpfileobj.writeline "Collector Hung" & vbtab & strComplete
				tmpfileobj.close
			Case 5
				writelog "Collector is hung, unable to restart it. SQL Agent stopped and was unable to restart it."
				mysendmail subject,"Unable to automatically restart it. SQL Agent stopped and was unable to restart it." & vbcrlf & strline
				Set tmpfileobj=fso.createtextfile(strtmpfilename,true) 
				tmpfileobj.writeline "Collector Hung" & vbtab & strComplete
				tmpfileobj.close
		End Select 
		
		writelog "Mail sent"
	Else
		writelog "Everything seems OK"
		Set tmpfileobj=fso.createtextfile(strtmpfilename,true) 
		tmpfileobj.writeline strline
		tmpfileobj.close
	End If 
End Sub 

Sub UpdateFiles(logmsg)
	Dim arr1, SrvName, ExcludeObj, OutFileObj, strLine, TargetId, INIFileObj, tmpfileobj
	
	If fso.FileExists(strOutFileName) Then
		Set OutFileObj = fso.OpenTextFile(strOutFileName, ForReading)
	Else
		writelog "Logfile " & strOutfilename & " not found"
		wscript.quit
	End If 
	
	writelog "Looking for current targetID"
	While not OutFileObj.atendofstream
		strLine = OutFileObj.readline
	Wend
	arr1=split(logmsg," ")
	SrvName = Left(arr1(2),Len(arr1(2))-3)

	arr1=split(strLine, vbtab)
	TargetID = Trim(arr1(1))
	writelog "Found it targetid of " & TargetID

	'wscript.echo "ServerName=" & srvname
	
	writelog "Updating exlude file..."
	Set ExcludeObj=fso.OpenTextFile(strExcludePath,ForAppending,true)
	ExcludeObj.writeline srvname
	
	If fso.FileExists(strINIFileName) Then
		Set INIFileObj = fso.OpenTextFile(strINIFileName, ForReading)
		Set tmpfileobj = fso.createtextfile(strINIFileName & ".tmp",true) 
	Else
		writelog "Logfile " & strOutfilename & " not found"
		wscript.quit
	End If 
	
	writelog "Updating ini file..."
	While not INIFileObj.atendofstream
		strLine = INIFileObj.readline
		If LCase(Left(strline,5)) <> "start" Then tmpFileObj.writeline strline
	Wend
	tmpFileObj.writeline "startat = " & TargetID + 2
	fso.copyfile strINIFileName & ".tmp", strINIFileName, true
End Sub
  
Sub MySendMail(StrSubject,msg)
	Dim iMsg,iConf,Flds 
	Exit Sub
	Set iMsg = CreateObject("CDO.Message") 
	Set iConf = CreateObject("CDO.Configuration") 
	Set Flds = iConf.Fields 
	
	With Flds 
	  .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort 
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MailServerName 
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = SMTPTimeout
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

Sub writelog (msg)
	
	wscript.echo msg
	logObj.writeline Now & vbtab & msg

End Sub

Function RestartJob()
Dim Job, oSQLServer, sUsername, sPasswd, Status, sec1, dtStart
	
	'wscript.echo "bUseIntegrated=" & bUseIntegrated
	'wscript.echo "txtServer=" & txtServer
	
	Set oSQLServer = CreateObject("SQLDMO.SQLServer")
	
	oSQLServer.LoginTimeout = -1 '-1 is the ODBC default (60) seconds
	'Connect to the Server
	If bUseIntegrated Then
	  With oSQLServer
	    .LoginSecure = True 'Use NT Authentication
	    .AutoReConnect = False
	    .Connect txtServer
	  End With
	Else
	  With oSQLServer
	    .LoginSecure = False 'Use SQL Server Authentication
	    .AutoReConnect = False
	    .Connect txtServer, sUsername, sPasswd
	  End With
	End If
	
	If oSQLServer.JobServer.Status <> 1 Then
		writelog "SQLAgent not running, attempting to start"
		oSQLServer.JobServer.start
		oSQLServer.JobServer.Jobs.refresh(false)
		dtstart = Now 
		Do While oSQLServer.JobServer.Status <> 1
			sec1=Second(Now)
			While sec1=Second(Now)
			Wend
			wscript.echo "Agent Status=" & oSQLServer.JobServer.Status
			oSQLServer.JobServer.Jobs.refresh(false)
			If DateAdd("s",JobStartTimeout,dtStart) <=Now Then 
				writelog "SQLAgent didn't start in time allowed"
				oSQLServer.DisConnect
				RestartJob = 5
				Exit Function
			End If 
		Loop
	End If 
	wscript.echo "SQL Agent Started"
	With oSQLServer.JobServer.Jobs(strJobName)		
		'wscript.echo .name & vbtab & .LastRunOutcome & vbtab & .LastRunDate
		If .CurrentRunStatus <> 4 Then
			writelog "Job currently running, attempting to stopp job."
			.stop()
			oSQLServer.JobServer.Jobs.refresh(false)
			dtStart = Now
			Do While .CurrentRunStatus <> 4
				sec1=Second(Now)
				While sec1=Second(Now)
				Wend
				wscript.echo "CurrentRunStatus=" & .CurrentRunStatus
				oSQLServer.JobServer.Jobs.refresh(false)
				
				'wscript.echo "JobStartTimeout=" & JobStartTimeout & "; dtStart=" & dtStart
				'wscript.echo "JobStartTimeout + dtStart = " & DateAdd("s",JobStartTimeout,dtStart)
				'wscript.echo "Now=" & Now
				
				If DateAdd("s",JobStartTimeout,dtStart) <=Now Then 
					writelog "Job didn't stop in time allowed"
					oSQLServer.DisConnect
					RestartJob = 1
					Exit Function
				End If 
			Loop
			writelog "Now restarting job"
			.start
		Else
			writelog "Collector not running, attempting to start the collector"
			.start
		End If 

		oSQLServer.JobServer.Jobs.refresh(false)				
		
		'sec1=Second(Now)
		'While sec1<Second(Now)+5
		'Wend
		
		dtStart = Now
		Do While .CurrentRunStatus = 4
			oSQLServer.JobServer.Jobs.refresh(false)				
			sec1=Second(Now)
			While sec1=Second(Now)
			Wend
			wscript.echo "CurrentRunStatus=" & .CurrentRunStatus
			oSQLServer.JobServer.Jobs.refresh(false)
			
			'wscript.echo "JobStartTimeout=" & JobStartTimeout & "; dtStart=" & dtStart
			'wscript.echo "JobStartTimeout + dtStart = " & DateAdd("s",JobStartTimeout,dtStart)
			'wscript.echo "Now=" & Now
			
			If DateAdd("s",JobStartTimeout,dtStart) <=Now Then 
				writelog "Job didn't start in time allowed"
				oSQLServer.DisConnect
				RestartJob = 2
				Exit Function
			End If 	
		Loop 
		oSQLServer.JobServer.Jobs.refresh(false)	
		
		dtStart = Now			
		Do While .CurrentRunStatus <> 1
			oSQLServer.JobServer.Jobs.refresh(false)				
			sec1=Second(Now)
			While sec1=Second(Now)
			Wend
			wscript.echo "CurrentRunStatus=" & .CurrentRunStatus
			oSQLServer.JobServer.Jobs.refresh(false)	
			If .CurrentRunStatus = 4 Then 
				wscript.echo "Job stopped"
				Exit Do 
			End If
			If DateAdd("s",JobStartTimeout,dtStart) <=Now Then 
				writelog "Job didn't start in time allowed"
				oSQLServer.DisConnect
				RestartJob = 3
				Exit Function
			End If 	
		Loop 
		oSQLServer.JobServer.Jobs.refresh(false)	

		If .CurrentRunStatus = 1 Then
			 writelog "Job started Successfully"
			 RestartJob = 0
		Else 
			writelog "Job Failed to start"
			RestartJob = 4
		End If 
	End With 
	
	If Not oSQLServer Is Nothing Then
		'When done with the connection to SQLServer you must Disconnect
	    oSQLServer.DisConnect
	End If
  
End Function