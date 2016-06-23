Option Explicit
Dim cn, rs,rsIDC, fso, OutputFileObject, dictCluster, dictDC, IDCEq, cmdText
Dim InputFileObject, INIFileObj, OutPutFile, DBServer, DefaultDB, ClusterEQ, OrderBy
Dim strInput, InputArray, strINIPath, ScriptPathStr, strINILine, INILineArray, LogFileObj
Dim EgressLabel, ClusterMIB, EgressMIB, ClusterExclude, UserID, PWD, ScriptName, strLogFilePath

Const ForReading = 1	' ADO Constant
Const adUseClient = 3	' ADO Constant
Const AdOpenKeyset= 1	' ADO Constant
Const adLockOptimistic = 3	' ADO Constant
Const FILEDEL = " = " 'How external files are deliminated
Const FUNCDESCDEL = ";" 'How the funcitional description is delimated

'Initializing global variables
UserID = ""
PWD = ""
OrderBy = ""
ScriptPathStr = Left (wscript.scriptFullname,InStr(wscript.scriptFullName,wscript.scriptname)-1)
Scriptname = Left (wscript.scriptname, Len(wscript.scriptname)-4)
strINIPath = ScriptPathStr & Scriptname & ".ini"
strLogFilePath = ScriptPathStr & Scriptname & ".log"

Set fso = CreateObject("Scripting.FileSystemObject")
Set LogFileObj = fso.CreateTextFile(strLogFilePath, True)
Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")
Set rsIDC = CreateObject("ADODB.Recordset")
Set dictCluster = CreateObject("Scripting.Dictionary")
Set dictDC = CreateObject("Scripting.Dictionary")

Sub LogMessage (strMessage)
	wscript.echo Now & vbtab & strMessage
	LogFileObj.writeline Now & vbtab & strMessage
End Sub

Sub LoadINI()
	If not fso.fileExists(strINIPath) Then
		LogMessage "Couldn't find required INI file " & strINIPath
		LogMessage "script full path name " &  wscript.scriptFullname
		LogMessage "script path only " & scriptpathstr
		wscript.quit
	Else
		LogMessage "Found the ini file " & strINIPath &", reading it in."
	End If
	
	Set INIFileObj = fso.opentextfile(strINIPath)
	
	Do While INIFileObj.AtEndOfStream = False
	    strINILine = INIFileObj.readline
	    INILineArray = Split(strINILine, FILEDEL)
	    Select Case Trim(INILineArray(0))
	        Case "OUTPUTFILE"
	            LogMessage "found Outputfilename: " & INILineArray(1)
	            OutPutFile = Trim(INILineArray(1))
	        Case "DBSERVER"
	            LogMessage "found DB Server name: " & INILineArray(1)
	            DBServer = Trim(INILineArray(1))
	        Case "DEFAULTDB"
	            LogMessage "found Default Database name: " & INILineArray(1)
	            DefaultDB = Trim(INILineArray(1))
	        Case "CLUSTEREQ"
	            LogMessage "found Cluster equivilance file name: " & INILineArray(1)
	            ClusterEQ = Trim(INILineArray(1))
	        Case "IDCEQ"
	            LogMessage "found IDC equivilance file name: " & INILineArray(1)
	            IDCEq = Trim(INILineArray(1))
	        Case "cmdText"
	            LogMessage "found SQL query text: " & INILineArray(1)
	            cmdText = Trim(INILineArray(1))
	        Case "EgressLabel"
	            LogMessage "found Egress Label: " & INILineArray(1)
	            EgressLabel = Trim(INILineArray(1))
	        Case "ClusterMIB"
	            LogMessage "found Cluster MIB: " & INILineArray(1)
	            ClusterMIB = Trim(INILineArray(1))
	        Case "EgressMIB"
	            LogMessage "found Egress MIB: " & INILineArray(1)
	            EgressMIB = Trim(INILineArray(1))
	        Case "ClusterExclude"
	            LogMessage "found Cluster Exclude label: " & INILineArray(1)
	            ClusterExclude = Trim(INILineArray(1))
	        Case "UserID"
	            LogMessage "found UserID: " & INILineArray(1)
	            UserID = Trim(INILineArray(1))
	       Case "PWD"
	            LogMessage "found Password: **********"
	            PWD = Trim(INILineArray(1))
	       Case "OrderBy"
	            LogMessage "found OrderBy string: " & INILineArray(1)
	            OrderBy = Trim(INILineArray(1))
	        Case Else
	            LogMessage "Don't know what to do with *" & strINILine & "*"            
	    End Select
	loop
End Sub 'Load INI


Sub LoadEqFiles()
	If fso.FileExists(ClusterEQ) Then
		LogMessage "Cluster equivilance file exists, loading it memory."
		Set InputFileObject = fso.OpenTextFile(ClusterEQ, ForReading)
		Do While InputFileObject.AtEndOfStream = False
			strInput = InputFileObject.ReadLine
			InputArray = split (strInput, FILEDEL)
			If not dictCluster.exists(InputArray(0)) Then
				dictCluster.add InputArray(0), InputArray(1)
			Else
				LogMessage "Duplicate Cluster equivlance: " & InputArray(0)
			End If
		loop
		InputFileObject.close
		Set InputFileObject = nothing	
	Else
		LogMessage "Cluster Equivilance file not found, no cluster substitutions will be made"
	End If
	
	If fso.FileExists(IDCEq) Then
		LogMessage "IDC equivilance file exists, loading it memory."
		Set InputFileObject = fso.OpenTextFile(IDCEq, ForReading)
		Do While InputFileObject.AtEndOfStream = False
			strInput = InputFileObject.ReadLine
			InputArray = split (strInput, FILEDEL)
			If not dictDC.exists(InputArray(0)) Then
				dictDC.add InputArray(0), InputArray(1)
			Else
				LogMessage "Duplicate IDC equivlance: " & InputArray(0)
			End If
		loop
		InputFileObject.close
		Set InputFileObject = nothing
	Else
		LogMessage "IDC Equivilance file not found, no IDC substitutions will be made"
	End If
End Sub 'LoadEQFiles

Function FindClusterEquivilance (strCluster)
	If dictCluster.exists(strCluster) Then
		FindClusterEquivilance = dictCluster.Item(strCluster)
	Else
		FindClusterEquivilance = strCluster
	End If
End Function

Function FindIDCEquivilance ()
Dim strIDC, strLastIDC

	strLastIDC=""
	While not rs.eof
		strIDC = Trim(rs.fields(0))
		If dictDC.exists(strIDC) Then
			rs.fields(0) = dictDC.Item(strIDC)
			rs.update
			strIDC = Trim(rs.fields(0))
			'rs.update
		End If
		If strIDC <> strLastIDC Then
			strLastIDC = strIDC
			LogMessage "Looking for equivilance for *" & strIDC & "*"
			LogMessage "Current IDC is " & strIDC
		End If 
		rs.movenext
	Wend
	rs.sort = OrderBy
	rs.movefirst
End Function


Function main()
Dim strCluster, strDeviceName, strMIB
Dim strOutput, strLastCluster, strSQL
Dim FuncStrArray, ClusterArray, DeviceNameArray
	
	LoadINI
	LoadEqFiles
	
	Set OutputFileObject = fso.CreateTextFile(OutPutFile, True)
	If OrderBy <> "" Then
		strSQL = cmdText & " order by " & OrderBy
	Else
		strSQL = cmdText
	End If
	
        'Set ADO connection properties.
        LogMessage "Setting up ADO properties"
        cn.Provider = "sqloledb"
        cn.Properties("Data Source").Value = DBServer
        cn.Properties("Initial Catalog").Value = DefaultDB
        If UserID = "" Then
        	cn.Properties("Integrated Security").Value = "SSPI"
        Else
        	cn.Properties("User ID").Value = UserID
		cn.Properties("Password").Value = PWD
	End If

        LogMessage "Opening ADO connection"
        cn.Open
	
	LogMessage "Opening recordset"
	rs.cursorlocation = adUseClient
	rs.CursorType = AdOpenKeyset
	rs.LockType = adLockOptimistic
        rs.Open strSQL, cn
        LogMessage "Recordset opended, Disconnecting recordset "
	rs.ActiveConnection=nothing
	cn.close
	LogMessage "Looking for IDC Equivilance"
        FindIDCEquivilance
	LogMessage "Done Processing IDC Equivilances, now starting to process clusters"
        strLastCluster=""
        strOutput = ""
        While not rs.eof
        	DeviceNameArray = split(Trim(rs.fields(1)),".")
        	strDeviceName = DeviceNameArray(0)
        	FuncStrArray=split(Trim(rs.fields(3)),FUNCDESCDEL)
        	Select Case FuncStrArray(0)
        		Case "cluster"
        			ClusterArray=split(FuncStrArray(1))
        			strCluster=FindClusterEquivilance(ClusterArray(0))
        			LogMessage Trim(rs.fields(0)) & " " & ClusterArray(0) & " displayed as " & strcluster & ". Device,Int: " & strDeviceName & "," & Trim(rs.fields(2))
        			strMIB=ClusterMIB
        		Case "egress"
        			strCluster=EgressLabel
        			LogMessage Trim(rs.fields(0)) & " Processing Egress on Device,Int: " & strDeviceName & "," & Trim(rs.fields(2))
        			strMIB=EgressMIB
        		Case Else
        			LogMessage Trim(rs.fields(0)) & " Unknown " & FuncStrArray(0) & ";" & FuncStrArray(1)
        			strCluster=ClusterExclude
        	End Select
        	If strCluster <> ClusterExclude Then
	        	If strLastCluster=strCluster Then
	        		strOutput=strOutPut & vbtab & strDeviceName & "," & Trim(rs.fields(2)) & "," & strMIB
	        	Else
	        		If strOutput <> "" Then	OutputFileObject.writeline strOutput
	        		strOutput=Trim(rs.fields(0)) & vbtab & strCluster & vbtab & strDeviceName & "," & Trim(rs.fields(2)) & "," & strMIB
	        		strLastCluster=strCluster
	        	End If
	        Else
	        	LogMessage "Excluding cluster " & ClusterArray(0)
	        End If
        	rs.movenext
        Wend        
        If strOutput <> "" Then	OutputFileObject.writeline strOutput
	rs.close
end function

main
LogMessage "Done Processing cleaning up."
'Cleaning up global variables
OutputFileObject.close
Set OutputFileObject=nothing
Set rs=nothing
Set rsIDC=nothing
Set fso = Nothing
Set LogFileObj = Nothing
Set INIFileObj = nothing
Set cn = Nothing
Set rs = Nothing
Set rsIDC = Nothing
Set dictCluster = Nothing
Set dictDC = Nothing
Set OutputFileObject = Nothing