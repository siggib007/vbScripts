' ImportArbitraryDataFromFileToSecureCRTSessions.txt
'   (Designed for use with SecureCRT 5.0 and later)
' 
' This sample script is designed to create sessions from a
' text file (.csv format by default, but this can be edited
' to fit the format you have).
'
' The first line of your data file should contain a comma-separated
' (or whatever you define as the g_strDelimiter below) list of
' supported "fields" designated by the following keywords:
' session_name: The name that should be used for the session. If this field
'               does not exist, the hostname field is used as the session_name.
'       folder: Relative path for session as displayed in the Connect dialog.
'     hostname: The hostname or IP for the remote server (For TAPI protocol,
'               specify the phone number here in the hostname field, for
'               example: 505-555-1212 or 5055551212 or 555-1212 or 5551212).
'               Alternatively, you could use the "port" field to specify the
'               area code, and use the "hostname" field to store the 7-digit
'               phone number.  If you choose to store all 10-digits in the
'               hostname field, make sure that you do not include a port field
'               in your datafile.
'     protocol: The protocol (SSH2, SSH1, telnet, rlogin, TAPI)
'         port: The port on which remote server is listening
'     username: The username for the account on the remote server
'    emulation: The emulation (vt100, xterm, etc.)
'  description: The comment/description. Multiple lines are separated with '\r'
'
' The first line of the data file instructs this script as to the format of the
' fields in your data file and their meaning.  It is not a requirement that all
' the options be used. For example, notice the first line of the following file
' only uses the "hostname", "username", and "protocol" fields.  Note also that
' the "protocol" field can be defaulted so that if a protocol field is empty it
' will use the default value.
'   hostname,username,protocol=SSH2
'   192.168.0.1,root,SSH1
'   192.168.0.2,administrator,SSH2
'   192.168.0.3,root,
'   192.168.0.4,root,
'   192.168.0.5,administrator,telnet
'   ... and so on

' The g_strDefaultProtocol variable will only be defined within the
' ValidateFieldDesignations function if the protocol field has a default value
' (e.g., protocol=TAPI), as read in from the first line of the data file.
Dim g_strDefaultProtocol


' If your data file uses spaces or a character other than comma as the
' delimiter, you would also need to edit the g_strDelimiter value a few lines
' below to indicate that fields are separated by spaces, rather than by commas.
' For example:
'   g_strDelimiter = " "

' If you are importing TAPI sessions, beware of using "," as your delimiter
' because the comma character is actually a valid dialing instruction (pause).
' Using a ";" might be a good alternative for a file that includes TAPI info.
Dim g_strDelimiter
g_strDelimiter = ","     ' comma
' g_strDelimiter = " "    ' space
' g_strDelimiter = ";"    ' semi-colon
' g_strDelimiter = chr(9) ' tab
' g_strDelimiter = "|||"  ' a more unique example of a delimiter.

' The g_strSupportedFields indicates which of all the possible fields, are
' supported in this example script.  If a field designation is found in a data
' file that is not listed in this variable, it will not be imported into the
' session configuration.
Dim g_strSupportedFields
g_strSupportedFields = _
    "description,emulation,folder,hostname,port,protocol,session_name,username"

' If you wish to overwrite existing sessions, set the
' g_bOverwriteExistingSessions to True; for this example script, we're playing
' it safe and leaving any existing sessions in place :).
Dim g_bOverwriteExistingSessions
g_bOverwriteExistingSessions = False

Dim g_fso, g_shell
Set g_fso = CreateObject("Scripting.FileSystemObject")
Set g_shell = CreateObject("WScript.Shell")

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Dim g_strHostsFile, g_strExampleHostsFile
g_strHostsFile = "C:\Temp\MyDataFile.csv"
g_strExampleHostsFile = _
    vbtab & "hostname,protocol,username,folder,emulation" & vbcrlf & _
    vbtab & "192.168.0.1,SSH2,root,Linux Machines,XTerm" & vbcrlf & _
    vbtab & "192.168.0.2,SSH2,root,Linux Machines,XTerm" & vbcrlf & _
    vbtab & "..." & vbcrlf & _
    vbtab & "10.0.100.1,SSH1,admin,CISCO Routers,VT100" & vbcrlf & _
    vbtab & "10.0.101.1,SSH1,admin,CISCO Routers,VT100" & vbcrlf & _
    vbtab & "..." & vbcrlf & _
    vbtab & "myhost.domain.com,SSH2,administrator,Windows Servers,VShell" & _
    vbtab & "..." & vbcrlf & _
    vbtab & "555-1212,tapi,n/a,TAPI Sessions,VT100" & vbcrlf & _
    vbtab & "555-1213,tapi,n/a,TAPI Sessions,VT100" & vbcrlf & _
    vbcrlf & _
    vbtab & "..."
g_strExampleHostsFile = Replace(g_strExampleHostsFile, ",", g_strDelimiter)

Dim g_strConfigFolder, strFieldDesignations, vFieldsArray, vSessionInfo

Dim strSessionName, strHostName, strPort
Dim strUserName, strProtocol, strEmulation
Dim strPathForSessions, strLine, nFieldIndex
Dim strSessionFileName, strFolder, nDescriptionLineCount, strDescription

Dim g_strLastError, g_strErrors
Dim g_nSessionsCreated, g_nDataLines

Import

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub Import()
   
    g_strHostsFile = BrowseForFile( _
        "Please select the host data file to be imported.", _
        g_fso.GetParentFolderName(g_strHostsFile))
    
    If g_strHostsFile = "" Then
        Exit Sub
    End If

    ' Open our data file for reading
    Dim objDataFile
    Set objDataFile = g_fso.OpenTextFile(g_strHostsFile, ForReading, False)

    ' Now read the first line of the data file to determine
    ' the field designations
    strFieldDesignations = LCase(objDataFile.ReadLine)
    
    ' Validate the data file
    If Not ValidateFieldDesignations(strFieldDesignations) Then
        objDataFile.Close
        Exit Sub
    End If
    
    ' Find out where the CRT/SecureCRT configuration lives.
    g_strConfigFolder = ReadRegKey("HKCU\Software\VanDyke\" & _
                                   "SecureCRT\Config Path")
                                      
    Do
        g_strConfigFolder = InputBox(_
            "Your current config folder is specified as seen below. " & _
            vbcrlf & vbcrlf & _
            "If you want to have the session files created in another " & _
            "folder, please specify the folder below.", _
            vbcrlf & vbcrlf & _
            "Please select/confirm Config folder", _
            g_strConfigFolder)
        
        If g_strConfigFolder = "" Then Exit Sub
        
        If Not g_fso.FolderExists(g_strConfigFolder & "\Sessions") Then
            Dim nAnswer
            nAnswer = MsgBox( _
                "This folder doesn't have a ""Sessions"" subfolder." & _
                vbcrlf & vbcrlf & _
                "Would you like to create one now?", vbYesNoCancel, _
                "Import Data To SecureCRT Sessions")
            Select Case nAnswer
                Case vbCancel
                    Exit Sub
                Case vbYes
                    If Not CreateFolderPath( _
                        g_strConfigFolder & "\Sessions") Then
                        MsgBox "Failed to create folder (" & _
                            g_strConfigFolder & "\Sessions" & "): " & _
                            vbcrlf & vbcrlf & g_strLastError, _
                            vbOkOnly, _
                            "Import Data To SecureCRT Sessions"
                    Else
                        Exit Do
                    End If
            End Select
        Else
            ' Folder already exists, so we know we can successfully continue
            Exit Do
        End If
    Loop
   
    ' Get a timer reading so that we can calculate how long it takes to
    ' import.
    nStartTime = Timer
    
    ' Here we create an array of the items that will be used to create
    ' the new session, based on the fields separated by the delimiter
    ' specified in g_strDelimiter
    vFieldsArray = Split(strFieldDesignations, g_strDelimiter)

    ' Loop through reading each line in the data file and creating a session
    ' based on the information contained on each line.
    Do While Not objDataFile.AtEndOfStream
        strLine = ""
        strLine = objDataFile.ReadLine

        ' This sets v_File Data array elements to
        ' each section of strLine, separated by the delimiter
        vSessionInfo = Split(strLine, g_strDelimiter)
        If UBound(vSessionInfo) < UBound(vFieldsArray) Then
            If Trim(strLine) <> "" Then
                g_strErrors = g_strErrors & vbcrlf & _
                    "Insufficient data on line #" & _
                    NN(objDataFile.Line - 1, 4) & ": " & strLine
            Else
                g_strErrors = g_strErrors & vbcrlf & _
                    "Insufficient data on line #" & _
                    NN(objDataFile.Line - 1, 4) & ": [Empty Line]"
            End If
        Else

            ' Variable used to determine if a session file should actually be
            ' created, or if there was an unrecoverable error (and the session
            ' should be skipped).
            Dim bWriteFile
            
            ' Now we will match the items from the new file array to the correct 
            ' variable for the session's ini file
            bWriteFile = True
            For nFieldIndex = 0 To UBound(vSessionInfo)
            
                Select Case vFieldsArray(nFieldIndex)
                    Case "session_name"
                        strSessionName = vSessionInfo(nFieldIndex)
                        ' Check folder name for any invalid characters
                        Dim re
                        Set re = New RegExp
                        re.Pattern = "[\\\|\/\:\*\?\""\<\>]"
                        If re.Test(strSessionName) Then
                            bWriteFile = False
                            If g_strErrors <> "" Then g_strErrors = _
                                vbcrlf & g_strErrors
                                
                            g_strErrors = _
                                "Error: " & _
                                "Invalid characters found in SessionName """ & _
                                strSessionName & """ specified on line #" & _
                                NN(objDataFile.Line - 1, 4) & _
                                ": " & strLine & g_strErrors
                        End If

                    Case "port"
                        strPort = Trim(vSessionInfo(nFieldIndex))
                        If Not IsNumeric(strPort) Then
                            bWriteFile = False
                            If g_strErrors <> "" Then g_strErrors = _
                                vbcrlf & g_strErrors
                                
                            g_strErrors = _
                                "Error: Invalid port """ & strPort & _
                                """ specified on line #" & _
                                NN(objDataFile.Line - 1, 4) & _
                                ": " & strLine & g_strErrors
                        End If
                      
                    Case "protocol"
                        strProtocol = Trim(lcase(vSessionInfo(nFieldIndex)))
                    
                        Select Case strProtocol
                            Case "ssh2"
                                strProtocol = "SSH2"
                            Case "ssh1"
                                strProtocol = "SSH1"
                            Case "telnet"
                                strProtocol = "Telnet"
                            Case "serial"
                                bWriteFile = False
                                g_strErrors = g_strErrors & vbcrlf & _
                                    "Warning: This sample script does " & _
                                    "not support creating sessions " & _
                                    "with protocol """ & _
                                    vSessionInfo(nFieldIndex) & _
                                    """ specified on line #" & _
                                    NN(objDataFile.Line - 1, 4) & _
                                    ": " & strLine
                            Case "tapi"
                                strProtocol = "TAPI"
                            Case "rlogin"
                                strProtocol = "RLogin"
                            Case Else
                                If g_strDefaultProtocol <> "" Then
                                    strProtocol = g_strDefaultProtocol
                                Else
                                    bWriteFile = False
                                    If g_strErrors <> "" Then g_strErrors = _
                                        vbcrlf & g_strErrors
                                        
                                    g_strErrors = _
                                        "Error: Invalid protocol """ & _
                                        vSessionInfo(nFieldIndex) & _
                                        """ specified on line #" & _
                                        NN(objDataFile.Line - 1, 4) & _
                                        ": " & strLine & g_strErrors
                                End If
                        End Select ' for protocols
                        
                    Case "hostname"
                        strHostName = Trim(vSessionInfo(nFieldIndex))
                        If strHostName = "" Then
                            bWriteFile = False
                            g_strErrors = g_strErrors & vbcrlf & _
                                "Warning: 'hostname' field on line #" & _
                                NN(objDataFile.Line - 1, 4) & _
                                " is empty: " & strLine
                        End If
                      
                    Case "username"
                        strUserName = Trim(vSessionInfo(nFieldIndex))
                      
                    Case "emulation"
                        strEmulation = LCase(Trim(vSessionInfo(nFieldIndex)))
                        Select Case strEmulation
                            Case "xterm"
                                strEmulation = "Xterm"
                            Case "vt100"
                                strEmulation = "VT100"
                            Case "vt102"
                                strEmulation = "VT102"
                            Case "vt220"
                                strEmulation = "VT220"
                            Case "ansi"
                                strEmulation = "ANSI"
                            Case "linux"
                                strEmulation = "Linux"
                            Case "scoansi"
                                strEmulation = "SCOANSI"
                            Case "vshell"
                                strEmulation = "VShell"
                            Case "wyse50"
                                strEmulation = "WYSE50"
                            Case "wyse60"
                                strEmulation = "WYSE60"
                            Case Else
                                bWriteFile = False
                                g_strErrors = g_strErrors & vbcrlf & _
                                    "Warning: Invalid emulation """ & _
                                    strEmulation & """ specified on line #" & _
                                    NN(objDataFile.Line - 1, 4) & _
                                    ": " & strLine
                        End Select
                    
                    Case "folder"
                        strFolder = Trim(vSessionInfo(nFieldIndex))
                        
                        ' Check folder name for any invalid characters
                        Set re = New RegExp
                        re.Pattern = "[\|\/\:\*\?\""\<\>]"
                        If re.Test(strFolder) Then
                            bWriteFile = False
                            If g_strErrors <> "" Then g_strErrors = _
                                vbcrlf & g_strErrors
                                
                            g_strErrors = _
                                "Error: Invalid characters in folder """ & _
                                strFolder & """ specified on line #" & _
                                NN(objDataFile.Line - 1, 4) & _
                                ": " & strLine & g_strErrors
                        End If
                        
                    Case "description"
                        strDescription = Trim(vSessionInfo(nFieldIndex))
                        If strDescription <> "" Then
                            Dim vDescriptionLines
                            vDescriptionLines = Split(strDescription, "\r")
                            nDescriptionLineCount = _
                                UBound(vDescriptionLines) + 1
                            strDescription = " " & _
                                Replace(strDescription, "\r", vbcrlf & " ")
                        Else
                            g_strErrors = g_strErrors & vbcrlf & _
                                "Warning: 'description' field on line #" & _
                                NN(objDataFile.Line - 1, 4) & _
                                " is empty: " & strLine
                        End If
                        
                    Case Else
                        ' If there is an entry that the script is not set to use
                        ' in strFieldDesignations, stop the script and display a
                        ' message
                        Dim strMsg1
                        strMsg1 = "Error: Unknown field designation: " & _
                            vFieldsArray(nFieldIndex) & vbcrlf & vbcrlf & _
                            "       Supported fields are as follows: " & _
                            vbcrlf & vbcrlf & vbtab & g_strSupportedFields & _
                            vbcrlf & _
                            vbcrlf & "       For a description of " & _
                            "supported fields, please see the comments in " & _
                            "the sample script file."
                        
                        If Trim(g_strErrors) <> "" Then
                            strMsg1 = strMsg1 & vbcrlf & vbcrlf & _
                                "Other errors found so far include: " & _
                                g_strErrors
                        End If
                        
                        MsgBox strMsg1, _
                            vbOkOnly, _
                            "Import Data To SecureCRT Sessions: Data File Error"
                        Exit Sub
                End Select
            Next
            
            If bWriteFile Then
                'Write the session file
                If strSessionName = "" Then
                    strSessionName = strHostName
                End If
                strPathForSessions = g_strConfigFolder & "\Sessions"
                'call function to check if a folder needs to be created
                If strFolder <> "" Then
                    strPathForSessions = strPathForSessions & "\" & strFolder
                End If
                
                If Not CreateFolderPath(strPathForSessions) Then
                    Dim strMsg
                    If g_nSessionsCreated > 0 Then
                        strMsg = "Error: We were able to create " & _
                            g_nSessionsCreated & _
                            ", but encountered the following fatal error:" & _
                            vbcrlf & vbcrlf
                    End If
                    
                    MsgBox strMsg & "Unable to create folder: " & _
                        strPathForSessions & vbcrlf & vbcrlf & vbtab & _
                        g_strLastError & vbcrlf & vbcrlf & _
                        "Other errors/warnings found so far include:" & _
                        vbcrlf & g_strErrors, _
                        vbOkOnly, _
                        "Import Data To SecureCRT Sessions"
                    Exit Sub
                End If

                strSessionFileName = strSessionName & ".ini"

                If g_fso.FileExists(strPathForSessions & "\" & _
                   strSessionFileName) And _
                   g_bOverwriteExistingSessions = False Then
                    g_strErrors = g_strErrors & vbcrlf & _
                        "Warning: Session already exists (and it was " & _
                        "left in place) for data found on line #" & _
                        NN(objDataFile.Line - 1, 4) & _
                        ": """ & strLine & """"
                Else
                
                    Dim objSessionFile
                    Set objSessionFile = g_Fso.OpenTextFile(_
                        strPathForSessions & "\" & _
                        strSessionFileName, _
                        ForWriting, _
                        True)

                    ' Convert port to Hexadecimal for every protocol except TAPI
                    ' and Serial
                    If strProtocol <> "TAPI"    And _
                       strProtocol <> "Serial"  And _
                       strPort     <> ""        Then
                        strPort = Hex(strPort)
                        strPort = NN(strPort, 8)
                    End If
                    
                    If strProtocol = "SSH2" Then
                        If strPort = "" Then strPort = NN("16", 8)
                        objSessionFile.Write "D:""[SSH2] Port""=" & strPort & _
                            vbcrlf
                    End If
                    If strProtocol = "SSH1" Then
                        If strPort = "" Then strPort = NN("16", 8)
                        objSessionFile.Write "D:""[SSH1] Port""=" & strPort & _
                            vbcrlf
                    End If
                    If strProtocol = "TAPI" Then
                        ' For TAPI sessions, we use the "hostname" field for the
                        ' phone number to dial, and we use the port field for
                        ' the area code (if port field is not empty)
                        If strPort <> "" Then
                            strHostName = "[1]+1 (" & strPort & ") " & _
                                strHostName
                        End If
                        objSessionFile.Write "S:""Dialing List""=" & _
                            strHostName & vbcrlf
                                                    
                        ' Also, if the protocol is TAPI, it doesn't make any
                        ' sense to store Hostname or Username fields, since
                        ' these are unused (it's OK to store them and it
                        ' shouldn't hurt, but we'll make a best effort to be
                        ' "clean" in our .ini file creation).
                        strUserName = ""
                        strHostName = ""
                    End If
                    If strProtocol = "Telnet" Then
                        If strPort = "" Then strPort = NN("17", 8)
                        objSessionFile.Write "D:""Port""=" & strPort & vbcrlf
                    End If
                    
                    objSessionFile.Write "S:""Protocol Name""=" & _
                        strProtocol & vbcrlf
                    objSessionFile.Write "S:""Emulation""=" & strEmulation & _
                        vbcrlf

                    If strHostName <> "" Then
                        objSessionFile.Write "S:""Hostname""=" & strHostName & _
                            vbcrlf
                    End If
                    If strUserName <> "" Then
                        objSessionFile.Write "S:""Username""=" & strUserName & _
                            vbcrlf
                    End If

                    If strDescription <> "" Then
                        objSessionFile.Write "Z:""Description""=" & _
                            NN(nDescriptionLineCount, 8) & vbcrlf
                        objSessionFile.Write strDescription & vbcrlf
                    End If
                    
                    objSessionFile.Close
                    g_nSessionsCreated = g_nSessionsCreated + 1
                End If
            End If
            
            
            strEmulation = ""
            strPort = ""
            strHostName = ""
            strFolder = ""
            strUserName = ""
            strSessionName = ""
            strDescription = ""
            nDescriptionLineCount = 0
        End If

    Loop

    g_nDataLines = objDataFile.Line
    objDataFile.Close
    
    Dim strResults
    strResults = "Import operation completed in " & _
        GetMinutesAndSeconds(Timer - nStartTime)
    
    If g_nSessionsCreated > 0 Then
        strResults = strResults & _
            vbcrlf & _
            String(70, "-") & vbcrlf & _
            "Number of Sessions created: " & g_nSessionsCreated
    Else
        strResults = strResults & vbcrlf & _
            String(70, "-") & vbcrlf & _
            "No sessions were created from " & g_nDataLines & " lines of data."
    End If
    
    ' Log activity information to a file for debugging purposes...
    strFilename = g_strConfigFolder & "\SecureCRT-Session-ImportLog-" & _
        Year(Now) & "-" & _
        NN(Month(Now), 2) & "-" & _
        NN(Day(Now), 2) & "--" & _
        NN(Hour(Now), 2) & "-" & _
        NN(Minute(Now), 2) & "-" & _
        NN(Second(Now), 2) & ".txt"
    Set objFile = g_fso.OpenTextFile(strFilename, ForWriting, True)
    objFile.Write strResults & vbcrlf & vbcrlf & _
        String(70, "-") & vbcrlf & _
        "Errors/warnings from this operation include:" & _
        g_strErrors
    objFile.Close

    ' Display the log file as an indication that the
    ' information has been imported.
    g_shell.Run Chr(34) & strFilename & Chr(34), 5, False
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'      Helper Methods and Functions
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function ValidateFieldDesignations(ByRef strFields)
    If Instr(strFieldDesignations, g_strDelimiter) = 0 Then
        Dim strErrorMsg, strDelimiterDisplay
        strErrorMsg = "Invalid header line in data file. " & _
            "Delimiter character not found: "
        If Len(g_strDelimiter) > 1 Then
            strDelimiterDisplay = g_strDelimiter
        Else
            If Asc(g_strDelimiter) < 33 or Asc(g_strDelimiter) > 126 Then
                strDelimiterDisplay = "ASCII[" & Asc(g_strDelimiter) & "]"
            Else
                strDelimiterDisplay = g_strDelimiter
            End If
        End If
        strErrorMsg = strErrorMsg & strDelimiterDisplay & vbcrlf & vbcrlf & _
            "The first line of the data file is a header line " & _
            "that must include" & vbcrlf & _
            "a '" & strDelimiterDisplay & _
            "' separated list of field keywords." & vbcrlf & _
            vbcrlf & "'hostname' and 'protocol' are required keywords." & _
            vbcrlf & vbcrlf & _
            "The remainder of the lines in the file should follow the " & _
            vbcrlf & _
            "pattern established by the header line " & _
            "(first line in the file)." & vbcrlf & "For example:" & vbcrlf & _
            g_strExampleHostsFile
        MsgBox strErrorMsg, _
               vbOkOnly, _
               "Import Data To SecureCRT Sessions"
        Exit Function
    End If
    
    If Instr(strFieldDesignations, "hostname") = 0 Then
        strErrorMsg = "Invalid header line in data file. " & _
            "'hostname' field is required."
        If Len(g_strDelimiter) > 1 Then
            strDelimiterDisplay = g_strDelimiter
        Else
            If Asc(g_strDelimiter) < 33 Or Asc(g_strDelimiter) > 126 Then
                strDelimiterDisplay = "ASCII[" & Asc(g_strDelimiter) & "]"
            Else
                strDelimiterDisplay = g_strDelimiter
            End If
        End If

        MsgBox strErrorMsg & vbcrlf & _
            "The first line of the data file is a header line " & _
            "that must include" & vbcrlf & _
            "a '" & strDelimiterDisplay & _
            "' separated list of field keywords." & vbcrlf & _
            vbcrlf & "'hostname' and 'protocol' are required keywords." & _
            vbcrlf & vbcrlf & _
            "The remainder of the lines in the file should follow the " & _
            vbcrlf & _
            "pattern established by the header line " & _
            "(first line in the file)." & vbcrlf & "For example:" & vbcrlf & _
            g_strExampleHostsFile, _
            vbOkOnly, _
            "Import Data To SecureCRT Sessions"
        Exit Function
    End If
    
    If Instr(strFieldDesignations, "protocol") = 0 Then
        MsgBox "Invalid data file header line: " & vbcrlf & vbcrlf & _
            vbtab & strFieldDesignations & vbcrlf & _
            vbcrlf & "--> 'protocol' field is required.", _
            vbOkOnly, _
            "Import Data To SecureCRT Sessions"
        Exit Function
    Else
        ' We found "protocol", now look for a default protocol designation
        vFields = Split(strFields,g_strDelimiter)
        For each strField In vFields
            If (InStr(strField, "protocol") > 0) And _
               (Instr(strField, "=") >0) Then
                    g_strDefaultProtocol = UCase(Split(strField, "=")(1))
                    
                    ' fix the protocol field since we know
                    ' the default protocol value
                    strFields = Replace(strFields, strField, "protocol")
            End If
        Next
    End If
    
    ValidateFieldDesignations = True
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function ReadRegKey(strKeyPath)
    On Error Resume Next
    Err.Clear
    ReadRegKey = g_shell.RegRead(strKeyPath)
    If Err.Number <> 0 Then
        ' Registry key must not have existed.
        ' ReadRegKey will already be empty, but
        ' for the sake of clarity, we'll set it
        ' to an empty string explicitly
        ReadRegKey = ""
    End If
    On Error Goto 0
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function CreateFolderPath(strPath)
' Recursive function
    If g_fso.FolderExists(strPath) Then
        CreateFolderPath = True
        Exit Function
    End If
    
    ' Check to see if we've reached the drive root
    If Right(strPath, 2) = ":\" Then
        CreateFolderPath = True
        Exit Function
    End If
    
    ' None of the other two cases were successful, so attempt to create the
    ' folder
    On Error Resume Next
    g_fso.CreateFolder strPath
    nError = Err.Number
    strErr = Err.Description
    On Error Goto 0
    If nError <> 0 Then
        ' Error 76 = Path not found, meaning that the full path doesn't exist.
        ' Call ourselves recursively until all the parent folders have been
        ' created:
        If nError = 76 Then _
            CreateFolderPath(g_fso.GetParentFolderName(strPath))
        
        On Error Resume Next
        g_fso.CreateFolder strPath
        nError = Err.Number
        strErr = Err.Description
        On Error Goto 0
        
        ' If the Error is not = 76, then we have to bail since we
        ' no longer have any hope of successfully creating each folder in the
        ' tree
        If nError <> 0 Then
            g_strLastError = strErr
            Exit Function
        End If
    End If
    
    CreateFolderPath = True
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function NN(nNumber, nDesiredDigits)
' Normalizes a number to have a number of zeros in front of it so that the
' total length of the number (displayed as a string) is nDesiredDigits.
    Dim nIndex, nOffbyDigits, strResult
    nOffbyDigits = nDesiredDigits - Len(nNumber)

    NN = nNumber
    
    If nOffByDigits = 0 Then Exit Function
    
    If nOffByDigits > 0 Then
        ' The number provided doesn't have enough digits
        strResult = String(nOffbyDigits, "0") & nNumber
    Else
        ' The number provided has too many digits.
        
        nOffByDigits = Abs(nOffByDigits)
        
        ' Only remove leading digits if they're all insignificant (0).
        If Left(nNumber, nOffByDigits) = String(nOffByDigits, "0") Then
            strResult = Mid(nNumber, nOffByDigits + 1)
        Else
            ' If leading digits beyond desired number length aren't 0,
            ' we'll return the number as originally passed in.
            strResult = nNumber
        End If
    End If

    NN = strResult
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function BrowseForFile(strTitle, strInitialDir)
    ' Since Windows XP is the only OS on which the UserAccounts.CommonDialog
    ' object is available, find out if we are using Windows XP or a different
    ' operating system (Vista and newer Windows versions are not guaranteed to
    ' have this control available, so for these operating systems, we will need
    ' to present a much simpler interface for choosing a file: an input box).
    Dim strOSName
    Set objWMIService = GetObject("winmgmts:" & _
        "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colSettings = _
        objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
    For Each objOS In colSettings
        ' Windows XP might look like this:
  ' Microsoft Windows XP Professional|C:\WINDOWS|\Device\Harddisk0\Partition1
        strOsName = Split(objOS.Name, "|")(0)
        Exit For
    Next
    
    If Instr(strOSName, "Windows XP") = 0 Then
        ' If not using Windows XP, we are limited in our ability to present
        ' an adequate file browser dialog, so we'll just use InputBox
        ' to prompt for the path to the file.
        strFilePath = g_strHostsFile
        Do
            strFilePath = InputBox(strTitle, _
                "SecureCRT Import Script", _
                strFilePath)
            If strFilePath = "" Then Exit Function
            If g_fso.FileExists(strFilePath) Then Exit Do
            MsgBox "Path not found: " & vbcrlf & vbcrlf & vbtab & _
                strFilePath & vbcrlf & vbcrlf & _
                "Please specify a valid file path", _
                vbOkOnly, _
                "Import Data To SecureCRT Sessions"
        Loop
        BrowseForFile = strFilePath
    Else
        ' Based on information obtained from
        ' http://blogs.msdn.com/gstemp/archive/2004/02/17/74868.aspx
        ' NOTE: Will only work with WindowsXP since other OS's
        '       don't have a UserAccounts.CommonDialog ActiveX
        '       object registered.
        Dim objDialog
        Set objDialog = CreateObject("UserAccounts.CommonDialog")
        'Set objDialog = CreateObject("MSComDlg.CommonDialog")
        objDialog.Filter = "CSV Files|*.csv;Text Files|*.txt;All Files|*.*"
        objDialog.FilterIndex = 1
        objDialog.InitialDir = strInitialDir
        'objDialog.InitDir = g_strMyDocs
        'objDialog.MaxFileSize = 512
        If MsgBox(strTitle, _
            vbOkCancel, _
            "Import Data To SecureCRT Sessions") <> vbok Then
            Exit Function
        End If
        'objDialog.DialogTitle = strTitle
        'objDialog.CancelError = False
        objDialog.ShowOpen

        BrowseForFile = objDialog.FileName
    End If
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function GetMinutesAndSeconds(nTotalSecondsElapsed)
    Dim nMinutesElapsed, nSecondsValue, nSecondsElapsed

    If nTotalSecondsElapsed = 0 Then
        GetMinutesAndSeconds = "less than a second."
        Exit Function
    End If

    ' convert seconds into a fractional minutes value
    nMinutesElapsed = nTotalSecondsElapsed / 60

    ' convert the decimal portion into the number of remaining seconds
    nSecondsValue = nMinutesElapsed - Fix(nMinutesElapsed)
    nSecondsElapsed = Fix(nSecondsValue * 60)

    ' Remove the decimal from the minutes value
    nMinutesElapsed = Fix(nMinutesElapsed)

    ' Get the number of Milliseconds, Seconds, and Minutes
    ' to return to the caller byref
    nMSeconds = fix(1000 * (nTotalSecondsElapsed - Fix(nTotalSecondsElapsed)))
    nSeconds = nSecondsElapsed
    nMinutes = nMinutesElapsed

    ' Form the final string to be returned
    GetMinutesAndSeconds = nMinutesElapsed & " minutes, " & _
        nSecondsElapsed & " seconds, and " & _
        nMSeconds & " ms"
End Function