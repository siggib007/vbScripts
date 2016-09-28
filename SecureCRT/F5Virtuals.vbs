' Connect to an SSH server using the SSH2 protocol. Specify the
' username and password and hostname on the command line as well as
' some SSH2 protocol specific options.
option Explicit

Sub Main
  const ForReading    = 1
  const ForWriting    = 2
  const ForAppending  = 8

  Dim host, cmd, strVersion, user, result, screenrow, strVirtual, strAvail, strState, strReason, strOutFile, fso, objFileOut
  host = "atrou051"
  strOutFile = "C:\Users\sbjarna\Documents\IP Projects\ESME\F5 Forklift\atrou051Virtuals.csv" ' The name of the output file, CSV file listing results
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  set objFileOut = fso.OpenTextFile(strOutFile, ForWriting, True)
  
  'Write a header for output file
  objFileOut.writeline "Virtual,Availability,State,Reason"

  If crt.Session.Connected Then
    crt.Session.Disconnect
  end if

  cmd = "/SSH2 "  & host
  crt.Session.Connect cmd
   '
  crt.Screen.Synchronous = True
  crt.Screen.WaitForString( "#" )
  crt.Screen.Send("show ltm virtual" & vbCR )
  result = crt.Screen.WaitForStrings ("(y/n)","#",5)
  if result = 1 then crt.screen.Send("y")
  do While true
    result=crt.Screen.WaitForStrings ("Ltm::Virtual Server: ","---(less","(END)","#",15)
    ' msgbox "result:" & result
    do while result = 2
      crt.Screen.Send(" ")
      result=crt.Screen.WaitForStrings("Ltm::Virtual Server: ","---(less","(END)","#",15)
    loop
    if result = 3 or result = 4 then exit do
    if result = 0 then 
      msgbox "Timeout waiting for virtual"
      exit do
    end if
    strVirtual = trim(crt.screen.Readstring(vbCR,15))

    result=crt.Screen.WaitForStrings ("Availability     : ","---(less","(END)","#",15)
    do while result = 2
      crt.Screen.Send(" ")
      result=crt.Screen.WaitForStrings ("Availability     : ","---(less","(END)","#",15)
    loop
    if result = 0 then 
      msgbox "Timeout waiting for Availability"
      exit do   
    end if 
    if result = 3 or result = 4 then exit do    
    strAvail = trim(crt.screen.Readstring(vbCR,5))
    result=crt.Screen.WaitForStrings ("State            : ","---(less","(END)","#",15)
    do while result = 2
      crt.Screen.Send(" ")
      result=crt.Screen.WaitForStrings ("State            : ","---(less","(END)","#",15)
    loop
    if result = 0 then 
      msgbox "Timeout waiting for state"
      exit do   
    end if 
    if result = 3 or result = 4 then exit do    
    strState = trim(crt.screen.Readstring(vbCR,5))
    result=crt.Screen.WaitForStrings ("Reason           : ","---(less","(END)","#",15)
    do while result = 2
      crt.Screen.Send(" ")
      result=crt.Screen.WaitForStrings ("Reason           : ","---(less","(END)","#",15)
    loop
    if result = 0 then 
      msgbox "Timeout waiting for reason"
      exit do   
    end if 
    if result = 3 or result = 4 then exit do    
    strReason = replace(trim(crt.screen.Readstring(vbCR,5)),",","")
    objFileOut.writeline strVirtual & "," & strAvail & "," & strState & "," & strReason
  loop
  crt.Screen.Synchronous = False
  crt.session.disconnect
  objFileOut.close
  Set objFileOut = Nothing

  Set fso = Nothing

  msgbox "All Done, Cleanup complete"

End Sub
