' Connect to an SSH server using the SSH2 protocol. Specify the
' username and password and hostname on the command line as well as
' some SSH2 protocol specific options.

Sub Main

  Dim host
  host = "argsnq21"
  Dim user
  user = "sbjarn"

  ' Prompt for a password instead of embedding it in a script...
  '
  Dim passwd
  passwd = crt.Dialog.Prompt("Enter " & user & "'s password for " & host, "Login", "", True)

  ' Build a command-line string to pass to the Connect method.
  '
  cmd = "/SSH2 /L " & user & " /PASSWORD " & passwd & " /C 3DES /M MD5 " & host

  on error resume next
  crt.Session.Connect cmd
  errcode = crt.GetLastError
  errmsg = crt.GetLastErrorMessage
  crt.Dialog.MessageBox "Error Code: " & errcode & " - Error Message: " & errmsg
  on error goto 0

End Sub
