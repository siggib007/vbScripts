option explicit

Sub main
  Dim host, user, passwd, app, wb, ws, wbName, waitStrs, HostArray(14), lineparts
  Dim row, screenrow, readline, objShell, cmd, intname, Hostline, strHeaders, x, IntShort
  
  user = "sbjarna"
  
	HostArray(0) 	= "DRCBOT11 FA3/82"
	HostArray(1) 	= "DRCAUS11 FA3/72"
	HostArray(2) 	= "DRCSTL11 F3/77"
	HostArray(3) 	= "DRCWAY11 F3/52"
	 
  crt.screen.synchronous = true
  
  passwd = crt.Dialog.Prompt("Enter password for " & host, "Login", "", True)

  for each hostline in hostarray
  	if hostline <> "" then
  		lineparts = split(hostline, " ")
  		host = lineparts(0)
  		intname = lineparts(1)
  			x = 1
			while (not IsNumeric(mid(intname,x,1))) and (x < len(intname))
				x = x + 1
			wend
			IntShort = mid(intname,x)
 
  		If crt.Session.Connected Then crt.Session.Disconnect
	  	cmd = "/SSH2 /L " & user & " /PASSWORD " & passwd & " " & host
  		crt.Session.Connect cmd
  		
  		crt.Screen.WaitForString "#"
  		crt.Screen.Send("show int " & intname & vbcr )
  		crt.Screen.WaitForString (vbcr)
  		crt.Screen.WaitForString (IntShort)
  		crt.Screen.WaitForString (vbcr)
  		
  		screenrow = crt.screen.CurrentRow
  		readline = trim(crt.Screen.Get(screenrow, 1, screenrow, crt.Screen.Columns ))
  		crt.dialog.messagebox "Status of " & host & " " & intname & " is " & readline
  		crt.Session.Disconnect
  		'crt.Screen.Send "exit" & vbcr
  		crt.dialog.messagebox "Disconnected from " & host
  	end if
  next
  crt.screen.synchronous = false
End Sub