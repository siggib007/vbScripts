public function ConvertTime(iTime)
dim ihour, iminute, isecond
	'172332 = 17:23:32
	ihour = int(iTime/10000)
	iminute = Int((itime-(iHour*10000))/100)
	iSecond = iTime - (iHour*10000) - (iMinute*100)
	ConvertTime = Right("0"&iHour,2) & ":" & Right("0"&iMinute,2) & ":" & Right("0"&iSecond,2)
end Function

public function ConvertDate(iDate)
dim iyear, imonth, iday
	'20010628 = 6/28/01
	iyear = int(iDate/10000)'left(idate,4)
	imonth = Int((iDate-(iyear*10000))/100)'mid(idate,5,2)
	iday = iDate - (iyear*10000) - (imonth*100)'right(idate,2)
	ConvertDate = Right("0"&imonth,2) & "/" & Right("0"&iday,2) & "/" & iyear
end function

wscript.echo ConvertDate(10628) & " " & converttime(172332)
wscript.echo ConvertDate(20010628) & " " & converttime(170302)