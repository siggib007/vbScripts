Option Explicit 

Const strIP = "210.205.244.156"
Const strMask = "255.255.255.128"


Main

Sub Main 
	Dim iBitCount, iHostCount, invMask, DecIP, binmask, HexMask, DecMask
	wscript.echo "starting"	
	wscript.echo "Dec 263 is " & dec2bin(263) & " in binary"
	decIP = ip2Int(strIP)
	wscript.echo "decip: " & decip
	invMask = ConvertMask(strMask)
	wscript.echo "invmask: " & invmask
	iBitCount = Mask2Bit (strMask)
	wscript.echo "ibit: " & ibitcount
	iHostCount = 2 ^ (32 - iBitCount)
	wscript.echo "ihost: " & ihostcount
	
	'wscript.echo iBitCount & " bits = " & DotDecGenerate(MaskGenerate(iBitCount))
	wscript.echo strIP & " = " & DecIP
	wscript.echo strmask & " = " & ip2Int(strmask) & " = " & iBitCount & " bits"
	binmask = maskgenerate(ibitcount)
	hexmask = bin2Hex(binmask)
	decmask = HextoDec(HexMask)
	'decmask = CDbl("&H" & hexmask)
	wscript.echo "Mask in Binary: " & binmask
	wscript.echo "Mask in Hex: " & hexmask
	wscript.echo "Mask in Dec: " & decmask & " = " & hextodec(decmask) & " in hex"
	wscript.echo "Mask in dotdec: " & dotdecgenerate(decmask)
	wscript.echo "InvMask = " & invmask
	wscript.echo "Host Count = " & ihostcount
	wscript.echo "Subnet: " & DotDecGenerate(DecIP - (DecIP Mod iHostCount))
	wscript.echo "Broadcast: " & DotDecGenerate(DecIP - (DecIP Mod iHostCount) + iHostCount - 1)
End Sub 

Function Mask2Bit (strMask)
Dim iBitCount, invMask

	invMask = ConvertMask(strMask)

	Select Case IP2Int(invMask)
	    Case 0
	        iBitCount = 32
	    Case 1
	        iBitCount = 31
	    Case Else
	        iBitCount = 32 - CInt(Log(IP2Int(invMask)) / Log(2))
	End Select
	Mask2Bit = iBitCount

End Function

Function MaskGenerate(iBitCount)
	Dim i, strTemp, iDecMask
	
	For i = 1 To iBitCount
	    strTemp = strTemp & "1"
	Next
	For i = iBitCount To 31
	    strTemp = strTemp & "0"
	Next
	'iDecMask = CInt("&B" & strTemp)
	MaskGenerate = strtemp 'iDecMask

End Function

Function DotDecGenerate(iDecMask)
	Dim strTemp, strHexMask
	
	strHexMask = Right("0000000" & dectohex(iDecMask), 8) ' Ensure the string is 8 characters
	'strHexMask = Right("0000000" & bin2Hex(dec2bin(iDecMask)), 8)
	strTemp = CStr(CInt("&H" & Mid(strHexMask, 1, 2))) & "."
	strTemp = strTemp & CStr(CInt("&H" & Mid(strHexMask, 3, 2))) & "."
	strTemp = strTemp & CStr(CInt("&H" & Mid(strHexMask, 5, 2))) & "."
	strTemp = strTemp & CStr (CInt("&H" & Mid(strHexMask, 7, 2)))
	DotDecGenerate = strTemp

End Function

Function ValidateIP(strIP, bMask) 
	Dim iQuads, x, bErr, strBinary, cBit, bFound
	
	iQuads = Split(strIP, ".")
	For x = 0 To UBound(iQuads)
	    If Not IsNumeric(iQuads(x)) Then
	        bErr = True
	        Exit For
	    Else
	        If iQuads(x) > 255 Or iQuads(x) < 0 Then
	            bErr = True
	            Exit For
	        End If
	    End If
	Next
	If UBound(iQuads) <> 3 Then bErr = True
	If false Then 'not bMask And Not bErr Then
	    strBinary = Dec2Bin(IP2Int(strIP))
	    wscript.echo "strbin: " & strbinary
	    cBit = Left(strBinary, 1)
	    bFound = False
	    For x = 2 To Len(strBinary)
	        If Mid(strBinary, x, 1) <> cBit Then
	            If Not bFound Then
	                cBit = Mid(strBinary, x, 1)
	                bFound = True
	            Else
	                bErr = True
	                Exit For
	            End If
	        End If
	    Next
	End If
	If bErr Then
	    ValidateIP = False
	Else
	    ValidateIP = True
	End If

End Function

Function ConvertMask(strMask)
	Dim iQuads, strTemp
	
	If ValidateIP(strMask, True) = False Then
	    ConvertMask = "Invalid mask"
	    Exit Function
	End If
	
	wscript.echo "valid mask"
	iQuads = Split(strMask, ".")
	
	strTemp = CStr(255 - CInt(iQuads(0))) & "."
	strTemp = strTemp & CStr(255 - CInt(iQuads(1))) & "."
	strTemp = strTemp & CStr(255 - CInt(iQuads(2))) & "."
	strTemp = strTemp & CStr(255 - CInt(iQuads(3)))
	ConvertMask = strTemp

End Function

Function IP2Int(strIP)
	Dim strIPQuads, HexIP
	
	If ValidateIP(strIP,false) = False Then
	    IP2Int = -1
	    Exit Function
	End If
	
	strIPQuads = Split(strIP, ".") 
	HexIP = Right("0" & Hex(strIPQuads(0)), 2) & Right("0" & Hex(strIPQuads(1)), 2) & Right("0" & Hex(strIPQuads(2)), 2) & Right("0" & Hex(strIPQuads(3)), 2)
	'wscript.echo hexip
	IP2Int = hextodec(HexIP)
	wscript.echo "Hex format: " & hexIP
	'wscript.echo "DecFormat: " & CLng(hexIP)

End Function

Function SubnetCompare(lIP1, lIP2, lMask)
	Dim lSubnet1, lSubnet2
	
	lSubnet1 = lIP1 Or lMask
	lSubnet2 = lIP2 Or lMask
	SubnetCompare = (lSubnet1 = lSubnet2)

End Function

Function Dec2Bin(lDecNum)
	Dim lPower, strTemp
	lpower = 1
	If ldecnum < 1 Then 
		dec2bin = 0 
		Exit Function 
	End If 
	
	wscript.echo "Converting " & ldecnum & " to bin"
	
	While lDecNum >= lPower
	    If (lDecNum And lPower) > 0 Then
	        strTemp = "1" & strTemp
	    Else
	        strTemp = "0" & strTemp
	    End If
	    wscript.echo "lpower: " & lpower
	    lPower = lPower * 2
	Wend
	
	Dec2Bin = strTemp

End Function

Function Bin2Hex(BinNum)
Dim NumLen, x, iNum, HexNum, i

	x = 0
	i = 0 
	inum = 0 
	hexnum = "" 
	NumLen = Len(BinNum)
	While x < numlen
		inum = inum + ((2^i)*CInt(Mid(binnum,numlen-x,1)))
		x = x + 1
		i = i + 1
		If x mod 4 = 0 Then
			HexNum = Hex(inum) & HexNum
			inum = 0 
			i = 0 
		End If 
	Wend
	'wscript.echo hexnum & " = " & CInt("&H" & hexnum ) & " = " & Hex(CInt("&H" & hexnum ))
	bin2Hex = hexnum
End Function 

Function Hex2Dec(HexNum)
Dim x, inum, numlen, i, HexByte, pos 

	x= 0
	inum=0
	numlen=Len(hexnum)
	wscript.echo "hexnum: " & hexnum
	wscript.echo "NumLen: " & numlen
	wscript.echo Mid(hexnum,5,4)
	wscript.echo Mid(hexnum,1,4)
	While x < numlen '- 4
		i = x + 3
		'If i => numlen Then i = 7
		pos = (2^(x*4))
		If pos = 4 Then pos = 1
		HexByte = Mid(hexnum,numlen-i,4)
		wscript.echo "Pos = " & pos & " i=" & i & " " & hexbyte & " = " & CLng("&H" & hexbyte)
		wscript.echo "inum=" & inum 
		inum = inum + (pos*CLng("&H" & hexbyte))
		'wscript.echo inum
		x = x + 4
	Wend 
	Hex2Dec = inum
	
End Function 
	

Function HexToDec(strHex)
  dim lngResult
  dim intIndex
  dim strDigit
  dim intDigit
  dim intValue
  lngResult = 0
  for intIndex = len(strHex) to 1 step -1
    strDigit = mid(strHex, intIndex, 1)
    intDigit = instr("0123456789ABCDEF", ucase(strDigit))-1
    if intDigit >= 0 then
      intValue = intDigit * (16 ^ (len(strHex)-intIndex))
      lngResult = lngResult + intValue
    else
      lngResult = 0
      intIndex = 0 ' stop the loop
    end if
  next
  HexToDec = lngResult
End Function
 
Function DecToHex(intDEC)
  dim strResult
  dim intValue
  dim intExp
  dim arrDigits
  arrDigits = array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F")
  strResult = ""
  intValue = intDEC
  ' modify below for maximum input number
  intExp = 1099511627776 '16^10
  while intExp >= 1
    if intValue >= intExp then
      strResult = strResult & arrDigits(int(intValue / intExp))
      intValue = intValue - intExp * int(intValue / intExp)
    else
      strResult = strResult & "0"
    end if    
    intExp = intExp / 16
  wend
  DecToHex = strResult
End Function
