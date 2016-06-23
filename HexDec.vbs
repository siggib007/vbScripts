hexmask = "D2CDF49C"
decmask = HexToDec(hexmask)
wscript.echo decmask
'convert back
wscript.echo DecToHex(decmask)

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
