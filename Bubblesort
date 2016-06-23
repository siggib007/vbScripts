Function SortIt(ByVal strValue)
  Dim arrUnsorted
  Dim arrSorted
  Dim bolSorted
  Dim strSorted

  arrUnsorted = Split(strValue, vbCrLf)

  bolSorted = False

  Do Until bolSorted
    bolSorted = True
    For I = 0 to UBound(arrUnsorted)
      ' Compare this entry to the next entry
      If I < UBound(arrUnsorted) Then
        If arrUnsorted(I+1) < arrUnsorted(i) Then
          strTemp = arrUnsorted(I+1)
          arrUnsorted(I+1) = arrUnsorted(I)
          arrUnsorted(I) = strTemp
          bolSorted = False
        End If
      End If
    Next
  Loop

  strSorted = ""

  For I = 1 to UBound(arrUnsorted)
    strSorted = strSorted & arrUnsorted(I) & vbCrLf
  Next

  SortIt = strSorted

End Function
