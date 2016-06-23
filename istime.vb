    Private Function IsTime(ByVal strTime As String) As Boolean
        Dim bCorrect As Boolean, strParts() As String, strTemp As String

        strTemp = Replace(strTime, " ", "")
        bCorrect = False
        strParts = Split(strTemp, ":")
        If UBound(strParts) = 1 Then
            If IsNumeric(strParts(0)) Then
                If CInt(strParts(0)) >= 0 And CInt(strParts(0) <= 23) Then
                    If Len(Trim(strParts(1))) < 5 Then
                        strTemp = Microsoft.VisualBasic.Left(strParts(1), 2)
                        If IsNumeric(strTemp) Then
                            If CInt(strTemp) >= 0 And CInt(strTemp) <= 59 Then
                                bCorrect = True
                            End If
                        End If
                        If Len(Trim(strParts(1))) = 4 Then
                            strTemp = (Microsoft.VisualBasic.Right(strParts(1), 2))
                            If (strTemp = "am" Or strTemp = "pm") And CInt(strParts(0)) < 13 Then
                                bCorrect = True
                            Else
                                bCorrect = False
                            End If
                        End If
                        If Len(Trim(strParts(1))) = 3 Then
                            bCorrect = False
                        End If
                    End If
                End If
            End If
        End If
        IsTime = bCorrect

    End Function
