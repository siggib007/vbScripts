Option Explicit
Dim CNfilename, PSSfilename, IDCfilename, Allfilename


   
'**************************************************
'Max
'Finds the maximum element of an array
'Input Param: array of data
'Return: maximum of array
'**************************************************

Public Function max(DataArray)
    Dim temp, n
    Dim i
    i = 0
    n = UBound(DataArray) - 1
    temp = DataArray(0)
    Do While IsNull(temp)
        temp = DataArray(i)
        i = i + 1
    Loop
    For i = 1 To n
        If DataArray(i) > temp Then
            temp = DataArray(i)
        End If
    Next
    max = temp
End Function

'**************************************************
'UtilIndex
'Performs calcs to create util(in), even entries,
'   and util(out), odd entries, indices
'Input Param: Kb arrays and BW array in orig order
'Return: Utilization index in array
'**************************************************

Public Function UtilIndex(KbMed, Kb95th, Bandwidth)
    Dim i, n
    n = UBound(KbMed) - 1
    Dim Util()
    ReDim Util(n)
    For i = 0 To n
        If (0.5 - ((KbMed(i) + (Kb95th(i) / 2)) / Bandwidth(i))) < -1 Then
            Util(i) = -1
        Else: Util(i) = (0.5 - ((KbMed(i) + (Kb95th(i) / 2)) / Bandwidth(i)))
        End If
    Next
    UtilIndex = Util
End Function

'**************************************************
'GrowthIndex
'Performs calcs to create growth index
'Input Param: 90D growth array
'Return: growth index in array
'**************************************************

Public Function GrowthIndex(DataArray)
    Dim i, n, maxofarray
    n = UBound(DataArray) - 1
    Dim growth()
    ReDim growth(n)
    maxofarray = max(DataArray)
    For i = 0 To n
        If (-(DataArray(i)) / maxofarray) > 1 Then
            growth(i) = 1
        Else: growth(i) = (-(DataArray(i)) / maxofarray)
        End If
    Next
    GrowthIndex = growth
End Function

'**************************************************
'QdelayIndex
'Performs calcs to create queue delay index
'Input Param: 30D min and 95th ms arrays
'Return: Queue delay index in array
'**************************************************

Public Function QdelayIndex(msMin, ms95th)
    Dim i, n
    n = UBound(msMin) - 1
    Dim QD()
    ReDim QD(n)
    For i = 0 To n
        QD(i) = (((msMin(i) - ms95th(i)) / (5000 - msMin(i))) + 0.5) * 2
    Next 
    QdelayIndex = QD
End Function

'**************************************************
'LossIndex
'Performs calcs to create loss index
'Input Param: Availability array
'Return: Loss index in array
'**************************************************

Public Function LossIndex(Avail)
    Dim i, n
    n = UBound(Avail) - 1
    Dim Loss
    ReDim Loss(n)
    For i = 0 To n
        If Avail(i) >= 99.9 Then
            Loss(i) = (10 * Avail(i)) - 999
        Else: Loss(i) = (Avail(i) / 99.9) - 1
        End If
    Next
    LossIndex = Loss
End Function

'**************************************************
'SumOfNeg
'Performs first step of final rank calculation.
'   Sums the negative weighted index values for
'   each circuit.
'Input Param: Util, Growth, QD, and Loss indexes
'Return: Array of sum of negative indexes per circuit
'**************************************************

Public Function SumOfNeg(Util, GrKbMed, GrKb95th, GrmsMed, Grms95th, QD, Loss)
    Dim i, n
    n = UBound(Util)
    Dim sum()
    ReDim sum(n)
    Dim Util2, GrKbMed2, GrKb95th2, GrmsMed2, Grms95th2, QD2, Loss2
    Util2 = Util
    GrKbMed2 = GrKbMed
    GrKb95th2 = GrKb95th
    GrmsMed2 = GrmsMed
    Grms95th2 = Grms95th
    QD2 = QD
    Loss2 = Loss
    For i = 0 To n - 1
        If Util2(i) > 0 Or IsNull(Util2(i)) Then
            Util2(i) = 0
        End If
        If Util2(i + 1) > 0 Or IsNull(Util2(i + 1)) Then
            Util2(i + 1) = 0
        End If
        If GrKbMed2(i) > 0 Or IsNull(GrKbMed2(i)) Then
            GrKbMed2(i) = 0
        End If
        If GrKbMed2(i + 1) > 0 Or IsNull(GrKbMed2(i + 1)) Then
            GrKbMed2(i + 1) = 0
        End If
        If GrKb95th2(i) > 0 Or IsNull(GrKb95th2(i)) Then
            GrKb95th2(i) = 0
        End If
        If GrKb95th2(i + 1) > 0 Or IsNull(GrKb95th2(i + 1)) Then
            GrKb95th2(i + 1) = 0
        End If
        If GrmsMed2(i) > 0 Or IsNull(GrmsMed2(i)) Then
            GrmsMed2(i) = 0
        End If
        If Grms95th2(i) > 0 Or IsNull(Grms95th2(i)) Then
            Grms95th2(i) = 0
        End If
        If QD2(i) > 0 Or IsNull(QD2(i)) Then
            QD2(i) = 0
        End If
        If Loss2(i) > 0 Or IsNull(Loss2(i)) Then
            Loss2(i) = 0
        End If
        If i / 2 <> Int(i / 2) Then
            sum(i) = sum(i - 1)
        Else: sum(i) = 0.15 * (Util2(i) + Util2(i + 1)) + 0.025 * (GrKbMed2(i) + GrKbMed2(i + 1) + GrKb95th2(i) + GrKb95th2(i + 1)) + 0.05 * (GrmsMed2(i) + Grms95th2(i)) + 0.25 * (QD2(i) + Loss2(i))
        End If
    Next
     sum(n) = sum(n - 1)
     SumOfNeg = sum
    'Dim j
End Function

'**************************************************
'MinOfIndex
'Performs the second step of the final rank calc.
'   Weighted sum of the most negative and one
'   tenth of the second most negative.
'Input Param: Util, Growth, QD, Loss, NegSum indexes
'Return: Min index "sum" of inputs
'**************************************************

Public Function MinOfIndex(Util, GrKbMed, GrKb95th, GrmsMed, Grms95th, QD, Loss, PI)
    Dim i, n
    n = UBound(Util)
    Dim minvalue
    ReDim minvalue(n)
    Dim Util2, GrKbMed2, GrKb95th2, GrmsMed2, Grms95th2, QD2, Loss2, PI2
    Util2 = Util
    GrKbMed2 = GrKbMed
    GrKb95th2 = GrKb95th
    GrmsMed2 = GrmsMed
    Grms95th2 = Grms95th
    QD2 = QD
    Loss2 = Loss
    PI2 = PI
    For i = 0 To n - 1
        Dim temparray
        ReDim temparray(11)
        If IsNull(Util2(i)) Then
            temparray(0) = 0
        Else: temparray(0) = 0.5 * (Util2(i))
        End If
        If IsNull(Util2(i + 1)) Then
            temparray(1) = 0
        Else: temparray(1) = 0.5 * (Util2(i + 1))
        End If
        If IsNull(GrKbMed2(i)) Then
            temparray(2) = 0
        Else: temparray(2) = 0.1 * (GrKbMed2(i))
        End If
        If IsNull(GrKbMed2(i + 1)) Then
            temparray(3) = 0
        Else: temparray(3) = 0.1 * (GrKbMed2(i + 1))
        End If
        If IsNull(GrKb95th2(i)) Then
            temparray(4) = 0
        Else: temparray(4) = 0.1 * (GrKb95th2(i))
        End If
        If IsNull(GrKb95th2(i + 1)) Then
            temparray(5) = 0
        Else: temparray(5) = 0.1 * (GrKb95th2(i + 1))
        End If
        If IsNull(GrmsMed2(i)) Then
            temparray(6) = 0
        Else: temparray(6) = 0.2 * (GrmsMed2(i))
        End If
        If IsNull(Grms95th2(i)) Then
            temparray(7) = 0
        Else: temparray(7) = 0.2 * (Grms95th2(i))
        End If
        If IsNull(QD2(i)) Then
            temparray(8) = 0
        Else: temparray(8) = QD2(i)
        End If
        If IsNull(Loss2(i)) Then
            temparray(9) = 0
        Else: temparray(9) = Loss2(i)
        End If
        If IsNull(PI2(i)) Then
            temparray(10) = 0
        Else: temparray(10) = PI2(i)
        End If
        If IsNull(QD2(i)) Then
            QD2(i) = 0
        End If
        If IsNull(Util2(i)) Then
            Util2(i) = 0
        End If
        If IsNull(Util2(i + 1)) Then
            Util2(i + 1) = 0
        End If
        temparray(11) = (0.5 * (Util2(i) + Util2(i + 1)) + QD2(i)) / 3
        
    Dim sortedtemparray
        sortedtemparray = SortArray(temparray)
        If i / 2 <> Int(i / 2) Then
            minvalue(i) = minvalue(i - 1)
        Else: minvalue(i) = sortedtemparray(11) + (sortedtemparray(10) / 10)
        End If
    Next
    minvalue(n) = minvalue(n - 1)
    MinOfIndex = minvalue
End Function

'**************************************************
'WhereMinOfArray
'Locates the minimum value of an array
'Input Param: Array of data
'Return: Location of min
'**************************************************

Public Function WhereMinOfArray(DataArray)
    Dim i, n
    n = UBound(DataArray)
    Dim location
    location = 0
    For i = 1 To n
        If DataArray(i) < DataArray(location) Then
            location = i
        End If
    Next
    WhereMinOfArray = location
End Function

'**************************************************
'WhereMaxOfArray
'Locates the maximum value of an array
'Input Param:
'Return:
'**************************************************

Public Function WhereMaxOfArray(DataArray)
    Dim i, n
    n = UBound(DataArray)
    Dim location
    location = 0
    For i = 1 To n
        If DataArray(i) > DataArray(location) Then
            location = i
        End If
    Next
    WhereMaxOfArray = location
End Function

'**************************************************
'SortArray
'Sorts an array from largest, at 0 location, to smallest
'Input Param: array of data
'Return: sorted array
'**************************************************

'sort from largest at 0 to smallest at n
Public Function SortArray(DataArray)
    Dim i, n, min, wheremin, wheremax
    n = UBound(DataArray)
    Dim Sorted()
    ReDim Sorted(n)
    Dim Dataarray2
    Dataarray2 = DataArray
    wheremin = WhereMinOfArray(Dataarray2)
    min = Dataarray2(wheremin)
    For i = 0 To n
        wheremax = WhereMaxOfArray(Dataarray2)
        Sorted(i) = Dataarray2(wheremax)
        Dataarray2(wheremax) = Dataarray2(wheremin) - 1
    Next
    SortArray = Sorted
End Function

'**************************************************
'MedianOfArray
'Finds median of a sorted array
'Input Param: sorted array
'Return: median of array
'**************************************************

Public Function MedianOfArray(SortedArray)
    Dim n
    n = UBound(SortedArray)
    MedianOfArray = SortedArray(n / 2)
End Function

'**************************************************
'Percentile95
'Finds 95th percentile value of a sorted array
'Input Param: Sorted array
'Return: 95th percentile of data
'**************************************************

Public Function Percentile95(SortedArray)
    Dim n, rounded
    n = UBound(SortedArray)
    rounded = Round(0.05 * n)
    Percentile95 = SortedArray(rounded)
End Function

'**************************************************
'Percentile5
'Finds 5th percentile value of a sorted array
'Input Param: Sorted array
'Return: 5th percentile of data
'**************************************************

Public Function Percentile5(SortedArray)
    Dim n, rounded
    n = UBound(SortedArray)
    rounded = Round(0.95 * n)
    Percentile5 = SortedArray(rounded)
End Function

'**************************************************
'ShiftAroundMedian
'Shifts array of values to -1 to 1 scale
'Input Param: Array of data
'Return: Array of data from -1 to 1
'**************************************************

Public Function ShiftAroundMedian(MinSumArray)
    Dim i, n, per95, per5, median
    n = UBound(MinSumArray)
    Dim Sorted
    ReDim Sorted(n)
    Sorted = SortArray(MinSumArray)
    Dim shifted()
    ReDim shifted(n)
    per95 = Percentile95(Sorted)
    per5 = Percentile5(Sorted)
    median = MedianOfArray(Sorted)
    For i = 0 To n
        If MinSumArray(i) < median Then
            shifted(i) = (MinSumArray(i) - median) / (median - per5)
        Else: shifted(i) = (MinSumArray(i) - median) / (per95 - median)
        End If
    Next
    For i = 0 To n
        If shifted(i) > 1 Then
            shifted(i) = 1
        ElseIf shifted(i) < -1 Then
            shifted(i) = -1
        End If
    Next
    ShiftAroundMedian = shifted
End Function


'**************************************************
'SortAllArrays
'Sorts a multidimensional array by one column
'Input Param: Arrays to be included in multi-
'   dimensional array
'Return: Multi-dimensional array of sorted data
'**************************************************

Public Function SortAllArrays(CircuitList, Network, DeviceID, HubIntID, TailIntID, HubYN, Circuit, HubSite, CCS, DeviceName, IntHWDesc, IntTypeID, BW, PortSpeed, FinalRank, IndexUtil, IndexMedKbGr, Index95thKbGr, IndexmsMedGr, Indexms95thGr, IndexQD, IndexLoss, NegSum, Yes)
    Dim i, n, max, wheremin, wheremax
    n = UBound(Network)
    Dim AllSorted()
    ReDim AllSorted(n / 2, 27)
    Dim FinalRank2
    FinalRank2 = FinalRank
     
    wheremax = WhereMaxOfArray(FinalRank2)
    max = FinalRank2(wheremax)
    For i = 0 To n - 1 Step 2
        wheremin = WhereMinOfArray(FinalRank2)
        If wheremin / 2 <> Int(wheremin / 2) Then
            wheremin = wheremin - 1
        End If
        Dim j
        j = i / 2
        AllSorted(j, 0) = CircuitList(wheremin)
        AllSorted(j, 2) = Network(wheremin)
        AllSorted(j, 3) = DeviceID(wheremin)
        AllSorted(j, 4) = HubIntID(wheremin)
        AllSorted(j, 5) = TailIntID(wheremin)
        AllSorted(j, 6) = HubYN(wheremin)
        AllSorted(j, 7) = Circuit(wheremin)
        AllSorted(j, 8) = HubSite(wheremin)
        AllSorted(j, 9) = CCS(wheremin)
        AllSorted(j, 10) = DeviceName(wheremin)
        AllSorted(j, 11) = IntHWDesc(wheremin)
        AllSorted(j, 12) = IntTypeID(wheremin)
        AllSorted(j, 13) = BW(wheremin)
        AllSorted(j, 14) = PortSpeed(wheremin)
        If not IsNull (FinalRank(wheremin)) Then
        	AllSorted(j, 15) = CStr(FormatNumber(FinalRank(wheremin), 2))
        End If
        If not IsNull (IndexUtil(wheremin)) Then
        	AllSorted(j, 16) = CStr(FormatNumber(IndexUtil(wheremin), 2))
        End If 
        If not IsNull (IndexUtil(wheremin + 1)) Then
        	AllSorted(j, 17) = CStr(FormatNumber(IndexUtil(wheremin + 1), 2))
        End If
        If not IsNull (IndexMedKbGr(wheremin)) Then
        	AllSorted(j, 18) = CStr(FormatNumber(IndexMedKbGr(wheremin), 2))
        End If 
	If not IsNull (IndexMedKbGr(wheremin + 1)) Then
        	AllSorted(j, 19) = CStr(FormatNumber(IndexMedKbGr(wheremin + 1), 2))
        Else
        	wscript.echo "Null"
        End If
        'AllSorted(j, 19) = CStr(FormatNumber(IndexMedKbGr(wheremin + 1), 2))
        If not IsNull (Index95thKbGr(wheremin)) Then
        	AllSorted(j, 20) = CStr(FormatNumber(Index95thKbGr(wheremin), 2))
        End If
        If not IsNull (Index95thKbGr(wheremin + 1)) Then
        	AllSorted(j, 21) = CStr(FormatNumber(Index95thKbGr(wheremin + 1), 2))
        Else
        	wscript.echo "null"
        End If
        If not IsNull (IndexmsMedGr(wheremin)) Then
        	AllSorted(j, 22) = CStr(FormatNumber(IndexmsMedGr(wheremin), 2))
        End If
        If not IsNull (Indexms95thGr(wheremin)) Then
        	AllSorted(j, 23) = CStr(FormatNumber(Indexms95thGr(wheremin), 2))
        End If
        If not IsNull (IndexQD(wheremin)) Then
        	AllSorted(j, 24) = CStr(FormatNumber(IndexQD(wheremin), 2))
        End If
        If not IsNull (IndexLoss(wheremin)) Then
        	AllSorted(j, 25) = CStr(FormatNumber(IndexLoss(wheremin), 2))
        End If 
        If not IsNull (NegSum(wheremin)) Then
        	AllSorted(j, 26) = CStr(FormatNumber(NegSum(wheremin), 2))
        End If
        AllSorted(j, 27) = Yes(wheremin)
        FinalRank2(wheremin) = FinalRank2(wheremax) + 1
        FinalRank2(wheremin + 1) = FinalRank2(wheremax) + 1
    Next
    Dim m
    Dim k
    For m = 0 To n / 2
        For k = 0 To 27
            If IsNull(AllSorted(m, k)) Then
                AllSorted(m, k) = ""
            End If
        Next
    Next
    SortAllArrays = AllSorted
End Function

Public Function PSSONSorted(AllSorted)
    Dim PSSONcount
    PSSONcount = 1
    Dim m
    Dim k
    Dim n
    n = UBound(AllSorted)
    Dim Sorted()
    ReDim Sorted(n, 27)
    For m = 0 To n
    If AllSorted(m, 0) = "PSSON" Then
        For k = 0 To 27
            Sorted(PSSONcount - 1, k) = AllSorted(m, k)
            Sorted(PSSONcount - 1, 1) = PSSONcount
        Next
        PSSONcount = PSSONcount + 1
    End If
    Next
    PSSONSorted = Sorted
End Function

Public Function CorpNetSorted(AllSorted)
    Dim CorpNetcount
    CorpNetcount = 1
    Dim m
    Dim k
    Dim n
    n = UBound(AllSorted)
    Dim Sorted()
    ReDim Sorted(n, 27)
    For m = 0 To n
    If AllSorted(m, 0) = "Corp Net" Then
        For k = 0 To 27
            Sorted(CorpNetcount - 1, k) = AllSorted(m, k)
            Sorted(CorpNetcount - 1, 1) = CorpNetcount
        Next
        CorpNetcount = CorpNetcount + 1
    End If
    Next
    CorpNetSorted = Sorted
End Function


Public Function IDCBackendSorted(AllSorted)
    Dim IDCBackendcount
    IDCBackendcount = 1
    Dim m
    Dim k
    Dim n
    n = UBound(AllSorted)
    Dim Sorted()
    ReDim Sorted(n, 27)
    For m = 0 To n
    If AllSorted(m, 0) = "IDC Backend" Then
        For k = 0 To 27
            Sorted(IDCBackendcount - 1, k) = AllSorted(m, k)
            Sorted(IDCBackendcount - 1, 1) = IDCBackendcount
        Next
        IDCBackendcount = IDCBackendcount + 1
    End If
    Next
    IDCBackendSorted = Sorted
End Function


Private Sub Main()

'**************************************************
'Main
'Reads data from NetID2000 db Joined Table (sorted
'   by circuit then by MIB)
'Input Param: None
'**************************************************

Dim CircuitList(), Network(), DeviceID(), TailIntID(), TargetID(), HubIntID(), Circuit(), HubSite(), IntTypeID()
Dim CCS(), BW(), PortSpeed(), Yes(), DeviceName(), IntHWDesc(), FuncDesc(), HubYN(), MIB()
Dim KbpsMed(), Kbps95th(), KbpsMedGr(), Kbps95thGr(), ms30DMin(), ms30D95th(), msMedGr(), ms95thGr()
Dim SiteDevice(), Avail30(), HrsOut30(), AvgLoss30(), Avail7(), HrsOut7(), AvgLoss7()
Dim HC_GA(), HC_PD(), HC_SM(), HC_Oper(), HC_Total(), BWperHC()
Dim IndexUtil, IndexMedKbGr, Index95thKbGr, IndexmsMedGr, Indexms95thGr, IndexQD, IndexLoss
Dim CorpNetList, PSSONList, IDCBackendList, AllList, DatedFilename
Dim dtDay, i, n, j, NegSum, MinValues, FinalRank, TotalSort, X
Dim fso, file1, file2, file3, file4, INIFileObj, strINILine, strINIPath, ScriptPathStr
Dim CNmerged, CNtaskID, PSSmerged, PSStaskID, IDCmerged, IDCtaskID, INILineArray
Dim rs, cn, strConfigSrv, strConfigDB, AllFileOldName, iAnswer, strMsg, bSaveOld

'dtDay = Format(Now, "mm-dd-yyyy")
dtDay = DatePart ("m",Now) & "-" & datepart ("d",Now) & "-" & datepart ("yyyy",Now)

ScriptPathStr = Left (wscript.scriptFullname,InStr(wscript.scriptFullName,wscript.scriptname)-1)

strINIPath = ScriptPathStr & "Perfindx.ini"

If wscript.arguments.count > 0 Then
	If LCase (Left (wscript.arguments(0),1)) = "y" Then
		wscript.echo "Detected yes"
		bSaveOld = true
	Else
		wscript.echo "Detected something other than yes"
		bSaveOld = false
	End If
Else
	wscript.echo "No input detected, switching to attended mode"
	bSaveOld = null
End If


Set fso = CreateObject("Scripting.FileSystemObject")
If not fso.fileExists(strINIPath) Then
	wscript.echo "Couldn't find required INI file " & strINIPath
	wscript.echo "script full path name " &  wscript.scriptFullname
	wscript.echo "script path only " & scriptpathstr
	wscript.quit
Else
	wscript.echo "Found the ini file, reading it in."
End If

Set INIFileObj = fso.opentextfile(strINIPath)

Do While INIFileObj.AtEndOfStream = False
    strINILine = INIFileObj.readline
    INILineArray = Split(strINILine, "=")
    Select Case Trim(INILineArray(0))
        Case "Allfilename"
            wscript.echo "found Allfilename"
            Allfilename = Trim(INILineArray(1))
            wscript.echo INILineArray(1)
        Case "DatedFilename"
            wscript.echo "found DatedFilename"
            DatedFilename = Trim(INILineArray(1)) & dtDay & ".txt"
            wscript.echo INILineArray(1)
        Case "CNfilename"
            wscript.echo "found CNfilename"
            CNfilename = Trim(INILineArray(1))
            wscript.echo INILineArray(1)
        Case "PSSfilename"
            wscript.echo "found PSSfilename"
            PSSfilename = Trim(INILineArray(1))
            wscript.echo INILineArray(1)
        Case "strConfigSrv"
            wscript.echo "found strConfigSrv"
            strConfigSrv = Trim(INILineArray(1))
            wscript.echo INILineArray(1)
        Case "strConfigDB"
            wscript.echo "found strConfigDB"
            strConfigDB = Trim(INILineArray(1))
            wscript.echo INILineArray(1)
	Case "AllFileOldName"
            wscript.echo "found AllFileOldName"
            AllFileOldName = Trim(INILineArray(1))
            wscript.echo INILineArray(1)	
        Case Else
            wscript.echo "Don't know what to do with *" & strINILine & "*"
    End Select
Loop

ReDim CircuitList(1), Network(1), DeviceID(1), TailIntID(1), TargetID(1), HubIntID(1), Circuit(1), HubSite(1), IntTypeID(1)
ReDim CCS(1), BW(1), PortSpeed(1), Yes(1), DeviceName(1), IntHWDesc(1), FuncDesc(1), HubYN(1), MIB(1)
ReDim KbpsMed(1), Kbps95th(1), KbpsMedGr(1), Kbps95thGr(1), ms30DMin(1), ms30D95th(1), msMedGr(1), ms95thGr(1)
ReDim SiteDevice(1), Avail30(1), HrsOut30(1), AvgLoss30(1), Avail7(1), HrsOut7(1), AvgLoss7(1)
ReDim HC_GA(1), HC_PD(1), HC_SM(1), HC_Oper(1), HC_Total(1), BWperHC(1)

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")
cn.Provider = "sqloledb"
cn.Properties("Data Source").Value = strConfigSrv
'cn.CursorLocation = adUseClient
cn.Properties("Initial Catalog").Value = strConfigDB
'cn.Provider = "Microsoft.Jet.OLEDB.4.0"
'cn.Properties("Data Source").Value = "C:\NetID2000.mdb"
'cn.Properties("User ID").Value = "admin"
'cn.Properties("Password").Value = ""
cn.Properties("Integrated Security").Value = "SSPI"
wscript.echo "Connection properties set about to open connection and record set"
cn.Open
rs.Open "select * from JoinedTables order by Circuit, MIB", cn ', adOpenStatic, adLockReadOnly
wscript.echo "Connection and recordset opened"

i = 0
    With rs
        '.Sort = "Circuit, MIB"
        While Not .EOF
        CircuitList(i) = .Fields("CircuitList")
        Network(i) = .Fields("NetworkID")
        DeviceID(i) = .Fields("DeviceID")
        TailIntID(i) = .Fields("InterfaceID")
        TargetID(i) = .Fields("TargetID")
        HubIntID(i) = .Fields("HubInterfaceID")
        Circuit(i) = .Fields("Circuit")
        HubSite(i) = .Fields("HubSite")
        IntTypeID(i) = .Fields("InterfaceTypeID")
        CCS(i) = .Fields("CountryCitySite")
        BW(i) = .Fields("Bandwidth")
        PortSpeed(i) = .Fields("PortSpeed")
        Yes(i) = .Fields("TrafficFlowEnabled")
        DeviceName(i) = .Fields("DeviceName")
        IntHWDesc(i) = .Fields("InterfaceHWDescription")
        FuncDesc(i) = .Fields("FunctionalDescription")
        HubYN(i) = .Fields("MultipleSubinterfaces")
        MIB(i) = .Fields("MIB")
        KbpsMed(i) = .Fields("iMedbits")
        Kbps95th(i) = .Fields("i95th")
        KbpsMedGr(i) = .Fields("MedKbpsG90D")
        Kbps95thGr(i) = .Fields("f95thKbpsG90D")
        ms30DMin(i) = .Fields("i30DMin")
        ms30D95th(i) = .Fields("i30D95th")
        msMedGr(i) = .Fields("MedmsG90D")
        ms95thGr(i) = .Fields("i95thmsG90D")
        SiteDevice(i) = .Fields("SiteDevice")
        Avail30(i) = .Fields("Availability")
        HrsOut30(i) = .Fields("HoursOut")
        AvgLoss30(i) = .Fields("AvgLoss")
        
        i = i + 1
        
        ReDim Preserve CircuitList(i), Network(i), DeviceID(i), TailIntID(i), TargetID(i), HubIntID(i), Circuit(i), HubSite(i), IntTypeID(i)
        ReDim Preserve CCS(i), BW(i), PortSpeed(i), Yes(i), DeviceName(i), IntHWDesc(i), FuncDesc(i), HubYN(i), MIB(i)
        ReDim Preserve KbpsMed(i), Kbps95th(i), KbpsMedGr(i), Kbps95thGr(i), ms30DMin(i), ms30D95th(i), msMedGr(i), ms95thGr(i)
        ReDim Preserve SiteDevice(i), Avail30(i), HrsOut30(i), AvgLoss30(i)

        .MoveNext
    Wend
    End With
    
    n = UBound(Network)
        
    IndexUtil = UtilIndex(KbpsMed, Kbps95th, BW)
    IndexMedKbGr = GrowthIndex(KbpsMedGr)
    Index95thKbGr = GrowthIndex(Kbps95thGr)
    IndexmsMedGr = GrowthIndex(msMedGr)
    Indexms95thGr = GrowthIndex(ms95thGr)
    IndexQD = QdelayIndex(ms30DMin, ms30D95th)
    IndexLoss = LossIndex(Avail30)
    
    NegSum = SumOfNeg(IndexUtil, IndexMedKbGr, Index95thKbGr, IndexmsMedGr, Indexms95thGr, IndexQD, IndexLoss)
    MinValues = MinOfIndex(IndexUtil, IndexMedKbGr, Index95thKbGr, IndexmsMedGr, Indexms95thGr, IndexQD, IndexLoss, NegSum)
    FinalRank = ShiftAroundMedian(MinValues)
    TotalSort = SortAllArrays(CircuitList, Network, DeviceID, HubIntID, TailIntID, HubYN, Circuit, HubSite, CCS, DeviceName, IntHWDesc, IntTypeID, BW, PortSpeed, FinalRank, IndexUtil, IndexMedKbGr, Index95thKbGr, IndexmsMedGr, Indexms95thGr, IndexQD, IndexLoss, NegSum, Yes)
    
    AllList = TotalSort
    CorpNetList = CorpNetSorted(TotalSort)
    PSSONList = PSSONSorted(TotalSort)
     
    'ask about saving the current list as the old list.  The old list is used
    'to create the "Circuits getting worse" list.
    strMsg = "Would you like to save the current Allperf.txt as Allperf " & dtDay & ".txt?  Click Yes if you are creating reports.  Click No if you are just running an update of the Perf Index."
    If IsNull(bSaveOld) Then
    	iAnswer = MsgBox(strMsg, vbYesNo + vbQuestion, "Save Current?")
    Else
    	If bSaveOld Then
    		iAnswer = vbYes
    	Else
    		iAnswer = vbNo
    	End If
    End If
    If iAnswer = vbYes Then
        fso.Copyfile Allfilename, DatedFilename
        fso.Copyfile DatedFilename, AllFileOldName
    End If
    
    Set file4 = fso.createtextfile(Allfilename, True, False)
    Set file1 = fso.createtextfile(CNfilename, True, False)
    Set file2 = fso.createtextfile(PSSfilename, True, False)
    Set file3 = fso.createtextfile(DatedFilename, True, False)
    
    file4.writeline "Rank" & vbTab & "Network" & vbTab & "Device ID" & vbTab & "Hub Interface ID" & vbTab & "Interface ID" & vbTab & "Multiple Subinterfaces" & vbTab & "Circuit" & vbTab & "Hub Site" & vbTab & "Country/City/Site" & vbTab & "Device Name" & vbTab & "Interface HW Description" & vbTab & "Interface Type ID" & vbTab & "Bandwidth/SCR/CIR" & vbTab & "Port Speed" & vbTab & "Rank Values(Performance Index)" & vbTab & "Index Util in" & vbTab & "Index Util out" & vbTab & "Index Med Kb G in" & vbTab & "Index Med kb G out" & vbTab & "Index 95th Kb G in" & vbTab & "Index 95th Kb G out" & vbTab & "Index Med ms" & vbTab & "Index 95th ms" & vbTab & "Index Queue Delay" & vbTab & "Index Loss" & vbTab & "Perf Index" & vbTab & "TrafficFlow Enabled"
    file3.writeline "Rank" & vbTab & "Network" & vbTab & "Device ID" & vbTab & "Hub Interface ID" & vbTab & "Interface ID" & vbTab & "Multiple Subinterfaces" & vbTab & "Circuit" & vbTab & "Hub Site" & vbTab & "Country/City/Site" & vbTab & "Device Name" & vbTab & "Interface HW Description" & vbTab & "Interface Type ID" & vbTab & "Bandwidth/SCR/CIR" & vbTab & "Port Speed" & vbTab & "Rank Values(Performance Index)" & vbTab & "Index Util in" & vbTab & "Index Util out" & vbTab & "Index Med Kb G in" & vbTab & "Index Med kb G out" & vbTab & "Index 95th Kb G in" & vbTab & "Index 95th Kb G out" & vbTab & "Index Med ms" & vbTab & "Index 95th ms" & vbTab & "Index Queue Delay" & vbTab & "Index Loss" & vbTab & "Perf Index" & vbTab & "TrafficFlow Enabled"
    For j = 0 To UBound(AllList)
        file4.writeline j + 1 & vbTab & AllList(j, 2) & vbTab & AllList(j, 3) & vbTab & AllList(j, 4) & vbTab & AllList(j, 5) & vbTab & AllList(j, 6) & vbTab & AllList(j, 7) & vbTab & AllList(j, 8) & vbTab & AllList(j, 9) & vbTab & AllList(j, 10) & vbTab & AllList(j, 11) & vbTab & AllList(j, 12) & vbTab & AllList(j, 13) & vbTab & AllList(j, 14) & vbTab & AllList(j, 15) & vbTab & AllList(j, 16) & vbTab & AllList(j, 17) & vbTab & AllList(j, 18) & vbTab & AllList(j, 19) & vbTab & AllList(j, 20) & vbTab & AllList(j, 21) & vbTab & AllList(j, 22) & vbTab & AllList(j, 23) & vbTab & AllList(j, 24) & vbTab & AllList(j, 25) & vbTab & AllList(j, 26) & vbTab & AllList(j, 27)
        file3.writeline j + 1 & vbTab & AllList(j, 2) & vbTab & AllList(j, 3) & vbTab & AllList(j, 4) & vbTab & AllList(j, 5) & vbTab & AllList(j, 6) & vbTab & AllList(j, 7) & vbTab & AllList(j, 8) & vbTab & AllList(j, 9) & vbTab & AllList(j, 10) & vbTab & AllList(j, 11) & vbTab & AllList(j, 12) & vbTab & AllList(j, 13) & vbTab & AllList(j, 14) & vbTab & AllList(j, 15) & vbTab & AllList(j, 16) & vbTab & AllList(j, 17) & vbTab & AllList(j, 18) & vbTab & AllList(j, 19) & vbTab & AllList(j, 20) & vbTab & AllList(j, 21) & vbTab & AllList(j, 22) & vbTab & AllList(j, 23) & vbTab & AllList(j, 24) & vbTab & AllList(j, 25) & vbTab & AllList(j, 26) & vbTab & AllList(j, 27)
    Next
    For j = 0 To UBound(CorpNetList)
        If CorpNetList(j, 1) <> "" Then
            file1.writeline CorpNetList(j, 1) & vbTab & CorpNetList(j, 2) & vbTab & CorpNetList(j, 3) & vbTab & CorpNetList(j, 4) & vbTab & CorpNetList(j, 5) & vbTab & CorpNetList(j, 6) & vbTab & CorpNetList(j, 7) & vbTab & CorpNetList(j, 8) & vbTab & CorpNetList(j, 9) & vbTab & CorpNetList(j, 10) & vbTab & CorpNetList(j, 11) & vbTab & CorpNetList(j, 12) & vbTab & CorpNetList(j, 13) & vbTab & CorpNetList(j, 14) & vbTab & CorpNetList(j, 15) & vbTab & CorpNetList(j, 16) & vbTab & CorpNetList(j, 17) & vbTab & CorpNetList(j, 18) & vbTab & CorpNetList(j, 19) & vbTab & CorpNetList(j, 20) & vbTab & CorpNetList(j, 21) & vbTab & CorpNetList(j, 22) & vbTab & CorpNetList(j, 23) & vbTab & CorpNetList(j, 24) & vbTab & CorpNetList(j, 25) & vbTab & CorpNetList(j, 26) & vbTab & CorpNetList(j, 27)
        End If
    Next
    For j = 0 To UBound(PSSONList)
        If PSSONList(j, 1) <> "" Then
            file2.writeline PSSONList(j, 1) & vbTab & PSSONList(j, 2) & vbTab & PSSONList(j, 3) & vbTab & PSSONList(j, 4) & vbTab & PSSONList(j, 5) & vbTab & PSSONList(j, 6) & vbTab & PSSONList(j, 7) & vbTab & PSSONList(j, 8) & vbTab & PSSONList(j, 9) & vbTab & PSSONList(j, 10) & vbTab & PSSONList(j, 11) & vbTab & PSSONList(j, 12) & vbTab & PSSONList(j, 13) & vbTab & PSSONList(j, 14) & vbTab & PSSONList(j, 15) & vbTab & PSSONList(j, 16) & vbTab & PSSONList(j, 17) & vbTab & PSSONList(j, 18) & vbTab & PSSONList(j, 19) & vbTab & PSSONList(j, 20) & vbTab & PSSONList(j, 21) & vbTab & PSSONList(j, 22) & vbTab & PSSONList(j, 23) & vbTab & PSSONList(j, 24) & vbTab & PSSONList(j, 25) & vbTab & PSSONList(j, 26) & vbTab & PSSONList(j, 27)
        End If
    Next
    If IsNull (bSaveOld) Then
    	MsgBox CNfilename & " and " & PSSfilename & " created successfully.", vbOKOnly + vbInformation, "Operation Successfull!!"
    End If
    'CNmerged = "notepad " & CNfilename
    'PSSmerged = "notepad " & PSSfilename
    'CNtaskID = Shell(CNmerged) ', vbNormalFocus)
    'PSStaskID = Shell(PSSmerged)', vbNormalFocus)
    rs.close
    cn.close
    Set rs = nothing
    Set cn = nothing
    Set inifileobj=nothing
    Set file1 = nothing
    Set file2 = nothing
    Set file3 = nothing
    Set file4 = nothing
    wscript.echo "Connection and recordset closed"
    wscript.echo "Objects created have been cleaned up. Script complete"
    
End Sub

Main