Option Explicit
Dim rs, rs1, rs2 'As ADODB.Recordset
Dim fld, Flds, Flds1 'As ADODB.Fields
Dim cn 'As ADODB.Connection
Dim cmdText1, cmdText2, cmdtext3, strOutput 'As String
Dim X , iFieldCount'as integer

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")
Set rs1 = CreateObject("ADODB.Recordset")
Set rs2 = CreateObject("ADODB.Recordset")

cn.Provider = "sqloledb"
cn.Properties("Data Source").Value = "."
cn.Properties("Initial Catalog").Value = "pubs"
cn.Properties("Integrated Security").Value = "SSPI"
cn.Open

cmdText1 = "select table_name from information_schema.tables where table_type='base table'"
wscript.echo
wscript.echo "Start Search Test at " & Now
rs.Open cmdText1, cn
If rs.EOF Then MsgBox "No tables found"
Do Until rs.EOF
    'wscript.echo
    'wscript.echo "Examining table named " & rs.Fields(0).Value
    cmdText2 = "select * from " & rs.Fields(0).Value & " where 1=0"
    rs1.Open cmdText2, cn
    'Set Flds1 = rs1.Fields
    For Each fld In rs1.Fields
        'field= fld.Name
         'wscript.echo fld.Name
         'wscript.echo fld.Type
         If fld.Type <> 205 And fld.Type <> 6 Then ' Exclude binary and dollar value
            'wscript.echo "select * from " & rs.Fields(0).Value & _
            '           " where " & fld.Name & " like '%17%'" ', cn
            rs2.Open "select * from " & rs.Fields(0).Value & _
                        " where " & fld.Name & " like '%17%'", cn
            If Not rs2.BOF Then
                wscript.echo "Table " & rs.Fields(0).Value
                wscript.echo "Field " & fld.Name
                iFieldCount = rs2.fields.count
                wscript.echo "Total Number of fields " & iFieldCount
                For x = 0 to iFieldCount - 1
                	strOutput = strOutput & ", " & rs2.fields(x).name
                Next 
                wscript.echo strOutput
                Do until rs2.eof
                	For x = 0 to iFieldCount - 1
                		strOutput = strOutput & ", " & rs2.fields(x).value
                		'wscript.echo rs2.fields(x).name
                		'wscript.echo rs2.fields(x).value
                	Next 
               		wscript.echo strOutput
               		rs2.movenext
                loop
            End If
             'wscript.echo fld.Value
             rs2.Close
        End If
    Next

        'cmdtext2 = "select * from " & fld.Value & " where au_lname like '%17%'"
        'wscript.echo cmdtext2
    rs1.Close
    'Set rs1 = Nothing
    rs.MoveNext
Loop
rs.Close
