Option Explicit 
Dim strLine, fso, strParts, ListFileObj, inFileObj, inFileName, FileCont, ListFileName, VLANoutFileObj, outFileObj

Const DefFolder = "x:\brix\"
Const VlanIDFileName = "sla_id.csv"

Set fso = CreateObject("Scripting.FileSystemObject")
ListFileName = DefFolder & VlanIDFileName
wscript.echo "Renaming files according to " & listfilename
Set ListFileObj = fso.opentextfile(listfilename)
Set VLANoutFileObj = fso.CreateTextFile(deffolder & "VLANAvailability.csv")
VLANoutFileObj.writeline "DateOfData,AppName,Datacenter,Devicename,DataName,DataValue,Summary"
While not listfileobj.atendofstream
	strLine = Trim(listFileObj.readline)
	If strline <> "" Then
		strparts = split(strline,",")
		'wscript.echo "UBound(strparts): " & UBound(strparts)
		If UBound(strparts) = 1 Then
			inFileName = DefFolder & strparts(0) & ".csv"
			If fso.FileExists(inFileName) Then
				Set inFileObj = fso.opentextfile(inFileName)
				FileCont = infileObj.readall
				infileobj.close
				FileCont = replace(filecont,"##",vbcrlf)
				Set outFileObj = fso.CreateTextFile(deffolder & strparts(1) & ".csv")
				outFileObj.write filecont
				outfileobj.close
				If InStr(strparts(1),"VLANAvailability") > 0 Then
					vlanoutfileobj.write filecont
				End If 
			End If
		End If
	End If 
Wend 
listfileobj.close
vlanoutfileobj.close
Set vlanoutfileobj = nothing
Set outfileObj = nothing
Set infileobj = nothing
Set listfileobj = nothing
Set fso = nothing

wscript.echo "Done"