Option Explicit 
Dim strLine, fso, strParts, ListFileObj, inFileObj, inFileName, FileCont, ListFileName
Dim VLANoutFileObj, outFileObj, dictSLA_ID, OutFileName, f, fc, f1

Const DefFolder = "d:\siggib\brix\"
'Const DefFolder = "Z:\Brix\"
Const VlanIDFileName = "sla_id.csv"

Set fso = CreateObject("Scripting.FileSystemObject")
ListFileName = DefFolder & VlanIDFileName
'wscript.echo "Renaming files according to " & listfilename
Set ListFileObj = fso.opentextfile(listfilename)
Set VLANoutFileObj = fso.CreateTextFile(deffolder & "VLANAvailability.csv")
Set f = fso.GetFolder(defFolder)
Set fc = f.Files
Set dictSLA_ID = CreateObject("Scripting.Dictionary")
'select distinct sla_id, sla_name from sld
While not listfileobj.atendofstream
	strLine = Trim(listFileObj.readline)
	If strline <> "" Then
		strparts = split(strline,",")
		dictSLA_ID.add strparts(0), strparts(1)
	End If 
Wend 

wscript.echo "Done Reading in SLA ID & Names"
		

VLANoutFileObj.writeline "DateOfData,AppName,Datacenter,Devicename,DataName,DataValue,SLAName"
For Each f1 in fc
	wscript.echo "Now processing " & f1.name
	strparts = split(f1.name,".")
	If UBound(strparts) = 1 Then 
		If IsNumeric(strparts(0)) and strparts(1) = "csv" Then 
			Set inFileObj = fso.opentextfile(deffolder & f1.name)
			If not infileobj.atendofstream Then 
				FileCont = infileObj.readall
			Else
				FileCont = ""
			End If
			infileobj.close
			If dictSLA_ID.exists(strparts(0)) Then 
				OutFileName = dictSLA_ID(strparts(0))
			Else
				outFileName = "UnknownSLA" & strparts(0)
			End If 
			wscript.echo outfilename
			FileCont = replace(filecont,",NULL##", "," & outFileName & vbcrlf)
			FileCont = replace(filecont,"##",vbcrlf)
			Set outFileObj = fso.CreateTextFile(deffolder & outFileName & ".csv")
			outFileObj.write filecont
			outfileobj.close
			If InStr(outFileName,"VLANAvailability") > 0 or InStr(outfilename, "UnknownSLA") > 0 Then
				vlanoutfileobj.write filecont
			End If 
		End If 
	End If 
Next 

listfileobj.close
vlanoutfileobj.close
Set vlanoutfileobj = nothing
Set outfileObj = nothing
Set infileobj = nothing
Set listfileobj = nothing
Set fso = nothing

'wscript.echo "Done"