' 
'
'  Name: 
'
' 
'
'    Detached.vbs 
'
' 
'
'  Description: 
'
' 
'
'    This script gives an example of how to create and use a detached recordset to sort a 
'    list of items. 
'
' 
'
'  Usage: 
'
' 
'
'    The text file to sort is specified on the command line. Output is to the console.       
' 
'
'  Audit: 
'
' 
'
'    2005-03-02  jdeg  original code 
'
' 
'

If WScript.Arguments.Count = 0 then
   Wscript.Echo "Detached sortfile"
   Wscript.Quit
end if

'get the file name from the command line and see if it exists

sortfile = WScript.Arguments(0)
set fso = CreateObject("Scripting.FileSystemObject")

if not fso.FileExists(sortfile) then
   Wscript.Echo "file",sortfile,"not found"
   Wscript.Quit
end if

'read the entire file and split into separate lines

file = fso.OpenTextFile(sortfile,1).ReadAll
file = Split(file,vbCrLf)

'create the recordset, define the data filed and open it for processing

set rec = CreateObject("ADODB.RecordSet")

rec.CursorLocation   = 3               '3 = adUseClient
rec.LockType         = 4               '4 = adLockBatchOptimistic
rec.CursorType       = 3               '3 = adOpenStatic
rec.ActiveConnection = Nothing
rec.Fields.Append "Field1",200,255     '200 = adVarChar
rec.Open

'add the lines into the recordset

for each line in file

   if len(line) > 0 then
      rec.AddNew
      rec.Fields(0) = line
      rec.Update
   end if

next

'sort it based on the only field

rec.Sort = "Field1"

'write out records in sorted order

rec.MoveFirst

Do While Not rec.EOF
   wscript.Echo rec.Fields(0)
   rec.MoveNext
Loop

rec.Close

