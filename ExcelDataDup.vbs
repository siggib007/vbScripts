Option Explicit 

  Dim app, wb, ws, wbName, wbin, wsin, wbOutName
  Dim row, NameParts, x
  Dim objShell
  Const xlHAlignCenter  = -4108
  Const xlVAlignBottom  = -4107
  Const xlVAlignCenter = -4108
  Const xlVAlignTop = -4160
  
  wbName = "C:\temp\testfile.xlsx"
  wscript.echo "Input file: " & wbName
  NameParts = split(wbName,"\")
  NameParts(UBound(NameParts)) = "Copy of " & NameParts(UBound(NameParts))
  wbOutName = Join(NameParts,"\")
  wscript.echo "output file: " & wbOutName
  Set objShell = CreateObject("WScript.Shell")
  Set app = CreateObject("Excel.Application")
  Set wb = app.Workbooks.Add
  Set ws = wb.Worksheets(1)
  Set wbin = app.Workbooks.Open (wbName)
  Set wsin = wbin.Worksheets(1)
  Set objShell = CreateObject("WScript.Shell")

  app.visible = True
  
  row = 1
  
  Do
  	If wsin.Cells(row,1).Value = "" Then Exit Do 
  	ws.Cells(row, 1) = wsin.Cells(row,1)
  	ws.Cells(row, 2) = wsin.Cells(row,2)
  	ws.Cells(row, 3) = wsin.Cells(row,3)
  	ws.Cells(row, 4) = wsin.Cells(row,4)
  	ws.Cells(row, 5) = wsin.Cells(row,5)
  	ws.Cells(row, 6) = wsin.Cells(row,6)
  	ws.Cells(row, 7) = wsin.Cells(row,7)
  	row = row + 1
  loop
  
 ' ws.select
'  wb.SaveAs(wbOutName)
'  wb.Close
'  wbin.Close
'  app.Quit

  Set ws = Nothing 
  Set wb = Nothing
  Set wsin = Nothing 
  Set wbin = Nothing  
  Set app = Nothing 
  
'  objShell.Run(wbOutName)
