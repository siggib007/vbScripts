Option Explicit
dim app, f1, strFullFileName, doc, sel, strSafeDate, strSafeTime

on error goto 0
const wdFindContinue = 1
const wdMatchAnyCharacter = 65599
const wdMatchAnyDigit = 65567
const wdMatchAnyLetter = 65583
const wdMatchCaretCharacter = 11
const wdMatchColumnBreak = 14
const wdMatchCommentMark = 5
const wdMatchEmDash = 8212
const wdMatchEnDash = 8211
const wdMatchEndnoteMark = 65555
const wdMatchField = 19
const wdMatchFootnoteMark = 65554
const wdMatchGraphic = 1
const wdMatchManualLineBreak = 65551
const wdMatchManualPageBreak = 65564
const wdMatchNonbreakingHyphen = 30
const wdMatchNonbreakingSpace = 160
const wdMatchOptionalHyphen = 31
const wdMatchParagraphMark = 65551
const wdMatchSectionBreak = 65580
const wdMatchTabCharacter = 9
const wdMatchWhiteSpace = 65655
const wdReplaceAll = 2
'const wdFindContinue = 1

on error resume next
Set app = CreateObject("Word.Application")
If Err.Number <> 0 Then
	WriteLog "Unable to start Word, probably not installed correctly."
	wscript.quit
end if
on error goto 0
f1 = "C:\Users\sbjarna\Documents\Projects\LTE\2015Q1\TemplateMMECorenet.docx"
strSafeDate = DatePart("yyyy",now) & Right("0" & DatePart("m",now), 2) & Right("0" & DatePart("d",now), 2)
strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)

wscript.echo "StrSaveDate: " & strSafeDate
strFullFileName = "C:\Users\sbjarna\Documents\Projects\LTE\2015Q1\Testing-" & strSafeDate & strSafeTime & ".docx"
wscript.echo "SaveAs name: " & strFullFileName
app.visible = true

on error goto 0
Set doc = app.Documents.Open (f1,0,true)
set sel = app.selection

sel.find.wrap = wdFindContinue
sel.Find.text = "$SB1$"
sel.Find.Replacement.Text = "SRGBLF01"
sel.Find.Execute ,,,,,,,,,,wdReplaceAll

sel.find.wrap = wdFindContinue
sel.Find.text = "$SB2$"
sel.Find.Replacement.Text = "SRBBLF02"
sel.Find.Execute ,,,,,,,,,,wdReplaceAll

sel.find.wrap = wdFindContinue
sel.Find.text = "$Baseline1$"
while sel.Find.execute
	sel.InsertFile "C:\Users\sbjarna\Documents\Projects\LTE\test\Configurations\WO2254840-ATMME002-SRGNRC01-Baseline.txt"
wend

sel.find.wrap = wdFindContinue
sel.Find.text = "$Baseline2$"
while sel.Find.execute
	sel.InsertFile "C:\Users\sbjarna\Documents\Projects\LTE\test\Configurations\WO2254840-ATMME002-SRGNRC02-Baseline.txt"
wend

sel.find.wrap = wdFindContinue
sel.Find.text = "$Implementation1$"
while sel.Find.execute
	sel.InsertFile "C:\Users\sbjarna\Documents\Projects\LTE\test\Configurations\WO2254840-ATMME002-SRGNRC01-Implementation.txt"
wend

sel.find.wrap = wdFindContinue
sel.Find.text = "$Implementation2$"
while sel.Find.execute
	sel.InsertFile "C:\Users\sbjarna\Documents\Projects\LTE\test\Configurations\WO2254840-ATMME002-SRGNRC02-Implementation.txt"
wend

wscript.echo "Now saving to " & strFullFileName
doc.SaveAs strFullFileName
