Option Explicit
Dim FileObj, fso, strOutFileName, objFileOut, HTTP, objLinks, objLogOut, strLogFileName, strInDirectory
Dim folder, files, file, strLine, strLineParts, bFoundURL, objRegExpr, Matches, Match, strURL

set objLinks = CreateObject("Scripting.Dictionary")
Set objRegExpr = new regexp
objRegExpr.Pattern = "HREF=""(.*)"" "

strLogFileName = "C:\Logs\FavoritesValidation.log"
strOutFileName = "C:\BrowswerFavorites\ConsolidatedFavorites.html"
strInDirectory = "C:\BrowswerFavorites\"

Set fso = CreateObject("Scripting.FileSystemObject")
Set objLogOut = fso.createtextfile(strLogFileName)
Set objFileOut = fso.createtextfile(strOutFileName)

Set folder = fso.GetFolder(strInDirectory)
Set files = folder.Files

For Each file in files
	If file.name <> "desktop.ini" and file.path <> strOutFileName then
		bFoundURL = false
		writeout "Opening " & file.path
		Set FileObj = fso.opentextfile(file.path)
		writeout file.path
		While not fileobj.atendofstream
			Set HTTP = CreateObject("Microsoft.XMLHTTP")
			strLine = FileObj.readline
			strLineParts = split (Trim(strLine), " ")
			If ubound(strLineParts) > 0 then 
				If instr(1,strLineParts(1),"HREF=",1) > 0 then	
					If instr(1,strLineParts(1),"HTTP",1) > 0 then	
						strURL = mid(strLineParts(1),7,len(strLineParts(1))-7)
						If not objLinks.Exists(strURL) then 
							objLinks.Add strURL, strURL
							writeout "URL: " & strURL
							On Error Resume Next
							writeout "preparing to test URL"
							HTTP.Open "GET", strURL, False
							writeout "testing URL"
							HTTP.Send
							writeout "evaluating test"
							If Err.Number > 0 then
								writeout strURL & "; " & Err.Number & Err.Description
							Else
								writeout strURL & "; " & HTTP.statusText
							End If
							On Error Goto 0
							If HTTP.statusText = "OK" Then
								writeout "copy line for " & strURL & " to " & strOutFileName
								objFileOut.writeline strline
							End If
						Else
							writeout "Dup: " & strURL
						End If
					End If
				Else
					objFileOut.writeline strline
				End If
			Else
				objFileOut.writeline strline
			End If
			Set HTTP = nothing
		Wend
	End If
Next

Sub writeout (msg)

	wscript.echo msg
	objLogOut.writeline msg

End Sub
