Option Explicit
Dim FileObj, fso, strOutFileName, objFileOut, HTTP, strDeadLinkFolder, strGoodLinkFolder
Dim WshShell, strFavorite, strDocuments, subfolder, subfiles, subFlds2, fld2, objLinks, strStripped

set WshShell = WScript.CreateObject("WScript.Shell")
strFavorite = WshShell.SpecialFolders("Favorites")
strDocuments = WshShell.SpecialFolders("MyDocuments")
set objLinks = CreateObject("Scripting.Dictionary")


strOutFileName = "FavoritesValidation.log"

strDeadLinkFolder = "c:\DeadLinks"
strGoodLinkFolder = "C:\ValidFavorites"

Set fso = CreateObject("Scripting.FileSystemObject")
Set objFileOut = fso.createtextfile(strDocuments & "\" & strOutFileName)
Set HTTP = CreateObject("Microsoft.XMLHTTP")

FolderContent strFavorite
	
For Each strLink in objLinks
	wscript.echo objLinks.item(strLink) & "; " & strLink
Next

Sub FolderContent (strCurrentFolder)
	Dim folder, files, file, subFlds, fld, strPath, strLine, strLineParts, bFoundURL

	Set folder = fso.GetFolder(strCurrentFolder)
	strPath = mid(folder.path, len(strFavorite)+2)
	if strPath <> "" then strPath = strPath & "\"
	Set files = folder.Files
	Set subFlds = folder.SubFolders
	
	For Each file in files
		if file.name <> "desktop.ini" then
			bFoundURL = false
'			wscript.echo "Opening " & file.path '& "\" & file.name
			Set FileObj = fso.opentextfile(file.path)
'			writeout strpath & file.name
			While not fileobj.atendofstream
				strLine = Trim(FileObj.readline)
				strLineParts = split (strLine, "=")
				If strLineParts(0) = "URL" Then
'					wscript.echo "testing URL " & strLineParts(1)
					If not objLinks.Exists(strLineParts(1)) then 
						objLinks.Add strLineParts(1), strpath & file.name
						on error resume next
						HTTP.Open "GET", strLineParts(1), False
						HTTP.Send
						if Err.Number > 0 then
							writeout strpath & file.name & "; " & strLineParts(1) & "; " & Err.Number & Err.Description
						else
							writeout strpath & file.name & "; " & strLineParts(1) & "; " & HTTP.statusText
						end if
						on error goto 0
						If HTTP.statusText = "OK" Then
							strStripped = Replace (strpath, " ", "", 1, -1, vbTextCompare)
							strStripped = Replace (strpath, ",", "", 1, -1, vbTextCompare)
							writeout strLineParts(1) & " is good, copying to " & strGoodLinkFolder & "\" & strStripped & file.name
							If Not fso.FolderExists(strGoodLinkFolder & "\" & strStripped) Then
								fso.CreateFolder(strGoodLinkFolder & "\" & strStripped)
							End If
							file.copy (strGoodLinkFolder & "\" & strStripped & file.name)
						else
							writeout strLineParts(1) & " is not OK" ', moving to " & strDeadLinkFolder & "\" & file.name
'							file.copy (strDeadLinkFolder & "\" & file.name)
'							file.move (strDeadLinkFolder & "\" & file.name)
'							file.delete (true)
'							fso.movefile file.path, "c:\" & strDeadLinkFolder & "\" & file.name
						End If
						bFoundURL = true
					End If
				End If 
			Wend
			if bFoundURL = false Then
				writeout "Couldn't find URL in " & file.path & "\" & file.name
			End If
		End if
	Next
	For Each fld in subFlds
		FolderContent fld
	Next
End Sub

Sub writeout (msg)
	
	wscript.echo msg
	objFileOut.writeline msg

End Sub
