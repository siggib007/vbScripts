set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")
strFavorite = WshShell.SpecialFolders("Favorites")
strDocuments = WshShell.SpecialFolders("MyDocuments")

wscript.echo "Path to Desktop is: " & strDestop
wscript.echo "Path to Favorites is: " & strFavorite
wscript.echo "Path to Documents is: " & strDocuments