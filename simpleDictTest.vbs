Option Explicit

Dim myDict, OutKey, oKey
Set myDict = CreateObject("Scripting.Dictionary")

myDict.Add "sgb", "Siggi"
myDict.Add "mk", "Mary"
myDict.Add "JS", "Joe"
myDict.Add "JD", "John"

OutKey = myDict.keys
for each oKey in OutKey
	wscript.echo oKey & "-" & myDict.item(okey)
	if oKey = "mk" then
		myDict.item(oKey) = myDict.item(oKey) & " Sue"
		wscript.echo oKey & "-" & myDict.item(okey)
	end if
next

