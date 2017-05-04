Option Explicit

Dim myDict, OutKey, InnKey, oKey, ikey
Set myDict = CreateObject("Scripting.Dictionary")

myDict.Add "OMW" , CreateObject("Scripting.Dictionary")
myDict.Add "Gi" , CreateObject("Scripting.Dictionary")

myDict("OMW").Add "sgb", "Siggi"
myDict("OMW").Add "mk", "Mary"

myDict("Gi").Add "JS", "Joe"
myDict("Gi").Add "JD", "John"

OutKey = myDict.keys
for each oKey in OutKey
	InnKey = myDict(oKey).keys
	for each ikey in InnKey
		wscript.echo oKey & "-" & ikey & "-" & myDict(oKey).item(ikey)
	next
next

