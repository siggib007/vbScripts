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

wscript.echo "now testing showkeys"
wscript.echo ShowKeys

Function ShowKeys
   Dim a, d, i, s   ' Create some variables.
   Set d = CreateObject("Scripting.Dictionary")
   d.Add "a", "Athens"   ' Add some keys and items.
   d.Add "b", "Belgrade"
   d.Add "c", "Cairo"
   a = d.Items
   For i = 0 To d.Count -1 ' Iterate the array.
      s = s & a(i) & "<BR>" ' Create return string.
   Next
   ShowKeys = s
End Function