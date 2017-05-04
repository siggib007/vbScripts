Option Explicit

dim myarr()

redim preserve myarr(2)
myarr(1) = "siggi"
myarr(2) = "geek"

wscript.echo myarr(1) & " " & myarr(2)

redim preserve myarr(3)
myarr(3) ="test"
wscript.echo myarr(3)

wscript.echo myarr(1) & " " & myarr(2) & "super"
