' Option Explicit
Class employeeclass
    Public first, last, salary
End Class
Dim employeedict, employee, testDict
Set employeedict = CreateObject("Scripting.Dictionary")
Set testDict = CreateObject("Scripting.Dictionary")

Set employee = new employeeclass
With employee
    .first = "John"
    .last = "Doe"
    .salary = 150000
End With
employeedict.Add "1", employee

Set employee = new employeeclass
With employee
    .first = "Mary"
    .last = "Jane"
    .salary = 50000
End With
employeedict.Add "3", employee

empID = employeedict.keys
for each emp in empID
	wscript.echo "ID:" & emp & " - " & employeedict.item(emp).first & " " & employeedict.item(emp).last
	wscript.echo employeedict.item(emp).salary
next

wscript.echo "ID: 3;" & employeedict.item("3").first & " " & employeedict.item("3").last & " salary: " & employeedict.item("3").salary

employeedict.item("3").last = employeedict.item("3").last & "-Joe"

wscript.echo "ID:3; " & employeedict.item("3").first & " " & employeedict.item("3").last & " salary: " & employeedict.item("3").salary

wscript.echo "testing empty dict"
test = testDict.keys
for each t in test
	wscript.echo "test:" & t & "-" & testDict.item(t)
next