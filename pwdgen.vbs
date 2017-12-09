Option Explicit
Dim ipwdlen, strpwd, iUBound, iLBound, iULDiff, x

'-------------------------------------------------------------------------------'
' Password generator                                                            '
' Written by Siggib on 5/31/05                                                  '
' Last Changed by Siggib at 01:33AM 5/31/05                                     '
' Usage: cscript pwdgen num                                                    '
'         where num is the number of characters you wish in your password.      '
'                                                                               '
' Generates a random password by repeatedly picking a random character out      '
' of the Ascii table. iUBound and iLBound variables specify the range of        '
' characters that are allowed. Allowing all printable characters will           '
' produce the strongest password. All printable characters have an ascii value  '
' between 33 and 127 inclusivly.                                                '
'-------------------------------------------------------------------------------'


'iUBound = 127 ' Upper bound of ascii table we want to use
'iLBound = 33  ' Lower bound of ascii table we want to use
'iLBound = 48  ' Lower bound of ascii table we want to use
iUBound = 90 ' Upper bound of ascii table we want to use
iLBound = 65  ' Lower bound of ascii table we want to use
iULDiff	= iUBound - iLBound + 1 ' The number of acceptable char
strpwd = ""

Randomize
If wscript.arguments.count > 0 Then
	ipwdlen = wscript.arguments(0)
Else
	wscript.echo "Please specify how many charactors you wish in your password"
	wscript.quit
End If

for x = 1 to ipwdlen
	strpwd = strpwd & chr(Int(iULDiff * Rnd + iLBound))
Next

wscript.echo "Your new password is: " & strpwd
wscript.echo
wscript.echo "Please verify that all characters used are valid in the intended environment"
wscript.echo "and complies with the appropriate password policy, if any."
wscript.echo "For example Cisco CLI passwords can not have ? in them."
wscript.echo "If the password contains undesired characters either remove the manually"
wscript.echo "or rerun this script for a different password."
