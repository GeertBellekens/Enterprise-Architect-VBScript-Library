'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
	if instr("redefineSomething", "redefine") > 0 then
		Session.Output "Found"
	else
		Session.Output "not found"
	end if
end sub

main