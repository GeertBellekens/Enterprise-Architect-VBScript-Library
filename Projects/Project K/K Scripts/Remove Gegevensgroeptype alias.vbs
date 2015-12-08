option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
	dim sqlupdate
	sqlupdate = "update t_object set Alias = null " & _
				" where [Stereotype] = 'gegevensgroeptype' and ALIAS is not null"
	Repository.Execute sqlupdate
end sub

main