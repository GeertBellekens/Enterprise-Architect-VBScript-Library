option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
	dim sqlUpdate
	sqlUpdate = "update t_attribute set upperbound = '*' " &_
			  "where upperbound not in ('1', '*') "
	Repository.Execute sqlUpdate
end sub

main