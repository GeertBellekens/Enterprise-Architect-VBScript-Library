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
	sqlUpdate =	"update t_attribute set Style = null " &_
				" where Style is not null " &_
				" and stereotype <> 'enum' "
	Repository.Execute sqlUpdate
end sub

main