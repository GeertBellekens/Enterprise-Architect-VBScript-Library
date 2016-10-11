'[path=\Projects\Project AC]
'[group=Acerta Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: Geert Bellekens
' Purpose: Sets all datatypes of columns to uppercase
' Date: 2016-1-10
'
sub main
	dim sqlUpdate
	sqlUpdate = "update a set a.type = upper(a.type) from t_attribute a where a.Stereotype = 'column'"
	Repository.Execute sqlUpdate
end sub

main