option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
	dim sqlDelete
	sqlDelete = "delete from [t_objectproblems]"
	dim response
	response = Msgbox("Alle issues en changes verwijderen?", _
        vbYesNo, "Maintenance opkuisen")

	If response = vbYes Then
		Repository.Execute sqlDelete
		msgbox "Klaar!"
	End If
end sub

main