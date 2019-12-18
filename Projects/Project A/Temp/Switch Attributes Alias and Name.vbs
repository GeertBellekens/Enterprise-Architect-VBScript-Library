'[path=\Projects\Project A\Temp]
'[group=Temp]

option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Switch Attributes Alias and Name
' Author: Geert Bellekens
' Purpose: Switch the alias with the name for each attribute of the selected element
' Date: 
'
sub main
	dim selecteElement as EA.Element
	set selectedElement = Repository.GetTreeSelectedObject()
	'check if element selected
	if selectedElement.ObjectType <> otElement then
		MsgBox "Please select an element"
		exit sub
	end if
	dim attribute as EA.Attribute
	for each attribute in selectedElement.Attributes
		'switch attribute name and alias
		dim temp
		temp = attribute.Alias
		attribute.Alias = attribute.Name
		attribute.Name = temp
		attribute.Update
	next
	MsgBox "finished"
end sub

main