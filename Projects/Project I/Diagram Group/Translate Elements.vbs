'[path=\Projects\Project I\Diagram Group]
'[group=Diagram Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

' Script Name: Translate Elements
' Author: Geert Bellekens
' Purpose: Translate the selected elements on the diagram (or all elements if nothing is selected)
' Date: 2025-08-27

const outPutName = "Translate Elements"


sub main
	'reset output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName

	dim diagram as EA.Diagram
	set diagram = Repository.GetCurrentDiagram
	if diagram is nothing then
		exit sub
	end if
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName & " for '"& diagram.Name &"'", 0
	'do the actual work
	translateDiagram diagram, "", false
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName & " for '"& diagram.Name &"'", 0
	
end sub

main

'function test
'	'reset output tab
'	Repository.CreateOutputTab outPutName
'	Repository.ClearOutput outPutName
'	Repository.EnsureOutputVisible outPutName
'
'	dim diagram as EA.Diagram
'	set diagram = Repository.GetDiagramByGuid("{244AFB03-094D-4e0f-9CA6-12F9D70EAF64}")
'	'set timestamp
'	Repository.WriteOutput outPutName, now() & " Starting " & outPutName & " for '"& diagram.Name &"'", 0
'	'do the actual work
'	translateDiagram diagram
'	'set timestamp
'	Repository.WriteOutput outPutName, now() & " Finished " & outPutName & " for '"& diagram.Name &"'", 0
'
'end function
'
'test