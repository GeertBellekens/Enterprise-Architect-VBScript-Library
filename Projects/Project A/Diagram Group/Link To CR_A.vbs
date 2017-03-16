'[path=\Projects\Project A\Diagram Group]
'[group=Diagram Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.Util
!INC Atrias Scripts.LinkToCRMain

'This script only calls the function defined in the main script.
'Ths script is to be copied in Diagram, Search and Project Browser groups

'Execute main function defined in LinkToCRMain
sub main
	dim diagram as EA.Diagram
	set diagram = Repository.GetCurrentDiagram
	if not diagram is nothing then
		dim selectedItems
		set selectedItems = getSelectedElements(diagram)
		if selectedItems.Count > 0 then
			linkItemToCR nothing, selectedItems
		end if
	end if
end sub

function getSelectedElements(diagram)
	dim selectedElements
	set selectedElements = CreateObject("System.Collections.ArrayList")
	dim selectedDiagramObjects
	set selectedDiagramObjects = diagram.SelectedObjects
	dim selectedDiagramObject as EA.DiagramObject
	for each selectedDiagramObject in selectedDiagramObjects
		dim selectedElement
		set selectedElement = Repository.GetElementByID(selectedDiagramObject.ElementID)
		if not selectedElement is nothing then
			selectedElements.Add selectedElement
		end if
	next	
	set getSelectedElements = selectedElements
end function

main