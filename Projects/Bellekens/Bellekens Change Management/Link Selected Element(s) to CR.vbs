'[path=\Projects\Bellekens\Bellekens Change Management]
'[group=Element Group]
option explicit

!INC Bellekens Change Management.LinkToCRMain

'This script only calls the function defined in the main script.
'Ths script is to be copied in Diagram, Search and Project Browser groups

'Execute main function defined in LinkToCRMain
sub main
	'check if called from diagram
	dim diagram as EA.Diagram
	set diagram = Repository.GetCurrentDiagram
	dim linkedToCR
	linkedToCR = false
	if not diagram is nothing then
		dim selectedItems
		set selectedItems = getSelectedElements(diagram)
		if selectedItems.Count > 0 then
			linkedToCR = true
			linkItemToCR nothing, selectedItems
		end if
	end if
	'if not from diagram then use the selection in the project browser
	if not linkedToCR then
		dim treeSelectedElements
		set treeSelectedElements = Repository.GetTreeSelectedElements()
		if treeSelectedElements.Count > 0 then
			linkItemToCR nothing, treeSelectedElements
		else
			dim selectedItem
			set selectedItem = Repository.GetContextObject
			linkItemToCR selectedItem, nothing
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