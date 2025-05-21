'[path=\Projects\Project B\Element Group]
'[group=Element Group]


!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Set all Attribute optional
' Author: Geert Bellekens
' Purpose: Set all attributes of the selected elements to optional (lowerbound = 0)
' Date: 2025-04-02
'
const outPutName = "Set all Attribute optional"

function main ()
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName , 0
	'actually do the work
	setAttributesOptional
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName , 0
end function

function setAttributesOptional()
	dim diagram as EA.Diagram
	set diagram = Repository.GetCurrentDiagram()
	dim elements
	set elements = CreateObject("System.Collections.ArrayList")
	if not diagram is nothing then
		dim diagramObjects
		set diagramObjects = diagram.SelectedObjects
		dim diagramObject as EA.DiagramObject
		for each diagramObject in diagramObjects
			dim element as EA.Element
			set element = Repository.GetElementByID(diagramObject.ElementID)
			elements.Add element
		next
		if elements.Count = 0 then
			set elements = Repository.GetTreeSelectedElements()
		end if
	else
		set elements = Repository.GetTreeSelectedElements()
	end if
	for each element in elements
		if element.Type = "Class" then
			Repository.WriteOutput outPutName, now() & " Processing element '" & element.Name & "'", 0
			dim attribute as EA.Attribute
			for each attribute in element.Attributes
				if attribute.LowerBound <> 0 then
					attribute.LowerBound = 0
					attribute.Update
				end if
			next
		end if
	next
end function

main
	
