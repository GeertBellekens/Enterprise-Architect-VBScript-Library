'[path=\Projects\Project I\Diagram Group]
'[group=Diagram Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

' Script Name: Add or Update Tagged Values
' Author: Geert Bellekens
' Purpose: Add new tagged values, or update existing tagged values on the selected elements
' Date: 2025-09-03

const outPutName = "Add or Update Tagged Values"


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
	addOrUpdateTagsForDiagram diagram
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName & " for '"& diagram.Name &"'", 0
	
end sub

function addOrUpdateTagsForDiagram(diagram)
	'ask tagname
	dim tagName
	tagName = InputBox("Please enter the tag name", "Tag Name", "" )
	if len(tagName) = 0 then
		exit function
	end if
	'ask tagValue
	dim tagValue
	tagValue = InputBox("Please enter the tag value", "Tag Value", "" )
	
	dim confirmedAdd
	confirmedAdd = 0
	'figure out if any element is selected
	dim elements
	set elements = getSelectedElementsOnDiagram(diagram)
	dim element as EA.Element
	for each element in elements
		confirmedAdd = addOrUpdateTaggedValue(element, tagName, tagValue, confirmedAdd)
	next
end function

function addOrUpdateTaggedValue(element, tagName, tagValue, confirmedAdd)
	Repository.WriteOutput outPutName, now() & " Processing element '"& element.Name &"'", 0
	dim found
	found = false
	dim tag as EA.TaggedValue
	for each tag in element.TaggedValues
		if lcase(tag.Name) = lcase(tagName) then
			found = true
			if not tag.Value = tagValue then
				tag.Value = tagValue
				tag.Update
			end if
			exit for
		end if
	next
	'add tag if needed
	if not found then
		'check if we need to ask permission to add tagged values
		if confirmedAdd = 0 then
			dim userInput
			userInput =  Msgbox("Add missing tags?", vbYesNo+vbQuestion, "Add missing tags?")
			if userInput = vbYes then
				confirmedAdd = 1
			else
				confirmedAdd = 0
			end if
		end if
		if confirmedAdd = 1 then
			set tag = element.TaggedValues.AddNew(tagName, "")
			tag.Value = tagValue
			tag.Update
		end if
	end if
	'return
	addOrUpdateTaggedValue = confirmedAdd
end function

'main

function test
	'reset output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName

	dim diagram as EA.Diagram
	set diagram = Repository.GetDiagramByGuid("{86139A0F-E575-8239-A67A-1A588A428FF6}")
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName & " for '"& diagram.Name &"'", 0
	'do the actual work
	addOrUpdateTagsForDiagram diagram
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName & " for '"& diagram.Name &"'", 0

end function

test