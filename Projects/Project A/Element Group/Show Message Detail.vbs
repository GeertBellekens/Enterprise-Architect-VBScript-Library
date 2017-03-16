'[path=\Projects\Project A\Element Group]
'[group=Element Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

' Script Name: Show Message Detail
' Author: Geert Bellekens
' Purpose: shows the message details of the selected message in the search window and asks 
' Date: 2017-03-15
'

'name of the output tab
const outPutName = "Create Message Detail"

sub main

	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	
	'get the selected element
	dim selectedElement as EA.Element
	set selectedElement = Repository.GetContextObject
	if selectedElement.ObjectType = otElement then
	
	'tell the user we are starting
	Repository.WriteOutput outPutName, now() & " Starting Creating Message Detail for '" & selectedElement.Name & "'", selectedElement.ElementID
	'do the actual work
	createMessageDetail(selectedElement)
	'tell the user we are finished
	Repository.WriteOutput outPutName, now() & " Finished Creating Message Detail for '" & selectedElement.Name & "'", selectedElement.ElementID
	else
		msgbox "This script only works on Elements. Please select an Element before executing this script"
	end if
end sub

function createMessageDetail(selectedElement)
	dim selectedMessage 
	set selectedMessage = new Message
	selectedMessage.loadMessage(selectedElement)
end function

main