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
	showMessageDetail(selectedElement)
	'tell the user we are finished
	Repository.WriteOutput outPutName, now() & " Finished Creating Message Detail for '" & selectedElement.Name & "'", selectedElement.ElementID
	else
		msgbox "This script only works on Elements. Please select an Element before executing this script"
	end if
end sub

function showMessageDetail(selectedElement)
	dim selectedMessage 
	set selectedMessage = new Message
	'first load the selected message
	selectedMessage.loadMessage(selectedElement)
	'get the headers
	dim messageHeaders
	set messageHeaders = selectedMessage.getHeaders()
	'get the content in the proper output format
	dim messageOutput
	set messageOutput = selectedMessage.createOuput()
	'show the output in the search window
	'create the output object
	dim searchOutput
	set searchOutput = new SearchResults
	searchOutput.Name = "Message detail"
	'put the headers in the output
	searchOutput.Fields = messageHeaders
	Session.Output "messageHeaders.Count: " & messageHeaders.Count
	'put the content in the output
	searchOutput.Results = messageOutput
	dim row
	for each row in messageOutput
		Session.Output "row.Count: " & row.Count
	next
	'show the output
	searchOutput.Show
end function


main