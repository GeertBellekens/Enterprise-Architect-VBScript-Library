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
	
	dim userinput
	userinput = MsgBox( "With Test Rules?", vbYesNo + vbQuestion, "Message Overview Diagram")
	dim includeRules
	if userinput = vbYes then
		'with test rules
		includeRules = true
	else	
		'without test rules
		includeRules = false
	end if
	set messageHeaders = selectedMessage.getHeaders(includeRules)
	
	'get the content in the proper output format
	dim messageOutput
	set messageOutput = selectedMessage.createOuput(includeRules)
	'show the output in the search window
	'create the output object
	dim searchOutput
	set searchOutput = new SearchResults
	searchOutput.Name = "Message detail"
	'put the headers in the output
	searchOutput.Fields = messageHeaders
	'put the content in the output
	searchOutput.Results = messageOutput
	dim row
	'show the output
	searchOutput.Show
	'export to excel file
	saveToExcelFile selectedMessage, messageOutput, messageHeaders
end function

function saveToExcelFile(message, messageOutput, messageHeaders)
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	'create tab for datatypes
	dim messageTypes
	set messageTypes = message.getMessageTypes()
	dim messageTypesArray 
	messageTypesArray = makeArrayFromArrayLists(messageTypes)
	excelOutput.createTab message.Prefix & " Types", messageTypesArray, true, "TableStyleMedium4"
	'create tab for message
	'merge headers with output
	messageOutput.Insert 0, messageHeaders
	'create a two-dimensional array from the messageOutput
	dim excelContents
	excelContents = makeArrayFromArrayLists(messageOutput)
	'add the output to a sheet in excel
	excelOutput.createTab message.Prefix & " Msg", excelContents, true, "TableStyleMedium4"
	'save the excel file
	excelOutput.save
end function


main