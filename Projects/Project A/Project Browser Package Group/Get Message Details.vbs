'[path=\Projects\Project A\Project Browser Package Group]
'[group=Project Browser Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

' Script Name: Show 
' Author: Geert Bellekens
' Purpose: Get Message Details for all messages in this folder and the subfolders and save them to excel
' Date: 2017-03-15
'

'name of the output tab
const outPutName = "Get Message Details"

sub main

	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	
		'get the selected element
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetContextObject
	if selectedPackage.ObjectType = otPackage then
		'tell the user we are starting
		Repository.WriteOutput outPutName, now() & " Starting Get Message Details for package '" & selectedPackage.Name & "'", selectedPackage.Element.ElementID
		'do the actual work
		getmessageDetails selectedPackage
		'tell the user we are finished
		Repository.WriteOutput outPutName, now() & " Finished Get Message Details for package '" & selectedPackage.Name & "'", selectedPackage.Element.ElementID
	else
		msgbox "This script only works on Packages. Please select a Package before executing this script"
	end if
end sub

function getmessageDetails(selectedPackage)
	'get the messages in the selected package (and it's subpackages)
	dim allMessages
	set allMessages = getMessages(selectedPackage)
	'add all messages to the Excel file
	saveToExcelFile allMessages
end function

function getMessages(selectedPackage)
	dim packageIDtree
	packageIDtree = getPackageTreeIDString(selectedPackage)
	dim sqlGetMessageElements
	sqlGetMessageElements =	"select o.Object_ID from t_object o " & _
							" where o.Stereotype = 'XSDtopLevelElement' " & _
							" and o.Package_ID in (" & packageIDtree & ")"
	dim messageElements
	set messageElements = getElementsFromQuery(sqlGetMessageElements)
	dim messageElement
	dim currentMessage
	dim messages
	set messages = CreateObject("System.Collections.ArrayList")
	'loop the message elements
	for each messageElement in messageElements
		set currentMessage = new Message
		currentMessage.loadMessage(messageElement)
		'add the message to the list of messages
		messages.add currentMessage
		'tell the user what we are doing
		Repository.WriteOutput outPutName, now() & " Processed Message '" & currentMessage.Name & "'", messageElement.ElementID
	next
	'return messages
	set getMessages = messages
end function

function saveToExcelFile(allMessages)
	dim message
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	'with or without test rules
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
	
	
	'loop al messages
	for each message in allMessages
		'create tab for types
		dim messageTypesList
		set messageTypesList = message.getMessageTypes()
		dim messageTypesArray
		messageTypesArray = makeArrayFromArrayLists(messageTypesList)
		excelOutput.createTab message.Prefix & " Types", messageTypesArray, true, "TableStyleMedium4"
		'create tab for message contents
		dim messageOutputList
		set messageOutputList = message.createFullOutput(includeRules)
		dim messageOutputArray
		messageOutputArray = makeArrayFromArrayLists(messageOutputList)
		'add the output to a sheet in excel
		excelOutput.createTab message.Prefix & " Msg", messageOutputArray, true, "TableStyleMedium4"
	next
	'only save if there is anything to save
	if allMessages.Count > 0 then
		'save the excel file
		excelOutput.save
	end if
end function

main