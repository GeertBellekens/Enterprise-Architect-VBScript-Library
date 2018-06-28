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
	sqlGetMessageElements =	"select o.Object_ID, x.Description from t_object o                                " & _
							" inner join t_xref x on x.Client = o.ea_guid                                     " & _
							" 					and x.Name = 'Stereotypes'                                    " & _
							" 					and ( x.Description like '%@STEREO;Name=XSDtopLevelElement;%' " & _
							" 						or x.Description like '%@STEREO;Name=MA;%')               " & _
							" where o.Package_ID in (" & packageIDtree & ")"
	dim messageElements
	set messageElements = getElementsFromQuery(sqlGetMessageElements)
	dim messageElement
	dim currentMessage
	dim messages
	set messages = CreateObject("System.Collections.ArrayList")
	'loop the message elements
	for each messageElement in messageElements
		Repository.WriteOutput outPutName, now() & " Processing Message '" & messageElement.Name & "'", messageElement.ElementID
		set currentMessage = new Message
		currentMessage.loadMessage(messageElement)
		'add the message to the list of messages
		messages.add currentMessage
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
	userinput = MsgBox( "With Test Rules?", vbYesNo + vbQuestion, "Include Test Rules")
	dim includeRules
	if userinput = vbYes then
		'with test rules
		includeRules = true
	else	
		'without test rules
		includeRules = false
	end if
	
	'unified or separate tabs
	userinput = MsgBox( "All in one sheet?", vbYesNo + vbQuestion, "Choose output style")
	dim unified
	if userinput = vbYes then
		'all in one sheet
		unified = true
	else	
		'one sheet per message
		unified = false
	end if
	
	if unified then
		saveUnifiedOutput allMessages, includeRules, excelOutput 
	else
		saveRegularOutput allMessages, includeRules, excelOutput 
	end if
	
end function

function saveUnifiedOutput(allMessages, includeRules, excelOutput)
	'first get the maximum depth
	dim maxDepth
	maxDepth = getMaximumDepth(allMessages)
	'create types list
	dim messageTypesList
	set messageTypesList = CreateObject("System.Collections.ArrayList")
	'add headers
	dim messageTypeHeaders
	set messageTypeHeaders = getMessageTypesHeaders(true)
	messageTypesList.Add messageTypeHeaders
	'create message output list
	dim messageOutputList
	set messageOutputList = CreateObject("System.Collections.ArrayList")
	'add headers
	dim customOrdering
	customOrdering = determineCustomOrdering(allMessages)
	dim messageHeaders
	set messageHeaders = getMessageHeaders(includeRules, maxDepth, customOrdering)
	messageOutputList.Add messageHeaders
	'then loop the messages and get their output
	dim message
	for each message in allMessages
		'add type info 
		messageTypesList.AddRange message.getUnifiedMessageTypes()
		'create tab for message contents
		messageOutputList.AddRange message.createUnifiedOutput(includeRules, maxDepth)
	next
	'only save if there is anything to save
	if allMessages.Count > 0 then
		'create types sheet
		dim messageTypesArray
		messageTypesArray = makeArrayFromArrayLists(messageTypesList)
		excelOutput.createTab "Message Types", messageTypesArray, true, "TableStyleMedium4"
		'create tab for message contents
		dim messageOutputArray
		messageOutputArray = makeArrayFromArrayLists(messageOutputList)
		'add the output to a sheet in excel
		excelOutput.createTab "Message Contents", messageOutputArray, true, "TableStyleMedium4"
		'save the excel file
		excelOutput.save
	end if
end function

function determineCustomOrdering(allMessages)
	determineCustomOrdering = false 'initialize
	dim message 
	for each message in allMessages
		determineCustomOrdering = message.CustomOrdering
		exit for 'break after first hit
	next
end function

function getMaximumDepth(allMessages)
	dim message
	dim maxDepth
	maxDepth = 0
	for each message in allMessages
		if message.MessageDepth > maxDepth then
			maxDepth = message.MessageDepth
		end if
	next
	getMaximumDepth = maxDepth
end function

function saveRegularOutput(allMessages, includeRules,excelOutput )
	dim messageDictionary
	set messageDictionary = CreateObject("Scripting.Dictionary")
	dim message
	'loop al messages
	for each message in allMessages
		dim messageAlias
		messageAlias = getMessageAlias(message, messageDictionary)
		'build a dictionary with unique names that can be used as tabs.
		messageDictionary.Add messageAlias, message
		'create tab for types
		dim messageTypesList
		set messageTypesList = message.getMessageTypes()
		dim messageTypesArray
		messageTypesArray = makeArrayFromArrayLists(messageTypesList)
		excelOutput.createTab messageAlias & " Types", messageTypesArray, true, "TableStyleMedium4"
		'create tab for message contents
		dim messageOutputList
		set messageOutputList = message.createFullOutput(includeRules)
		dim messageOutputArray
		messageOutputArray = makeArrayFromArrayLists(messageOutputList)
		'add the output to a sheet in excel
		excelOutput.createTab messageAlias & " Msg", messageOutputArray, true, "TableStyleMedium4"
	next
	'only save if there is anything to save
	if allMessages.Count > 0 then
		'create index
		createIndexSheet excelOutput, messageDictionary
		'save the excel file
		excelOutput.save
	end if
end function 

function createIndexSheet(excelOutput, messageDictionary)
	dim indexContent
	set indexContent = CreateObject("System.Collections.ArrayList")
	dim key
	dim messageName
	'add the header
	dim header
	set header = CreateObject("System.Collections.ArrayList")
	header.add "Index"
	indexContent.Add header
	'add the rows
	for each key in messageDictionary.Keys
		dim row
		set row = CreateObject("System.Collections.ArrayList")
		'=HYPERLINK("#'MessageWithAVeryLongN_2 Msg'!A2";"MessageWithVeryLongNametThatCannotFitIntoTheSheetName")
		dim formula
		formula = "=HYPERLINK(""#'" & key & " Msg'!A2"",""" & messageDictionary(key).Name & """)"
		row.Add formula
		'add the row data to the indexcontent (each time inserting right behind the header
		indexContent.Insert 1, row
	next
	'make it into an array
	dim indexArray
	indexArray = makeArrayFromArrayLists(indexContent)
	'create sheet
	excelOutput.createTabWithFormulas "Index", indexArray, true, "TableStyleMedium4"
end function

function getMessageAlias(message, messageDictionary)
	dim namePartLenght
	namePartLenght = 21
	dim namePart
	'check only left 21 characters to stay below the maximum 31 for excel worksheet names
	if len(message.Prefix) > 0 then
		namePart = left(message.Prefix,namePartLenght)
	else
		namePart = left(message.Name,namePartLenght)
	end if
	'find conflicts and add counter if needed
	dim existingCounter
	existingCounter = 1
	dim key
	for each key in messageDictionary.Keys
		 if lcase(namePart) = lcase(left(key,namePartLenght)) then
			existingCounter = existingCounter + 1
		 end if
	next
	'check if existing found
	if existingCounter > 1 then
		'add counter after name
		getMessageAlias = namePart & "_" & existingCounter
	else
		getMessageAlias = namePart
	end if
end function

main