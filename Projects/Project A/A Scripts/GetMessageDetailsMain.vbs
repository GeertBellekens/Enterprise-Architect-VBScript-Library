'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

' Script Name: GetMessageDetailsMain 
' Author: Geert Bellekens
' Purpose: Main script for the GetMessageDetails export
' Date: 2017-03-15
'

'the location of the template to be used
const MIGDGOTemplate = "G:\Projects\80 Enterprise Architect\Output\UMIG - IM - Message Details template - Copy.xltx"
const MIG6Template = "G:\Projects\80 Enterprise Architect\Output\UMIG DGO - IM - Message Details template.xltx"

'colors used for formatting the excel file
dim atriasRed, white, yellow, blue, black, l1, l2, l3, l4, l5, l6, l7, l8, l9, l10
atriasRed =  RGB(153, 20, 37)
white = RGB(255,255,255)
yellow = RGB(255,230,153)
blue = RGB(189,215,238)
black = RGB(0,0,0)
l1 = RGB(235,103,119)
l2 = RGB(238,123,139)
l3 = RGB(241,143,159)
l4 = RGB(244,163,179)
l5 = RGB(247,183,199)
l6 = RGB(250,203,219)
l7 = RGB(253,223,239)
l8 = RGB(255,243,255)

dim isTechnical
dim customOrdering

sub getMessageDetailsMain(technical)
	'set the technical flag
	isTechnical = technical
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
	customOrdering = determineCustomOrdering(allMessages)
	'add all messages to the Excel file
	saveToExcelFile allMessages
end function

function getMessages(selectedPackage)
	dim packageIDtree
	packageIDtree = getPackageTreeIDString(selectedPackage)
	dim sqlGetMessageElements
	sqlGetMessageElements =	" select o.Object_ID                                       " & _
							" from t_object o                                          " & _
							" inner join t_xref x on x.Client = o.ea_guid              " & _                       
							" 					and x.Name = 'Stereotypes'             " & _                      
							" inner join t_package p on p.Package_ID = o.Package_ID    " & _
							" inner join t_object po on po.ea_guid = p.ea_guid		   " & _		              
							" where                                                    " & _
							" (x.Description like '%@STEREO;Name=XSDtopLevelElement;%' " & _
							" 	or  x.Description like '%@STEREO;Name=MA;%'            " & _
							" 		and po.Stereotype = 'DOCLibrary' )                 " & _
							" and o.Package_ID in (" & packageIDtree & ")"
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
		currentMessage.IncludeDetails = isTechnical
		currentMessage.loadMessage(messageElement)
		'add the message to the list of messages
		messages.add currentMessage
	next
	'return messages
	set getMessages = messages
end function

function saveToExcelFile(allMessages)
	dim message
	dim includeRules
	includeRules = false
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	'set template
	if not isTechnical then
		if customOrdering then
			excelOutput.NewFile MIG6Template
		else
			excelOutput.NewFile MIGDGOTemplate
			'with or without test rules
			dim userinput
			userinput = MsgBox( "With Test Rules?", vbYesNo + vbQuestion, "Include Test Rules")
			if userinput = vbYes then
				'with test rules
				includeRules = true
			end if
		end if
		saveRegularOutput allMessages, includeRules, excelOutput 
	else
		saveUnifiedOutput allMessages, includeRules, excelOutput 
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
	set messageHeaders = getMessageHeaders(includeRules, maxDepth, customOrdering, isTechnical, nothing)
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
		'create tab for message contents
		dim messageOutputArray
		messageOutputArray = makeArrayFromArrayLists(messageOutputList)
		'add the output to a sheet in excel
		excelOutput.createTab "Message Contents", messageOutputArray, true, "TableStyleMedium4"
		'create types sheet
		dim messageTypesArray
		messageTypesArray = makeArrayFromArrayLists(messageTypesList)
		excelOutput.createTab "Message Types", messageTypesArray, true, "TableStyleMedium4"
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
	'let the user know we are working on it
	Repository.WriteOutput outPutName, now() & " Formatting output..." , 0
	dim messageDictionary
	set messageDictionary = CreateObject("Scripting.Dictionary")
	dim message
	'loop al messages
	for each message in allMessages
		dim messageAlias
		messageAlias = getMessageAlias(message, messageDictionary)
		'build a dictionary with unique names that can be used as tabs.
		messageDictionary.Add messageAlias, message
'		'no types sheet for MIG6
'		if not message.CustomOrdering then
'			'create tab for types
'			dim messageTypesList
'			set messageTypesList = message.getMessageTypes()
'			dim messageTypesArray
'			messageTypesArray = makeArrayFromArrayLists(messageTypesList)
'			excelOutput.createTab messageAlias & " Types", messageTypesArray, true, "TableStyleMedium4"
'		end if
		'create tab for message contents
		dim messageOutputList
		set messageOutputList = message.createFullOutput(includeRules)
		dim messageOutputArray
		messageOutputArray = makeArrayFromArrayLists(messageOutputList)
		dim worksheet
		'add the output to a sheet in excel (no need for Msg suffix as there is no Types suffix
		set worksheet = excelOutput.createTab(messageAlias , messageOutputArray, true, "TableStyleLight11")
		'format the tab
		formatSheet excelOutput, worksheet, message, includeRules
	next
	'only save if there is anything to save
	if allMessages.Count > 0 then
		'create index
		createIndexSheet excelOutput, messageDictionary, allMessages
		'save the excel file
		excelOutput.save
	end if
end function 

function formatSheet(excelOutput, worksheet, message, includeRules)
	'add a new row at the top
	worksheet.Range("A1").EntireRow.Insert
	'remove the first two columns (order + message)
	worksheet.Columns(1).Delete
	worksheet.Columns(1).Delete
	'remove the third row (contains the message root level)
	worksheet.Rows(3).Delete 
	'merge columns till constraint
	worksheet.Range("A1", worksheet.Cells(1, message.MessageDepth + 3)).Merge
	'add the name of the message to Field A1
	worksheet.Cells(1,1).Value = "XSD: " & message.Name
	'merge colums for LDM mapping
	if message.CustomOrdering then
		worksheet.Range(worksheet.Cells(1,message.MessageDepth + 4 ), worksheet.Cells(1,message.MessageDepth + 5 )).Merge
		worksheet.Cells(1,message.MessageDepth + 4 ).Value = "LDM"
		'merge columns for Business Usage
		worksheet.Range(worksheet.Cells(1,message.MessageDepth + 6 ), worksheet.Cells(1,worksheet.UsedRange.Columns.Count)).Merge
		worksheet.Cells(1,message.MessageDepth + 6 ).Value = "Business Usage"
	elseif includeRules then
		'merge columns for Message Test rules
		worksheet.Range(worksheet.Cells(1,message.MessageDepth + 4 ), worksheet.Cells(1,worksheet.UsedRange.Columns.Count)).Merge
		worksheet.Cells(1,message.MessageDepth + 4 ).Value = "Test Rules"
	end if
	'freeze panes
	excelOutput.freezePanes worksheet, 2, 0
	'format headers
	formatHeaders excelOutput, worksheet, message, includeRules
	'format levels
	formatLevels excelOutput, worksheet, message
	'group rows per level
	groupRows excelOutput, worksheet, message
end function

function formatHeaders(excelOutput, worksheet, message, includeRules)
	'formatting Headers
	dim range
	set range = worksheet.Cells(1,1) 'Title for xsd
	excelOutput.formatRange range, atriasRed, white, "default", 14 , true, xlCenter
	set range = worksheet.Range("A2", worksheet.Cells(2, message.MessageDepth + 3)) 'headers for XSD
	excelOutput.formatRange range, atriasRed, white, "default", "default" , true, xlLeft
	if message.CustomOrdering then
		'formatting LDM headers (yellow)
		set range = worksheet.Cells(1,message.MessageDepth + 4 ) 'title for LDM
		excelOutput.formatRange range, yellow, black, "default", 14, true, xlCenter
		set range = worksheet.Range(worksheet.Cells(2,message.MessageDepth + 4), worksheet.Cells(2, message.MessageDepth + 5)) 'LDM headers
		excelOutput.formatRange range, yellow, black, "default", "default", true, xlLeft
		'formatting Business Fields (blue)
		set range = worksheet.Cells(1,message.MessageDepth + 6 )'Business Usage Title
		excelOutput.formatRange range, blue, black, "default", 14, true, xlCenter
		set range = worksheet.Range(worksheet.Cells(2,message.MessageDepth + 6), worksheet.Cells(2, worksheet.UsedRange.Columns.Count)) 'Business Usage headers
		excelOutput.formatRange range, blue, black, "default", "default", true, xlLeft
	elseif includeRules then
		'formatting Business Fields (blue)
		set range = worksheet.Cells(1,message.MessageDepth + 4 )'Test Rules
		excelOutput.formatRange range, blue, black, "default", 14, true, xlCenter
		set range = worksheet.Range(worksheet.Cells(2,message.MessageDepth + 4), worksheet.Cells(2, worksheet.UsedRange.Columns.Count)) 'Business Usage headers
		excelOutput.formatRange range, blue, black, "default", "default", true, xlLeft
	end if
	'set autofit to account for changes in size/bold etc..
	worksheet.UsedRange.Columns.Autofit
	'set the constraints column to be larger
	worksheet.Columns(message.MessageDepth + 3).ColumnWidth = 35
	worksheet.UsedRange.Rows.Autofit
	'set width of level colums to 5 except for the last one
	dim i
	for i = 1 to message.MessageDepth -2
		worksheet.Columns(i).ColumnWidth  = 5
	next
end function

function formatLevels (excelOutput, worksheet, message)
	'create array with colors
	dim levelColors
	set levelColors = CreateObject("System.Collections.ArrayList")
	levelColors.add l1
	levelColors.add l2
	levelColors.add l3
	levelColors.add l4
	levelColors.add l5
	levelColors.add l6
	levelColors.add l7
	levelColors.add l8 'maximum 8 levels of grouping
	'loop levels
	dim i
	for i = 1 to message.MessageDepth -2
		dim color
		if i < 8 then
			color = levelColors(i)
		end if
		dim row
		dim range
		for row = 3 to worksheet.UsedRange.Rows.Count
			if len(worksheet.Cells(row, i).Value) > 0  then 
				'check if this is the last level in this row and if there is a next level
				if len(worksheet.Cells(row, i + 1).Value) = 0 _
				AND len(worksheet.Cells(row + 1, i + 1).Value) > 0 then
					'set background color for whole row
					set range = worksheet.Range(worksheet.Cells(row, i), worksheet.Cells(row, worksheet.UsedRange.Columns.Count))
					excelOutput.formatRange range, color, "default", "default", "default", "default", "default"
				elseif len(worksheet.Cells(row, i + 1).Value) > 0 then
					set range = worksheet.Cells(row, i)
					'set background and fontcolor for this column
					excelOutput.formatRange range, color, color, "default", "default", "default", "default"
				end if
			end if
		next
	next
end function

function groupRows(excelOutput, worksheet, message)
	'set outline setting to abow
	worksheet.Outline.SummaryRow = xlAbove
	dim i
	for i = 1 to message.MessageDepth -2
		'check if this field is filled in and the next is not
		dim row
		dim range
		dim startRow
		startRow = 0
		dim currentName
		currentName = ""
		dim startName
		startName = ""
		for row = 3 to worksheet.UsedRange.Rows.Count + 1
			currentName = worksheet.Cells(row, i).Value
			if len(startName) = 0 then
				startName = currentName
				startRow = row
			end if
			if currentName <> startName then
				if startRow + 1 < row _
				  and i < 8 then 'there should be at least two rows to group and the maximum level of grouping in excel is 8
					on error resume next 'in case the grouping doesn't work
					set range = worksheet.Range(worksheet.Cells(startRow + 1,1), worksheet.Cells(row -1,1))
					range.EntireRow.Group
					Err.Clear
					on error goto 0
				end if
				'reset startname
				startName = currentName
				startRow = row
			end if
		next
	next
end function



function createIndexSheet(excelOutput, messageDictionary, allMessages)
	dim indexContent
	set indexContent = CreateObject("System.Collections.ArrayList")
	dim key
	dim messageName
	'determine custom ordering
	dim customOrdering
	customOrdering = determineCustomOrdering(allMessages)
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
		formula = "=HYPERLINK(""#'" & key & "'!A2"",""" & messageDictionary(key).Name & """)"
		row.Add formula
		'add the row data to the indexcontent (each time inserting right behind the header
		indexContent.Insert 1, row
	next
	'make it into an array
	dim indexArray
	indexArray = makeArrayFromArrayLists(indexContent)
	'create sheet
	excelOutput.createTabWithFormulas "Index", indexArray, true, "TableStyleMedium4", 1
end function

function getMessageAlias(message, messageDictionary)
	dim namePartLenght
	if message.CustomOrdering then
		namePartLenght = 27 'no suffix needed so larger name allowed
	else
		namePartLenght = 21 'need to account for suffix as well
	end if
	dim namePart
	'check only left 21/27 characters to stay below the maximum 31 for excel worksheet names
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