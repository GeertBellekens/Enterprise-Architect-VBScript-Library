'[path=\Projects\Project B\Baloise Scripts]
'[group=Baloise Scripts]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

' Script Name: MessageDetailsMain
' Author: Geert Bellekens
' Purpose: library for exporting Message Details
' Date: 2023-12-01
'


dim isTechnical

sub getMessageDetailsMain(exportType)
	'set the technical flag
	isTechnical = false
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
		getmessageDetails selectedPackage, exportType
		'tell the user we are finished
		Repository.WriteOutput outPutName, now() & " Finished Get Message Details for package '" & selectedPackage.Name & "'", selectedPackage.Element.ElementID
	else
		msgbox "This script only works on Packages. Please select a Package before executing this script"
	end if
end sub

function getmessageDetails(selectedPackage, exportType)
	'get the messages in the selected package (and it's subpackages)
	dim allMessages
	set allMessages = getMessages(selectedPackage)
	'add all messages to the Excel file
	saveToExcelFile allMessages, selectedPackage, exportType
end function

function getMessages(selectedPackage)
	dim packageIDtree
	packageIDtree = getPackageTreeIDString(selectedPackage)
	dim sqlGetMessageElements
	sqlGetMessageElements =	"select o.Object_ID                                                            " & vbNewLine & _
							" from t_object o                                                              " & vbNewLine & _
							" where o.Stereotype  = 'XSDtopLevelElement'                                   " & vbNewLine & _
							" and o.Package_ID in (" & packageIDtree & ")                                  "
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
	'check if we found any messages. If not we try to get the message for the subset
	if not messages.Count > 0 then
		set currentMessage = getMessageForSubset(selectedPackage)
		'add the message to the list of messages
		messages.add currentMessage
	end if
	'return messages
	set getMessages = messages
end function

function getMessageForSubset(selectedPackage)
	dim message
	'create the message
	set message = new Message
	message.IncludeDetails = isTechnical
	message.loadMessage(selectedPackage)
	'return
	set getMessageForSubset = message
end function

function saveToExcelFile(allMessages, selectedPackage, exportType)
	dim message
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	excelOutput.openUserSelectedFile
	
	dim includeRules
	'without test rules
	includeRules = false
	'add business rules
	addBusinessRules selectedPackage, excelOutput
	
	select case exportType
		case functionalDesign
			saveFSOutput allMessages, excelOutput 
		case mappingDocument
			saveMappingOutput allMessages, excelOutput 
		case else
			saveRegularOutput allMessages, includeRules, excelOutput 
	end select
	
end function

function addBusinessRules(selectedPackage, excelOutput)
	dim BRspecification
	BRspecification = "" 'initialize
	'get tagged value 'BusinessRules'
	dim taggedValue as EA.TaggedValue
	for each taggedValue in selectedPackage.Element.TaggedValues
		if taggedValue.Name = "BusinessRules" then
			BRspecification = taggedValue.Value
			exit for
		end if
	next
	'check if any BRspecification found
	if len(BRspecification) <= 0 then
		exit function
	end if
	'BRspecification is fileName:SheetName, so we split on ":"
	dim parts
	parts = split(BRspecification,"::")
	if Ubound(parts) <= 0 then
		exit function
	end if
	'get filename and sheetname
	dim fileName
	fileName = parts(0)
	dim sheetName
	sheetName = parts(1)
	
	'copy sheet to this excel file
	excelOutput.copyWorksheet fileName, sheetName, "Business Rules"
	
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
	set messageHeaders = getMessageHeaders(includeRules, maxDepth, customOrdering, isTechnical)
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
		excelOutput.createTab "Message Types", messageTypesArray, false, ""
		'create tab for message contents
		dim messageOutputArray
		messageOutputArray = makeArrayFromArrayLists(messageOutputList)
		'add the output to a sheet in excel
		excelOutput.createTab "Message Contents", messageOutputArray, false, ""
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

function saveMappingOutput(allMessages, excelOutput)
	dim enumTypesColumn
	enumTypesColumn = 4
	'report progress
	Repository.WriteOutput outPutName, now() & " Updating Excel file", 0
	dim messageDictionary
	set messageDictionary = CreateObject("Scripting.Dictionary")
	

	dim messageTypesList
	set messageTypesList = CreateObject("System.Collections.ArrayList")
	dim valuesTabName
	valuesTabName = "ValueList"
	dim message
	'loop all messages to get MessageTypesList
	for each message in allMessages	
		messageTypesList.AddRange message.getMessageTypes()
	next
	'report progress
	Repository.WriteOutput outPutName, now() & " Updating sheet '" & valuesTabName & "'", 0
	'remove duplicate types
	removeDuplicateMessageTypes(messageTypesList)
	'create tab for the enums
	dim valuesWorksheet
	dim messageTypesArray
	messageTypesArray = makeArrayFromArrayLists(messageTypesList)
	set valuesWorksheet = excelOutput.createTabWithOffset(valuesTabName, messageTypesArray, false, "", 5, enumTypesColumn)
	'format value lists tab (underline enum rows)
	formatValuesSheet valuesWorksheet, messageTypesList, 5, enumTypesColumn
	
	'loop al messages to write the contents
	for each message in allMessages
		dim structureTabName
		dim structureRowStart
		dim structureColumnStart
		'create tab for FS contents
		dim contentWorksheet
		dim tabName
		tabName = message.Name
		'report progress
		Repository.WriteOutput outPutName, now() & " Updating sheet '" & tabName & "'", 0
		dim messageOutputList
		set messageOutputList = message.createStructureOutput(mappingDocument)
		dim messageOutputArray
		messageOutputArray = makeArrayFromArrayLists(messageOutputList)
		'debug
'		dim i
'		for i = 0 to Ubound(messageOutputArray) -1
'			Session.Output "0: '" & messageOutputArray(i,0) & "' 1: '" &  messageOutputArray(i,1) & "' 2: '" &  messageOutputArray(i,2) & "' 3: '" &  messageOutputArray(i,3) & "'" & "' 4: '" &  messageOutputArray(i,4) & "'"
'		next
		set contentWorksheet = excelOutput.createTabWithOffset(tabName, messageOutputArray, false, "", 16, 2)
		'format sheet (grouping + hyperlinks to enum types)
		formatFSsheet excelOutput, contentWorksheet, message, messageOutputList, 16, messageTypesList, valuesTabName, enumTypesColumn
	next
	'only save if there is anything to save
	if allMessages.Count > 0 then
		'save the excel file
		excelOutput.save
	end if
end function 


function saveFSOutput(allMessages, excelOutput)
	dim enumTypesColumn
	enumTypesColumn = 2
	'report progress
	Repository.WriteOutput outPutName, now() & " Updating Excel file", 0
	dim messageDictionary
	set messageDictionary = CreateObject("Scripting.Dictionary")
	dim message
	'report progress
	Repository.WriteOutput outPutName, now() & " Updating sheet '" & valuesTabName & "'", 0
	'loop al messages to write the contents
	for each message in allMessages
		dim structureTabName
		dim structureRowStart
		dim structureColumnStart
		'create tab for FS contents
		dim contentWorksheet
		dim tabName
		tabName = message.Name
		dim valuesTabName
		valuesTabName = tabName & "_VL"
		dim messageTypesList
		set messageTypesList = message.getMessageTypes()
		'report progress
		Repository.WriteOutput outPutName, now() & " Updating sheet '" & valuesTabName & "'", 0
		'create sheet for the enums
		dim valuesWorksheet
		dim messageTypesArray
		messageTypesArray = makeArrayFromArrayLists(messageTypesList)
		set valuesWorksheet = excelOutput.createTabWithOffset(valuesTabName, messageTypesArray, false, "", 5, enumTypesColumn)
		'format value lists tab (underline enum rows)
		formatValuesSheet valuesWorksheet, messageTypesList, 5, enumTypesColumn
		'report progress
		Repository.WriteOutput outPutName, now() & " Updating sheet '" & tabName & "'", 0
		dim messageOutputList
		set messageOutputList = message.createStructureOutput(functionalDesign)
		dim messageOutputArray
		messageOutputArray = makeArrayFromArrayLists(messageOutputList)
		set contentWorksheet = excelOutput.createTabWithOffset(tabName, messageOutputArray, false, "", 16, 2)
		'format sheet (grouping + hyperlinks to enum types)
		formatFSsheet excelOutput, contentWorksheet, message, messageOutputList, 16, messageTypesList, valuesTabName, enumTypesColumn
		

	next
	'only save if there is anything to save
	if allMessages.Count > 0 then
		'save the excel file
		excelOutput.save
	end if
end function 

function removeDuplicateMessageTypes(messageTypesList)
	dim processedTypes
	set processedTypes = CreateObject("Scripting.Dictionary")
	dim indexesToRemove
	set indexesToRemove = CreateObject("System.Collections.ArrayList")
	dim row
	dim i
	dim toRemove
	toRemove = false
	'record all indexes that need to be removed
	for i = 0 to messageTypesList.Count - 1
		set row = messageTypesList(i)
		dim enumName
		dim currentRowName
		currentRowName = row(0)
		if len(currentRowName) > 0 then
			enumName = currentRowName
			'if the enum already exists we need to remove it (and all it's values)
			if processedTypes.Exists(enumName) then
				toRemove = true
			else
				toRemove = false
				processedTypes.Add enumName, enumName
			end if
		end if
		if toRemove then
			indexesToRemove.Add i
		end if
	next
	'loop all indexes backwards, and remove all of them from the messageTypesList
	for i = indexesToRemove.Count - 1 to 0 step -1
		dim indexToRemove
		indexToRemove = indexesToRemove(i)
		messageTypesList.RemoveAt indexToRemove
	next
end function

function saveRegularOutput(allMessages, includeRules,excelOutput )
	dim enumTypesColumn
	enumTypesColumn = 4
	'report progress
	Repository.WriteOutput outPutName, now() & " Updating Excel file", 0
	dim messageDictionary
	set messageDictionary = CreateObject("Scripting.Dictionary")
	
	dim message
	'loop al messages
	for each message in allMessages
		dim valuesTabName
		dim structureTabName
		dim structureRowStart
		dim structureColumnStart
		if lcase(right(message.Name,4)) = "_out" then
			valuesTabName = "Value Lists_OUT"
			structureTabName = "Structure OUT"
			'make sure the content starts correctly structure (without header)
			structureRowStart = 5
			structureColumnStart = 4
		elseif lcase(right(message.Name,3)) = "_in" then
			valuesTabName = "Value Lists_IN"
			structureTabName = "Structure IN"
			'make sure the content starts correctly structure (with header)
			structureRowStart = 8
			structureColumnStart = 5
		else
			valuesTabName = "Value Lists"
			structureTabName = "Structure"
			'make sure the content starts correctly structure (without header)
			structureRowStart = 5
			structureColumnStart = 4
		end if
		'report progress
		Repository.WriteOutput outPutName, now() & " Updating sheet '" & valuesTabName & "'", 0
		'create tab for the enums
		dim valuesWorksheet
		dim messageTypesArray
		dim messageTypesList
		set messageTypesList = message.getMessageTypes()
		messageTypesArray = makeArrayFromArrayLists(messageTypesList)
		set valuesWorksheet = excelOutput.createTabWithOffset(valuesTabName, messageTypesArray, false, "", 5, enumTypesColumn)
		'format value lists tab (underline enum rows)
		formatValuesSheet valuesWorksheet, messageTypesList, 5, enumTypesColumn
		
		'report progress
		Repository.WriteOutput outPutName, now() & " Updating sheet '" & structureTabName & "'", 0
		'create tab for the structure
		dim structureWorksheet
		dim structureOutput
		set structureOutput = message.createStructureOutput(regularMessageContent)
		dim structureArray
		structureArray = makeArrayFromArrayLists(structureOutput)
		set structureWorksheet = excelOutput.createTabWithOffset(structureTabName, structureArray, false, "", structureRowStart, structureColumnStart)
		
		'create tab for message contents
		dim contentWorksheet
		dim tabName
		tabName = message.Name
		'report progress
		Repository.WriteOutput outPutName, now() & " Updating sheet '" & tabName & "'", 0
		dim messageOutputList
		set messageOutputList = message.createOutput(includeRules)
		dim messageOutputArray
		messageOutputArray = makeArrayFromArrayLists(messageOutputList)
		'debug
'		dim i
'		for i = 0 to Ubound(messageOutputArray) -1
'			Session.Output "0: '" & messageOutputArray(i,0) & "' 1: '" &  messageOutputArray(i,1) & "' 2: '" &  messageOutputArray(i,2) & "' 3: '" &  messageOutputArray(i,3) & "'" & "' 4: '" &  messageOutputArray(i,4) & "'"
'		next
		set contentWorksheet = excelOutput.createTabWithOffset(tabName, messageOutputArray, false, "", 5, 4)
		'format sheet (grouping + hyperlinks to enum types)
		formatContentSheet excelOutput, contentWorksheet, message, messageOutputList, 5, messageTypesList, valuesTabName, enumTypesColumn
		
	next
	'only save if there is anything to save
	if allMessages.Count > 0 then
		'save the excel file
		excelOutput.save
	end if
end function 


function formatValuesSheet(valuesWorksheet, messageTypesList, rowOffset, columnOffset)
	dim i
	for i = 0 to messageTypesList.count -1
		dim range
		set range = valuesWorksheet.Range(valuesWorksheet.Cells(i + rowOffset, columnOffset), valuesWorksheet.Cells(i + rowOffset,columnOffset + 3))
		dim typeName
		typeName = messageTypesList(i)(0)
		if len(typeName) > 0 then
			'set bottom border
			with range.Borders(xlEdgeBottom)
				.Color = RGB(0,112,192) 'dark blue 'REBRANDING: Color changed from RGB(0,51,153) to RGB(0,13,110)'
				.LineStyle = xlContinuous
				.Weight = xlThin
			end with
		else
			'remove bottom border for all others
			range.Borders(xlEdgeBottom).LineStyle = xlNone
		end if
	next
	dim remarkColumn
	remarkColumn = columnOffset + 3
	'if the API remark says "not in use" then we set the color of the row to grey
	formatNotInUseRows valuesWorksheet, messageTypesList, rowOffset, columnOffset
	'create hyperlinks to business rules sheet
	setHyperlinksForBusinessRules valuesWorksheet, messageTypesList, rowOffset, columnOffset, remarkColumn
end function

function formatNotInUseRows(valuesWorksheet,messageTypesList, rowOffset,columnOffset)
	dim row
	for row = 0 to messageTypesList.Count - 1
		'get the remark
		dim currentCell
		set currentCell = valuesWorksheet.Cells(row +rowOffset, columnOffset + 3)
		dim apiRemark
		apiRemark = currentCell.Value
		dim range
		set range = valuesWorksheet.Range(valuesWorksheet.Cells(row +rowOffset, columnOffset), valuesWorksheet.Cells(row +rowOffset, columnOffset + 3))
		if lcase(apiRemark) = "not in use" then
			'format whole row in grey text
			range.Font.Color = RGB(128,128,128)
		else
			range.Font.Color = RGB(0,112,192)
		end if
	next
end function

function setHyperlinksForBusinessRules(valuesWorksheet, messageTypesList, rowOffset, typeColumn, remarkColumn)
	'loop types column values
	dim brSheetName
	brSheetName = "Business Rules"
	dim row
	for row = 0 to messageTypesList.Count - 1
		dim typeName
		dim secondaryName
		dim formatted
		formatted = false
		'get the typeName
		typeName = valuesWorksheet.Cells(row +rowOffset, typeColumn).Value
		'get the secondaryName
		secondaryName = valuesWorksheet.Cells(row +rowOffset, remarkColumn-1).Value
		'get the remark
		dim currentCell
		set currentCell = valuesWorksheet.Cells(row +rowOffset, remarkColumn)
		dim apiRemark
		apiRemark = currentCell.Value
		'In case of enumeration the typeName will contain no spaces and 
		if len(typeName) > 0 and UCase(apiRemark) = "SEE BUSINESS RULES" then
			'find the value in the messageTypesList
			dim j
			for j = 0 to messageTypesList.count -1
				if typeName = messageTypesList(j)(0) then
					'found it, set hyperlink formula and 
					'example formula: =HYPERLINK(CONCATENATE("#'Business Rules'!D";MATCH("VehicleDetailUseCodeType";'Business Rules'!$D$1:$D$3018;0));"SEE BUSINESS RULES")
					'new formula: =HYPERLINK(CONCATENATE("#'Business Rules'!D";IF(ISNA(MATCH("OtherCompanyNamee";'Business Rules'!$D$1:$D$9999;0));MATCH("OtherCompanyName";'Business Rules'!$D$1:$D$9999;0);MATCH("OtherCompanyNamee";'Business Rules'!$D$1:$D$9999;0)));"SEE BUSINESS RULES")
					dim formula
					'formula = "=HYPERLINK(CONCATENATE(""#'" & brSheetName & "'!D"",MATCH("""& trim(typeName) &""",'" & brSheetName & "'!$D$1:$D$9999,0)),""" & apiRemark & """)"
					formula = "=HYPERLINK(CONCATENATE(""#'" & brSheetName & "'!D"",IF(ISNA(MATCH("""& trim(typeName) &""",'" & brSheetName & "'!$D$1:$D$9999,0)),MATCH("""& trim(secondaryName) &""",'" & brSheetName & "'!$D$1:$D$9999,0),MATCH("""& trim(typeName) &""",'" & brSheetName & "'!$D$1:$D$9999,0))),""" & apiRemark & """)"

					currentCell.Formula = formula
					'set formatting bold + underline
					currentCell.Font.Bold = true
					currentCell.Font.Color = RGB(0,112,192)
					currentCell.Font.Underline = xlUnderlineStyleSingle
					'currentCell.Interior.Color = RGB(255,255,153) 'yellow 'REBRANDING: This line has been commented'
					'remember that we formatted
					formatted = true
					exit for
				end if
			next
		end if
		'undo formatting
		if not formatted and not lcase(apiRemark) = "not in use" then
			currentCell.Font.Bold = false
			currentCell.Font.Color = RGB(0,112,192)
			currentCell.Font.Underline = xlUnderlineStyleNone
		end if
	next
end function

function formatFSsheet(excelOutput, worksheet, message, messageOutputList, rowOffset, messageTypesList, valuesWorksheetName, enumTypesColumn)
	'set outline setting to abow
	worksheet.Outline.SummaryRow = xlAbove
	'remove all grouping
	excelOutput.ungroupAll worksheet
	'group the levels in the message
	dim i
	for i = 1 to message.MessageDepth - 2
		'maximum number of nested groups in Excel is 8
		if i >= 8 then 
			exit for
		end if
		dim j
		dim startRow
		startRow = -1
		for j = 0 to messageOutputList.Count - 1
			dim row
			row = j + rowOffset
			if j = messageOutputList.Count - 1 then
				row = row + 1 'the end row should be included in the grouping
			end if
			dim level
			dim currentName
			currentName = messageOutputList(j)(0)
			level = getLevel (currentName)
			if level <= i or j = messageOutputList.Count - 1 then
'				'debug
'				Session.Output "Level: " & level & " i: " & i & " j: " & j & " startrow: " & startRow & " currentName: " & currentName
				'check if we have a startrow
				if startRow >= 0  and startRow + 1 < row -1 then
					'group here
					on error resume next 'in case the grouping doesn't work
					dim range
					set range = worksheet.Range(worksheet.Cells(startRow + 1, 1), worksheet.Cells(row - 1,1))
'					'debug
'					Session.Output "grouping range: " & range.Address
					range.EntireRow.Group
					Err.Clear
					on error goto 0
				end if
				'remember the new startRow
				startRow = row
			end if
		next
	next
	'start at the rowoffset for column B. Set Root element and complex types bold.
	'set Choice elements bold, italic and in a different color
	i = 0
	for each row in messageOutputList
		dim currentCell
		set currentCell = worksheet.Cells(rowOffset + i, 2)
		'reset format
		'currentCell.Font.Name = "Arial" 'REBRANDING: font name changed from Calibri to Arial'
		currentCell.Font.Color = RGB(0,0,0) 'black
		currentCell.Font.Bold = false
		currentCell.Font.Italic = false
		'Set Root element and complex types bold.
		if i = 0 _
		  or lcase(row(5)) = "complex" then
			currentCell.Font.Bold = true  
		end if
		'set Choice elements bold, italic and in a different color
		if lcase(right(row(0), len("choice"))) = "choice" then
			currentCell.Font.Bold = true
			currentCell.Font.Italic = true
			currentCell.Font.Color = RGB(0,112,192) 'dark blue 'REBRANDING: Color changed from RGB(226,107,10) to RGB(0,13,110)'
		end if
		i = i  + 1
	next
	'set hyperlinks to enumerations
	setHyperlinksForEnums worksheet, messageOutputList, rowOffset, messageTypesList, valuesWorksheetName, enumTypesColumn, RGB(0,112,192), RGB(0,0,0) 
end function

function formatContentSheet(excelOutput, worksheet, message, messageOutputList, rowOffset, messageTypesList, valuesWorksheetName, enumTypesColumn)
	'set the name of the message
	worksheet.Cells(2,2).Value = message.Name
	'set outline setting to abow
	worksheet.Outline.SummaryRow = xlAbove
	'remove all grouping
	excelOutput.ungroupAll worksheet
	'group the levels in the message
	dim i
	for i = 1 to message.MessageDepth - 2
		'maximum number of nested groups in Excel is 8
		if i >= 8 then 
			exit for
		end if
		dim j
		dim startRow
		startRow = -1
		for j = 0 to messageOutputList.Count - 1
			dim row
			row = j + rowOffset
			if j = messageOutputList.Count - 1 then
				row = row + 1 'the end row should be included in the grouping
			end if
			dim level
			dim currentName
			currentName = messageOutputList(j)(0)
			level = getLevel (currentName)
			if level <= i or j = messageOutputList.Count - 1 then
'				'debug
'				Session.Output "Level: " & level & " i: " & i & " j: " & j & " startrow: " & startRow & " currentName: " & currentName
				'check if we have a startrow
				if startRow >= 0  and startRow + 1 < row -1 then
					'group here
					on error resume next 'in case the grouping doesn't work
					dim range
					set range = worksheet.Range(worksheet.Cells(startRow + 1, 1), worksheet.Cells(row - 1,1))
'					'debug
'					Session.Output "grouping range: " & range.Address
					range.EntireRow.Group
					Err.Clear
					on error goto 0
				end if
				'remember the new startRow
				startRow = row
			end if
		next
	next
	'set hyperlinks to enumerations
	setHyperlinksForEnums worksheet, messageOutputList, rowOffset, messageTypesList, valuesWorksheetName, enumTypesColumn, RGB(0,112,192), RGB(0,112,192)
	'set hyperlinks for business rules
	setHyperlinksForBusinessRules worksheet, messageOutputList, rowOffset, 4, 8
end function


function setHyperlinksForEnums(worksheet, messageOutputList, rowOffset, messageTypesList, valuesWorksheetName, enumTypesColumn, formattedColor, unformattedColor)
	dim enumTypesColumnLetter
	enumTypesColumnLetter = chr(64 + enumTypesColumn)
	'loop types column values
	dim column
	column = 7
	dim row
	for row = 0 to messageOutputList.Count - 1
		dim typeName
		dim formatted
		formatted = false
		'get the typeName
		dim currentCell
		set currentCell = worksheet.Cells(row +rowOffset, column)
		typeName = currentCell.Value
		'In case of enumeration the typeName will contain no spaces and 
		if len(typeName) > 0 and len(trim(typeName)) = len(typeName) then
			'find the value in the messageTypesList
			dim j
			for j = 0 to messageTypesList.count -1
				if typeName = messageTypesList(j)(0) then
					'found it, set hyperlink formula and 
					'example formula:  =HYPERLINK(CONCATENATE("#'Value Lists_IN'!D",MATCH("ValidationAnswerYesNoType",'Value Lists_IN'!$D$1:$D$1739,0)),"ValidationAnswerYesNoType")
					dim formula
					formula = "=HYPERLINK(CONCATENATE(""#'" & valuesWorksheetName & "'!" & enumTypesColumnLetter & """,MATCH("""& typeName &""",'" & valuesWorksheetName & "'!$" & enumTypesColumnLetter & "$1:$" & enumTypesColumnLetter & "$9999,0)),""" & typeName & """)"
					currentCell.Formula = formula
					'set formatting
					currentCell.Font.Color = formattedColor 'RGB(0,112,192) 'dark blue 
					currentCell.Font.Bold = true
					currentCell.Font.Underline = xlUnderlineStyleSingle
					'remember that we formatted
					formatted = true
					exit for
				end if
			next
		end if
		if not formatted then
			'default formatting
			currentCell.Font.Color = unformattedColor 'dark blue 'REBRANDING: Color changed from RGB(0,0,0) to RGB(0,13,110)'
			currentCell.Font.Bold = false
			currentCell.Font.Underline = xlUnderlineStyleNone
		end if
	next
end function

function getLevel(nodeName)
	'count the number of prefixing spaces * 6
	dim nbrSpaces
	nbrSpaces = len(nodeName) - len(ltrim(nodename))
	getLevel = nbrSpaces / 6
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