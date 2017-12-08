'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Import Message Rules
' Author: Geert Bellekens
' Purpose: Import Message Rules from Excel files
' Date: 2017-03-28
'

'name of the output tab
const outPutName = "Import message Rules"

sub main
	
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	
	'let the user select the package to store the rules in
	msgbox "Please select the package to store the message rules"
	dim userSelectedPackage
	set userSelectedPackage = selectPackage()
	if not userSelectedPackage is nothing then
		if isRequireUserLockEnabled() then
			if not userSelectedPackage.ApplyUserLock() then
				msgbox "Please apply user lock to the selected package",vbOKOnly+vbExclamation,"Selected Package not locked!"
				exit sub
			end if
		end if
		'get the message rules and save them in the given package
		getRulesFromExcel userSelectedPackage 
	end if
end sub


function getRulesFromExcel(userSelectedPackage)
	dim sourceExcelFile
	set sourceExcelFile = new ExcelFile
	sourceExcelFile.openUserSelectedFile()
	if len(sourceExcelFile.FileName)> 0 then
		'tell the user we are starting
		Repository.WriteOutput outPutName, now() & " Starting Import message Rules from file'" & sourceExcelFile.FileName & "'",0
		dim sheet
		for each sheet in sourceExcelFile.worksheets
			dim currentPackage as EA.Package
			'create new package per sheet
			set currentPackage = userSelectedPackage.Packages.AddNew(sheet.Name,"")
			currentPackage.Update
			'get the contents of the sheet
			dim contents
			contents = sourceExcelFile.getContents(sheet)
			dim indexes
			set indexes = getIndexesBasedOnHeaders(contents)
			dim i
			dim j
			for i = 2 to Ubound(contents,1)  step +1
				'get the path
				dim path
				dim ruleName
				dim ruleID
				dim ruleReason
				dim action
				'initialize fields
				set path = CreateObject("System.Collections.ArrayList")
				ruleName = ""
				ruleID = ""
				ruleReason = ""
				action = ""
				'loop the contents
				for j = 1 to Ubound(contents,2)  step +1
					dim currentField
					'if it is one of the level fields add it to the path
					currentField = contents(i,j)
					if j < Ubound(indexes.Keys()) then
						if len(currentField) > 0 _
						  AND len(indexes.Keys()(j)) < 4 _
						  AND Ucase(Left(indexes.Keys()(j),1)) = "L" then
							path.Add currentField
						end if
					end if
					'get the name
					if j = indexes("NEW Test Rule") then
						ruleName = currentField
					end if
					'get the ID
					if j = indexes("Test Rule ID") then
						ruleID = currentField
					end if
					'get the ruleReason )Error reason
					if j = indexes("Error reason") then
						ruleReason = currentField
					end if
					'get the action
					if j = indexes("Action") then
						action = currentField
					end if					
					if path.Count > 0 _
					AND len(ruleName) > 0 _
					AND len(ruleID) > 0 _
					AND len(ruleReason) > 0 then
						'not if marked to delete
						if lcase(action) <> "delete" then
							'create Message Rule
							createMessageRule path, ruleName, ruleID, ruleReason, currentPackage
						end if
						'found everything we need, exit the for loop
						exit for
					else
						if j = Ubound(contents,2) _
						AND lcase(action) <> "delete" _
						AND NOT (path.Count = 0 _
						AND len(ruleName) = 0 _
						AND len(ruleID) = 0 _
						AND len(ruleReason) = 0) then
							'report error (not for all blanks)
							Repository.WriteOutput outPutName, now() & " ERROR: Could not create Message Rule for row :" & i & " in sheet '" & sheet.Name & "'",0
						end if
					end if
				next
			next
		next
		'tell the user we are finished
		Repository.WriteOutput outPutName, now() & " Finished Import message Rules from file'" & sourceExcelFile.FileName & "'",0
	end if
end function

Function getIndexesBasedOnHeaders(contents)
	dim j
	dim indexes
	set indexes = CreateObject("Scripting.Dictionary")
	for j = 1 to Ubound(contents,2)  step +1
		if not indexes.Exists(contents(1,j)) then
			indexes.Add contents(1,j) , j
		end if
	next
	'return indexes
	set getIndexesBasedOnHeaders = indexes
end function

function createMessageRule(path, ruleName, ruleID, ruleReason, ownerPackage)
	Repository.WriteOutput outPutName, now() & " Adding Rule'" & ruleID & "'",0
	'create MessageRule
	dim messageRule as EA.Element
	set messageRule = ownerPackage.Elements.AddNew(ruleID,"Test")
	messageRule.StereotypeEx = "Message Test Rule"
	messageRule.Notes = ruleName
	messageRule.Update
	'add the tagged for the rule reason
	dim reasonTaggedValue as EA.TaggedValue
	set reasonTaggedValue = getExistingOrNewTaggedValue(messageRule, "Error Reason")
	reasonTaggedValue.Notes = ruleReason
	reasonTaggedValue.Update
	'add the tagged value with the concatenated path
	dim pathTaggedValue as EA.TaggedValue
	set pathTaggedValue = getExistingOrNewTaggedValue(messageRule, "Constraint Path")
	pathTaggedValue.Value = Join(path.ToArray(),".")
	pathTaggedValue.Update
	'link the message rule to the message it is related to	
	dim relatedMessageObjects
	set relatedMessageObjects = getRelatedMessageObjects(path)
	dim relatedMessage as EA.Element
	for each relatedMessage in relatedMessageObjects
		dim messageLink as EA.Connector
		set messageLink = messageRule.Connectors.AddNew("","Dependency")
		messageLink.SupplierID = relatedMessage.ElementID
		messageLink.Update
	next
end function

function getRelatedMessageObjects(path)
	dim getMessageObjectsSQL 
	getMessageObjectsSQL = "select o.Object_ID from t_object o                     " & _
							" inner join t_package p on o.Package_ID = p.Package_ID" & _
							" where o.Stereotype = 'XSDtopLevelElement'             " & _
							" and p.name = '" & path(0) & " '                      "
	dim messageObjects
	set messageObjects = getElementsFromQuery(getMessageObjectsSQL)
	set getRelatedMessageObjects = messageObjects
end function
main