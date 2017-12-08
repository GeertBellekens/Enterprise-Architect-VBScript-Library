'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: FixOCLConstraints
' Author: Geert Bellekens
' Purpose: Some entities are renamed (removed suffixe _Type) and the OCL constraints need to be adapted to that. 
' next to the updates the errors where . notations references a non existing element will be reported
' Date: 2017-02-07
'
'name of the output tab
const outPutName = "Fix OCL Constraints"
const messagingModelGUID = "{52A8DE61-6FAC-46f7-89E9-55700CE04977}"


sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'set timestamp for start
	Repository.WriteOutput outPutName,now() & " Starting fixing OCL constraints"  , 0
	'get the elements with OCL constraints
	dim OCLElements
	set OCLElements = getElementWithOCLConstraints()
	'if nothing found then we don't need to bother
	if OCLElements.Count > 0 then
		'remember the strings we don't want to fix
		dim skippedStrings
		set skippedStrings = CreateObject("System.Collections.ArrayList")
		'get the package ID string of the messaging model
		dim messagingPackageIDs
		dim messagingRoot as EA.Package
		set messagingRoot = Repository.GetPackageByGuid(messagingModelGUID)
		if not messagingRoot is nothing then
			messagingPackageIDs = getPackageTreeIDString(messagingRoot)
			'fix the OCLElements
			dim OCLElement as EA.Element
			for each OCLElement in OCLElements
				fixOCLConstraints OCLElement, skippedStrings,messagingPackageIDs
			next
		else
			msgbox "GUID " & messagingModelGUID & " is not a valid GUID for the messaging model root package!" _
				,vbOKOnly+vbExclamation ,"Wrong messagingModelGUID!"
		end if 
	end if
	'set timestamp for end
	Repository.WriteOutput outPutName,now() & " Finished fixing OCL constraints"  , 0
end sub



function getElementWithOCLConstraints()
	'initialize empty
	set getElementWithOCLConstraints = CreateObject("System.Collections.ArrayList")
	'ask the user to select the messages package
	msgbox "Please select the Messages root package"
	dim messagesPackage
	set messagesPackage = selectPackage()
	'ask confirmation
	dim response
	if not messagesPackage is nothing then
		response = Msgbox("Fix OCL constraints on messages in package '" & messagesPackage.Name & "'?", vbYesNo+vbQuestion, "Fix OCL Constraints?")
		if response = vbYes then
			dim packageIDString
			packageIDString = getCurrentPackageTreeIDString()
			dim sqlGetOCLElements
			sqlGetOCLElements = "select distinct o.Object_ID from t_objectconstraint ocl " & _
								" inner join t_object o on ocl.Object_ID = o.Object_ID " & _
								" where ocl.ConstraintType = 'OCL' " & _
								" and o.Stereotype = 'XSDtopLevelElement' " & _
								" and o.Package_ID in (" & packageIDString & ")"
			set getElementWithOCLConstraints = getElementsFromQuery(sqlGetOCLElements)
		end if
	end if
end function

'fix the OCL constraints for this element
function fixOCLConstraints(OCLElement, skippedStrings, messagingPackageIDs)
	dim OCLConstraint as EA.Constraint
	for each OCLConstraint in OCLElement.Constraints
		'remember all fixed strings
		dim fixedStrings
		set fixedStrings = CreateObject("System.Collections.ArrayList")
		'remove all _Type suffixes
		dim constraintText
		constraintText = Repository.GetFormatFromField("TXT",OCLConstraint.Notes)
		'create a regular expression to get the elements with _Type as suffix
		Dim regExp  
		Set regExp = CreateObject("VBScript.RegExp")
		regExp.Global = True   
		regExp.IgnoreCase = False
		'execut the regex pattern	
		regExp.Pattern = "\b[\w]*_Type\b"
		dim matches
		set matches = regExp.Execute(constraintText)
		dim match
		'loop the matches and replace the _Type with empty string
		For each match in matches
			dim matchText
			matchText = match.Value
			if not fixedStrings.Contains(matchText) and not skippedStrings.Contains(matchText) then
				dim replacementText
				'remove the last 5 characters (_Type)
				replacementText = left(matchText,len(matchText) - 5) 
				'check if we need to fix it
				if needsFixing(matchText,replacementText,messagingPackageIDs,OCLConstraint,OCLElement) then
				'tell the user what we are doing
				Repository.WriteOutput outPutName,now() & " Removing '_Type' suffix on for '" & matchText & "' in constraint '" & OCLConstraint.Name & "' on element '" &  OCLElement.Name & "'"  , OCLElement.ElementID
				'replace in the constraint
				constraintText = replace(constraintText,matchText,replacementText)
				'add the text to the list of fixed strings
				fixedStrings.Add matchText
				else
					'doesn't need fixing, add it to the skipped strings
					skippedStrings.Add matchText
				end if
			end if
		next
		'check if we fixed something
		if fixedStrings.Count > 0 then
			'put the constraint back
			OCLConstraint.Notes = Repository.GetFieldFromFormat("TXT",constraintText)
			'save the constraint
			OCLConstraint.Update
		end if
		'TODO: report possible wrong references?
	next
end function

function needsFixing(matchText,replacementText,messagingPackageIDs,OCLConstraint,OCLElement)
	dim exactMatch
	dim replacementMatch
	exactMatch = hasMatch(matchText,messagingPackageIDs)
	replacementMatch = hasMatch(replacementText,messagingPackageIDs)
	if replacementMatch then
		'needs fixing
		needsFixing = true
		if exactMatch then
			'Issue warning because both exist
			needsFixing = false
			Repository.WriteOutput outPutName, now() & " ERROR found both '" & matchText & "' and '" & replacementText & "' OCL constraints have not been changed for this type"  , OCLElement.ElementID
		end if
	else
		'doesn't need fixing
		needsFixing = false
		if not exactMatch then
			Repository.WriteOutput outPutName, now() & " ERROR '" & matchText & "' was found in constraint '" & OCLConstraint.Name & "' on element '" &  OCLElement.Name & "' and does not match an existing element (with or without '_Type')"  , OCLElement.ElementID
		end if
	end if
end function

function hasMatch(nameString,messagingPackageIDs)
	dim sqlFindMatch 
	sqlFindMatch = "select o.Object_ID from t_object o " & _
						" where o.Object_Type in ('Class','Enumeration') " & _
						" and o.name = '" & nameString & "' " & _
						" and o.Package_ID in (" & messagingPackageIDs & ") "
	dim matches
	set matches = getElementsFromQuery(sqlFindMatch)
	if matches.Count > 0 then
		hasMatch = true
	else
		hasMatch = false
	end if
end function

main