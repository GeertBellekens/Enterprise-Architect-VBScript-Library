'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: PreventDuplicates
' Author: Geert Bellekens
' Purpose: Prevents duplicate elements based on a list of element types and stereotypes
' Date: 2020-10-10
'
'EA-Matic


Const combinationsToCheck =  "Class,Table;Activity,null;Requirement,Term"

function EA_OnPostNewElement(Info)
	'get the elementID from Info
	dim elementID
	elementID = Info.Get("ElementID")
	dim element as EA.Element
	set element = Repository.GetElementByID(elementID)
	dim renamed
	renamed = renameDuplicate(element)
	'return true
	EA_OnPostNewElement = true
end function

function renameDuplicate(element)
	'only if needed
	if not elementNeedsChecking(element) then
		exit function
	end if
	'default false
	renameDuplicate = false
	dim dup
	dup = hasDuplicateElements(element)
	if dup then
		MsgBox "An element with this the same name and stereotype already exists!", vbOKOnly + vbExclamation, "Duplicate Element detected"
		element.Name = element.Name & "_" & element.ElementID
		element.update
		'return true
		renameDuplicate = true
	end if
end function

function elementNeedsChecking(element)
	elementNeedsChecking = false
	dim combos
	combos = Split(combinationsToCheck, ";" )
	dim combo
	for each combo in combos
		dim typeStereo 
		typeStereo = Split(combo, ",")
		dim elementType
		elementType = typeStereo(0)
		'check type
		if lcase(element.Type) = lcase(elementType) _
			or lcase(elementType) = "null" then
			'check stereotype
			dim stereo
			stereo = typeStereo(1)
			if lcase(element.Stereotype) = lcase(stereo) _
			  or lcase(stereo) = "null" then
				'Found one that needs checking
				elementNeedsChecking = true
				exit function
			end if
		end if
	next
	
end function

function hasDuplicateElements(element)
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o               " & vbNewLine & _
				" where o.Name = '" & element.Name & "'            " & vbNewLine & _
				" and ( o.Stereotype = '" & element.Stereotype & "'" & vbNewLine & _
				" or o.Stereotype is null ) 					   " & vbNewLine & _
				" and o.Object_ID <> " & element.ElementID
	dim duplicates
	set duplicates = Repository.GetElementSet(sqlGetData,2)
	'return boolean
	if duplicates.Count > 0 then
		hasDuplicateElements = true
	else
		hasDuplicateElements = false
	end if
end function

Dim contextName
Dim contextStereo
Dim contextGUID

function EA_OnContextItemChanged(GUID, ot)
	 if ot = otElement then
		Dim contextElement 
		set contextElement = Repository.GetElementByGuid(GUID)
		if not contextElement is nothing then
			contextName = contextElement.Name
			contextStereo = contextElement.Stereotype
			contextGUID = contextElement.ElementGUID
		end if
	end if
end function


function EA_OnNotifyContextItemModified(GUID, ot)
	 'check if the name has been changed
	 if GUID = contextGUID then
		Dim contextElement 
		set contextElement = Repository.GetElementByGuid(GUID)
		'only check if the element has been renamed
		if contextName <> contextElement.Name _
		  or contextStereo <> contextElement.Stereotype then
			renameDuplicate(contextElement)
		end if
	 end if
end function