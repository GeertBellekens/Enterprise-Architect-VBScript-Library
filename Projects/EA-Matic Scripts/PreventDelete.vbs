'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
option explicit

'!INC Wrappers.Include

'EA-Matic
'Author: Geert Bellekens
'This script will prevent any element to be deleted if it is still used as a type in either a parameter
'or an attribute. The can be overridden by first prepending the name with DELETED_

function EA_OnPreDeleteElement(Info)
     'Start by setting false
     EA_OnPreDeleteElement = false
     dim usage
     'get the elementID from Info
     dim elementID
     elementID = Info.Get("ElementID")
     'get the element being deleted
     dim element as EA.Element
     set element = Repository.GetElementByID(elementID)
     'Manual override is triggered by the name. If it starts with DELETED_ then the element may be deleted.
     if Left(element.name,LEN("DELETED_")) = "DELETED_" then
        'OK the element may be deleted
        EA_OnPreDeleteElement = true
     else
		dim usedInSchema
		usedInSchema = checkUsedInSchema(element, usage)
		dim usedAsType
		usedAsType = checkUsedAsAttributeType(element, usage)
		if usedInSchema or usedAsType then
			'don't allow delete if still used in schema or type
			EA_OnPredeleteElement = false
		else
			EA_OnPredeleteElement = true
		end if
     end if
     if EA_OnPredeleteElement = false then
          'NO the element cannot be deleted
          MsgBox "Element '" & element.Name & "' is still used by:" & vbNewLine _
			& usage , vbExclamation, "Cannot delete element"
     end if
end function

function checkUsedInSchema(element, usage)
	dim used
	'check if element still used in a shema
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_document d                    " & vbNewLine & _
				" inner join t_object o on o.ea_guid = d.ElementID        " & vbNewLine & _
				" where d.StrContent like '%" & element.ElementGUID & "%' " & vbNewLine & _
				" and d.ElementType = 'SC_MessageProfile'                 "
	dim schemaElements
	set schemaElements = Repository.GetElementSet(sqlGetData, 2)
	dim schemaElement as EA.Element
	if schemaElements.Count = 0 then
		used = false
	else 
		used = true
		for each schemaElement in schemaElements
			usage = usage & vbNewLine & schemaElement.Name & " - " & schemaElement.ElementGUID
		next
	end if
	'return
	checkUsedInSchema = used
end function

function checkUsedAsAttributeType(element, usage)
	dim used
	'check if element still used in as type in an attribute from an object in another folder
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_attribute a                    " & vbNewLine & _
				" inner join t_object o on o.Object_ID = a.Object_ID      " & vbNewLine & _
				" inner join t_object oo on oo.Object_ID = a.Classifier   " & vbNewLine & _
				" 					and o.Package_ID != oo.Package_ID     " & vbNewLine & _
				" where a.Classifier = " & element.ElementID
	dim schemaElements
	set schemaElements = Repository.GetElementSet(sqlGetData, 2)
	dim schemaElement as EA.Element
	if schemaElements.Count = 0 then
		used = false
	else 
		used = true
		for each schemaElement in schemaElements
			usage = usage & vbNewLine & schemaElement.Name & " - " & schemaElement.ElementGUID
		next
	end if
	'return
	checkUsedAsAttributeType = used
end function

'function test
'	dim EA_OnPreDeleteElement
'	'Start by setting false
'	EA_OnPreDeleteElement = false
'	dim usage
'	'get the elementID from Info
'	dim elementID
'	'elementID = Info.Get("ElementID")
'	elementID = 32848
'	'get the element being deleted
'	dim element as EA.Element
'	set element = Repository.GetElementByID(elementID)
'	'Manual override is triggered by the name. If it starts with DELETED_ then the element may be deleted.
'	if Left(element.name,LEN("DELETED_")) = "DELETED_" then
'		'OK the element may be deleted
'		EA_OnPreDeleteElement = true
'	else
'		'check if element still used in a shema
'		dim sqlGetData
'		sqlGetData = "select o.Object_ID from t_document d                    " & vbNewLine & _
'					" inner join t_object o on o.ea_guid = d.ElementID        " & vbNewLine & _
'					" where d.StrContent like '%" & element.ElementGUID & "%' " & vbNewLine & _
'					" and d.ElementType = 'SC_MessageProfile'                 "
'		dim schemaElements
'		set schemaElements = Repository.GetElementSet(sqlGetData, 2)
'		dim schemaElement as EA.Element
'		if schemaElements.Count = 0 then
'			EA_OnPreDeleteElement = true
'		else 
'			EA_OnPreDeleteElement = false
'			for each schemaElement in schemaElements
'				usage = usage & vbNewLine & "-" & schemaElement.Name & " - " & schemaElement.ElementGUID
'			next
'		end if
'	end if
'	if EA_OnPredeleteElement = false then
'		'NO the element cannot be deleted
'		MsgBox "I'm sorry Dave, I'm afraid I can't do that" & vbNewLine _
'		& element.name & " is used as type in: " & usage , vbExclamation, "Cannot delete element"
'	end if
'end function
'
'test

