'[path=\Projects\Project B\Baloise Scripts]
'[group=Baloise Scripts]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Copy notes down main
' Author: Geert Bellekens
' Purpose: Copy the notes of the selected item down to all derived items
' Date: 2019-07-05
'
const elementTag = "sourceElement"
const attributeTag = "sourceAttribute"
const associationTag = "sourceAssociation"
const splitCodesEnumGUID = "{2BC858A9-A976-47d1-BDE0-ED04D1204E30}"

const outPutName = "Clear All Notes"
	

	

function copyNotesDown()
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'inform user
	Repository.WriteOutput outPutName, now() & " Starting copy notes down", 0
	'get split tag notes
	dim splitTagNames
	set splitTagNames = getSplitTagNames
	'get selected item
	dim contextItemType
	contextItemType = Repository.GetContextItemType
	dim contextItem
	set contextItem = Repository.GetContextObject
	dim response
	select case contextItemType
		case otElement
			'get confirmation from user
			response = msgbox("Copy notes to derived elements for " & contextItem.Type & " '" & contextItem.Name & "'?"_
							, vbYesNo+vbQuestion, "Copy notes to derived Elements?")
			if response = vbYes then
				copyElementNotes contextItem, splitTagNames
			end if
		case otAttribute
			'get confirmation from user
			response = msgbox("Copy notes to derived attributes for attribute '" & contextItem.Name & "'?"_
							, vbYesNo+vbQuestion, "Copy notes to derived Attributes?")
			if response = vbYes then
				copyAttributeNotes contextItem, splitTagNames
			end if
		case otConnector
			'get confirmation from user
			response = msgbox("Copy notes to derived associations?"_
							, vbYesNo+vbQuestion, "Copy notes to derived Associations?")
			if response = vbYes then
				copyConnectorNotes contextItem, splitTagNames
			end if
		case else
			MsgBox "Please select an Element, Attribute or Relation before executing this script", vbExclamation, "Wrong selection"
	end select
	'inform user
	Repository.WriteOutput outPutName, now() & " Finished copy notes down", 0
end function

function copyElementNotes(element, splitTagNames)
	'find derived elements
	dim sqlGetDerivedElements
	sqlGetDerivedElements = "select oo.Object_ID from t_object o                          " & vbNewLine & _
							" inner join t_objectproperties tv on tv.Value = o.ea_guid   " & vbNewLine & _
							"					and tv.Property = '" & elementTag & "'   " & vbNewLine & _
							" inner join t_object oo on oo.Object_ID = tv.Object_ID      " & vbNewLine & _
							" where o.ea_guid = '" & element.ElementGUID &"' "
	dim derivedElements
	set derivedElements = getElementsFromQuery(sqlGetDerivedElements)
	'loop derived elemnents
	dim derivedElement as EA.Element
	for each derivedElement in derivedElements
		'update notes
		derivedElement.Notes = removeElementSplitTags(element.Notes, splitTagNames) 
		derivedElement.Update
		'inform user
		Repository.WriteOutput outPutName, now() & " Updated notes for '" & derivedElement.FQName & "'" , 0
		'recurse further	
		copyElementNotes derivedElement, splitTagNames
	next
end function
function removeElementSplitTags(notes, splitTagNames)
	'default return the same
	removeElementSplitTags = notes
	dim cleanNotes
	Dim regEx
	Set regEx = CreateObject("VBScript.RegExp")
	regEx.IgnoreCase = True
	regEx.Multiline = True
	'first with parentheses => (all) + newline
	regEx.Pattern = "^\((" & Join(splitTagNames.ToArray(),"|") & ")\)[\r\n]+"
	cleanNotes = regEx.Replace(notes, "")
	'then without parentheses => all + newline
	regEx.Pattern = "^(" & Join(splitTagNames.ToArray(),"|") & ")[\r\n]+"
	cleanNotes = regEx.Replace(cleanNotes, "")
	'return
	removeElementSplitTags = trim(cleanNotes)
end function


function copyAttributeNotes(attribute,splitTagNames)
	'find derived attributes
	dim sqlGetDerivedAttributes
	sqlGetDerivedAttributes = "select aa.ID from t_attribute a                               " & vbNewLine & _
							" inner join t_attributetag tv on tv.VALUE = a.ea_guid          " & vbNewLine & _
							" 					and tv.Property = '" & attributeTag & "'    " & vbNewLine & _
							" inner join t_attribute aa on aa.ID = tv.ElementID             " & vbNewLine & _
							" where a.ea_guid = '" & attribute.AttributeGUID &"'           "
	dim derivedAttributes
	set derivedAttributes = getAttributesFromQuery(sqlGetDerivedAttributes)
	'loop derived elemnents
	dim derivedAttribute as EA.Attribute
	for each derivedAttribute in derivedAttributes
		'update notes
		derivedAttribute.Notes = removeAttributeSplitTags(attribute.Notes, splitTagNames) 
		derivedAttribute.Update
		'get owner
		dim derivedElement as EA.Element
		set derivedElement = Repository.GetElementByID(derivedAttribute.ParentID)
		'inform user
		Repository.WriteOutput outPutName, now() & " Updated notes for '" & derivedElement.FQName & "." & derivedAttribute.Name & "'" , 0
		'recurse further	
		copyAttributeNotes derivedAttribute, splitTagNames
	next
end function

function removeAttributeSplitTags(notes, splitTagNames)
	'default return the same
	removeAttributeSplitTags = notes
	dim cleanNotes
	dim firstPart
	dim dashPosition
	dashPosition = InStr(notes,"-")
	if dashPosition <= 0 then
		exit function
	end if
	firstPart =  left(notes,dashPosition)
	Dim regEx
	Set regEx = CreateObject("VBScript.RegExp")
	regEx.IgnoreCase = True
	regEx.Multiline = True
	regEx.Pattern = "^(?:\[\s[0-9]{1,3}\s)-\s(" & Join(splitTagNames.ToArray(),"|") & ")[\s]+-"
	cleanNotes = regEx.Replace(notes, firstPart)
	'return
	removeAttributeSplitTags = cleanNotes
end function

function getSplitTagNames()
	dim splitTagNames
	set splitTagNames = CreateObject("System.Collections.ArrayList")
	dim splitTagEnum as EA.Element
	'get enumeration
	set splitTagEnum = Repository.GetElementByGuid(splitCodesEnumGUID)
	'get values
	dim tagValue as EA.Attribute
	for each tagValue in splitTagEnum.Attributes
		splitTagNames.Add tagValue.Name
	next
	'return
	set getSplitTagNames = splitTagNames
end function

function copyConnectorNotes(connector, splitTagNames)
	'find derived connectors
	dim sqlGetDerivedConnectors
	sqlGetDerivedConnectors =   "select cc.Connector_ID from t_connector c                      " & vbNewLine & _
								" inner join t_connectortag tv on tv.VALUE = c.ea_guid          " & vbNewLine & _
								" 					and tv.Property = '" & associationTag & "'  " & vbNewLine & _
								" inner join t_connector cc on cc.Connector_ID = tv.ElementID   " & vbNewLine & _
								" where c.ea_guid = '" & connector.ConnectorGUID &"'                " 
	dim derivedConnectors
	set derivedConnectors = getConnectorsFromQuery(sqlGetDerivedConnectors)
	'loop derived elemnents
	dim derivedConnector as EA.Connector
	for each derivedConnector in derivedConnectors
		'update notes
		derivedConnector.Notes = connector.Notes
		derivedConnector.Update
		'get source
		dim derivedElement as EA.Element
		set derivedElement = Repository.GetElementByID(derivedConnector.ClientID)
		'get target
		dim targetElement
		set targetElement = Repository.GetElementByID(derivedConnector.SupplierID)
		'inform user
		Repository.WriteOutput outPutName, now() & " Updated notes on relation from '" & derivedElement.FQName & "' to '" & targetElement.Name & "'" , 0
		'recurse further	
		copyConnectorNotes derivedConnector, splitTagNames
	next
end function
