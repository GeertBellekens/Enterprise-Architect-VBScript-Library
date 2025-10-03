'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
option explicit
'EA-Matic
!INC Local Scripts.EAConstants-VBScript


'
' Script Name: Default Business Process Block diagram type
' Author: Geert Bellekens
' Purpose: Update the applied metamodel for the diagrams under the Business Process Block element
' Date: 2025-08-29

function EA_OnPostNewDiagram(Info)
	msgbox "EA_OnPostNewDiagram"
	dim returnValue
	returnValue = false
	 'Add code here
	dim diagram as EA.Diagram
	set diagram = Repository.GetDiagramByID(Info.Get("DiagramID"))
	msgbox diagram.MetaType
	'get owner
	if diagram.ParentID = 0 _
	  or not diagram.MetaType = "BPMN2.0::Business Process" then
		exit function
	end if
	dim owner as EA.Element 
	set owner = Repository.GetElementByID(diagram.ParentID)
	if owner.Stereotype = "ArchiMate_BusinessProcess" then
		diagram.ExtendedStyle = setValueForKey(diagram.ExtendedStyle, "MDGView", "Elia Modelling::Business Process Diagram")
		diagram.Update
		msgbox "updated"
		returnValue = true
	end if
	msgbox "Finished EA_OnPostNewDiagram"
	'return
	EA_OnPostNewDiagram = returnValue
end function

function setValueForKeyEx(searchString, key, value, separator)
	dim keyValues
	set keyValues = getKeyValuePairsEx(searchString, separator)
	if not keyValues.exists(key) then
		dim newKeyValues
		set newKeyValues = CreateObject("Scripting.Dictionary")
		newKeyValues.Add key, value
		'add all existing kevalues
		dim oldKey
		for each oldKey in keyValues.Keys
			newKeyValues.Add oldKey, keyValues(oldKey)
		next
		'replace dictionary
		set keyValues = newKeyValues
	else
		'set value
		keyValues(key) = value
	end if
	'join keyValuePairs
	dim joinedString
	joinedString = joinKeyValuePairsEx(keyValues, separator)
	'return
	setValueForKeyEx = joinedString
end function

function setValueForKey(searchString, key, value)
	setValueForKey = setValueForKeyEx(searchString, key, value, ";")
end function

function joinKeyValuePairsEx(keyValuePairs, separator)
	dim joinedString
	joinedString = ""
	dim key
	for each key in keyValuePairs.Keys
		joinedString = joinedString & key & "=" & keyValuePairs(key) & separator
	next
	'return
	joinKeyValuePairsEx = joinedString
end function

function getKeyValuePairsEx(keyValueString, separator)
	dim keyValuePairDictionary
	Set keyValuePairDictionary = CreateObject("Scripting.Dictionary")
	dim keyValuePairs
	'first split in keyvalue pairs using ";"
	keyValuePairs = split(keyValueString,separator)
	'then loop the key value pairs
	dim keyValuePairString
	for each keyValuePairString in keyValuePairs
		'and split them usign "=" as delimiter
		dim keyValuePair
		if instr(keyValuePairString,"=") > 0 then
			keyValuePair = split(keyValuePairString,"=")
			if UBound(keyValuePair) = 1 then
				'set the value, don't care about duplicate keys
				keyValuePairDictionary(keyValuePair(0)) = keyValuePair(1)
			elseif UBound(keyValuePair) > 1 then
				'get the part after the first =
				dim start
				start = instr(keyValuePairString, "=") + 1
				keyValuePairDictionary(keyValuePair(0)) = mid(keyValuePairString,start)
			end if
		end if
	next
	'return
	set getKeyValuePairsEx = keyValuePairDictionary
end function