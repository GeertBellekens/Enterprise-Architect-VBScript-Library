'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Atrias Scripts.Util

'
' Script Name: Merge Packages
' Author: Geert Bellekens
' Purpose: Merge packages containing conceptually the same elements moving all references to references to the master element
' Date: 2016-07-28
'

dim outputTabName
outputTabName = "Merge Packages"

sub main
	Repository.CreateOutputTab outputTabName
	Repository.ClearOutput outputTabName
	Repository.EnsureOutputVisible outputTabName
	'select source logical
	dim DMPackage as EA.Package
	msgbox "select the DM package"
	set DMPackage = selectPackage()
	'select master package
	dim LDMPackage as EA.Package
	msgbox "select the LDM package"
	set LDMPackage = selectPackage()
	if not DMPackage is nothing and not LDMPackage is nothing then 
		dim response
		response = Msgbox("Merge package """ & DMPackage.Name & """ to package """ & LDMPackage.Name & """?", vbYesNoCancel+vbQuestion, "Merge Package")
		if response = vbYes then
			Repository.WriteOutput outputTabName, now() & ": Starting merge package """ & DMPackage.Name & """ to package """ & LDMPackage.Name & """" ,0
			mergePackages DMPackage, LDMPackage
			Repository.WriteOutput outputTabName, now() & ": Finished merge package """ & DMPackage.Name & """ to package """ & LDMPackage.Name & """" ,0
		end if
	end if
	msgbox "finished"
end sub


function mergePackages(DMPackage, LDMPackage)
	
	dim LDMDictionary
	Set LDMDictionary = CreateObject("Scripting.Dictionary")
	'Logical Data Model
	Repository.WriteOutput outPutName, "Creating dictionary from Logical Data Model", 0
	addClassesToDictionary LDMPackage, LDMDictionary
	'Domain model
	dim DMDictionary
	Set DMDictionary = CreateObject("Scripting.Dictionary")
	'Logical Data Model
	Repository.WriteOutput outPutName, "Creating dictionary from Domain Model", 0
	addClassesToDictionary DMPackage, DMDictionary
	'create list of (recursive) owned elements in LDM
	for each DMClassName in DMDictionary.Keys
		if LDMDictionary.Exists(DMClassName) then
			mergeElements DMDictionary.Item(DMClassName), LDMDictionary.Item(DMClassName), DMDictionary, LDMDictionary
		else
			Repository.WriteOutput outputTabName, "Moving " & DMClassName ,0
			'Move the DM class to the LDM package
			DMDictionary.Item(DMClassName).PackageID = LDMPackage.PackageID
			DMDictionary.Item(DMClassName).Update
		end if
	next
end function

function addClassesToDictionary(package, dictionary)
	dim classElement as EA.Element
	dim subpackage as EA.Package
	'process owned elements
	for each classElement in package.Elements
		'this works for FISSES as well because they are classes with stereotype Message
		if (classElement.Type = "Class" OR classElement.Type = "Datatype" OR classElement.Type = "Enumeration") _
			AND len(classElement.Name) > 0 AND not dictionary.Exists(classElement.Name) then
			Repository.WriteOutput outPutName, "Adding element: " & classElement.Name, 0
			dictionary.Add classElement.Name,  classElement
		end if	
	next
	'process subpackages
	for each subpackage in package.Packages
		addClassesToDictionary subpackage, dictionary
	next
end function


'merging elements is the fastest if we use database updates
function mergeElements(fromElement, masterElement, DMDictionary, LDMDictionary)
	Repository.WriteOutput outputTabName, "Merging """ & fromElement.Name & """ to """ & masterElement.Name & """" ,0
	'attribute datatypes
	dim sqlUpdateDatatypes 
	slqUpdateDatatypes = "update t_attribute set classifier = " & masterElement.ElementID & " where classifier = " & fromElement.ElementID
	Repository.Execute slqUpdateDatatypes
	'connectors
	mergeConnectors fromElement, masterElement, DMDictionary, LDMDictionary
	'diagramObjects
	dim slqUpdateDiagramObjects
	slqUpdateDiagramObjects = "update t_diagramobjects set Object_ID= " & masterElement.ElementID & " where Object_ID = " & fromElement.ElementID
	Repository.Execute slqUpdateDiagramObjects 
	'parameter types
	dim sqlUpdateParameters
	sqlUpdateParameters = "update t_operationparams set Classifier =  " & masterElement.ElementID & "  where Classifier = " & fromElement.ElementID
	Repository.Execute sqlUpdateParameters
	'tagged values (elements)
	updateTaggedValues "t_objectproperties", masterElement.ElementGUID , fromElement.ElementGUID
	'tagged values (attributes)
	updateTaggedValues "t_attributetag", masterElement.ElementGUID , fromElement.ElementGUID
	'tagged values (operations)
	updateTaggedValues "t_operationtag", masterElement.ElementGUID , fromElement.ElementGUID
	'tagged values (connectors)
	updateTaggedValues "t_connectortag", masterElement.ElementGUID , fromElement.ElementGUID
	'merge references to attributes
	mergeAttributeReferences fromElement, masterElement
end function

function mergeConnectors (fromElement, masterElement, DMDictionary, LDMDictionary)
	'connectors (not for connectors to self)
	dim DMElementIDs, LDMElementIDs
	DMElementIDs = makeIDString(DMDictionary.Items)
	LDMElementIDs = makeIDString(LDMDictionary.Items)
	dim sqlUpdateConnectorSource, sqlUpdateConnectorTarget
	'not for connectors from DM elements to DM elements
	sqlUpdateConnectorSource = "update t_connector set Start_Object_ID = " & masterElement.ElementID & " where Start_Object_ID <> End_Object_ID and Start_Object_ID = " & fromElement.ElementID & " and End_Object_ID not in (" & DMElementIDs & ")"
	Repository.Execute sqlUpdateConnectorSource
	'not for connectors coming from DM or LDM elements
	sqlUpdateConnectorTarget = "update t_connector set End_Object_ID = " & masterElement.ElementID & " where Start_Object_ID <> End_Object_ID and End_Object_ID = " & fromElement.ElementID & " and Start_Object_ID not in (" & DMElementIDs & "," & LDMElementIDs & ")"
	Repository.Execute sqlUpdateConnectorTarget
end function




function mergeAttributeReferences(fromElement, masterElement)
	dim fromAttribute as EA.Attribute
	dim masterAttribute as EA.Attribute
	for each fromAttribute in fromElement.Attributes
		set masterAttribute = getCorrespondingAttribute(fromAttribute,masterElement)
		if not masterAttribute is nothing then
			'tagged values (attributes)
			updateTaggedValues "t_attributetag", masterAttribute.AttributeGUID, fromAttribute.AttributeGUID 
		end if	
	next
end function

function getCorrespondingAttribute(fromAttribute,masterElement)
	'initialize empty
	set getCorrespondingAttribute = nothing
	dim candidateAttribute as EA.Attribute
	for each candidateAttribute in masterElement.Attributes
		if candidateAttribute.Name = fromAttribute.Name then
			set getCorrespondingAttribute = candidateAttribute
			exit for
		end if
	next
end function

function updateTaggedValues (tableName, newValue, oldValue)
	dim sqlUpdateTaggedValues
	sqlUpdateTaggedValues = "update " & tableName & " set value = '" & newValue & "' where value = '" & oldValue & "'"
	Repository.Execute sqlUpdateTaggedValues 
end function


function getCorrespondingElement(masterPackage, fromElement)
	dim sqlGetElement
	'initialize to nothing
	set getCorrespondingElement = nothing
	sqlGetElement = "select o.Object_ID from t_object o " & _
					" where o.Package_ID = " & masterPackage.PackageID & _
					" and o.Object_Type = '" & fromElement.Type & "' " & _
					" and o.Name = '" & fromElement.Name & "' " & _
					" and (o.Stereotype is null or o.Stereotype = '" & fromElement.Stereotype & "') "
	dim masterElement as EA.Element
	dim elementCollection
	set elementCollection = getElementsFromQuery(sqlGetElement)
	for each masterElement in elementCollection
		'return the first element
		set getCorrespondingElement = masterElement
		exit for
	next
end function

main