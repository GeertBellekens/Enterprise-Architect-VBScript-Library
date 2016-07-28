'[path=\Projects\Project AC]
'[group=Acerta Scripts]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

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
	dim fromPackage as EA.Package
	msgbox "select the FROM package"
	set fromPackage = selectPackage()
	'select master package
	dim masterPackage as EA.Package
	msgbox "select the TO (master) package"
	set masterPackage = selectPackage()
	if not fromPackage is nothing and not masterPackage is nothing then 
		dim response
		response = Msgbox("Merge package """ & fromPackage.Name & """ to package """ & masterPackage.Name & """?", vbYesNoCancel+vbQuestion, "Merge Package")
		if response = vbYes then
			Repository.WriteOutput outputTabName, now() & ": Starting merge package """ & fromPackage.Name & """ to package """ & masterPackage.Name & """" ,0
			mergePackages fromPackage, masterPackage
			Repository.WriteOutput outputTabName, now() & ": Finished merge package """ & fromPackage.Name & """ to package """ & masterPackage.Name & """" ,0
		end if
	end if
	msgbox "finished"
end sub


function mergePackages(fromPackage, masterPackage)
	dim fromElement as EA.Element
	'merge owned elements
	for each fromElement in fromPackage.Elements
		dim masterElement as EA.Element
		set masterElement = getCorrespondingElement(masterPackage, fromElement)
		if not masterElement is nothing then
			'found corresponding element, merge te two
			mergeElements fromElement, masterElement
		else
			Repository.WriteOutput outputTabName, "Moving """ & fromElement.Name & """ to the master package" ,0
			'no corresponding element found, move the element to the master package
			fromElement.PackageID = masterPackage.PackageID
			fromElement.Update
		end if
	next
	'merge owned packages
	dim fromSubPackage as EA.Package
	for each fromSubPackage in fromPackage.Packages
		dim masterSubPackage as EA.Package
		set masterSubPackage = getCorrespondingPackage(fromSubPackage, masterPackage)
		if not masterSubPackage is nothing then
			mergePackages fromSubPackage, masterSubPackage
		end if
	next
end function

function getCorrespondingPackage(fromSubPackage, masterPackage)
	dim candidatePackage as EA.Package
	'initialize empty
	set getCorrespondingPackage = nothing
	for each candidatePackage in masterPackage.Packages
		if candidatePackage.Name = fromSubPackage.Name then
			set getCorrespondingPackage = candidatePackage
			exit for
		end if
	next
end function

'merging elements is the fastest if we use database updates
function mergeElements(fromElement, masterElement)
	Repository.WriteOutput outputTabName, "Merging """ & fromElement.Name & """ to """ & masterElement.Name & """" ,0
	'attribute datatypes
	dim sqlUpdateDatatypes 
	slqUpdateDatatypes = "update t_attribute set classifier = " & masterElement.ElementID & " where classifier = " & fromElement.ElementID
	Repository.Execute slqUpdateDatatypes
	'connectors (not for connectors to self)
	dim sqlUpdateConnectorSource, sqlUpdateConnectorTarget
	sqlUpdateConnectorSource = "update t_connector set Start_Object_ID = " & masterElement.ElementID & " where Start_Object_ID <> End_Object_ID and Start_Object_ID = " & fromElement.ElementID
	Repository.Execute sqlUpdateConnectorSource
	sqlUpdateConnectorTarget = "update t_connector set End_Object_ID = " & masterElement.ElementID & " where Start_Object_ID <> End_Object_ID and End_Object_ID = " & fromElement.ElementID
	Repository.Execute sqlUpdateConnectorTarget
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