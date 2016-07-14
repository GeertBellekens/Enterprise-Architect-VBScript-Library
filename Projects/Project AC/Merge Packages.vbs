'[path=\Projects\Project AC]
'[group=Acerta Scripts]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Import Database mappings
' Author: Geert Bellekens
' Purpose: Import the database mappings from a csv file exported from MEGA
' Date: 2016-07-07
'
sub main
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
			mergePackages fromPackage, masterPackage
		end if
	end if
end sub


function mergePackages(fromPackage, masterPackage)
	dim fromElement as EA.Element
	for each fromElement in fromPackage.Elements
		dim masterElement as EA.Element
		set masterElement = getCorrespondingElement(masterPackage, fromElement)
		if not masterElement is nothing then
			'found corresponding element, merge te two
			mergeElements fromElement, masterElement
		else
			'no corresponding element found, move the element to the master package
			fromElement.PackageID = masterPackage.PackageID
			fromElement.Update
		end if
	next
end function

'merging elements is the fastest if we use database updates
function mergeElements(fromElement, masterElement)
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
	'TODO
end function

function updateTaggedValues (tableName, newValue, oldValue)
	dim sqlUpdateTaggedValues
	sqlUpdateTaggedValues = "update " & tableName & " set value = '" & newValue & "' where value = '" & oldValue & "'"
	Repository.Execute sqlUpdateTaggedValues 
end function


function getCorrespondingElement(masterPackage, fromElement)
	dim sqlGetElement
	sqlGetElement = "select o.Object_ID from t_object o " & _
					" where o.Package_ID = " & masterPackage.PackageID & _
					" and o.Object_Type = '" & fromElement.Type & "' " & _
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