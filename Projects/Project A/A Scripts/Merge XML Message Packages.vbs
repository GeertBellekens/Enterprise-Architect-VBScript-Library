'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]

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
dim deleteElementsList

sub main
	'initialize delete list
	set deleteElementsList = CreateObject("Scripting.Dictionary")
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
			Repository.WriteOutput outputTabName, now() & ": Starting removing '_Type' suffix" ,0
			removeTypeSuffix fromPackage
			Repository.WriteOutput outputTabName, now() & ": Starting rename Payloads" ,0
			renamePayLoads fromPackage
			Repository.WriteOutput outputTabName, now() & ": Starting merge package """ & fromPackage.Name & """ to package """ & masterPackage.Name & """" ,0
			mergePackages fromPackage, masterPackage
			Repository.WriteOutput outputTabName, now() & ": Starting deleting merged elements" ,0
			deleteMergedElements
			Repository.WriteOutput outputTabName, now() & ": Finished!" ,0
			'reload
			Repository.RefreshModelView(0)
		end if
	end if
	msgbox "finished"
end sub

function deleteMergedElements()
	dim element as EA.Element
	dim package as EA.Package
	dim i
	dim currentElement as EA.Element
	for each element in deleteElementsList.Keys
		set package = deleteElementsList(element)
		for i = package.Elements.Count -1 to 0 step -1
			set currentElement = package.Elements.GetAt(i)
			if currentElement.ElementID = element.ElementID then
				Repository.WriteOutput outputTabName, "Deleting element: " & currentElement.Name ,0
				package.Elements.DeleteAt i,false
				exit for
			end if
		next
	next
end function

function removeTypeSuffix(fromPackage)
	dim sqlRemoveSuffix
	'first update the attribute types
	sqlRemoveSuffix = "update a set a.Type = replace (o.name, '_Type','') " & _
					" from t_attribute a " & _
					" inner join t_object o on o.Object_ID = a.Classifier " & _
					" where o.name like '%_Type' " & _
					" and o.Package_ID =" & fromPackage.PackageID
	Repository.Execute sqlRemoveSuffix		
	'then update the actual element names
	sqlRemoveSuffix = "update o set o.Name = replace (o.name, '_Type','') " & _
						" from t_object o  " & _
						" where o.name like '%_Type' " & _
						" and o.Package_ID =" & fromPackage.PackageID
    Repository.Execute sqlRemoveSuffix
end function

function renamePayLoads(fromPackage)
	dim packageTree
	set packageTree = getPackageTree(fromPackage)
	dim packageIDList
	packageIDList = makePackageIDString(packageTree)
	'rename the payloads
	dim sqlRenamePayloadName
	sqlRenamePayloadName = "update o set o.name = o.name + '_' + tv.Value " & _
							" from t_object o " & _
							" inner join t_package p on o.Package_ID = p.Package_ID " & _
							" inner join t_object po on po.ea_guid = p.ea_guid " & _
							" inner join t_objectproperties tv on tv.Object_ID = po.Object_ID " & _
							" 						and tv.Property = 'TargetNamespacePrefix' " & _
							" where o.Name = 'Payload' " & _
							" and o.Object_Type = 'Class'  " & _
							" and o.Stereotype = 'XSDcomplexType' " & _
							" and p.Package_ID in (" & packageIDList & ")"
    Repository.Execute sqlRenamePayloadName	
	'rename the types of the attributes
	dim sqlRenamePayloadAttributes
	sqlRenamePayloadAttributes = "update a set a.type = o.name " & _
								" from t_attribute a " & _
								" inner join t_object o on o.Object_ID = a.Classifier " & _
								" 						and o.Object_Type = 'Class'  " & _
								" 						and o.Stereotype = 'XSDcomplexType' " & _
								"						and o.name like 'Payload%' " & _
								" 						and o.name <> a.Type " & _
								" where a.type = 'Payload' " & _
								" and o.Package_ID in (" & packageIDList & ")"
	Repository.Execute sqlRenamePayloadAttributes
end function

function mergePackages(fromPackage, masterPackage)
	dim fromElement as EA.Element
	'merge owned elements
	for each fromElement in fromPackage.Elements
		if fromElement.Type = "Class" AND fromElement.Stereotype <> "XSDtopLevelElement" then
			dim masterElement as EA.Element
			set masterElement = getCorrespondingElement(masterPackage, fromElement)
			if not masterElement is nothing then
				'add the from element to the list to be deleted
				deleteElementsList.Add fromElement, fromPackage
				'found corresponding element, merge te two
				mergeElements fromElement, masterElement
			else
				Repository.WriteOutput outputTabName, "Moving """ & fromElement.Name & """ to the master package" ,0
				'no corresponding element found, move the element to the master package
				fromElement.PackageID = masterPackage.PackageID
				fromElement.Update
			end if
		end if
	next
	'merge owned packages
	dim fromSubPackage as EA.Package
	for each fromSubPackage in fromPackage.Packages
		dim masterSubPackage as EA.Package
		'merge the subPackage also to the master package
		mergePackages fromSubPackage, masterPackage
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
	'move the missing attributes to the master element
	moveMissingattributes fromElement, masterElement
end function

function moveMissingattributes(fromElement, masterElement)
	dim sqlMoveAttributes
	sqlMoveAttributes = "update a set a.Object_ID = " & masterElement.ElementID & _
						" from t_attribute a " & _
						" where a.Object_ID = " & fromElement.ElementID & _
						" and not exists " & _
						" (select a.ID from t_attribute am  " & _
						" where am.Object_ID = " & masterElement.ElementID & _
						" and am.Name = a.Name  " & _
						" and am.Type = a.Type) "
	Repository.Execute sqlMoveAttributes
end function

function mergeAttributeReferences(fromElement, masterElement)
	dim fromAttribute as EA.Attribute
	dim masterAttribute as EA.Attribute
	for each fromAttribute in fromElement.Attributes
		set masterAttribute = getCorrespondingAttribute(fromAttribute,masterElement)
		if not masterAttribute is nothing then
			'tagged values (attributes)
			updateTaggedValues "t_attributetag", masterAttribute.AttributeGUID, fromAttribute.AttributeGUID 
			'if the fromElement attribute is optional and the masterAttribute not then we make the master attribute optional
			setOptionality fromAttribute , masterAttribute
		end if	
	next
end function

function setOptionality(fromAttribute , masterAttribute)
	if fromAttribute.LowerBound <> "1" and masterAttribute.LowerBound = "1" then
		masterAttribute.LowerBound = "0"
		masterAttribute.Update
	end if
end function

function getCorrespondingAttribute(fromAttribute,masterElement)
	'initialize empty
	set getCorrespondingAttribute = nothing
	dim candidateAttribute as EA.Attribute
	for each candidateAttribute in masterElement.Attributes
		if candidateAttribute.Name = fromAttribute.Name _
			and candidateAttribute.Type = fromAttribute.Type then
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