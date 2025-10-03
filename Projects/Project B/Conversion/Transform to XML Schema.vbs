'[path=\Projects\Project B\Conversion]
'[group=Conversion]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Conversion.Transformation utils

'
' Script Name: Transform to XML Schema
' Author: Geert Bellekens
' Purpose: Transforms the current package into a new package, transforming the stereotypes
' Date: 2025-05-16
'

const generateAttributeDependencies = false 'set to true in order to generate attribute dependencies

const outPutName = "Transform to XSD"
const packageStereotype = "UML Profile for XSD Schema::XSDschema"
const classStereotype = "UML Profile for XSD Schema::XSDcomplexType"
const choiceStereotype = "UML Profile for XSD Schema::XSDchoice"
const attributeStereotype = "UML Profile for XSD Schema::XSDelement"
const enumerationStereotype = ""
const datatypeStereotype = "UML Profile for XSD Schema::XSDsimpleType"
const messageStereotype = "UML Profile for XSD Schema::XSDtopLevelElement"
const baseXSDTypesPackageGUID = "{9047E8CB-6D6A-47ec-82B9-16FA22D288D1}"


sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get the selected package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	'let the user know we started
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName & " for package '"& package.Name &"'", 0
	if right(package.Name, len("-XSD") ) = "-XSD" then
		'do only the conversion
		convertPackageToXSDProfile package, false
	else
		'first clone, and then convert
		transformToXSD package
	end if
	'let the user know it is finished
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName & " for package '"& package.Name &"'", 0
end sub



function transformToXSD (package)
	'select the target package
	dim targetPackage as EA.Package
	set targetPackage = selectPackage()
	if targetPackage is nothing then
		Repository.WriteOutput outPutName, now() & " No target package selected by user", 0
		exit function
	end if
	dim userIsSure
	userIsSure = Msgbox("Do you really want to transform package '" & package.Name & "' to target package '" & targetPackage.Name &  "'?", vbYesNo+vbQuestion, "Transform to XSD?")
	if not userIsSure = vbYes then
		Repository.WriteOutput outPutName, now() & " Script cancelled by user", 0
		exit function
	end if
	Repository.WriteOutput outPutName, now() & " Creating a clone of the package '" & package.Name &"'", 0
	'create a copy of this package, and then transform that package to JSON
	dim clonedPackage as EA.Package
	set clonedPackage = package.Clone()
	clonedPackage.Name = clonedPackage.Name & "-XSD"
	clonedPackage.ParentID = targetPackage.PackageID
	clonedPackage.Update
	'convert to XSD profile
	convertPackageToXSDProfile clonedPackage, true

end function

function convertPackageToXSDProfile (package, userConfirmed)
	if not userConfirmed then
		dim userIsSure
		userIsSure = Msgbox("Do you really want to transform package '" & package.Name & "' to XSD schema?", vbYesNo+vbQuestion, "Transform to XSD?")
			if not userIsSure = vbYes then
				Repository.WriteOutput outPutName, now() & " Script cancelled by user", 0
			exit function
		end if
	end if
	'set package stereotype
	setPackageSteretoypes package
	'get packageTreeID to get a list of all classes
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	'remove any ignoreXSD classes
	removeIgnoredClasses(packageTreeIDString)
	dim superclasses
	'get the superClasses and remember the inheritance strategy
	set superclasses = getSuperclasses(package, packageTreeIDString)
	'convert the classes to XSD profile
	convertPackageToXSDProfileElements package
	'convert the specializations
	dim element
	for each element in superclasses
		convertSpecializations element, packageTreeIDString
	next
	'set to camelCase
	'convertPackageToCamelCase package
	'reset packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	'delete schema object
	deleteSchemaArtifact packageTreeIDString
	'delete all associations and generalizations
	deleteRelations packageTreeIDString
	'add attribute dependencies
	if generateAttributeDependencies then
		addAttributeDependencies packageTreeIDString
	end if
	'order attributes and associations
	orderAttributesAndAssociations packageTreeIDString
	'reload package
	Repository.ReloadPackage package.PackageID
end function

function orderAttributesAndAssociations(packageTreeIDString)
	'infrom user
	Repository.WriteOutput outPutName, now() & " Ordering attributes and associations", 0
	'get the id's and order numbers of all attributes and associations in this package
	dim sqlGetData
	sqlGetData = "select a.id, a.itemType                                                                              " & vbNewLine & _
				" , ROW_NUMBER() OVER (                                                                               " & vbNewLine & _
				"         PARTITION BY a.Object_ID                                                                    " & vbNewLine & _
				"         ORDER BY a.Name                                                                             " & vbNewLine & _
				"     ) AS sequenceNumber                                                                             " & vbNewLine & _
				" from                                                                                                " & vbNewLine & _
				" (                                                                                                   " & vbNewLine & _
				" 	select o.name as className, a.Object_ID, a.name, a.id, 'Attribute' as itemType                    " & vbNewLine & _
				" 	from t_attribute a                                                                                " & vbNewLine & _
				" 	inner join t_object o on o.Object_ID = a.Object_ID                                                " & vbNewLine & _
				" 	where a.Stereotype = 'XSDElement'                                                                 " & vbNewLine & _
				" 	and o.Package_ID in (" & packageTreeIDString & ")                                                 " & vbNewLine & _
				" 	union                                                                                             " & vbNewLine & _
				" 	select o.name as className, o.Object_ID, '_' + oe.Name, c.Connector_ID, 'Connector' as itemType   " & vbNewLine & _
				" 	from t_connector c                                                                                " & vbNewLine & _
				" 	inner join t_object o on o.Object_ID = c.Start_Object_ID                                          " & vbNewLine & _
				" 	inner join t_object oe on oe.Object_ID = c.End_Object_ID                                          " & vbNewLine & _
				" 	where c.Connector_Type = 'Association'                                                            " & vbNewLine & _
				" 	and c.StyleEx like '%alias=choice;%'                                                              " & vbNewLine & _
				" 	and o.Package_ID in (" & packageTreeIDString & ")                                                 " & vbNewLine & _
				" ) a                                                                                                 "
	dim result
	set result = getArrayListFromQuery(sqlGetData)
	dim row
	for each row in result
		dim itemID
		itemID = Clng(row(0))
		dim itemType
		itemType = row(1)
		dim sequenceNumber
		sequenceNumber = row(2)
		if itemType = "Attribute" then
			setAttributeXSDPosition itemID, sequenceNumber
		else
			setConnectorXSDPosition itemID, sequenceNumber
		end if
	next
end function

function setAttributeXSDPosition(attributeID, position)
	dim attribute as EA.Connector
	set attribute = Repository.GetAttributeByID(attributeID)
	'set the tagged value for the position
	setTagValue attribute, "position", position
end function

function setConnectorXSDPosition(connectorID, position)
	dim connector as EA.Connector
	set connector = Repository.GetConnectorByID(connectorID)
	'set the tagged value on the source role
	dim positionTag as EA.TaggedValue
	set positionTag = getExistingOrNewTaggedValue(connector.ClientEnd, "position")
	positionTag.Value = position
	positionTag.Update
end function

function removeIgnoredClasses(packageTreeIDString)
	'delete all classes that have tagged value IgnoreXSD = true
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o                                " & vbNewLine & _
				" inner join t_objectproperties tv on tv.Object_ID = o.Object_ID   " & vbNewLine & _
				" where tv.Property = 'IgnoreXSD'                                  " & vbNewLine & _
				" and tv.Value = 'true'                                            " & vbNewLine & _
				" and o.Package_ID in (" & packageTreeIDString & ")                "
	dim results
	set results = getElementsFromQuery(sqlGetData)
	dim element
	for each element in results
		deleteElement element
	next
end function

function setPackageSteretoypes(package)
	'set it for the current package
	package.StereotypeEx = packageStereotype
	package.Update
	'process subPackages
	dim subPackage as EA.Package
	for each subPackage in package.Packages
		setPackageSteretoypes subPackage
	next
end function

function deleteRelations(packageTreeIDString)
	'get all associations and generalizations in this package
	dim sqlGetData
	sqlGetData = "select o.Object_ID, c.Connector_ID from t_connector c                         " & vbNewLine & _
				" inner join t_object o on o.Object_ID = c.Start_Object_ID                     " & vbNewLine & _
				"                     and o.Object_Type = 'Class'                              " & vbNewLine & _
				" inner join t_object o2 on o2.Object_ID = c.End_Object_ID                     " & vbNewLine & _
				" where c.Connector_Type in ('Generalization', 'Aggregation', 'Association')   " & vbNewLine & _
				" and isnull(o.stereotype, '') <> 'xsdTopLevelElement'                         " & vbNewLine & _
				" and isnull(c.styleEx, '') not like '%alias=choice;%'                         " & vbNewLine & _
				" and o.Package_ID in (" & packageTreeIDString & ")                            " & vbNewLine & _
				" order by 1                                                                   "
	'delete all of them
	dim results
	set results = getArrayListFromQuery(sqlGetData)
	dim row
	dim owner as EA.Element
	'get the first owner
	set owner = nothing
	for each row in results
		dim ownerID
		ownerID = Clng(row(0))
		if owner is nothing then
			set owner = Repository.GetElementByID(ownerID)
			Repository.WriteOutput outPutName, now() & " Deleting connectors from '" & owner.Name &"'", 0
		end if
		if owner.ElementID <> ownerID then
			set owner = Repository.GetElementByID(ownerID)
			Repository.WriteOutput outPutName, now() & " Deleting connectors from '" & owner.Name &"'", 0
		end if
		dim connectorID
		connectorID = Clng(row(1))
		'make sure we have the last connectors collection
		owner.Connectors.Refresh
		'loop connectors and delete the one that corresponds to the id
		dim i
		for i = owner.Connectors.Count -1 to 0 step -1
			dim connector as EA.Connector
			set connector = owner.Connectors.GetAt(i)
			if connector.ConnectorID = connectorID then
				owner.Connectors.DeleteAt i, false
				exit for
			end if
		next
	next
end function

function deleteSchemaArtifact(packageTreeIDString)
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o where     " & vbNewLine & _
				" o.Style like '%MessageProfile=1;%'          " & vbNewLine & _
				" and o.package_id in (" & packageTreeIDString & ") "
	dim results
	set results = getElementsFromQuery(sqlGetData)
	dim element as EA.Element
	for each element in results
		deleteElement element
	next
end function

function convertPackageToCamelCase(package)
	'make everything camelCase
	package.Elements.Refresh
	dim element
	for each element in package.Elements
		convertElementToCamelCase element
	next
	'camelCase subPackages
	dim subPackage
	for each subPackage in package.Packages
		convertPackageToCamelCase subPackage
	next
end function

function getSuperclasses(package, packageTreeIDString )
	Repository.WriteOutput outPutName, now() & " Processing package '" & package.Name &"'", 0
	'get all elements with specializations
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o                              " & vbNewLine & _
				" where o.Package_ID in (" & packageTreeIDString & ")            " & vbNewLine & _
				" and o.Object_Type = 'Class'                                    " & vbNewLine & _
				" and exists	                                                 " & vbNewLine & _
				" 	(select Connector_ID from t_connector c                      " & vbNewLine & _
				" 	inner join t_object so on so.Object_ID = c.Start_Object_ID   " & vbNewLine & _
				" 		and so.Package_ID in (" & packageTreeIDString & ")       " & vbNewLine & _
				" 	where c.End_Object_ID = o.Object_ID                          " & vbNewLine & _
				" 	and c.Connector_Type = 'Generalization')                     "
	dim elements
	set elements = getElementsFromQuery(sqlGetData)
	dim element
	for each element in elements
		'get strategy
		dim strategy
		strategy = getTaggedValueValue(element, "Inheritance Strategy")
		dim tempTag as EA.TaggedValue
		set tempTag = getExistingOrNewTaggedValue(element, "Temp Inheritance Strategy")
		tempTag.Value strategy
		tempTag.Update
	next
	'return
	set getSuperclasses = elements
end function

function convertSpecializations(element, packageTreeIDString)
	'reload element because the in-memory object could be out of date
	set element = Repository.GetElementByID(element.ElementID)
	'get subClasses
	dim subClasses
	set subClasses = getSubClasses(element, packageTreeIDString)
	'get strategy
	'refresh tagged values
	element.taggedValues.Refresh
	dim strategy
	strategy = getTaggedValueValue(element, "Temp Inheritance Strategy")
	select case lcase(strategy)
		case "flatten up"
			flattenUp element, subclasses
		case "oneof"
			oneOf element, subclasses
		case else
			'flatten down = default
			flattenDown element, subClasses
	end select
	'delete the temporary tag
	'TODO?
end function

function flattenUp(element, subclasses)
	'move all association to the superclass
	redirectIncomingAssociations element, subclasses
	moveOutgoingAssociations element, subclasses
	
	dim subClass as EA.Element
	for each subClass in subclasses
		'copy all attributes from the subClasses to the superclass
		copyAttributes subClass, element
		'move all using attributes to the superclass
		moveUsingAttributes subClass, element
		'then delete element
		deleteElement subClass
	next
end function

function moveUsingAttributes(source, target)
	dim sqlGetData
	sqlGetData = "select a.ID from t_attribute a    " & vbNewLine & _
				" where a.Classifier = " & source.ElementID
	dim results
	set results = getAttributesFromQuery(sqlGetData)
	dim attribute as EA.Attribute
	for each attribute in results
		attribute.ClassifierID = target.ElementID
		attribute.Type = target.Name
		attribute.Update
	next
end function

function oneOf(element, subclasses)
	if subclasses.Count = 0 then
		'nothing to do if there are no subclasses
		exit function
	end if
	Repository.WriteOutput outPutName, now() & " Oneof for '" & element.Name &"'", 0
	'make element not abstract
	if element.Abstract then
		element.Abstract = false
		element.Update
	end if
	if subClasses.Count > 1 then
		'create new OneOf class
		dim ownerPackage as EA.Package
		set ownerPackage = Repository.GetPackageByID(element.packageID)
		dim oneOfClass as EA.Element
		set oneOfClass = ownerPackage.Elements.AddNew(element.Name & "OneOf", choiceStereotype)
		oneOfClass.Update
		'add to diagrams
		addToSameDiagrams element, oneOfClass
		'make it a oneOf class
		makeOneOf oneOfClass, subClasses
	else
		'if there is only one subclass, we don't need the oneOfClass, but simply redirect to the subClass
		set oneOfClass = subClasses(0)
	end if
	'addd association in main element
	dim oneOfAssociation as EA.Connector
	set oneOfAssociation = element.Connectors.AddNew("", "Association")
	oneOfAssociation.SupplierID = oneOfClass.ElementID
	oneOfAssociation.SupplierEnd.Cardinality = "1..1"
	oneOfAssociation.Alias = "choice"
	oneOfAssociation.Update
	'redirect incoming associations to subclasses to the main class
	 redirectIncomingAssociations element, subClasses
end function

function addToSameDiagrams(element, otherElement)
	'get all diagrams for the element
	dim sqlGetData
	sqlGetData = "select dod.Diagram_ID from t_diagramobjects dod    " & vbNewLine & _
				 " where dod.Object_ID = " & element.ElementID & "   "
	dim diagrams
	set diagrams = getDiagramsFromQuery(sqlGetData)
	'add the other element to all these diagrams
	dim diagram as EA.Diagram
	for each diagram in diagrams
		dim diagramObject as EA.DiagramObject
		set diagramObject = diagram.DiagramObjects.AddNew("", "")
		diagramObject.ElementID = otherElement.ElementID
		diagramObject.Update
	next
end function

function redirectIncomingAssociations(target, sources)
	dim source as EA.Element
	for each source in sources
		dim sqlGetData
		sqlGetData = "select c.Connector_ID from t_connector c    " & vbNewLine & _
					" where c.Connector_Type = 'Association'     " & vbNewLine & _
					" and c.End_Object_ID = " & source.ElementID
		dim results
		set results = getConnectorsFromQuery(sqlGetData)
		dim association as EA.Connector
		for each association in results
			association.SupplierID = target.elementID
			association.Update
		next
	next
end function

function moveOutgoingAssociations(target, sources)
	dim source as EA.Element
	for each source in sources
		dim sqlGetData
		sqlGetData = "select c.Connector_ID from t_connector c    " & vbNewLine & _
					" where c.Connector_Type = 'Association'     " & vbNewLine & _
					" and c.Start_Object_ID = " & source.ElementID
		dim results
		set results = getConnectorsFromQuery(sqlGetData)
		dim association as EA.Connector
		for each association in results
			association.ClientID = target.elementID
			association.Update
		next
	next
end function

function flattenDown(element, subClasses)
	Repository.WriteOutput outPutName, now() & " Flatten down for '" & element.Name &"'", 0
	'copy all attributes to the subclasses
	dim subclass
	for each subClass in subClasses
		copyAttributes element, subClass
	next
	'check if there are attributes that have this element as type. 
	'if this is the case, then make this class a oneOf with all subclasses
	'else delete this class
	dim used
	used = isUsedAsAttributeType(element)
	if used then
		makeOneOf element, subClasses
	else
		deleteElement element
	end if	
end function

function deleteElement(element)
	dim package as EA.Package
	set package = Repository.GetPackageByID(element.PackageID)
	dim i
	for i = package.Elements.Count - 1 to 0 step -1
		dim tempElement as EA.Element
		set tempElement = package.Elements.GetAt(i)
		if tempElement.ElementID = element.ElementID then
			package.Elements.DeleteAt i,false
			exit function
		end if
	next
end function

function isUsedAsAttributeType(element)
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o             " & vbNewLine & _
				" where exists	                                " & vbNewLine & _
				" 	(select a.ID from t_attribute a             " & vbNewLine & _
				" 	where a.Classifier = o.Object_ID)           " & vbNewLine & _
				" and o.Object_ID = " & element.ElementID & "   "
	dim results
	set results = getArrayListFromQuery(sqlGetData)
	if results.Count > 0 then
		isUsedAsAttributeType = true
	else
		isUsedAsAttributeType = false
	end if
end function

function makeOneOf(element, subClasses)
	'if there is only one subclass then we delete the element and redirect all using attributes to the subclass
	if subClasses.Count = 1 then
		'get the subclass
		dim subClass as EA.Element
		set subClass = subClasses(0)
		'find all using attributes
		dim sqlGetData
		sqlGetData = "select a.ID from t_attribute a where a.Classifier = " & element.ElementID
		dim results
		set results = getAttributesFromQuery(sqlGetData)
		dim attribute as EA.Attribute
		for each attribute in results
			'redirect type to subclass
			attribute.ClassifierID = subClass.ElementID
			attribute.Type = subClass.Name
			attribute.Update
		next
		'delete element
		deleteElement element
	else
		'delete all attributes
		deleteAllAttributes element
		'create the attributes for the subclasses
		for each subClass in SubClasses
			set attribute = element.Attributes.AddNew(subClass.Name, subClass.Name)
			attribute.StereotypeEx = attributeStereotype
			attribute.Visibility = "Public"
			attribute.ClassifierID = subClass.ElementID
			attribute.Type = subclass.Name
			attribute.Update
			setTagValue attribute, "anonymousType", "true"
		next
		dim addPostFix
		'add OneOf postfix
		if len(element.Name) > len("OneOf") then
			if not right(element.Name, len("OneOf")) = "OneOf" then
				addPostFix = true
			else
				addPostFix = false
			end if
		else
			addPostFix = true
		end if
		if addPostFix then
			element.Name = element.Name & "OneOf"
			element.Update
		end if
		'set xsdChoice stereotype
		element.StereotypeEx = choiceStereotype
		element.Update
		'create the complex type if needed
		if hasUsingAttributesAsClassifier(element) then
			dim package
			dim complexType
			set complexType = createComplexTypeForOneOf(element)
			'add replace the using attributes
			replaceUsingAttributesClassifier element, complexType
		end if
	end if
end function

function createComplexTypeForOneOf(element)
	dim package as EA.Package
	set package = Repository.GetPackageByID(element.PackageID)
	dim complexType as EA.Element
	set complexType = package.Elements.AddNew(element.Name & "Type", classStereotype)
	complexType.Update
	'add to same diagrams
	addToSameDiagrams element, complexType
	'create association to choice element
	dim oneOfAssociation as EA.Connector
	set oneOfAssociation = complexType.Connectors.AddNew("", "Association")
	oneOfAssociation.SupplierID = element.ElementID
	oneOfAssociation.SupplierEnd.Cardinality = 1 & ".." & 1
	oneOfAssociation.Alias = "choice"
	oneOfAssociation.Update
	'return 
	set createComplexTypeForOneOf = complexType
end function

function hasUsingAttributesAsClassifier(element)
	dim sqlGetData
	sqlGetData = "select a.ID from t_attribute a    " & vbNewLine & _
					" where a.Classifier = " & element.ElementID
	dim result
	set result = getArrayListFromQuery(sqlGetData)
	if result.Count = 0 then
		hasUsingAttributesAsClassifier = false
	else
		hasUsingAttributesAsClassifier = true
	end if
end function

function replaceUsingAttributesClassifier(element, complexType)
	dim sqlGetData
	sqlGetData = "select a.ID from t_attribute a    " & vbNewLine & _
					" where a.Classifier = " & element.ElementID
	dim attributes
	set attributes = getAttributesFromQuery(sqlGetdata)
	dim attribute as EA.Attribute
	for each attribute in attributes
		attribute.ClassifierID = complexType.ElementID
		attribute.Type = complexType.Name
		attribute.Update
	next
end function

function deleteAllAttributes(element)
	dim i
	for i = element.Attributes.Count -1 to 0 step - 1
		element.Attributes.DeleteAt i, false
	next
	'refresh attributes collection
	element.Attributes.Refresh
end function

function copyAttributes(source, target)
	'get a dictionary of all target attributes to make sure we don't get duplicates
	'create attributesDictionary
	target.Attributes.Refresh
	dim attributesDictionary
	set attributesDictionary = CreateObject("Scripting.Dictionary")
	dim attribute as EA.Attribute
	for each attribute in target.Attributes
		if not attributesDictionary.Exists(attribute.Name) then
			attributesDictionary.Add attribute.Name, attribute
		end if
	next
	source.Attributes.Refresh
	'then loop the source attributes, and add them to the target class
	for each attribute in source.Attributes
		if not attributesDictionary.Exists(attribute.Name) then
			attributesDictionary.Add attribute.Name, attribute
			dim newAttribute as EA.Attribute
			set newAttribute = target.Attributes.AddNew(attribute.Name, attribute.Type)
			newAttribute.ClassifierID = attribute.ClassifierID
			newAttribute.Alias = attribute.Alias
			newAttribute.Notes = attribute.Notes
			newAttribute.Visibility = attribute.Visibility
			newAttribute.LowerBound = attribute.LowerBound
			newAttribute.StereotypeEx = attribute.FQStereotype
			newAttribute.Update
			'copy tagged values
			copyTaggedValues attribute, newAttribute
		end if
	next
end function

function getSubClasses(element, packageTreeIDString)
		dim sqlGetData
		sqlGetData = "select o.Object_ID from t_object o                               " & vbNewLine & _
				" inner join t_connector c on c.Start_Object_ID = o.Object_ID     " & vbNewLine & _
				" 				and c.Connector_Type = 'Generalization'           " & vbNewLine & _
				" 				and c.End_Object_ID = " & element.ElementID & "   " & vbNewLine & _
				" where o.Package_ID in (" & packageTreeIDString & ")             "	
		dim elements
		set elements = getElementsFromQuery(sqlGetData)
		'return
		set getSubClasses = elements
end function



function convertPackageToXSDProfileElements(package)
	Repository.WriteOutput outPutName, now() & " Processing package '" & package.Name &"'", 0
	dim element as EA.Element
	'convert elements
	for each element in package.Elements
		convertElementToXSDProfile element
	next
	'convert diagrams
	dim diagram as EA.Diagram
	for each diagram in package.Diagrams
		if diagram.Type = "Logical" _
		  and not instr(diagram.StyleEx, "MDGDgm=JSON::JSON;") > 0 then
			'report progress
			Repository.WriteOutput outPutName, now() & " Migrating diagram '"& package.Name & "." & diagram.Name &"'", 0
			diagram.Metatype = "JSON::JSON"
			diagram.StyleEx = setValueForKey(diagram.StyleEx, "MDGDgm", "JSON::JSON")
			diagram.StyleEx = setValueForKey(diagram.StyleEx, "HideConnStereotype", "1")
			diagram.ExtendedStyle = setValueForKey(diagram.ExtendedStyle, "HideStereo", "1")
			diagram.ExtendedStyle = setValueForKey(diagram.ExtendedStyle, "HideEStereo", "1")
			diagram.Update
		end if
	next
	'process subPackages
	dim subPackage as EA.Package
	for each subPackage in package.Packages
		convertPackageToXSDProfileElements subPackage
	next
end function

function convertElementToXSDProfile(element)
	Repository.WriteOutput outPutName, now() & " Processing element '"& element.Name &"'", 0
	dim targetStereo
	targetStereo = ""
	'convert element
	select case lcase(element.Type)
		case "class"
			if element.Stereotype = "LDM_Message" then
				convertRootElement element
			else
				updateElementStereotype element, classStereotype
			end if
		case "datatype"
			element.Type = "Class"
			element.Update
			set element = Repository.GetElementByID(element.ElementID)
			updateElementStereotype element, datatypeStereotype
			updateToBaseXSDTypes element
			exit function
		case "enumeration"
			updateElementStereotype element, enumerationStereotype
			switchNameAndAliasForEnumerations element
	end select
	if element.Type = "Class"  _
	  or element.Type = "DataType" then
		'remove abstract property
		if element.Abstract then
			element.Abstract = false
			element.Update
		end if
		'create attributesDictionary
		dim attributesDictionary
		set attributesDictionary = CreateObject("Scripting.Dictionary")
		'convert attributes
		dim attribute as EA.Attribute
		for each attribute in element.Attributes
			if not attributesDictionary.Exists(attribute.Name) then
				attributesDictionary.Add attribute.Name, attribute
			end if
			if attribute.FQStereotype <> attributeStereotype then
				attribute.StereotypeEx = attributeStereotype
				attribute.Visibility = "Public"
				attribute.Update
				setTagValue attribute, "anonymousType", "true"
			end if
		next
		'convert associations
		convertAssociationsToXSD element, attributesDictionary
	end if
end function

function convertRootElement(element)
	dim package
	set package = Repository.GetPackageByID(element.PackageID)
	'set version on package
	setTagValue package.Element, "version", element.Version
	'set targetNamespace on package
	dim targetNameSpace
	targetNameSpace = getTaggedValueValue(element, "targetNamespace")
	setTagValue package.Element, "targetNamespace", targetNameSpace
	setTagValue package.Element, "defaultNamespace", targetNameSpace
	'create a xsdTopLevelElement with the same name as the message element
	dim topLevelElement as EA.Element
	set topLevelElement  = package.Elements.AddNew(element.Name, messageStereotype)
	topLevelElement.Update
	'add to diagrams
	addToSameDiagrams element, topLevelElement
	'set the xsdComplexType stereotype 
	updateElementStereotype element, classStereotype
	'add generalization from xsdTopLevelElement to complex type
	dim generalization as EA.Connector
	set generalization = topLevelElement.Connectors.AddNew("", "Generalization")
	generalization.SupplierID = element.ElementID
	generalization.Update
end function

function switchNameAndAliasForEnumerations(element)
	dim attribute
	for each attribute in element.Attributes
		if len(attribute.Alias) > 0 then
			dim temp
			temp = attribute.Alias
			attribute.Alias = attribute.Name
			attribute.Name = temp
			attribute.Notes = attribute.Alias 'set annotation for XSD
			attribute.Update
		end if
	next
end function


function updateToBaseXSDTypes(datatype)
	'check if there is a base XSD type with the same name
	dim baseType
	set baseType = getBaseXSDType(datatype.Name)
	if baseType is nothing then
		exit function
	end if
	'replace all attributes that use this datatype with the xsdBaseType
	dim usingAttributes
	set usingAttributes = getUsingAttributes(datatype)
	dim attribute as EA.Attribute
	for each attribute in usingAttributes
		attribute.ClassifierID = baseType.ElementID
		attribute.Type = baseType.Name
		attribute.Update
	next
	'redirect all generalizations to the base XSD type
	dim generalizations
	set generalizations = getUsingGeneralizations(datatype)
	dim generalization as EA.Connector
	for each generalization in generalizations
		generalization.SupplierID = baseType.ElementID
		generalization.Update
	next
	'delete the datatype
	deleteElement datatype
end function

function getUsingGeneralizations(datatype)
	dim sqlGetData
	sqlGetData = "select c.Connector_ID from t_connector c     " & vbNewLine & _
				" where c.Connector_Type = 'Generalization'   " & vbNewLine & _
				" and c.End_Object_ID = " & datatype.ElementID
	dim result
	set result = getConnectorsFromQuery(sqlGetData)
	'return
	set getUsingGeneralizations = result
end function

function getUsingAttributes(datatype)
	dim sqlGetData
	sqlGetData = "select a.ID from t_attribute a where a.Classifier = " & datatype.ElementID
	dim result
	set result = getAttributesFromQuery(sqlGetData)
	set getUsingAttributes = result
end function

function getBaseXSDType(typeName)
	dim baseType as EA.Element
	set baseType = nothing
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o                                           " & vbNewLine & _
				" inner join t_package p on p.Package_ID = o.Package_ID                        " & vbNewLine & _
				" 					and p.ea_guid = '" & baseXSDTypesPackageGUID & "'          " & vbNewLine & _
				" where o.Stereotype = 'XSDsimpleType'                                         " & vbNewLine & _
				" and o.name = '" & typeName & "'                                              "
	dim result
	set result = getElementsFromQuery(sqlGetData)
	if result.Count > 0 then
		set baseType = result(0)
	end if
	'retunr
	set getBaseXSDType = baseType
end function

function convertElementToCamelCase(element)
	Repository.WriteOutput outPutName, now() & " CamelCasing '" & element.Name &"'", 0
	'convert element
	convertNamingtoCamelCase element
	'do not convert enum values
	if lcase(element.Type) = "enumeration" then
		exit function
	end if
	'convert attributes
	element.Attributes.Refresh
	dim attribute as EA.Attribute
	for each attribute in element.Attributes
		convertNamingtoCamelCase attribute
	next
end function

function convertNamingtoCamelCase(item)
	dim camelCaseName
	camelCaseName = getCamelCaseName(item.Name)
	if item.Name <> camelCaseName then
		item.Name = camelCaseName
		item.Update
	end if	
end function

function getCamelCaseName(name)
	dim camelCaseName
	camelCaseName = name
	if len(name) = 0 then
		getCamelCaseName = camelCaseName
		exit function
	end if
	if lcase(left(name, 1)) <> left(name, 1) then
		camelCaseName = lcase(left(name, 1)) & right(name, len(name) -1)
	end if
	'return
	getCamelCaseName = camelCaseName
end function

function convertAssociationsToXSD(element, attributesDictionary)
	'convert associations
	dim connector as EA.Connector
	dim i
	for i = element.Connectors.Count - 1 to 0 step -1
		set connector = element.Connectors.GetAt(i)
		if connector.Stereotype = "LDM_Association" _
		  and connector.ClientID = element.ElementID then
			'get source and target element
			dim sourceElement as EA.Element
			dim targetElement as EA.Element
			if connector.ClientID = element.ElementID then
				set sourceElement = element
				set targetElement = Repository.GetElementByID(connector.SupplierID)
			else
				set sourceElement = Repository.GetElementByID(connector.ClientID)
				set targetElement = element
			end if
			'set multiplicity
			dim multiplicity
			multiplicity = getMultiplicityFromCardinality(connector.SupplierEnd.Cardinality)
			dim lowerBound 
			lowerBound = multiplicity(0)
			dim upperBound
			upperBound = multiplicity(1)
			'create attribute
			dim attributeName
			attributeName = connector.SupplierEnd.Role
			if len(attributeName) = 0 then
				attributeName  = targetElement.Name
			end if
			if upperBound <> "1" then
				attributeName = attributeName & "List"
			end if
			'check if attribute doesn't exist yet
			if not attributesDictionary.Exists(attributeName) then
				'report progress
				Repository.WriteOutput outPutName, now() & " Creating attribute '"& sourceElement.Name & "." & attributeName &"'", 0
				dim attribute
				set attribute = sourceElement.Attributes.AddNew(attributeName, targetElement.Name)
				attribute.StereotypeEx = attributeStereotype
				attribute.Visibility = "Public"
				attribute.Notes = connector.Notes
				attribute.lowerBound = multiplicity(0)
				if upperBound <> "1" then
					'create complextype for list element
					dim listObject as EA.Element
					dim package as EA.Package
					set package = Repository.GetPackageByID(element.PackageID)
					set listObject = createListObject(attributeName, package, sourceElement, targetElement, multiplicity)
					'set type on main attribute
					attribute.ClassifierID = listObject.ElementID
					attribute.Type = listObject.Name
				else
					attribute.UpperBound = multiplicity(1)
					attribute.ClassifierID = targetElement.ElementID
				end if
				'save the atribute
				attribute.Update
				'set anonymousType property
				setTagValue attribute, "anonymousType", "true"
				'add to dictionary
				attributesDictionary.Add attribute.Name, attribute
			end if
		end if
	next
end function

function createListObject(name, package, sourceElement, targetElement, multiplicity)
	dim listObject as EA.Element
	set listObject = nothing
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o       " & vbNewLine & _
				" where o.Stereotype = 'xsdComplexType'   " & vbNewLine & _
				" and o.Package_ID = " & package.PackageID  & vbNewLine & _
				" and o.name = '" & name & "'             "
	dim result
	set result = getElementsFromQuery(sqlGetData)
	if result.Count > 0 then
		set listObject = result(0)
	else
		set listObject = package.Elements.AddNew(name, classStereotype)
		listObject.update
		'add to diagram
		addToSameDiagrams sourceElement, listObject
		'add attribute to target
		dim targetAttribute as EA.Attribute
		set targetAttribute = listObject.Attributes.AddNew(targetElement.Name, targetElement.Name)
		targetAttribute.StereotypeEx = attributeStereotype
		targetAttribute.ClassifierID = targetElement.ElementID
		targetAttribute.lowerBound = "1"
		targetAttribute.UpperBound = multiplicity(1)
		targetAttribute.Update
		'set anonymousType
		setTagValue targetAttribute, "anonymousType", "true"
	end if
	'return
	set createListObject = listObject
end function

function getMultiplicityFromCardinality(cardinality)
	dim multiplicity(2)
	'default 1..1
	multiplicity(0) = "1"
	multiplicity(1) = "1"
	if instr(cardinality, "..") > 0 then
		dim parts
		parts = split(cardinality, "..")
		if Ubound(parts) > 0 then
			multiplicity(0) = parts(0)
			multiplicity(1) = parts(1)
		end if
	else
		multiplicity(0) = cardinality
		multiplicity(1) = cardinality
	end if
	'return
	getMultiplicityFromCardinality = multiplicity
end function


function updateElementStereotype(element, targetStereotype)
	'skip if already the correct stereotype
	if element.FQStereotype = targetStereotype then
		exit function
	end if
	dim oldTagValues
	set oldTagValues = CreateObject("Scripting.Dictionary")
	dim tag as EA.TaggedValue
	for each tag in element.TaggedValues
		oldTagValues(lcase(tag.Name)) = tag.Value
	next
	'set correct stereotype
	element.StereotypeEx = targetStereotype
	element.Update
	'refresh tags
	element.TaggedValues.Refresh
	'copy the old values into the new tags
	for each tag in element.TaggedValues
		if oldTagValues.Exists(lcase(tag.Name)) then
			tag.Value = oldTagValues(lcase(tag.Name))
			tag.Update
		end if
	next
end function

main
'test

'sub test
'	'create output tab
'	Repository.CreateOutputTab outPutName
'	Repository.ClearOutput outPutName
'	Repository.EnsureOutputVisible outPutName
'	'get the selected package
'	dim package as EA.Package
'	set package = Repository.GetTreeSelectedPackage()
'	'let the user know we started
'	Repository.WriteOutput outPutName, now() & " Starting " & outPutName & " for package '"& package.Name &"'", 0
'	dim packageTreeIDString
'	packageTreeIDString = getPackageTreeIDString(package)
'	'order attributes
'	orderAttributesAndAssociations packageTreeIDString
'	'let the user know it is finished
'	Repository.WriteOutput outPutName, now() & " Finished " & outPutName & " for package '"& package.Name &"'", 0
'end sub
