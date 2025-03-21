'[path=\Projects\Project B\Conversion]
'[group=Conversion]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Transform to JSON
' Author: Geert Bellekens
' Purpose: Transforms the current package into a new package, transforming the stereotypes
' Date: 2025-01-16
'

const outPutName = "Transform to JSON"
const classStereotype = "JSON::JSON_Element"
const attributeStereotype = "JSON::JSON_Attribute"
const enumerationStereotype = ""
const datatypeStereotype = "JSON::JSON_Datatype"
const messageStereotype = "JSON::JSON_Schema"


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
	if right(package.Name, len("-JSON") ) = "-JSON" then
		'do only the conversion
		convertPackageToJSONProfile(package)
	else
		'first clone, and then convert
		transformToJSON package
	end if
	'reload package
	Repository.ReloadPackage package.PackageID
	'let the user know it is finished
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName & " for package '"& package.Name &"'", 0
end sub



function transformToJSON (package)
	'select the target package
	dim targetPackage as EA.Package
	set targetPackage = selectPackage()
	Repository.WriteOutput outPutName, now() & " Creating a clone of the package '" & package.Name &"'", 0
	'create a copy of this package, and then transform that package to JSON
	dim clonedPackage as EA.Package
	set clonedPackage = package.Clone()
	clonedPackage.Name = clonedPackage.Name & "-JSON"
	clonedPackage.ParentID = targetPackage.PackageID
	clonedPackage.Update
	'convert to JSON profile
	convertPackageToJSONProfile clonedPackage
end function

function convertPackageToJSONProfile(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	dim superclasses
	'get the superClasses and remember the inheritance strategy
	set superclasses = getSuperclasses(package, packageTreeIDString)
	'convert the classes to JSON profile
	convertPackageToJSONProfileElements package
	'convert the specializations
	dim element
	for each element in superclasses
		convertSpecializations element, packageTreeIDString
	next
	'set to camelCase
	convertPackageToCamelCase package
	'delete schema object
	deleteSchemaArtifact package
end function

function deleteSchemaArtifact(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
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
		set oneOfClass = ownerPackage.Elements.AddNew(element.Name & "OneOf", "JSON::JSON_Element")
		oneOfClass.Update
		'make it a oneOf class
		makeOneOf oneOfClass, subClasses
	else
		'if there is only one subclass, we don't need the oneOfClass, but simply redirect to the subClass
		set oneOfClass = subClasses(0)
	end if
	'addd attribute in main element
	dim oneOfAttribute as EA.Attribute
	set oneOfAttribute = element.Attributes.AddNew(oneOfClass.Name, oneOfClass.Name)
	oneOfAttribute.stereotypeEx = attributeStereotype
	oneOfAttribute.ClassifierID = oneOfClass.ElementID
	oneOfAttribute.Visibility = "Private"
	oneOfAttribute.Update
	'redirect incoming associations to subclasses to the main class
	 redirectIncomingAssociations element, subClasses
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
			attribute.Update
		next
		'delete element
		deleteElement element
	else
		'delete all attributes
		deleteAllAttributes element
		'set oneOf 
		setTagValue element, "compositionType", "oneOf"
		'create the attributes for the subclasses
		for each subClass in SubClasses
			set attribute = element.Attributes.AddNew(subClass.Name, subClass.Name)
			attribute.StereotypeEx = attributeStereotype
			attribute.Visibility = "Private"
			attribute.ClassifierID = subClass.ElementID
			attribute.Update
		next
	end if
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



function convertPackageToJSONProfileElements(package)
	Repository.WriteOutput outPutName, now() & " Processing package '" & package.Name &"'", 0
	dim element as EA.Element
	'convert elements
	for each element in package.Elements
		convertElementToJSONProfile element
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
		convertPackageToJSONProfileElements subPackage
	next
end function

function convertElementToJSONProfile(element)
	Repository.WriteOutput outPutName, now() & " Processing element '"& element.Name &"'", 0
	dim targetStereo
	targetStereo = ""
	'convert element
	select case lcase(element.Type)
		case "class"
			if element.Stereotype = "LDM_Message" then
				updateElementStereotype element, messageStereotype
			else
				updateElementStereotype element, classStereotype
			end if
		case "datatype"
			updateElementStereotype element, datatypeStereotype
			addStringInheritanceToDateAndTime(element)
			exit function
		case "enumeration"
			updateElementStereotype element, enumerationStereotype
	end select
	if element.Type = "Class"  _
	  or element.Type = "DataType" then
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
				attribute.Update
			end if
		next
		'convert associations
		convertAssociationsToJSON element, attributesDictionary
	end if
end function

function addStringInheritanceToDateAndTime(element)
	if lcase(element.Name) = "date" _
	  or lcase(element.Name) = "datetime" _
	  or lcase(element.Name) = "time" then
		'find string datatype
		dim sqlGetData
		sqlGetData = "select o.Object_ID from t_object o    " & vbNewLine & _
						" where o.Object_Type = 'Datatype'     " & vbNewLine & _
						" and o.Name = 'String'                " & vbNewLine & _
						" and o.Package_ID = "  & element.PackageID
		dim results
		set results = getElementsFromQuery(sqlGetData)
		dim stringDataType as EA.Element
		if results.Count > 0 then
			set stringDataType = results(0)
		else
			dim package as EA.Package
			set package = Repository.GetPackageByID(element.PackageID)
			set stringDataType = package.Elements.AddNew("String", datatypeStereotype)
			stringDataType.Update
		end if
		'set generalization to stringDatatype
		dim connector as EA.Connector
		dim found
		found = false
		for each connector in element.Connectors
			if connector.Type = "Generalization" _
			  and connector.SupplierID = stringDataType.ElementID then
				found = true
				exit for
			end if
		next
		'create if needed
		if not found then
			dim generalization as EA.Connector
			set generalization = element.Connectors.AddNew("", "Generalization")
			generalization.SupplierID = stringDataType.ElementID
			generalization.Update
		end if
	end if
end function

function convertElementToCamelCase(element)
	Repository.WriteOutput outPutName, now() & " CamelCasing '" & element.Name &"'", 0
	'convert element
	convertNamingtoCamelCase element
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

function convertAssociationsToJSON(element, attributesDictionary)
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
			'create attribute
			dim attributeName
			attributeName = connector.SupplierEnd.Role
			if len(attributeName) = 0 then
				attributeName  = targetElement.Name
			end if
			'check if attribute doesn't exist yet
			if not attributesDictionary.Exists(attributeName) then
				'report progress
				Repository.WriteOutput outPutName, now() & " Creating attribute '"& sourceElement.Name & "." & attributeName &"'", 0
				dim attribute
				set attribute = sourceElement.Attributes.AddNew(attributeName, targetElement.Name)
				attribute.StereotypeEx = attributeStereotype
				attribute.ClassifierID = targetElement.ElementID
				'set multiplicity
				dim multiplicity
				multiplicity = getMultiplicityFromCardinality(connector.SupplierEnd.Cardinality)
				attribute.lowerBound = multiplicity(0)
				attribute.UpperBound = multiplicity(1)
				if attribute.UpperBound <> "1" then
					attribute.Name = attribute.Name & "List"
				end if
				attribute.Visibility = "Private"
				attribute.Notes = connector.Notes
				attribute.Update
				'add to dictionary
				attributesDictionary.Add attributeName, attribute
			end if
			'delete association? or transform into dependency? or leave it?
			'element.Connectors.DeleteAt i, false
		end if
	next
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
