'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.Util

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
Dim XSDBaseTypes
XSDBaseTypes = Array("string","boolean","decimal","float","double","duration","dateTime","time","date","gYearMonth","gYear","gMonthDay","gDay","gMonth","hexBinary","base64Binary","anyURI","QName","integer","long","int")

sub main
	dim response
	response = Msgbox("This script will move all underlying elements into one package!" & vbnewLine & "This should only be done when making an XSD from the LDM." & vbnewLine & " Are you sure?", vbYesNo+vbExclamation, "Post XSD transformation")
	if response = vbYes then
		'Create new package for the whole of the schema
		dim package as EA.Package 
		set package = Repository.GetTreeSelectedPackage()
		dim schemaPackage as EA.Package
		set schemaPackage = package.Packages.AddNew(package.Name,"package")
		schemaPackage.Update
		schemaPackage.Element.Stereotype = "XSDSchema"
		schemaPackage.Update
		' move all elements from the subpackages to the newly create package
		'-------------------------------------------------------------------
		mergeToSchemaPackage package, schemaPackage 'uncomment for production
		'set schemaPackage = package 'comment out for production
		'-------------------------------------------------------------------
		' fix the elements
		fixElements schemaPackage 
		'fix the connectors
		fixConnectors schemaPackage
		'fix the attributes with a primitive type
		fixAttributePrimitives schemaPackage
		'set lowerbound of attributes to 0
		fixAttributeLowerBound schemaPackage
		'remove abstract classes
		fixAbstractClasses schemaPackage		
		'reload
		Repository.RefreshModelView(package.PackageID)
		msgbox "Finished!"
	end if
end sub

function fixElements(schemaPackage)
	dim element as EA.Element
	for each element in schemaPackage.Elements
		if element.Stereotype = "XSDsimpleType" then
			if element.Attributes.Count = 0 then
				' fix xsdSimpleTypes
				fixXSDsimpleType element
			else
				element.Stereotype = "XSDcomplexType"
				element.Update
			end if
		end if
		'process complex types
		if element.Stereotype = "XSDcomplexType" then
			fixComplexTypes element
		end if
	next	
end function

'fix the complex type.
'copy all attributes and associations from the parent classes to this class (flatten)
'add versioning and timeslicing attributes
'set cardinality to optional for all attributes (and associations ?)
'remove abstract classes
function fixComplexTypes(element)
	dim sourceElement as EA.Element
	dim parentClass as EA.Element
	set sourceElement = getTransformationSource(element)
	'check for tagged values versioned and timesliced
	setTimeSlicingAndVersioning sourceElement, element
	if not element is nothing and element.Abstract = "0" then
		'loop parent classes
		for each parentClass in element.BaseClasses
			copyFromBaseClasses parentClass, element
		next
	end if
end function

function copyFromBaseClasses(sourceElement, targetElement)			
	'copy attributes
	copyAttributes sourceElement, targetElement
	'copy associations
	copyAssociations sourceElement, targetElement
	'copy from parents of source as well
	dim parent as EA.Element
	for each parent in sourceElement.BaseClasses
		copyFromBaseClasses parent, targetElement
	next
end function

'check if the source element has the timeslicing or versioning tagged value set to "yes".
'if so it will apply the pattern to the element
function setTimeSlicingAndVersioning(sourceElement, targetElement)
'		dim sourceElement as EA.Element
'		dim targetElement as EA.Element
	dim taggedValue as EA.TaggedValue
	dim attribute as EA.Attribute
	for each taggedValue in sourceElement.TaggedValues
		if taggedValue.Name = "Timesliced" then
			'add attributes
			set attribute = targetElement.Attributes.AddNew("StartDate","dateTime")
			attribute.Update
			set attribute = targetElement.Attributes.AddNew("EndDate","dateTime")
			attribute.Update
		elseif taggedValue.Name = "Versioned" then
			'add attributes
			set attribute = targetElement.Attributes.AddNew("ValidFromDate","dateTime")
			attribute.Update
			set attribute = targetElement.Attributes.AddNew("ValidUntilDate","dateTime")
			attribute.Update
		end if
	next
end function

'copies all attributes from the sourceElement to the target element
function copyAttributes (sourceElement, targetElement)
'		dim sourceElement as EA.Element
'		dim targetElement as EA.Element
	dim attribute as EA.Attribute
	dim newAttribute as EA.Attribute
	for each attribute in sourceElement.Attributes
		'don't copy the timeslicing and versioning attributes
		if attribute.Name <> "StartDate" and attribute.Name <> "EndDate" and attribute.Name <> "ValidFromDate" and attribute.Name <> "ValidUntilDate" then
			set newAttribute = targetElement.Attributes.AddNew(attribute.Name,attribute.Type)
			'newAttribute.Type = attribute.Type
			newAttribute.ClassifierID = attribute.ClassifierID
			newAttribute.LowerBound = attribute.LowerBound
			newAttribute.UpperBound = attribute.UpperBound
			newAttribute.Update
		end if
	next
end function

'copies all associations from the sourceElement to the targetElement
function copyAssociations (sourceElement, targetElement)
'		dim sourceElement as EA.Element
'		dim targetElement as EA.Element
	dim supplierElement as EA.Element
	dim association as EA.Connector
	dim newAssociation as EA.Connector
	for each association in sourceElement.Connectors
		if association.Type = "Association" or association.Type = "Aggregation" then
			if association.ClientID = sourceElement.ElementID then
				set newAssociation = targetElement.Connectors.AddNew(association.Name, association.Type)
				newAssociation.SupplierID = association.SupplierID
			else
				set supplierElement = Repository.GetElementByID(association.ClientID)
				set newAssociation = supplierElement.Connectors.AddNew(association.Name, association.Type)
				newAssociation.SupplierID = targetElement.ElementID
			end if
			newAssociation.Update
			'set the ends
			'source
			newAssociation.ClientEnd.Cardinality = association.ClientEnd.Cardinality
			newAssociation.ClientEnd.Role = association.ClientEnd.Role
			newAssociation.ClientEnd.Update
			'target
			newAssociation.SupplierEnd.Cardinality = association.SupplierEnd.Cardinality
			newAssociation.SupplierEnd.Role = association.SupplierEnd.Role
			newAssociation.SupplierEnd.Update
		end if
	next
end function

'finds the element from which the given element was tranformed
function getTransformationSource(element)
	dim sourceElement as EA.Element
	set sourceElement = nothing
	dim sqlFindSource
	dim sourceElements
	sqlFindSource = "select o.[Object_ID] from  t_object o " & _
					"inner join t_xref x on x.[Supplier] = o.[ea_guid] " & _
					"where x.TYPE = 'Transformation' " & _
					"and x.[Client] =  '" & element.ElementGUID & "'"
	set sourceElements = getElementsFromQuery(sqlFindSource)
	if sourceElements.Count > 0 then
		set sourceElement = sourceElements(0)
	end if
	set getTransformationSource = sourceElement
end function

function fixXSDsimpleType(element)
	'find the element it was transformed from
	dim sourceElement as EA.Element
	set sourceElement = getTransformationSource(element)
	if not sourceElement is nothing then
		'copy the tagged values
		copyTaggedValues sourceElement, element
		'determine the "parent" type
		dim baseClass as EA.Element
		dim baseXSDType
		baseXSDType = "string" 'default
		for each baseClass in sourceElement.BaseClasses
			if Ubound(Filter(XSDBaseTypes, baseClass.Name )) > -1 then
				'found the base class
				baseXSDType = baseClass.Name
			end if
		next
		'set the base type
		element.Genlinks = "Parent=" & baseXSDType & ";"
		element.Update
	end if
end function


function mergeToSchemaPackage (package, schemaPackage)
	dim subPackage as EA.Package
	dim i
	for i = package.Packages.Count -1 to i = 1 step -1
		set subPackage = package.Packages.GetAt(i)
		'should only be done on XSDschema packages
		if subPackage.Element.Stereotype = "XSDschema" then
			dim element as EA.Element
			'move elements
			for each element in subPackage.Elements
				element.PackageID = schemaPackage.PackageID
				element.Update
			next
			'remove original package
			package.Packages.DeleteAt i,false
		end if
	next
end function

function fixConnectors(package)
	dim SQLgetConnectors
	SQLgetConnectors = "select distinct c.Connector_ID from " & _
						" ( " & _
						" select source.StartID, source.EndID from  " & _
						" ( " & _
						" select o.[Object_ID] AS StartID, con.[End_Object_ID] AS EndID " & _
						" from (t_object o " & _
						" inner join t_connector con on con.[Start_Object_ID] = o.[Object_ID]) " & _
						" where o.package_ID = "& package.PackageID & _
						" union all " & _
						" select o.[Object_ID] AS StartID, con.[Start_Object_ID] AS EndID " & _
						" from (t_object o " & _
						" inner join t_connector con on con.[End_Object_ID] = o.[Object_ID]) " & _
						" where o.package_ID = "& package.PackageID & _
						" ) source " & _
						" group by source.StartID, source.EndID " & _
						" having count(*) > 1 " & _
						" ) grouped, t_connector c " & _
						" where (c.[Start_Object_ID] = grouped.StartID  " & _
						"       and c.[End_Object_ID] = grouped.EndID) " & _
						"       or " & _
						"       (c.[Start_Object_ID] = grouped.EndID  " & _
						"       and c.[End_Object_ID] = grouped.StartID) "
	dim connectors
	set connectors = getConnectorsFromQuery(SQLgetConnectors)
	dim connector as EA.Connector
	dim changed
	changed = false
	for each connector in connectors
		'set the connector's rolenames
		if len(connector.ClientEnd.Role) < 1 then
			connector.ClientEnd.Role = replace(connector.Name, " ", "_")
			connector.ClientEnd.Update
			changed = true
		end if
		if len(connector.SupplierEnd.Role) < 1 then
			connector.SupplierEnd.Role = replace(connector.Name, " ", "_")
			connector.SupplierEnd.Update
			changed = true
		end if
		if changed then
			connector.Update
		end if
	next
	'now get *all* associations between elements in this package in order to copy it in the other direction
	SQLgetConnectors = "select c.Connector_ID from ((t_connector c " & _
						" inner join t_object sob on c.Start_Object_ID = sob.Object_ID) " & _
						" inner join t_object tob on c.End_Object_ID = tob.Object_ID) " & _
						" where  " & _
						" c.Connector_Type in ('Association','Aggregation') " & _
						" and sob.Package_ID = "& package.PackageID & _
						" and tob.Package_ID = "& package.PackageID
	set connectors = getConnectorsFromQuery(SQLgetConnectors)
	Session.Output "selectedconnectors: " & connectors.Count
	for each connector in connectors
		'set the lower bound to 0 on both ends
		setEndOptional connector.ClientEnd
		setEndOptional connector.SupplierEnd
	next
end function

function setEndOptional (associationEnd)
'		dim associationEnd as EA.ConnectorEnd
	select case associationEnd.Cardinality
		case "1"
			associationEnd.Cardinality = "0..1"
		case "1..1"
			associationEnd.Cardinality = "0..1"
		case "1..*"
			associationEnd.Cardinality = "0..*"
	end select
	associationEnd.Update
end function

function copyAssociationEnd(source, target)
'		dim source as EA.ConnectorEnd
'		dim target as EA.ConnectorEnd
	target.Aggregation = source.Aggregation
	target.Alias = source.Alias
	target.AllowDuplicates = source.AllowDuplicates
	target.Cardinality = source.Cardinality
	target.Constraint = source.Constraint
	target.Containment = source.Containment
	target.Derived = source.Derived
	target.DerivedUnion = source.DerivedUnion
	target.IsChangeable = source.IsChangeable
	target.Navigable = source.Navigable
	target.Ordering = source.Ordering
	target.OwnedByClassifier = source.OwnedByClassifier
	target.Qualifier = source.Qualifier
	target.Role = source.Role
	target.RoleNote = source.RoleNote
	target.StereotypeEx = source.StereotypeEx
	target.Visibility = source.Visibility
	'save changes
	target.Update
end function

function fixAttributePrimitives(package)	
	
	dim sqlUpdate
	sqlUpdate = "update attr set attr.Classifier = 0  " & _
				" from t_attribute attr  " & _
				" inner join t_object o on attr.object_id = o.object_id " & _
				" where attr.[TYPE] in ('string','boolean','decimal','float','double','duration','dateTime','time','date','gYearMonth','gYear','gMonthDay','gDay','gMonth','hexBinary','base64Binary','anyURI','QName','integer','long','int') " & _
				" and attr.Classifier > 0 " & _
				" and o.[Package_ID] =  " & package.PackageID
	Repository.Execute sqlUpdate
end function

function fixAttributeLowerBound(package)
	dim sqlUpdate
	sqlUpdate = "update attr set attr.LowerBound = '0'  " & _
				" from t_attribute attr  " & _
				" inner join t_object o on attr.object_id = o.object_id  " & _
				" where o.[Package_ID] =  " & package.PackageID
	Repository.Execute sqlUpdate
end function

function fixAbstractClasses (package)
	dim i
	dim element as EA.Element
	for i = package.Elements.Count -1 to 0 step -1
		set element = package.Elements(i)
		if element.Abstract = "1" then
			package.Elements.DeleteAt i,true
		end if
	next
end function
main