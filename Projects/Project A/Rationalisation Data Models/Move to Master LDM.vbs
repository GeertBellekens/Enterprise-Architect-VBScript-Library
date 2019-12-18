'[path=\Projects\Project A\Rationalisation Data Models]
'[group=Rationalisation Data Models]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Move to Master LDM
' Author: Geert Bellekens
' Purpose: Move a class or attribute from the "subset" LDM to the master LDM
' Date: 2019-02-21
'
const masterLDMGuid = "{D66A3533-4125-4175-975B-6B4D9C90A3B2}"

sub main
	'get master LDM package
	dim masterLDMPackage as EA.Package
	set masterLDMPackage = Repository.GetPackageByGuid(masterLDMGuid)
	if masterLDMPackage is nothing then
		Msgbox "No master LDM package found for guid " & masterLDMGuid
		exit sub
	end if
	dim master
	'get the master package id tree string
	dim masterPackageIDs
	masterPackageIDs = getPackageTreeIDString(masterLDMPackage)
	'get selected item
	dim selectedItem
	set selectedItem = Repository.GetContextObject()
	if selectedItem.ObjectType = otElement then
		moveElement selectedItem, masterPackageIDs
	elseif selectedItem.ObjectType = otAttribute  then
		moveAttribute selectedItem, masterPackageIDs
	elseif selectedItem.ObjectType = otConnector  then
		'TODO
		Msgbox "ERROR: moving connectors is not supported (yet)"
	end if
	'tell the user we are finished
	Msgbox "Finished!"
end sub

function moveAttribute(attribute, masterPackageIDs)
	'get equivalent owner
	dim targetOnwers
	set targetOnwers = findEquivalents(attribute.ParentID , masterPackageIDs )
	if targetOnwers.Count <> 1 then
		Msgbox "ERROR: multiple or no target owners found for " & attribute.Name 
		exit function
	end if
	dim targetOwner as EA.Element
	set targetOwner = targetOnwers(0)
	'actually copy the attribute to the target owner
	copyAttribute attribute, targetOwner, masterPackageIDs
end function

public function moveElement( element, masterPackageIDs)
	dim parentPackage as EA.Package
	set parentPackage = Repository.GetPackageByID(element.PackageID)
	'find master LDM package
	dim sqlFindMasterPackage
	sqlFindMasterPackage = "select p.Package_ID from t_package p " & _
							"where p.Parent_ID in (" & masterPackageIDs & ")" & _
							" and p.Name = '" & parentPackage.Name & "'"
	dim masterPackages
	set masterPackages = getPackagesFromQuery(sqlFindMasterPackage)
	if masterPackages.Count <> 1 then
		Msgbox "ERROR: multiple or no master packages found for element " & element.Name 
		exit function
	end if
	dim masterPackage
	set masterPackage = masterPackages(0)
	'duplicate element
	copyElement element, masterPackage, masterPackageIDs
end function

function copyElement(element, targetPackage, masterPackageIDs)
	dim duplicateElement as EA.Element
	set duplicateElement = targetPackage.Elements.AddNew(element.Name, element.Type)
	'copy properties
	duplicateElement.Notes = element.Notes
	duplicateElement.Stereotype = element.Stereotype
	duplicateElement.Update
	'copy attributes
	dim attribute as EA.Attribute
	for each attribute in element.Attributes
		copyAttribute attribute, duplicateElement, masterPackageIDs
	next
	'copy tagged values
	copyAllTaggedValues element, duplicateElement
	'copy connectors
	dim connector as EA.Connector
	for each connector in element.Connectors
		'only connectors that start at the element
		if connector.ClientID = element.ElementID then
			copyConnector connector, duplicateElement, masterPackageIDs
		end if
	next
	'add traceability
	dim trace as EA.Connector
	set trace = element.Connectors.AddNew("","Abstraction")
	trace.Stereotype = "trace"
	trace.SupplierID = duplicateElement.ElementID
	trace.Update
end function

function copyConnector(connector, targetElement, masterPackageIDs)
	dim duplicateConnector as EA.Connector
	'find the otherEnd
	dim otherends
	set otherEnds = findEquivalents(connector.SupplierID , masterPackageIDs)
	if otherEnds.count <> 1 then
		MsgBox "ERROR: other ends found <> 1 for connector " & targetElement.Name & "." & connector.Name
		exit function
	end if
	dim otherend as EA.Element		
	set otherend = otherends(0)
	'create duplicate connector
	set duplicateConnector = targetElement.connectors.AddNew(connector.Name, connector.Type)
	'set properties
	duplicateConnector.Notes = connector.Notes
	duplicateConnector.Stereotype = connector.Stereotype
	duplicateConnector.SupplierID = otherend.ElementID
	duplicateConnector.Direction = connector.Direction
	duplicateConnector.Update
	'set connector end properties
	copyConnectorEnd connector.ClientEnd, duplicateConnector.ClientEnd
	copyConnectorEnd connector.SupplierEnd, duplicateConnector.SupplierEnd
	'copy taggedvalues
	copyAllTaggedValues connector, duplicateConnector
	'set traceability
	if connector.Type = "Association" _
	  or connector.Type = "Aggregation" then	
		dim traceabilityTag as EA.TaggedValue
		set traceabilityTag = getExistingOrNewTaggedValue(connector, "sourceAssociation")
		traceabilityTag.Value = duplicateConnector.ConnectorGUID
		traceabilityTag.Update
	end if
end function

function copyConnectorEnd(originalEnd, duplicateEnd)
	duplicateEnd.Aggregation = originalEnd.Aggregation
	duplicateEnd.Cardinality = originalEnd.Cardinality
	duplicateEnd.Role = originalEnd.Role
	duplicateEnd.Navigable = originalEnd.Navigable
	duplicateEnd.Update
end function


function findEquivalents(elementID, masterPackageIDs)
	dim sqlFindEquivalent
	sqlFindEquivalent = "select o.Object_ID from t_object o 							  " & _
						" inner join t_object oo on oo.Name = o.Name 					  " & _
						"					and oo.Object_ID = " & elementID   & _
						" where o.Package_ID in (" & masterPackageIDs & ") "
	set findEquivalents = getElementsFromQuery(sqlFindEquivalent)
end function

function copyAttribute(attribute, targetElement, masterPackageIDs)
	dim duplicateAttribute as EA.Attribute
	set duplicateAttribute = targetElement.Attributes.AddNew(attribute.Name, attribute.Type)
	duplicateAttribute.Notes = attribute.Notes
	duplicateAttribute.Stereotype = attribute.Stereotype
	duplicateAttribute.Visibility = attribute.Visibility
	duplicateAttribute.IsID = attribute.IsID
	duplicateAttribute.LowerBound = attribute.LowerBound
	duplicateAttribute.UpperBound = attribute.UpperBound
	duplicateAttribute.Pos = attribute.Pos
	duplicateAttribute.Update
	'find the corresponding datatype
	if attribute.ClassifierID > 0 then
		dim datatypes
		set datatypes = findEquivalents(attribute.ClassifierID , masterPackageIDs)
		if datatypes.count <> 1 then
			MsgBox "ERROR: datatypes found <> 1 for attribute " & attribute.Name
			exit function
		end if
		dim datatype as EA.Element		
		set datatype = datatypes(0)
		duplicateAttribute.ClassifierID = datatype.ElementID
		duplicateAttribute.Update
	end if
	'copy tagged values
	copyAllTaggedValues attribute, duplicateAttribute
	'add traceability
	dim traceabilityTag as EA.TaggedValue
	set traceabilityTag = getExistingOrNewTaggedValue(attribute, "sourceAttribute")
	traceabilityTag.Value = duplicateAttribute.AttributeGUID
	traceabilityTag.Update
end function

main