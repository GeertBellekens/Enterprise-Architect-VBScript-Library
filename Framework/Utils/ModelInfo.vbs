'[path=\Framework\Utils]
'[group=Utils]
'Author: Geert Bellekens
'Date: 2016-01-08

'sub test
'	dim selectedElement, fqn
'	set selectedElement = Repository.GetContextObject()
'	fqn = getFullyQualifiedName(selectedElement)
'	Session.Output "FQN: " & fqn
'end sub
'test

'TODO: use caching to speed up the process
Dim elementCache 	
dim packageCache
'initialise the cache objects
init

Private Sub Module_Initialize()
	'the cache contains the ID (package or element ID) and the fully qualified name of the element or package.
	'this only seems to work in debug mode for some reason?
'	set elementCache = CreateObject("Scripting.Dictionary")
'	set packageCache = CreateObject("Scripting.Dictionary")
	init
End Sub

Private Sub Module_Terminate()
	set elementCache = nothing
	set packageCache = nothing
End Sub

private sub init()
	if not IsObject(elementCache) then
		set elementCache = CreateObject("Scripting.Dictionary")
	end if
	if not IsObject(packageCache) then
		set packageCache = CreateObject("Scripting.Dictionary")
	end if
end sub
'Group of functions related to information from the model

'returns the fully qualified name for the given item.
'this is the full path of the element divided by dots e.g. "Root.GrandParent.Parent.Item"
function getFullyQualifiedName(item)
	dim fqn, parentfqn
	fqn = ""
	'add the parent part
	parentFQN = getParentFQN(item)
	if len(parentFQN) > 0 then
		fqn = parentFQN & "."
	end	if
	fqn = fqn & getItemName(item)
	getFullyQualifiedName = fqn
end function

'returns the parent object for the given object
function getParentFQN(item)
	dim itemType, parentID, parent, parentFQN, packageID
	parentID = 0
	packageID = 0
	parentFQN = ""
	set parent = nothing
	itemType = TypeName(item)
	select case itemType
		case "IDualElement"
			parentID = item.ParentID
			packageID = item.PackageID
		case "IDualConnector"
			parentID = item.ClientID
		case "IDualAttribute"
			parentID = (item.parentID)
		case "IDualDiagram"
			parentID = item.ParentID
			packageID = item.PackageID
		case "IDualPackage"
			packageID = item.ParentID
	end select
	if parentID > 0 then
		'the item is owned by an element
		'first check if the element is in the cache already
		if elementCache.Exists(parentID) then
			'get the FQN from the cache
			parentFQN = elementCache(parentID)
		else
			'not in the cache, get the element and its FQN
			set parent = Repository.GetElementByID(parentID)
			parentFQN = getFullyQualifiedName(parent)
			'add it to the cache
			elementCache.Add parentID, parentFQN
		end if
	elseif packageID > 0 then
		'the item is owned by a package
		'first check if it is in the cache already
		if packageCache.Exists(packageID) then
			'get the FQN from the cache
			parentFQN = packageCache(packageID)
		else
			'not in the cache
			set parent = Repository.GetPackageByID(packageID)
			parentFQN = getFullyQualifiedName(parent)
			'add it to the cache
			packageCache.Add packageID, parentFQN
		end if
	end if
	getParentFQN = parentFQN
end function

function getItemName(item)
	dim itemName
	itemName = item.Name
	if len(itemName) = 0 then
		itemName = "[Anonymous]"
	end if
	getItemName = itemName
end function

'gets the attributes by the id's returned by the given query
function getAttributesByQuery(sqlQuery)
	dim xmlResult
	xmlResult = Repository.SQLQuery(sqlQuery)
	dim attributeIDs
	attributeIDs = convertQueryResultToArray(xmlResult)
	dim attributes 
	set attributes = CreateObject("System.Collections.ArrayList")
	dim attributeID
	dim attribute as EA.Attribute
	for each attributeID in attributeIDs
		if attributeID > 0 then
			set attribute = Repository.GetAttributeByID(attributeID)
			if not attribute is nothing then
				attributes.Add(attribute)
			end if
		end if
	next
	set getAttributesByQuery = attributes
end function

function getConnectorsFromQuery(sqlQuery)
	dim xmlResult
	xmlResult = Repository.SQLQuery(sqlQuery)
	dim connectorIDs
	connectorIDs = convertQueryResultToArray(xmlResult)
	dim connectors 
	set connectors = CreateObject("System.Collections.ArrayList")
	dim connectorID
	dim connector as EA.Connector
	for each connectorID in connectorIDs
		if connectorID > 0 then
			set connector = Repository.GetConnectorByID(connectorID)
			if not connector is nothing then
				connectors.Add(connector)
			end if
		end if
	next
	set getConnectorsFromQuery = connectors
end function



function selectObjectFromQualifiedName(rootPackage,rootElement, qualifiedName, seperator)
	dim foundObject
	set foundObject = nothing
	'devide qualified name into parts
	dim parts
	parts = Split(qualifiedName,seperator)
	if ubound(parts) >= 0 then
		dim rootPart
		dim rootOK, rootName
		rootOK = false
		dim hasElement
		hasElement = false
		rootPart = parts(0)
		'check if we have a root element
		if not rootElement is nothing then
			rootName = rootElement.Name
			hasElement = true
		else
			rootName = rootPackage.Name
		end if
		'check the rootname
		if lcase(rootName) = lcase(rootPart) then
			rootOK = true
		end if
		if rootOK then
			if ubound(parts) > 0 then
				dim childPart
				childPart = parts(1)
				'check attributes if the childpart is the last part
				'debug
				'Session.Output "Searching root = " & rootName & " child = " & childPart & " qualifiedName = " & qualifiedName
				if hasElement AND ubound(parts) = 1 then
					set foundObject = getAttributeByName(rootElement, childPart)
					'if no attribute found we try to find an operation
					if foundObject is nothing then
						set foundObject = getOperationByName(rootElement, childPart)
					end if
				end if
				
				'nothing found we go deeper
				if foundObject is nothing then
					dim subElement as EA.Element
					dim subPackage as EA.Package
					set subPackage  = nothing
					if hasElement then
						set subElement = getSubElementByName(rootElement,childPart)
					else
						set subElement = getSubElementByName(rootPackage,childPart)
						if subElement is nothing then
							set subPackage = getSubPackageByName(rootPackage, childPart)
						end if
					end if
					'go deeper
					dim substring
					substring = mid(qualifiedName, len(rootPart) + len(seperator) +1)
					if not subElement is nothing then
						set foundObject = selectObjectFromQualifiedName(nothing,subElement, substring, seperator)
					elseif not subPackage is nothing then
						set foundObject = selectObjectFromQualifiedName(subPackage,nothing, substring, seperator)
					end if
				end if
			else
				'only one part is given, return root
				if hasElement then
					set foundObject = rootElement
				else
					set foundObject = rootPackage
				end if
			end if
		end if
	end if
	set selectObjectFromQualifiedName = foundObject
end function

function getAttributeByName(element, attributeName)
	set getAttributeByName = nothing
	if not element is nothing then
		dim attribute as EA.Attribute
		for each attribute in element.Attributes
			if lcase(attribute.Name) = lcase(attributeName) then
				set getAttributeByName = attribute
				exit for
			end if
		next
	end if
end function

function getOperationByName(element, operationName)
	set getOperationByName = nothing
	if not element is nothing then
		dim operation as EA.Method
		for each operation in element.Methods
			if lcase(operation.Name) = lcase(operationName) then
				set getOperationByName = operation
				exit for
			end if
		next
	end if
end function

function getSubElementByName(owner, elementName)
	set getSubElementByName = nothing
	if not owner is nothing then
		dim subElement as EA.Element
		for each subElement in owner.Elements
			if lcase(subElement.Name) = lcase(elementName) then
				set getSubElementByName = subElement
				exit for
			end if
		next
	end if
end function

function getSubPackageByName(package, packageName)
	set getSubPackageByName = nothing
	if not package is nothing then
		dim subPackage as EA.Package
		for each subPackage in package.Packages
			if lcase(subPackage.Name) = lcase(packageName) then
				set getSubPackageByName = subPackage
				exit for
			end if
		next
	end if
end function


'returns an ArrayList with the elements accordin tot he ObjectID's in the given query
function getElementsFromQuery(sqlQuery)
	dim elements 
	set elements = Repository.GetElementSet(sqlQuery,2)
	dim result
	set result = CreateObject("System.Collections.ArrayList")
	dim element
	for each element in elements
		result.Add Element
	next
	set getElementsFromQuery = result
end function

'returns a comma separated string with the package id's the given package and all subpackages recusively
function getPackageTreeIDString(package)
	dim packageTree
	set packageTree = getPackageTree(package)
	getPackageTreeIDString = makePackageIDString(packageTree)
end function

'returns an ArrayList of the given package and all its subpackages recursively
function getPackageTree(package)
	dim packageList
	set packageList = CreateObject("System.Collections.ArrayList")
	addPackagesToList package, packageList
	set getPackageTree = packageList
end function

'add the given package and all subPackges to the list (recursively
function addPackagesToList(package, packageList)
	dim subPackage as EA.Package
	'add the package itself
	packageList.Add package
	'add subpackages
	for each subPackage in package.Packages
		addPackagesToList subPackage, packageList
	next
end function

'make an id string out of the package ID of the given packages
function makePackageIDString(packages)
	dim package as EA.Package
	dim idString
	idString = ""
	dim addComma 
	addComma = false
	for each package in packages
		if addComma then
			idString = idString & ","
		else
			addComma = true
		end if
		idString = idString & package.PackageID
	next 
	'if there are no packages then we return "0"
	if packages.Count = 0 then
		idString = "0"
	end if
	'return idString
	makePackageIDString = idString
end function

'Returns the value of the tagged value with the given name (case insensitive)
'If there is no tagged value with the given name, an empty string is returned
'This function can be used with anything that can have tagged values
function getTaggedValueValue(owner, taggedValueName)
	dim taggedValue as EA.TaggedValue
	getTaggedValueValue = ""
	for each taggedValue in owner.TaggedValues
		if lcase(taggedValueName) = lcase(taggedValue.Name) then
			getTaggedValueValue = taggedValue.Value
			exit for
		end if
	next
end function

function getOrCreateTaggedValue(owner,taggedValueName)
	'Initialize
	set getOrCreateTaggedValue = nothing
	'check if tagged value exists
	for each taggedValue in owner.TaggedValues
		if lcase(taggedValueName) = lcase(taggedValue.Name) then
			set getOrCreateTaggedValue = taggedValue
			exit function
		end if
	next
	'if not found create new one
	set getOrCreateTaggedValue = owner.TaggedValues.addNew(taggedValueName,"")
end function

function getRootPackage(selectedElement)
	'initialize
	set getRootPackage = nothing
	dim selectedPackage as EA.Package
	if selectedElement is nothing then
		set selectedPackage = Repository.GetTreeSelectedPackage
	else
		if selectedElement.ObjectType = otElement _
		  OR selectedElement.ObjectType = otDiagram then
			set selectedPackage = Repository.GetPackageByID(selectedElement.PackageID)
		elseif selectedElement.ObjectType = otPackage then
			set selectedPackage = selectedElement
		end if
	end if
	if not selectedPackage is nothing then
		if selectedPackage.ParentID = 0 then
			set getRootPackage = selectedPackage
		else
			dim parentPackage
			set parentPackage = Repository.GetPackageByID(selectedPackage.ParentID)
			set getRootPackage = getRootPackage(parentPackage)
		end if 
	end if
end function