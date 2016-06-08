'[path=\Framework\Utils]
'[group=Utils]
'Author: Geert Bellekens
'Date: 2016-01-08

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

'sub test
'	dim selectedElement, fqn
'	set selectedElement = Repository.GetContextObject()
'	fqn = getFullyQualifiedName(selectedElement)
'	Session.Output "FQN: " & fqn
'end sub
'test