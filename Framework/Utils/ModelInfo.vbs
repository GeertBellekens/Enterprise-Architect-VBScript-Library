'[path=\Framework\Utils]
'[group=Utils]
'Author: Geert Bellekens
'Date: 2016-01-08

'TODO: use caching to speed up the process
Dim elementCache 	
dim packageCache

Private Sub Module_Initialize()
	set elementCache = CreateObject("Scripting.Dictionary")
	set packageCache = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Module_Terminate()
	set elementCache = nothing
	set packageCache = nothing
End Sub

'Group of functions related to information from the model

'returns the fully qualified name for the given item.
'this is the full path of the element divided by dots e.g. "Root.GrandParent.Parent.Item"
function getFullyQualifiedName(item)
	dim fqn, parent
	fqn = item.name
	set parent = getParent(item)
	if not parent is nothing then
		fqn = getFullyQualifiedName(parent) & "." & fqn
	end if
	'return fully qualified name
	getFullyQualifiedName = fqn
end function

'returns the parent object for the given object
function getParent(item)
	dim itemType, parentID, parent
	set parent = nothing
	itemType = TypeName(item)
	select case itemType
		case "IDualElement"
			parentID = item.ParentID
			if parentID > 0 then
				set parent = Repository.GetElementByID(parentID)
			else
				set parent = Repository.GetPackageByID(item.PackageID)
			end if
		case "IDualConnector"
			set parent = Repository.GetElementByID(item.ClientID)
		case "IDualAttribute"
			dim attribute as EA.Attribute
			set parent = Repository.GetElementByID(item.parentID)
		case "IDualDiagram"
			parentID = item.ParentID
			if parentID > 0 then
				set parent = Repository.GetElementByID(parentID)
			else
				set parent = Repository.GetPackageByID(item.PackageID)
			end if
		case "IDualPackage"
			parentID = item.ParentID
			if parentID > 0 then
				set parent = Repository.GetPackageByID(parentID)
			end if
	end select
	set getParent = parent
end function

'sub test
'	dim selectedElement, fqn
'	set selectedElement = Repository.GetContextObject()
'	fqn = getFullyQualifiedName(selectedElement)
'	Session.Output "FQN: " & fqn
'end sub
'test