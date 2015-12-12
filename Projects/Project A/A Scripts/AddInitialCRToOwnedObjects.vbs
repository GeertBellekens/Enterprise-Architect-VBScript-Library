'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
'
' Script Name: AddInitialCRtoOwnedObject
' Author: Geert Bellekens
' Purpose: adds the initial CR as a tagged value to all owned elements int this package
' and its subpackackages.
' It also add's the tagged value to all attributes of those elements
' Date: 30/01/2015
'
' Project Browser Script main function
'

Dim CRGuid
CRGuid = "{03DE6415-5FD7-40e3-8ADB-9B8FDE574B95}"  'Change this GUID into the GUID of the actual initial CR

sub OnProjectBrowserScript()
	
	' Get the type of element selected in the Project Browser
	dim treeSelectedType
	treeSelectedType = Repository.GetTreeSelectedItemType()
	select case treeSelectedType
	case otPackage
			' We only do something for a package
			dim thePackage as EA.Package
			set thePackage = Repository.GetTreeSelectedObject()
			AddCRToOwnedElements thePackage
			MsgBox "Ready!"
		case else
			' Error message
			Session.Prompt "This script does not support items of this type.", promptOK			
	end select
end sub

'adds the initial CR to all elements owned by this package and its subpackages.
'The CR is also added to all attributes of the elements found.
Function AddCRToOwnedElements(package)
    Dim element
    Dim taggedValue
    For Each element In package.Elements
        'first check if the tagged value already exists
		set taggedValue = Nothing
		dim existingTag
		set existingTag = nothing
        For Each existingTag In element.TaggedValues
            If existingTag.name = "CR" Then
				Set taggedValue = existingTag
				Exit For
			end if
        Next
        'if it doesn't exist yet then add it
        If taggedValue Is Nothing Then
            Set taggedValue = element.TaggedValues.AddNew("CR", "")
            taggedValue.Value = CRGuid
            taggedValue.Update
        End If
		'Do the same for all attributes
		AddCRtoOwnedAttributes element
    Next
	'then recursively do all sub-packages
	dim subPackage
	For Each subPackage In package.Packages
		AddCRToOwnedElements subPackage
	Next
End Function

'adds the CR as tagged value to all attributes owned by the given element
function AddCRtoOwnedAttributes(element)
	Dim attribute
    Dim taggedValue
    For Each attribute In element.Attributes
        'first check if the tagged value already exists
		set taggedValue = Nothing
		dim existingTag
		set existingTag = nothing
        For Each existingTag In attribute.TaggedValues
            If existingTag.name = "CR" Then
				Set taggedValue = existingTag
				Exit For
			end if
        Next
        'if it doesn't exist yet then add it
        If taggedValue Is Nothing Then
            Set taggedValue = attribute.TaggedValues.AddNew("CR", "")
            taggedValue.Value = CRGuid
            taggedValue.Update
        End If
    Next
end function

OnProjectBrowserScript