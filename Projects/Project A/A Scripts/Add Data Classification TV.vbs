'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit
!INC Atrias Scripts.Util

!INC Local Scripts.EAConstants-VBScript
'
' Script Name: Add Data classification TV
' Author: Geert Bellekens
' Purpose: Adds the Data Classification tagged value to all elements and attributes owned by the selected package (recursive)
' Date: 2016-10-07
'
' Project Browser Script main function
'
const outPutName = "Add Data classification"

sub main()
	
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage
	if not selectedPackage is nothing then
		'create output tab
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		'set timestamp
		Repository.WriteOutput outPutName, "Starting Add Data classification " & now(), 0
		'start processing
		AddTaggedValuetoOwnedElements selectedPackage, "Atrias::Data Classification", "NP-NBC"
		'set timestamp
		Repository.WriteOutput outPutName, "Finished Add Data classification " & now(), 0
	end if
end sub

'adds the initial CR to all elements owned by this package and its subpackages.
'The CR is also added to all attributes of the elements found.
Function AddTaggedValuetoOwnedElements(package, tvName, tvValue)
    Dim element as EA.Element
    Dim taggedValue
    For Each element In package.Elements
		if element.Type = "Class" then 'only for classes
			Repository.WriteOutput outPutName, "Processing element  " & element.Name, 0
			'first check if the tagged value already exists
			set taggedValue = Nothing
			dim existingTag
			set existingTag = nothing
			For Each existingTag In element.TaggedValues
				If existingTag.name = tvName Then
					Set taggedValue = existingTag
					Exit For
				end if
			Next
			'if it doesn't exist yet then add it
			If taggedValue Is Nothing Then
				Set taggedValue = element.TaggedValues.AddNew(tvName, "")
			end if
			'now update the value
			If not taggedValue Is Nothing Then
				taggedValue.Value = tvValue
				taggedValue.Update
			End If
			'Do the same for all attributes
			AddTaggedValuetoOwnedAttributes element, tvName, tvValue
		end if
    Next
	'then recursively do all sub-packages
	dim subPackage
	For Each subPackage In package.Packages
		AddTaggedValuetoOwnedElements subPackage, tvName, tvValue
	Next
End Function

'adds the CR as tagged value to all attributes owned by the given element
function AddTaggedValuetoOwnedAttributes(element, tvName, tvValue)
	Dim attribute
    Dim taggedValue
    For Each attribute In element.Attributes
		Repository.WriteOutput outPutName, "Processing attribute  " & element.Name & "." & attribute.Name, 0
        'first check if the tagged value already exists
		set taggedValue = Nothing
		dim existingTag
		set existingTag = nothing
        For Each existingTag In Attribute.TaggedValues
            If existingTag.name = tvName Then
				Set taggedValue = existingTag
				Exit For
			end if
        Next
        'if it doesn't exist yet then add it
        If taggedValue Is Nothing Then
            Set taggedValue = attribute.TaggedValues.AddNew(tvName, "")
		end if
		'now update the value
		If not taggedValue Is Nothing Then
            taggedValue.Value = tvValue
            taggedValue.Update
        End If
    Next
end function

main