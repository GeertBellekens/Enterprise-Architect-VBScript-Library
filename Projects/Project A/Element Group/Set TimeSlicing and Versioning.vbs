'[path=\Projects\Project A\Element Group]
'[group=Element Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.Util

' Script Name: Set Timeslicing and versioning tagged values
' Author: Geert Bellekens
' Purpose: Sets the timeslicing and versioning tagged values to no if they are still TBD
' Date: 29/01/2016
'

'Constants
dim TimeSlicingValue, VersioningValue, TimeSlicingName, VersioningName, TBDValue
TimeSlicingName = "Timesliced"
VersioningName = "Versioned"

TimeSlicingValue = "no"
VersioningValue = "no"
TBDValue = "TBD"

sub main
	dim response
	response = Msgbox("Click on the value for timeslicing and versioning.", vbYesNoCancel+vbQuestion, "Set Timeslicing and Versioning")
	if response = vbYes then
		TimeSlicingValue = "yes"
		VersioningValue = "yes"
	elseif response = vbNo then
		TimeSlicingValue = "no"
		VersioningValue = "no"
	else
		exit sub
	end if
	' get the selected element
	
	dim classElement as EA.Element
	set classElement = Repository.GetContextObject
	if classElement.ObjectType = otElement then
		'make sure the tags are there
		addTaggedValuesToElement classElement
		'set on element
		setTags classElement.TaggedValues
		'set on attributes
		dim attribute as EA.Attribute
		for each attribute in classElement.Attributes
			setTags attribute.TaggedValues
		next
		'set on connectors for which this element is the source
		dim connector as EA.Connector
		for each connector in classElement.Connectors
			if connector.ClientID = classElement.ElementID then
				setTags connector.TaggedValues
			end if
		next
	end if
end sub

sub setTags (taggedValues)
	taggedValues.Refresh
	dim tag as EA.TaggedValue
	for each tag in taggedValues
		if tag.Value = TBDValue then
			if tag.Name = TimeSlicingName then
				tag.Value = TimeSlicingValue
				tag.Update
			elseif tag.Name = VersioningName then
				tag.Value = VersioningValue
				tag.Update
			end if
		end if
	next
end sub


function addTaggedValuesToElement(element)
	'log progress
	'Repository.WriteOutput "FixLDM","Processing: " & element.Name,0
	'add tagged values to this element
	addTaggedValues element
	'add tagged values to all attributes of this element
	dim attribute as EA.Attribute
	for each attribute in element.Attributes
		addTaggedValues attribute
	next
	'add tagged values to all associations or aggregations of this element
	dim connector as EA.Connector
	for each connector in element.Connectors
		if connector.Type = "Association" or connector.Type = "Aggregation" then
			addTaggedValues connector
		end if
	next
	'process all subElements
	dim subElement as EA.Element
	for each subElement in element.Elements
		addTaggedValues subElement
	next
end function

function addTaggedValues (item)
	dim tv
	dim TSExist
	dim VExist
	TSExist = false
	VExist = false
	'first check if it exists
	for each tv in item.TaggedValues
		if tv.Name = "Versioned" then
			VExist = true
		elseif tv.Name = "Timesliced" then
			TSExist = true
		end if
	next
	'if not create the tagged values
	if not VExist then
		set tv = item.TaggedValues.AddNew("Versioned","")
		tv.Value = "TBD"
		tv.Update
		item.Update
	end if
	if not TSExist then
		set tv = item.TaggedValues.AddNew("Timesliced","")
		tv.Value = "TBD"
		tv.Update
		item.Update
	end if
end function

function switchAttrAliasAndNameOnElement(element)
	dim attribute as EA.Attribute
	dim tempAlias
	dim subElement as EA.Element
	'log progress
	'Repository.WriteOutput "FixLDM","Processing: " & element.Name,0
	'first do all owned attributes
	for each attribute in element.Attributes
		if len(attribute.Alias) > 0 then
			tempAlias = attribute.Alias
			attribute.Alias = attribute.Name
			attribute.Name = tempAlias
			attribute.Update
		end if
	next
	'then do the owned element
	for each subElement in element.Elements
		switchAttrAliasAndNameOnElement subElement
	next
end function

main