'[path=\Projects\Project EL\Conversion]
'[group=Conversion]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
'
' Script Name: Convert Archimate Stereotypes
' Author: Geert Bellekens
' Purpose: Convert ArchiMate Capabilities from standard to Elia MDG
' Date: 2023-01-20
'
const outPutName = "Convert Model"


dim stereotypeMapping
set stereotypeMapping = CreateObject("Scripting.Dictionary")

dim taggedValuesMapping
set taggedValuesMapping = CreateObject("Scripting.Dictionary")
dim taggedValuesCopy
set taggedValuesCopy = CreateObject("Scripting.Dictionary")
'----------------CONFIGURATION---------------------------------------
'Stereotypes
'Elements
stereotypeMapping.Add "ArchiMate3::ArchiMate_BusinessProcess", "Elia Modelling::ArchiMate_BusinessProcess"
stereotypeMapping.Add "ArchiMate3::ArchiMate_BusinessActor", "Elia Modelling::ArchiMate_BusinessActor"
stereotypeMapping.Add "ArchiMate3::ArchiMate_BusinessRole", "Elia Modelling::ArchiMate3::ArchiMate_BusinessRole"
stereotypeMapping.Add "ArchiMate3::ArchiMate_BusinessEvent", "Elia Modelling::ArchiMate3::ArchiMate_BusinessEvent"
stereotypeMapping.Add "BPMN2.0::BusinessProcess", "Elia Modelling::ArchiMate_BusinessProcess"
stereotypeMapping.Add "ArchiMate3::ArchiMate_Capability", "Elia Modelling::ArchiMate_Capability"
stereotypeMapping.Add "ArchiMate3::ArchiMate_ArchiMate_Product", "Elia Modelling::ArchiMate_ArchiMate_Product"
'Diagrams
stereotypeMapping.Add "BPMN2.0::Business Process", "BPMN2.0::Business Process;Elia Modelling::Business Process Diagram" 
'Connectors


'stereotypeMapping.Add "Elia Modelling::ArchiMate_BusinessFunction", "Elia Modelling::ArchiMate_Capability"
'stereotypeMapping.Add "ArchiMate3::ArchiMate_Capability", "Elia Modelling::ArchiMate_Capability"
'stereotypeMapping.Add "ArchiMate3::ArchiMate_BusinessFunction", "Elia Modelling::ArchiMate_Capability"
'stereotypeMapping.Add "ArchiMate3::ArchiMate_BusinessObject", "Elia Modelling::Elia_BIM_Class"
'stereotypeMapping.Add "Class;<null>", "Elia Modelling::Elia_BIM_Class"
'stereotypeMapping.Add "Attribute;<null>", "Elia Modelling::Elia_BIM_Attribute"
'stereotypeMapping.Add "Association;<null>", "Elia Modelling::Elia_BIM_Association"
'stereotypeMapping.Add "Aggregation;<null>", "Association;Elia Modelling::Elia_BIM_Association"
'stereotypeMapping.Add "Elia Modelling::Elia_BIM_Association", "Association;Elia Modelling::Elia_BIM_Association" 'temporary
'stereotypeMapping.Add "Logical;<null>", "Elia Modelling::BIM View"

'stereotypeMapping.Add "Elia Modelling::BIM View", "Elia Modelling::CBIM View" 'temporary
'stereotypeMapping.Add "Elia Modelling::Elia_BIM_Class", "Elia Modelling::Elia_CBIM_Class"
'stereotypeMapping.Add "Elia Modelling::Elia_BIM_Association", "Elia Modelling::Elia_CBIM_Association"
'TaggedValues (replace by)
'taggedValuesMapping.Add "Capa Maturity", "Maturity"
'taggedValuesMapping.Add "Capa Level", "Level"

'TaggedValues (copy value to)
'taggedValuesCopy.Add "Maturity", "Maturity 50Hz"
'----------------CONFIGURATION---------------------------------------

sub main
	'reset output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'report progress
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName, 0
	'do the actual work
	convertCapabilities()
	'report progress
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName, 0
end sub

function convertCapabilities()
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage
	convertPackage package

end function

function convertPackage(package)
	'report progress
	Repository.WriteOutput outPutName, now() & " Processing '" & package.Name & "'", 0
	'convert elements
	dim element as EA.Element
	for each element in package.Elements
		convertElement element, -1
	next
	'process diagrams
	convertDiagrams(package)
	'process subPackages
	dim subPackage as EA.Package
	for each subPackage in package.Packages
		convertPackage subPackage
	next
end function

function convertElement(element, level)
	dim mappingKey
	if len(element.FQStereotype) > 0 then
		mappingKey = element.FQStereotype
	else
		mappingKey = element.Type & ";<null>"
	end if
	dim userMessage
	userMessage = " Converting element '" & element.Name & "'"
	'Repository.WriteOutput outPutName, now() & " debug before convertItem: " & element.Name, 0
	'convert the element
	convertItem element, userMessage, mappingKey, level
	'Repository.WriteOutput outPutName, now() & " debug after convertItem: " & element.Name, 0
	'process attributes
	dim attribute as EA.Attribute
	for each attribute in element.Attributes
		convertAttribute element, attribute
	next
	'process (outgoing) relations
	dim connector as EA.Connector
	for each connector in element.Connectors
		if connector.ClientID = element.ElementID then
			convertConnector element, connector
		end if
	next
	'process diagrams
	convertDiagrams(element)
	'process subElements
	dim subElement as EA.Element
	for each subElement in element.Elements
		convertElement subElement, level + 1
	next
end function

function convertConnector(element, connector)
	dim mappingKey
	if len(connector.FQStereotype) > 0 then
		mappingKey = connector.FQStereotype
	else
		mappingKey = connector.Type & ";<null>"
	end if
	dim userMessage
	userMessage = " Converting connector '" & element.Name & "." & connector.Name & "'"
	'convert the connector
	convertItem connector, userMessage, mappingKey, ""
end function



function convertAttribute(element, attribute)
	dim mappingKey
	if len(attribute.FQStereotype) > 0 then
		mappingKey = attribute.FQStereotype
	else
		mappingKey = "Attribute;<null>"
	end if
	dim userMessage
	userMessage = " Converting attribute '" & element.Name & "." & attribute.Name & "'"
	'convert the attribute
	convertItem attribute, userMessage, mappingKey, ""
end function

function convertDiagrams(diagramOwner)
	dim diagram as EA.Diagram
	for each diagram in diagramOwner.Diagrams
		'get mapping key
		dim mappingKey
		if len(diagram.MetaType) > 0 then
			mappingKey = diagram.MetaType
		else
			mappingKey = diagram.Type & ";<null>"
		end if
		'convert
		if stereotypeMapping.Exists(mappingKey) then
		'if true then 'temporary for CBIM conversion
			Repository.WriteOutput outPutName, now() & " Converting diagram '" & diagram.Name & "'", 0
			dim styleEx
			'hide connector stereotypes
			styleEx = setValueForKey(diagram.StyleEx, "HideConnStereotype", "1")
			'disable fully scoped object names
			styleEx = setValueForKey(styleEx, "NoFullScope", "1")

			dim extendedStyle
			'hide attribute stereotypes
			extendedStyle = setValueForKey(diagram.ExtendedStyle, "HideStereo", "1")
			'hide element stereotypes
			extendedStyle = setValueForKey(extendedStyle, "HideEStereo", "1")
			
			diagram.StyleEx = styleEx
			diagram.ExtendedStyle = extendedStyle
			dim mappingValue
			mappingValue = stereotypeMapping(mappingKey)
			dim metaType
			dim mdgView
			mdgView = ""
			if instr(mappingValue, ";") then
				dim mappingParts
				mappingParts = split(mappingValue, ";")
				metaType = mappingParts(0)
				mdgView = mappingParts(1)
			else
				metaType = mappingValue
			end if
			diagram.MetaType = metaType
			diagram.StyleEx = setValueForKey(diagram.StyleEx, "MDGView", mdgView )
			'diagram.MetaType = "Elia Modelling::CBIM View" 'temporary for CBIM conversion
			diagram.Update
		end if
	next
end function

function convertItem(item,  userMessage, mappingKey, level)
	if stereotypeMapping.Exists(mappingKey) then
		dim tagsDictionary
		set tagsDictionary = getTagsDictionary(item)
		dim setLevel3
		setLevel3 = false
		if instr(mappingKey, "ArchiMate_BusinessFunction") > 0 then
			setLevel3 = true
		end if
		'report progress
		Repository.WriteOutput outPutName, now() & userMessage, 0
		dim dirty
		dirty = false
		'check if we have a type to convert to
		dim mappingTarget
		dim targetSteretoype
		mappingTarget = stereotypeMapping(mappingKey)
		if instr(mappingTarget, ";") > 0 then
			'fix the aggregation direction if needed
			'aggregation with Direction "Source -> Destination" don't have arrows. 
			'If we change the type to Association, they do, so we change Direction to Unspecified in order to remove the arrow
			if item.ObjectType = otConnector then
				if item.Type = "Aggregation" _ 
				  and item.Direction = "Source -> Destination" then
					item.Direction = "Unspecified"
					dirty = true
				end if
			end if
			'get type and stereotype
			dim mappingTargetParts
			mappingTargetParts = split(mappingTarget, ";")
			if item.Type <> mappingTargetParts(0) then
				item.Type = mappingTargetParts(0) 'type
				dirty = true
			end if
			targetSteretoype= mappingTargetParts(1)
		else
			targetSteretoype = mappingTarget
		end if
		if item.FQStereotype <> targetSteretoype then
			item.StereotypeEx = targetSteretoype 'stereotype
			dirty = true
		end if
		if dirty then
			item.Update
		end if
		'Repository.WriteOutput outPutName, now() & " debug after item update: " & item.Name, 0
		if setLevel3 then 
			'convert tagged values
			convertTaggedValues item, "3", tagsDictionary
		else
			'convert tagged values
			convertTaggedValues item, level, tagsDictionary
		end if	
	end if
end function

function getTagsDictionary(item)
	dim tagsDictionary
	set tagsDictionary = createObject("Scripting.Dictionary")
	dim tag as EA.TaggedValue
	for each tag in item.TaggedValues
		if not tagsDictionary.Exists(tag.Name) then
			tagsDictionary.Add tag.Name, tag.Value
			'Repository.WriteOutput outPutName, now() & " adding tag to dictionary " & tag.Name & " value: " & tag.Value , 0
		end if
	next
	'return
	set getTagsDictionary = tagsDictionary
end function

function convertTaggedValues(element, level, tagsDictionary)
	'Repository.WriteOutput outPutName, now() & " debug before update existing tags: " & element.Name, 0
	dim tag as EA.TaggedValue
	element.TaggedValues.Refresh
	'Repository.WriteOutput outPutName, now() & " debug after taggedvalues refresh: " & element.Name, 0
	'first loop tagged values to copy the pre-existing tag values
	for each tag in element.TaggedValues
		'Repository.WriteOutput outPutName, now() & " debug processing tag: " & tag.Name & " tagsDictionary.Count: " & tagsDictionary.Count , 0
		'check the pre-existing tags and copy their value
		if tagsDictionary.Exists(tag.Name) then
			'Repository.WriteOutput outPutName, now() & " debug updating tag: " & tag.Name, 0
			tag.Value = tagsDictionary(tag.Name)
			tag.Update
		end if
	next
	'refresh to make sure we have the correct tags and values
	element.TaggedValues.Refresh
	'Repository.WriteOutput outPutName, now() & " debug after update existing tags: " & element.Name, 0
	'then convert if needed
	for each tag in element.TaggedValues
		dim convertedTag
		'convert tags
		if taggedValuesMapping.Exists(tag.Name) then
			set convertedTag = getExistingOrNewTaggedValue(element, taggedValuesMapping(tag.Name))
			convertedTag.Value = tag.Value
			convertedTag.Update
			deleteTag element, tag.Name
		end if
		'copy tag values
		if taggedValuesCopy.Exists(tag.Name) then
			'Repository.WriteOutput outPutName, now() & " debug tag to copy: " & tag.Name, 0
			set convertedTag = getExistingOrNewTaggedValue(element, taggedValuesCopy(tag.Name))
			convertedTag.Value = tag.Value
			convertedTag.Update
		end if
		
		'check level tag
		if lcase(tag.Name) = "level" then
			dim correctLevel
			correctLevel = "L" & level
			if not lcase(tag.Value) = lcase(correctLevel) then
				if lcase(tag.Value) = "tbd" or len(tag.Value) = 0 then
					tag.Value = correctLevel
					tag.Update
				else
					'report issue
					Repository.WriteOutput outPutName, now() & " ERROR: Level tag on '" & element.Name & "' with GUID '" & element.ElementGUID & "' is '" & tag.Value & "' and should be '" & correctLevel  & "'", 0
				end if
			end if
		end if
	next
	
end function

main