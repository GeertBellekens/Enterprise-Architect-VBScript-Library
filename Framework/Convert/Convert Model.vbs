'[path=\Framework\Convert]
'[group=Convert]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
'
' Script Name: Conversion AA model
' Author: Geert Bellekens
' Purpose: Convert the AA model to the new Metamodel profile
' Date: 2023-11-09
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
stereotypeMapping.Add "Class;<null>", "TestProfile::T_TestStereo"
'stereotypeMapping.Add "AA::Applicatie Interface", "Metamodel::Applicatie Interface"
stereotypeMapping.Add "TestProfile::T_TestStereo", "ArchiMate3::ArchiMate_Capability"
'TaggedValues (replace by)
'taggedValuesMapping.Add "Capa Maturity", "Maturity"
'taggedValuesMapping.Add "Capa Level", "Level"

'TaggedValues (copy value to)
'taggedValuesCopy.Add "Maturity", "Maturity SP"
'----------------CONFIGURATION---------------------------------------

sub main
	'reset output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'report progress
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName, 0
	'do the actual work
	convert
	'report progress
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName, 0
end sub

function convert()
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
		convertElement element
	next
	'process diagrams
	convertDiagrams(package)
	'process subPackages
	dim subPackage as EA.Package
	for each subPackage in package.Packages
		convertPackage subPackage
	next
end function

function convertElement(element)
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
	convertItem element, userMessage, mappingKey
	'copy the alias to the ID tagged value
	copyAliasToID element
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
		convertElement subElement
	next
end function

function copyAliasToID(item)
	'check if alias is filled in
	if len(item.Alias) = 0 then
		exit function
	end if
	'make sure we have the correct list of tagged values
	item.TaggedValues.Refresh
	'get the ID tag
	dim IDTag as EA.TaggedValue
	set IDTag = getExistingOrNewTaggedValue(item, "ID")
	'set the value
	IDTag.Value = item.Alias
	IDTag.Update
	'TODO: Delete alias?
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
	convertItem connector, userMessage, mappingKey
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
	convertItem attribute, userMessage, mappingKey
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
			Repository.WriteOutput outPutName, now() & " Converting diagram '" & diagram.Name & "'", 0
			dim styleEx
			'hide connector stereotypes
			styleEx = setValueForKey(diagram.StyleEx, "HideConnStereotype", "1")
			'set view to empty
			styleEx = setValueForKey(styleEx, "MDGView", "")
			'disable fully scoped object names
			styleEx = setValueForKey(styleEx, "NoFullScope", "1")

			dim extendedStyle
			'hide attribute stereotypes
			extendedStyle = setValueForKey(diagram.ExtendedStyle, "HideStereo", "1")
			'hide element stereotypes
			extendedStyle = setValueForKey(extendedStyle, "HideEStereo", "1")
			diagram.StyleEx = styleEx
			diagram.ExtendedStyle = extendedStyle
			diagram.MetaType = stereotypeMapping(mappingKey)
			'diagram.MetaType = "Elia Modelling::CBIM View" 'temporary for CBIM conversion
			diagram.Update
		end if
	next
end function

function convertItem(item,  userMessage, mappingKey)
	if stereotypeMapping.Exists(mappingKey) then
		dim tagsDictionary
		set tagsDictionary = getTagsDictionary(item)
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
		'convert the tagged values
		convertTaggedValues item,  tagsDictionary
	end if
	if mappingKey = "AA::Integratielijn" then
		'report progress
		Repository.WriteOutput outPutName, now() & userMessage, 0
		convertIntegratieLijn item
	end if
end function

function convertIntegratieLijn(element)
	'get sources and targets elements
	dim sqlGetData
	sqlGetData = "select o2.Object_ID from t_object o                                             " & vbNewLine & _
				" inner join t_connector c on c.End_Object_ID = o.Object_ID                      " & vbNewLine & _
				"       and c.Stereotype = 'ArchiMate_Flow'                      " & vbNewLine & _
				" inner join t_object o2 on o2.Object_ID = c.Start_Object_ID                     " & vbNewLine & _
				"      and o2.Stereotype in ('Applicatie','Applicatie Interface')   " & vbNewLine & _
				" where o.Object_ID = " & element.ElementID & "                                  "
	dim sources
	set sources = getElementsFromQuery(sqlGetData)
	sqlGetData = "select o2.Object_ID from t_object o                                             " & vbNewLine & _
				" inner join t_connector c on c.Start_Object_ID = o.Object_ID                    " & vbNewLine & _
				"       and c.Stereotype = 'ArchiMate_Flow'                      " & vbNewLine & _
				" inner join t_object o2 on o2.Object_ID = c.End_Object_ID                       " & vbNewLine & _
				"      and o2.Stereotype in ('Applicatie','Applicatie Interface')   " & vbNewLine & _
				" where o.Object_ID = " & element.ElementID & "                                  "
	dim targets
	set targets = getElementsFromQuery(sqlGetData)
	'loop through all sources and targets
	dim source as EA.Element
	for each source in sources
		'create relation to ea target (except to itself)
		dim target as EA.Element
		for each target in targets
			if target.ElementID <> source.ElementID then
				createIntegration source, target, element
			end if
		next
	next
end function

function createIntegration (source, target, integrationElement)
	'get the IntegrationID
	dim integrationID
	if len(integrationElement.Name) > 5 then
		integrationID = mid(integrationElement.Name,4,3)
	else
		integrationID = integrationElement.Name & "_ERROR"
	end if
	'first check if it doesn't exist yet
	dim sqlGetData
	sqlGetData = "select c.Connector_ID from t_connector c                         " & vbNewLine & _
				" inner join t_connectortag tv on tv.ElementID = c.Connector_ID   " & vbNewLine & _
				"        and tv.Property = 'ID'                " & vbNewLine & _
				" where c.Stereotype like 'Integratie%'                           " & vbNewLine & _
				" and c.Start_Object_ID = " & source.ElementID & "                " & vbNewLine & _
				" and c.End_Object_ID = " & target.ElementID & "                  " & vbNewLine & _
				" and tv.VALUE = '" & integrationID & "'                          "
	dim existingIntegrations
	set existingIntegrations = getConnectorsFromQuery(sqlGetData)
	if existingIntegrations.Count > 0 then
		exit function
	end if
	Repository.WriteOutput outPutName, now() & " Creating Integration between '" & source.Name & "' and '" & target.Name & "'", 0
	'doesn't exist yet, so create it.
	'if source or target is an applicationInterface, then we need to create an Integratie between the owning applications, and an IntegrationLijn between current source and target
	dim createIntegratieLijn
	createIntegratieLijn = false
	dim sourceApplication
	if source.Stereotype = "Applicatie Interface" then
		createIntegratieLijn = true
		set sourceApplication = getApplicationForInterface(source)
	else
		set sourceApplication = source
	end if
	
	dim targetApplication
	if target.Stereotype = "Applicatie Interface" then
		createIntegratieLijn = true
		set targetApplication = getApplicationForInterface(target)
	else
		set targetApplication = target
	end if
	'check if we need to create an Integratie, or an IntegratieLijn
	dim integrationStereotype
	if createIntegratieLijn then
		'create Integration between source and target Application
		if not sourceApplication is nothing and not targetApplication is nothing then
			createIntegration sourceApplication, targetApplication, integrationElement
		end if
		integrationStereotype = "Metamodel::IntegratieLijn"
	else
		integrationStereotype = "Metamodel::Integratie"
	end if
	'actually create the new connector
	dim integrationConnector as EA.Connector
	set integrationConnector = source.Connectors.AddNew(mid(integrationElement.Name,7),integrationStereotype)
	integrationConnector.SupplierID = target.ElementID
	integrationConnector.Update
	'make sure we have the correct set of tagged values
	integrationConnector.TaggedValues.Refresh
	'copy the other tagged values including setting the ID
	dim tagsDictionary
	set tagsDictionary = getTagsDictionary(integrationElement)
	dim tag as EA.ConnectorTag
	for each tag in integrationConnector.TaggedValues
		if tagsDictionary.Exists(tag.Name) then
			if len(tagsDictionary(tag.Name)) > 0 then
				tag.Value = tagsDictionary(tag.Name)
				tag.Update
			end if
		elseif tag.Name = "ID" then
			tag.Value = integrationID
			tag.Update
		end if
	next
end function

function getApplicationForInterface(interface)
	dim application as EA.Element
	set application = nothing
	if interface.ParentID > 0 then
		set application = Repository.GetElementByID(interface.ParentID)
		if application.Stereotype <> "Applicatie" then
			application = nothing
		end if
	end if
	'return
	set getApplicationForInterface = application
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

function convertTaggedValues(element, tagsDictionary)
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
	next
end function

main
