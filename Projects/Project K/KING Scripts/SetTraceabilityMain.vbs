'[path=\Projects\Project K\KING Scripts]
'[group=KING Scripts]
!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Set traceability
' Author: Geert Bellekens
' Purpose: Adds traceability from the copy package selected in the project browser to the original package selected by the user.
' Date: 2016-02-08
'

dim outputTabName
outputTabName = "Set Traceability"
dim isTransform
isTransform = false

sub SetTraceability(withTransformation)

	Repository.CreateOutputTab outputTabName
	Repository.ClearOutput outputTabName
	Repository.EnsureOutputVisible outputTabName
	Repository.EnableUIUpdates = false
	
	isTransform = withTransformation
	dim copyPackage as EA.Package
	dim originalPackage as EA.Package
	'the copy packag is the one selected in the project browser
	set copyPackage = Repository.GetTreeSelectedPackage
	msgbox "Please select the original package"
	'the original package is the one selected by the user
	dim originalPackageID
	dim originalPackageElementID
	originalPackageElementID = Repository.InvokeConstructPicker("IncludedTypes=Package;Selection=" & copyPackage.PackageGUID)
	if originalPackageElementID > 0 then
		dim originalPackageElement as EA.Element
		set originalPackageElement = Repository.GetElementByID(originalPackageElementID)
		set originalPackage = Repository.GetPackageByGuid(originalPackageElement.ElementGUID)
		'tell the user we are starting
		Repository.WriteOutput outputTabName, now() & " Starting adding traces",0
		'start the trace		
		tracePackageElements originalPackage, copyPackage
		Repository.WriteOutput outputTabName, now() & " Finished adding traces",0
	end if
	Repository.EnableUIUpdates = true
	Repository.RefreshModelView copyPackage.PackageID
end sub

function tracePackageElements(originalPackage, copyPackage)
	dim originalElement as EA.Element
	dim copyElement as EA.Element
	'find corresponding element
	for each originalElement in originalPackage.Elements
		'only process elements that have a name
		if len(originalElement.Name) > 0 then
			'Repository.WriteOutput outputTabName, now() & " Processing " & originalElement.Type & ": " & originalElement.Name ,0
			dim matchFound
			matchFound = false
			for each copyElement in copyPackage.Elements
				if copyElement.Name = originalElement.Name then
					'found a match
					traceElements originalElement,copyElement
					if isTransform then
						transformElement originalElement,copyElement
					end if
					matchFound = true
					exit for
				end if
			next
			if matchFound then
				Repository.WriteOutput outputTabName, now() & " Match found for " & originalElement.Type & ": " & originalElement.Name ,0
			else
				Repository.WriteOutput outputTabName, now() & " Match NOT found for " & originalElement.Type & ": " & originalElement.Name ,0
			end if
		end if
	next
	'process subpackages
	dim originalSubPackage
	dim copySubpackage
	for each originalSubPackage in originalPackage.Packages
		for each copySubpackage in copyPackage.Packages
			if originalSubPackage.Name = copySubpackage.Name then
				'found a match
				tracePackageElements originalSubPackage, copySubpackage
				exit for
			end if
		next
	next
end function

function traceElements (originalElement,copyElement)
'	dim originalElement as EA.Element
'	dim copyElement as EA.Element
	'add trace relation
	dim trace as EA.Connector
	dim connector as EA.Connector
	
	'delete the original traces if they exist
	dim keepDeleting
	keepDeleting = true
	do
		 keepDeleting = deleteOriginalTraces(originalElement,copyElement)
	loop while keepDeleting
	
	dim traceExists
	traceExists = false
	'check if the connector exists already
	for each connector in copyElement.Connectors
		if connector.SupplierID = originalElement.ElementID _
		AND connector.Type = "Abstraction" _
		AND connector.Stereotype = "trace" then
			traceExists = true
			exit for
		end if
	next
	
	'if it doesn't exist yet we create a new one
	if traceExists = false then
		set trace = copyElement.Connectors.AddNew("","Abstraction")
		trace.SupplierID = originalElement.ElementID
		trace.Stereotype = "trace"
		trace.Update
	end if
	'trace attributes
	traceAttributes originalElement,copyElement
	'trace associations
	traceAssociations originalElement,copyElement
end function

function deleteOriginalTraces(originalElement,copyElement)
	dim i
	dim copyConnector as EA.Connector
	deleteOriginalTraces = false
	'make sure the connectors are refreshed
	copyElement.Connectors.Refresh
	originalElement.Connectors.Refresh
	'remove all the traces to domain model classes
	for each copyConnector in copyElement.Connectors
		if copyConnector.Type = "Abstraction" AND copyConnector.Stereotype = "trace" then
			'check if the original element has the same trace
			dim originalConnector as EA.Connector
			for each originalConnector in originalElement.Connectors
				if copyConnector.Type = "Abstraction" AND _
				copyConnector.Stereotype = "trace" AND _
				copyConnector.SupplierID = originalConnector.SupplierID then
					deleteConnector copyElement, copyConnector
					'refresh again and exit function
					copyElement.Connectors.Refresh
					originalElement.Connectors.Refresh
					'found on, try again
					deleteOriginalTraces = true
					exit function
				end if
			next
		end if
	next
end function

function traceAttributes(originalElement,copyElement)
	dim originalAttribute as EA.Attribute
	dim copyAttribute as EA.Attribute
	for each originalAttribute in originalElement.Attributes
		for each copyAttribute in copyElement.Attributes
			if copyAttribute.Name = originalAttribute.Name then
				if isTransform then
					transformAttribute originalAttribute,originalElement, copyAttribute 
				end if
				'found match, add trace tag
				dim traceTag as EA.AttributeTag
				set traceTag = getExistingOrNewTaggedValue(copyAttribute,"SourceAttribute")
				traceTag.Value = originalAttribute.AttributeGUID
				traceTag.Update
				exit for
			end if
		next
	next
end function

function traceAssociations (originalElement,copyElement)
	'make sure the connectors are refreshed
	copyElement.Connectors.Refresh
	originalElement.Connectors.Refresh
	
	dim originalConnector as EA.Connector
	dim copyConnector as EA.Connector
	for each originalConnector in originalElement.Connectors
		'we process only associations that start from the original element
		if (originalConnector.Type = "Association" or originalConnector.Type = "Aggregation") _
			AND originalConnector.ClientID =  originalElement.ElementID then
			for each copyConnector in copyElement.Connectors
				if copyConnector.Type = originalConnector.Type _
					AND copyConnector.Name = originalConnector.Name _
					AND copyConnector.ClientEnd.Cardinality = originalConnector.ClientEnd.Cardinality _
					AND copyConnector.ClientEnd.Role = originalConnector.ClientEnd.Role _
					AND copyConnector.ClientEnd.Aggregation = originalConnector.ClientEnd.Aggregation _
					AND copyConnector.SupplierEnd.Cardinality = originalConnector.SupplierEnd.Cardinality _
					AND copyConnector.SupplierEnd.Role = originalConnector.SupplierEnd.Role _
					AND copyConnector.SupplierEnd.Aggregation = originalConnector.SupplierEnd.Aggregation then
					'connector properties match, now check the other ends
					dim originalOtherEnd as EA.Element
					dim copyOtherEnd as EA.Element
					set originalOtherEnd = Repository.GetElementByID(originalConnector.SupplierID)
					set copyOtherEnd = Repository.GetElementByID(copyConnector.SupplierID)
					if copyOtherEnd.Name = originalOtherEnd.Name _
						or isTransformedFrom(originalOtherEnd, copyOtherEnd) then
						'associations with stereotype 'tekentechnisch' need to be removed
						if isTransform AND copyConnector.Stereotype = "Tekentechnisch" then
							deleteConnector copyElement, copyConnector
							copyElement.Connectors.Refresh
						else
							'transform
							if isTransform then
								transformAssociations originalConnector,copyConnector
							end if
							'found a match, add trace tag
							dim traceTag as EA.ConnectorTag
							set traceTag = getExistingOrNewTaggedValue(copyConnector,"SourceAssociation")
							traceTag.Value = originalConnector.ConnectorGUID
							traceTag.Update
						end if
						exit for
					end if
				end if
			next
		end if
	next
end function

function deleteConnector (owner, connector)
	dim currentConnector as EA.Connector
	dim i
	for i = 0 to owner.Connectors.Count -1
		set currentConnector = owner.Connectors.GetAt(i)
		if currentConnector.ConnectorID = connector.ConnectorID then
			owner.Connectors.DeleteAt i,false
			exit for
		end if
	next
end function

' the copyElement is transformed from the originalElement if it has a trace relationship to the originalElement
function isTransformedFrom(originalElement, copyElement)
	dim connector as EA.Connector
	isTransformedFrom = false
	for each connector in copyElement.Connectors
		if connector.Stereotype = "trace" and connector.SupplierID = originalElement.ElementID then
			isTransformedFrom = true
			exit for
		end if
	next
end function

function getExistingOrNewTaggedValue(owner, tagname)
	dim taggedValue as EA.TaggedValue
	dim returnTag as EA.TaggedValue
	set returnTag = nothing
	'check if a tag with that name alrady exists
	for each taggedValue in owner.TaggedValues
		if taggedValue.Name = tagName then
			set returnTag = taggedValue
			exit for
		end if
	next
	'create new one if not found
	if returnTag is nothing then
		set returnTag = owner.TaggedValues.AddNew(tagname,"")
	end if
	'return
	set getExistingOrNewTaggedValue = returnTag
end function

function transformElement(originalElement,copyElement)
	'set stereotype
	select case originalElement.Stereotype
		case "Objecttype"
			copyElement.StereotypeEx = "Entiteittype"
			copyElement.Name = getCamelCase(originalElement.Name)
'			copyElement.Name = originalElement.Alias
'			copyElement.Alias = getCamelCase(originalElement.Name)
		case "Gegevensgroeptype"
			copyElement.StereotypeEx = "Groep"
			copyElement.Name = getCamelCase(originalElement.Name)
		case "Referentielijst"
			copyElement.StereotypeEx = "Tabel-entiteit"
			copyElement.Name = getCamelCase(originalElement.Name)
'			copyElement.Name = originalElement.Alias
'			copyElement.Alias = getCamelCase(originalElement.Name)
		case "Relatieklasse"
			copyElement.StereotypeEx = "Relatie-entiteit"
			copyElement.Name = getCamelCase(originalElement.Name)
		case "Complex datatype"
			copyElement.StereotypeEx = "MUG Complex datatype"
		case "Union"
			copyElement.StereotypeEx = "MUG Union"
	end select
	'remove notes for all elements
'	copyElement.Notes = "" 'not yet
	copyElement.Notes = vbNewLine & "--" & vbNewLine & originalElement.Notes
	copyElement.Update
	'copy tagged values
	copyTaggedValueValues originalElement, copyElement
end function

function transformAttribute(originalAttribute,originalElement, copyAttribute)
	'set stereotype
	select case originalAttribute.Stereotype
		case "Attribuutsoort"
			copyAttribute.StereotypeEx = "Element"
			copyAttribute.Name = getCamelCase(originalAttribute.Name)

		case "Referentie element"
			copyAttribute.StereotypeEx = "Tabel element"
			'copyAttribute.Name = getCamelCase(originalAttribute.Name)
		case "Union element"
			copyAttribute.StereotypeEx = "MUG Union Element"
		case "Data element"
			copyAttribute.StereotypeEx = "MUG Data element"
	end select
	
	'not for enum values. Enum Values have a parent of type enumeration and have not "IsLiteral=0" in the styleEx field
	if NOT(originalElement.Type = "Enumeration" _
		or instr(originalAttribute.StyleEx, "IsLiteral=1;") > 0 _
		or originalElement.Stereotype = "Enumeration") then
		'name => camelCase
		copyAttribute.Name = getCamelCase(originalAttribute.Name)
	end if
	'notes
	copyAttribute.Notes = vbNewLine & "--" & vbNewLine & originalAttribute.Notes
	copyAttribute.Update
	'cop tagged values
	copyTaggedValueValues originalAttribute, copyAttribute
	'kerngegeven stereotype
'	if isKernGegeven(originalAttribute) then
'		copyAttribute.StereotypeEx = copyAttribute.StereotypeEx  & ",Kerngegeven"
'	end if
end function

function isKernGegeven(attribute)
	dim taggedValue as EA.TaggedValue
	isKernGegeven = false
	for each taggedValue in attribute.TaggedValues
		if taggedValue.Name = "Indicatie kerngegeven" then
			if taggedValue.Value = "Ja" then
				isKernGegeven = true
			end if
			exit for
		end if
	next
end function

function transformAssociations (originalConnector,copyConnector)
	'set stereotype
	select case originalConnector.Stereotype
		case "Relatiesoort"
			copyConnector.StereotypeEx = "Relatie"
		case "Relatieklasse"
			copyConnector.StereotypeEx = "Relatie-entiteit"
	end select
	'name => camelCase
	copyConnector.Name = getCamelCase(originalConnector.Name)
	'notes
	copyConnector.Notes = vbNewLine & "--" & vbNewLine & originalConnector.Notes
	copyConnector.Update
	'cop tagged values
	copyTaggedValueValues originalConnector,copyConnector
end function

function getCamelCase(nameToConvert)
	  Dim arr, i
	  arr = Split(nameToConvert, " ")
	  For i = LBound(arr) To UBound(arr)
		if i = 0 then
			arr(i) = LCase(arr(i))
		else
			arr(i) = UCase(Left(arr(i), 1)) & LCase(Mid(arr(i), 2))
		end if
	  Next
  getCamelCase = Join(arr, "")
end function

function copyTaggedValueValues(originalElement, copyElement)
	dim copyTV as EA.TaggedValue
	dim originalTV as EA.TaggedValue
	for each copyTV in copyElement.TaggedValues
		for each originalTV in originalElement.TaggedValues
			if copyTV.Name = originalTV.Name then
				copyTV.Value = originalTV.Value
				copyTV.Notes = originalTV.Notes
				copyTV.Update
				exit for
			end if
		next
	next
end function

function removeTaggedValuesExcept(item, tvsToKeep)
	dim taggedValue as EA.TaggedValue
	dim i
	for i = item.TaggedValues.Count -1 to 0 step -1
		set taggedValue = item.TaggedValues(i)
		if not tvsToKeep.contains(taggedValue.Name) then
			item.TaggedValues.DeleteAt i, false
		end if	
	next
end function