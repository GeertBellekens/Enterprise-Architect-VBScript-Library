'[path=\Projects\Project BRO\BRO Scripts]
'[group=BRO Scripts]

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
dim isTranslate
isTranslate = true 'always translate
dim translatedProfile
translatedProfile = "NEN3610 Grouping (EN)"
dim originalPackage as EA.Package

'dictionary with stereotype translations
dim stereotypeTranslations


sub SetTraceability(withTransformation)

	Repository.CreateOutputTab outputTabName
	Repository.ClearOutput outputTabName
	Repository.EnsureOutputVisible outputTabName
	Repository.EnableUIUpdates = false
	
	'Create translations dictionary
	createTranslationsDictionary
	
	isTransform = withTransformation
	dim copyPackage as EA.Package
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
		'create tagged value types if needed
		createTaggedValueTypes
'		'ask the user if we need to translate to English
'		dim translateResponse
'		translateResponse = msgbox("Translate to English?", vbYesNo+vbQuestion, "Translate?")
'		if translateResponse = vbYes then
'			isTranslate = true
'		end if
		'tell the user we are starting
		Repository.WriteOutput outputTabName, now() & " Starting adding traces",0
		'start the trace		
		tracePackageElements originalPackage, copyPackage
		'tell the users we are finished
		Repository.WriteOutput outputTabName, now() & " Finished adding traces",0
	end if
	Repository.EnableUIUpdates = true
	'Repository.ReloadPackage copyPackage.PackageID 'only available from v13.0.1305
	Repository.RefreshModelView copyPackage.PackageID
end sub

function createTaggedValueTypes()
	dim tagDetail
	if not taggedValueTypeExists("SourceAttribute") then
		tagDetail = "Type=RefGUID; " & vbNewLine & _
					"Values=Attribute;" & vbNewLine & _
					"AppliesTo=Attribute;" & vbNewLine
		addTaggedValueType "SourceAttribute", "is derived from this Attribute", tagDetail
	end if
	if not taggedValueTypeExists("SourceAssociation") then
		tagDetail = "Type=String; " & vbNewLine & _
					"AppliesTo=Association,Aggregation;" & vbNewLine
		addTaggedValueType "SourceAssociation", "is derived from this Association", tagDetail
	end if
end function

function addTaggedValueType(tagName, tagDescription, tagDetail)
	dim taggedValueType
	set taggedValueType = Repository.PropertyTypes.AddNew(tagName,"")
	taggedValueType.Description = tagDescription
	taggedValueType.Detail = tagDetail
	taggedValueType.Update
end function

function taggedValueTypeExists(tagName)
	'inital false
	taggedValueTypeExists = false
	'refresh tagged value types
	Repository.PropertyTypes.Refresh
	dim taggedValueType
	'find tagged value type with the given name
	for each taggedValueType in Repository.PropertyTypes
		dim tagTypeName
		tagTypeName = taggedValueType.Tag
		'ignore case or EA will complain
		if lcase(tagTypeName) = lcase(tagName) then
			taggedValueTypeExists = true
			exit for
		end if
	next
end function


function tracePackageElements(originalPackage, copyPackage)
	dim originalElement as EA.Element
	dim copyElement as EA.Element
	'process package transformation
	if isTransform and originalPackage.ParentID > 0 and copyPackage.ParentID > 0 then
		transformElement originalPackage.Element, copyPackage.Element
	end if
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
				Repository.WriteOutput outputTabName, now() & " Match found for " & originalElement.Type & ": " & originalElement.Name ,copyElement.ElementID
			else
				Repository.WriteOutput outputTabName, now() & " Match NOT found for " & originalElement.Type & ": " & originalElement.Name ,originalElement.ElementID
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
	if isTranslate then
		translateStereotype copyElement
		'switch name and alias (if alias is filled in)
		if copyElement.Type <> "Package" then
			'not for packages
			swithNameAndAlias copyElement
		else
			'copy certain properties for packages
			copyProperties originalElement, copyElement
		end if
	end if
	'delete al enum literals
	deleteAllEnumLiterals copyElement
	'add -- to notes
	addNotesDisclaimer originalElement, copyElement
	'save element
	copyElement.Update
end function

function transformAttribute(originalAttribute,originalElement, copyAttribute)
	'set stereotype
	if isTranslate then
		translateStereotype copyAttribute
		'switch name and alias (if alias is filled in)
		swithNameAndAlias copyAttribute
	end if
	'add -- to notes
	addNotesDisclaimer originalAttribute, copyAttribute
	'save attribute
	copyAttribute.Update
end function

function transformAssociations (originalConnector,copyConnector)
	'set stereotype
	if isTranslate then
		translateStereotype copyConnector
		'translate also roles stereotypes
		translateStereotype copyConnector.ClientEnd
		translateStereotype copyConnector.SupplierEnd
		'switch name and alias (if alias is filled in)
		swithNameAndAlias copyConnector
		'also for roles
		swithNameAndAlias copyConnector.ClientEnd
		swithNameAndAlias copyConnector.SupplierEnd
	end if
	'add -- to notes
	addNotesDisclaimer originalConnector, copyConnector
	'save attribute
	copyConnector.Update
end function

function deleteAllEnumLiterals(copyElement)
	if lcase(copyElement.Stereotype) = "codelist" then	
		dim i
		for i = copyElement.Attributes.Count -1  to 0 step -1
			copyElement.Attributes.DeleteAt i, false
		next
	end if
end function

function copyProperties(originalElement,copyElement)
	select case copyElement.Stereotype
		case "Basemodel"
			copyElement.Version = "1.0.0"
			copyElement.Phase = "Draft"
			copyElement.Alias = originalElement.Alias
			copyElement.Name = originalElement.Name
			'set tagged values
			copyElement.TaggedValues.Refresh
			copyTaggedValue originalElement, copyElement, "Release", "Release"
			setTaggedValue copyElement,  "Supplier-project", "Conceptual Model"
			setTaggedValue copyElement, "Supplier-name", originalPackage.Name
			copyTaggedValue originalElement, copyElement, "Release", "Supplier-release"
			setTaggedValue copyElement, "Imvertor", "model"
		case "Domain"
			copyElement.Version = "1.0.0"
			copyElement.Phase = "Draft"
			'set tagged values
			copyElement.TaggedValues.Refresh
			copyTaggedValue originalElement, copyElement, "Release", "Release"
	end select
end function

function setTaggedValue(element, tagName, tagValue)
	dim tv as EA.TaggedValue
	for each tv in element.TaggedValues
		if lcase(tv.Name) = lcase(tagName) then
			tv.Value = tagValue
			tv.Update
			exit for
		end if
	next
end function

function copyTaggedValue(originalElement, copyElement, originalTagName, copyTagName)
	dim originalTV as EA.TaggedValue
	dim copyTV as EA.TaggedValue
	'first find the correct original tag
	for each originalTV in originalElement.TaggedValues
		if lcase(originalTV.Name) = lcase(originalTagName) then
			'then look for the copy tag
			for each copyTV in copyElement.TaggedValues
				if lcase(copyTV.Name) = lcase(copyTagName) then
					copyTV.Value =  originalTV.Value
					copyTV.Update
					exit for
				end if
			next
			exit for
		end if
	next
end function

function addNotesDisclaimer(originalItem, copyItem)
		'add -- to notes
	copyItem.Notes = "---" & vbNewLine & _
		"Documentation extracted from conceptual model. This may be outdated." & vbNewLine & vbNewLine & originalItem.Notes
end function

function swithNameAndAlias(copyItem)
	dim tempAlias
	'only if Alias is filled in
	if len(copyItem.Alias) > 0 then
		tempAlias = copyItem.Alias
		copyItem.Alias = copyItem.Name
		copyItem.Name = tempAlias
	end if
end function

function deleteAllTaggedValues(copyItem)
	dim i
	for i = copyItem.TaggedValues.Count -1  to 0 step -1
		copyItem.TaggedValues.DeleteAt i, false
	next
	copyItem.TaggedValues.Refresh
end function

function translateStereotype(copyItem)
	'check if stereotype is in list of translated stereotypes
	'if not then report error?
	if stereotypeTranslations.Exists(lcase(copyItem.Stereotype)) then
		'inform user
		if copyItem.ObjectType = otConnectorEnd then
			'connector roles do not support fully qualified stereotypes
			Repository.WriteOutput outputTabName, now() & " Translating stereotype «" & copyItem.Stereotype & "» to «" & stereotypeTranslations.Item(lcase(copyItem.Stereotype)) & "» on item " & copyItem.Role, 0
		else
			Repository.WriteOutput outputTabName, now() & " Translating stereotype «" & copyItem.Stereotype & "» to «" & stereotypeTranslations.Item(lcase(copyItem.Stereotype)) & "» on item " & copyItem.Name, 0
		end if
		'delete all tagged values
		deleteAllTaggedValues copyItem
		'translate stereotype
		copyItem.StereotypeEx = translatedProfile & "::" & stereotypeTranslations.Item(lcase(copyItem.Stereotype))
		'save stereotype
		copyItem.Update
	else
		if len(copyItem.Stereotype) > 0 then
			if copyItem.ObjectType = otConnectorEnd then
				Repository.WriteOutput outputTabName, now() & " ERROR: No translation found for stereotype «" & copyItem.Stereotype & "» on item " & copyItem.Role, 0
			else
				Repository.WriteOutput outputTabName, now() & " ERROR: No translation found for stereotype «" & copyItem.Stereotype & "» on item " & copyItem.Name, 0
			end if
		end if
	end if
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

function createTranslationsDictionary()
	set stereotypeTranslations = CreateObject("Scripting.Dictionary")
	stereotypeTranslations.Add "attribuutsoort"				,"AttributeType"
	stereotypeTranslations.Add "codelist"					,"Codelist"
	stereotypeTranslations.Add "data element"				,"DataElement"
	stereotypeTranslations.Add "extern"						,"External"
	stereotypeTranslations.Add "externe koppeling"			,"ExternalLink"
	stereotypeTranslations.Add "gegevensgroep"				,"AttributeGroup"
	stereotypeTranslations.Add "gegevensgroeptype"			,"AttributeGroupType"
	stereotypeTranslations.Add "gestructureerd datatype"	,"StructuredDataType"
	stereotypeTranslations.Add "objecttype"					,"FeatureType"
	stereotypeTranslations.Add "primitief datatype"			,"PrimitiveDataType"
	stereotypeTranslations.Add "referentie element"			,"ReferenceElement"
	stereotypeTranslations.Add "referentielijst"			,"ReferenceList"
	stereotypeTranslations.Add "relatieklasse"				,"AssociationClass"
	'stereotypeTranslations.Add "relatierol"					,"FeatureAssociationRole" 'FeatureAssociatonRole does not exist in the profile NEN3610 
	stereotypeTranslations.Add "relatierol"					,"AssociationRole"
	stereotypeTranslations.Add "relatiesoort"				,"FeatureAssociationType"
	stereotypeTranslations.Add "union"						,"Union"
	stereotypeTranslations.Add "union element"				,"UnionElement"
	stereotypeTranslations.Add "view"						,"View"
	stereotypeTranslations.Add "toepassing"					,"Application"
	stereotypeTranslations.Add "basismodel"					,"Basemodel"
	stereotypeTranslations.Add "domein"						,"Domain"
	stereotypeTranslations.Add "intern"						,"Internal"
	stereotypeTranslations.Add "project"					,"Project"
	stereotypeTranslations.Add "prullenbak"					,"Recycle bin"
	stereotypeTranslations.Add "formele historie"			,"Formal history"
	stereotypeTranslations.Add "materiële historie"			,"Material history"
	stereotypeTranslations.Add "formele levensduur"			,"Formal lifecycle"
	stereotypeTranslations.Add "materiële levensduur"		,"Material lifecycle"
	stereotypeTranslations.Add "identificatie"				,"Identification"
	stereotypeTranslations.Add "voidable"					,"Voidable"
	'also add each translated term as a translation of itself
	dim translatedTerm
	for each translatedTerm in stereotypeTranslations.Items
		if not stereotypeTranslations.Exists(lcase(translatedTerm)) then
			stereotypeTranslations.Add lcase(translatedTerm), translatedTerm
		end if
	next
end function