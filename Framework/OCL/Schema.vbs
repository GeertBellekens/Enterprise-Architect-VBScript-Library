'[path=\Framework\OCL]
'[group=OCL]


'Author: Geert Bellekens
'Date: 2017-12-06
'Purpose: Class representing a Schema Composer Schema

Class Schema 
	'private variables
	Private m_Artifact
	Private m_Context
	Private m_Elements
	Private m_Owner
	'constructor
	Private Sub Class_Initialize
		set m_Artifact = nothing
		set m_Context = nothing
		set m_Elements = CreateObject("Scripting.Dictionary")
	End Sub
	
	' Artifact property (EA.Element)
	Public Property Get Artifact
	 set Artifact = m_Artifact
	End Property
	Public Property Let Artifact(value)
	  set m_Artifact = value
	End Property
	
	' Context property (EA.Element)
	Public Property Get Context
	 set Context = m_Context
	End Property
	Public Property Let Context(value)
		set m_Context = value
		'add the context as a SchemaElement
		dim contextSchemaElement
		set contextSchemaElement = addSchemaElement(me.Context, me.Context.Name, false)
		contextSchemaElement.IsRoot = true
	End Property
	
	' Elements property (Dictionary of source element GUID and SchemaElements)
	Public Property Get Elements
	 set Elements = m_Elements
	End Property
	Public Property Let Elements(value)
	  set m_Elements = value
	End Property
	
	' Owner property (Schema)
	Public Property Get Owner
	 set Owner = m_Owner
	End Property
	Public Property Let Owner(value)
	  set m_Owner = value
	End Property
	
	'saves the schema to the artifact
	Public function save()
		'create artifact
		dim artifact as EA.Element
		set artifact = createArtifact()
		'get XML
		dim xmlString
		xmlString = getXML()
		'create new record in t_document
		createTDocument artifact, xmlString
	end function
	
	function createTDocument(artifact, xmlString)
		dim timestamp
		timestamp = Year(now()) & "-" & Month(now()) & "-" & Day(now()) & " " & Hour(now()) & ":" & Minute(now) & ":" & Second(now())
		dim sqlCreateSchemaDocument
		sqlCreateSchemaDocument = " INSERT INTO [dbo].[t_document]             " & vbNewLine & _
								"            ([DocID]                        " & vbNewLine & _
								"            ,[DocName]                      " & vbNewLine & _
								"            ,[Notes]                        " & vbNewLine & _
								"            ,[Style]                        " & vbNewLine & _
								"            ,[ElementID]                    " & vbNewLine & _
								"            ,[ElementType]                  " & vbNewLine & _
								"            ,[StrContent]                   " & vbNewLine & _
								"            ,[BinContent]                   " & vbNewLine & _
								"            ,[DocType]                      " & vbNewLine & _
								"            ,[Author]                       " & vbNewLine & _
								"            ,[Version]                      " & vbNewLine & _
								"            ,[IsActive]                     " & vbNewLine & _
								"            ,[Sequence]                     " & vbNewLine & _
								"            ,[DocDate])                     " & vbNewLine & _
								"      VALUES                                " & vbNewLine & _
								"            ('" & CreateGuid() & "'         " & vbNewLine & _
								"            ,'" & artifact.Name & "'        " & vbNewLine & _
								"            ,NULL                           " & vbNewLine & _
								"            ,NULL                           " & vbNewLine & _
								"            ,'" & artifact.ElementGUID & "' " & vbNewLine & _
								"            ,'SC_MessageProfile'            " & vbNewLine & _
								"            ,N'" & xmlString & "'           " & vbNewLine & _
								"            ,NULL                           " & vbNewLine & _
								"            ,'SC_MessageProfile'            " & vbNewLine & _
								"            ,'OCL to Schema Script'         " & vbNewLine & _
								"            ,NULL                           " & vbNewLine & _
								"            ,1                              " & vbNewLine & _
								"            ,0                              " & vbNewLine & _
								"            ,'" & timestamp & "')                "
			Repository.Execute sqlCreateSchemaDocument
	end function
	
	private function createArtifact()
		dim ownerPackage as EA.Package
		set ownerPackage = Repository.GetPackageByID(me.Context.PackageID)
		'add new artifact in owner package
		dim artifact as EA.Element
		set artifact = ownerPackage.Elements.AddNew(me.context.Name & "_Schema", "Artifact")
		artifact.Update
		'save the Schemacomposer property in the Style settings
		Repository.Execute "update t_object set Style = 'MessageProfile=1;' where ea_guid = '" & artifact.ElementGUID & "'"
		set createArtifact = artifact
	end function
	

	
	'processes the given ocl constraint and adds the required elements and properties to the schema
	public function processOCLs(OCLs, outputName)
		dim ocl
		on error resume next
		for each ocl in OCLs
			dim sourceProperty
			set sourceProperty = nothing
			'figure out which association end or attribute we are talking about
			dim sourceProperties
			set sourceProperties = findSource(ocl.LeftHand)
			if ocl.IsFacet  then
				'facets are defined on the content, but should result in a tagged value on the attribute that uses the BDT as type.
				if sourceProperties.Count > 1 then
					set sourceProperty = sourceProperties(1)
				end if
			else
				'other constraints are actually defined on the actual property
				if sourceProperties.Count > 0 then
					set sourceProperty = sourceProperties(0)
				end if
			end if
			'set the constraint (if any) on the source
			if not sourceProperty is nothing then
				processConstraint sourceProperty, ocl
			end if
			'report any errors
			if Err.Number <> 0 then
				Repository.WriteOutput outPutName, now() &  " Error processing OCL statement:'" & ocl.Statement & "' ->" & Err.Description, 0
				Err.Clear
			end if
		next
		'set error handling back to standard
		on error goto 0
		'add the missing attributes
		addMissingAttributes
		'remove the properties with maxOccurs = 0
		deleteNotNeededProperties
		'merge duplicate refines
		dim element
		for each element in me.Elements.Items
			element.mergeAllRedefines
		next
		
	end function
	
	'add the missing attributes for all «BDT» elements as they should always be part of the schema
	function addMissingAttributes()
		dim element
		for each element in me.Elements.Items
			if element.Source.Stereotype = "BDT" then 'TODO: check if really only needed for BDT's
				'loop all attributes and add them.
				dim attribute as EA.Attribute
				for each attribute in element.Source.Attributes
					createCorrespondingProperty attribute.Name, element
				next
			end if
		next
	end function
	
	'delete the properties that have a zero maxOccurs
	function deleteNotNeededProperties()
		dim element
		for each element in me.Elements.Items
			'then delete the properties that have a maxOccurs of 0
			dim schemaProperty
			for each schemaProperty in element.Properties.Items
				if schemaProperty.maxOccurs = "0" then
					schemaProperty.Delete
				end if
			next
		next
	end function

	
	
	private function processConstraint(sourceProperty, ocl)
		if not sourceProperty is nothing _
			and not ocl is nothing then
			'check the type of constraint
			select case ocl.ConstraintType
				case OCLEqual
					setEqualValue sourceProperty, ocl.rightHand
					'process next statement
					processConstraint sourceProperty, ocl.NextOCLStatement
				case OCLMultiplicity
					if trim(lcase(ocl.Operator)) = "->isempty()" then
						setMultiplicityValue sourceProperty, 0
					else
						setMultiplicityValue sourceProperty, ocl.rightHand
					end if
				case OCLChoice
					'report choice
					if not ocl.NextOCLStatement is nothing then
						Repository.WriteOutput outPutName, now() &  " Choice constraint found between :'" & ocl.leftHand & "' with GUID " & sourceProperty.GUID &_
									" with attribute '" & ocl.NextOCLStatement.LeftHand & "'",0
					else
						'invalid choice
						Err.Raise vbObjectError + 13, "processConstraint", " Invalid choice statement in OCL constraint '" & ocl.Statement & "'"
					end if
				case else
					'process facets
					if ocl.IsFacet then
						processFacetConstraint sourceProperty, ocl
					else
						'unknown constraint, should never happen
						Err.Raise vbObjectError + 14, "processConstraint", " ConstraintType:'" & ocl.Operator & "' is Unknown in OCL constraint '" & ocl.Statement & "'"
					end if
			end select
		end if
	end function
	
	private function processFacetConstraint(sourceProperty, ocl)
		'check if a tagged value for the facet already exists. If it does check if the value corresponds. If the value is different hen alert the user
		'get the taggedValueName for the facet
		dim facetTag as EA.TaggedValue
		set facetTag = getExistingOrNewTaggedValue(sourceProperty.Source, ocl.FacetName)
		'if the value is blank then simply fill it in
		if facetTag.Value = "" then
			facetTag.Value = replace(trim(ocl.RightHand), """","") 'we don't need any double quotes
			facetTag.Update
		elseif facetTag.Value <> replace(trim(ocl.RightHand), """","")  then
			Repository.WriteOutput outPutName, now() &  " Duplicate facet found for attribute with with GUID " & sourceProperty.GUID &_
									" Original facet: " & facetTag.Name & " -> " & facetTag.Value & _
									" OCL statement = '" & ocl.Statement & "'",0
		end if
	end function
	
	private function setEqualValue(sourceProperty, rightHandValue)
		if not sourceProperty.ClassifierSchemaElement is nothing then
			'remove any parentheses left
			rightHandValue = replace(rightHandValue, ")", "")
			'get the name of the property
			dim rightHandParts
			dim identifier
			rightHandParts = split(rightHandValue,"::")
			'gets the last value of the array
			identifier = rightHandParts(Ubound(rightHandParts))
			processIdentifierPart identifier, sourceProperty.ClassifierSchemaElement, true
		else
			Err.Raise vbObjectError + 11, "setEqualValue", "No ClassiferSchemaElement found for Property: '" & sourceProperty.Name & "' with GUID: '" & sourceProperty.GUID & "'"
		end if
	end function
	private function setMultiplicityValue(sourceProperty, rightHandValue)
		'if the right hand value is empty then it is an ->notEmpty() constraint -> minOccurs = 1
		if len(rightHandValue) = 0 then
			sourceProperty.minOccurs = "1"
		else
			if rightHandValue = "1" then
				sourceProperty.minOccurs = "1"
				sourceProperty.maxOccurs = "1"
			elseif lcase(trim(rightHandValue)) = "optional" then
				sourceProperty.minOccurs = "0"
			elseif len(rightHandValue) >= 3 _
				and left(rightHandValue,2) = "<=" then
				'find the actual value
				sourceProperty.maxOccurs = mid(rightHandValue,3)
			elseif trim(rightHandValue) = "0" then
				'set maxOccurs to 0
				sourceProperty.maxOccurs = "0"
			else
				Err.Raise vbObjectError + 12, "setMultiplicityValue", "Value '" & rightHandValue & "' not valid as multiplicityValue for Property: '" & sourceProperty.Name & "' with GUID: '" & sourceProperty.GUID & "'" 
			end if
		end if
	end function
	
	public function removeSchemaElement(element)
		if me.Elements.Exists(element.GUID) then
			if element.IsRedefinition then
				dim parentElement
				set parentElement = me.Elements.Item(element.GUID)
				'must be a redefined element
				parentElement.removeRedefine(element)
			else
				if element.Redefines.count = 0 then
					'do not delete elements that still have redefines
					me.Elements.Remove element.GUID
				end if
			end if
		end if	
	end function
	
	private function addSchemaElement(source, name, isNew)
		'add the context as a SchemaElement
		if not me.Elements.Exists(source.ElementGUID) then
			'create new schema Element
			dim schemaElement
			set schemaElement = new SchemaElement
			schemaElement.Source = source
			schemaElement.Name = name
			schemaElement.Schema = me
			'add it to the list
			me.Elements.Add schemaElement.GUID, schemaElement
			'return it
			set addSchemaElement = schemaElement
		else
			'element exists already
			dim existingElement
			set existingElement = me.Elements.Item(source.ElementGUID)
			dim test as EA.Element
			if isNew and existingElement.Source.Type = "Enumeration" then
				'add redefine and return it
				set addSchemaElement = existingElement.addNewRedefine()
			else
				'return it
				set addSchemaElement = existingElement
			end if
		end if
	end function
	
	private function findSource(identifierString)
		'split into parts 
		dim identifierParts
		identifierParts = split(identifierString, ".")
		dim identifierPart
		dim localContext as EA.Element
		set localContext = me.Context 'start with he MA as context
		dim contextSchemaElement
		set contextSchemaElement = addSchemaElement(localContext, localContext.Name, true)
		dim i
		dim correspondingProperty
		dim correspondingProperties
		set correspondingProperties = CreateObject("System.Collections.ArrayList")
		'start from the second one as the first one will be "self"
		for i = 1 to Ubound(identifierParts)
			identifierPart = identifierParts(i)
			set correspondingProperty = createCorrespondingProperty(identifierPart,contextSchemaElement)
			'set the context schema element
			set contextSchemaElement = correspondingProperty.ClassifierSchemaElement
			'add it to the list
			correspondingProperties.Add correspondingProperty
		next
		'reverse the list to make the last corresponding property the first
		correspondingProperties.Reverse
		'return
		set findSource = correspondingProperties
	end function
	function createCorrespondingProperty(identifierPart,contextSchemaElement)
		dim correspondingProperty
		dim localContext
		dim isNew
		set correspondingProperty = processIdentifierPart(identifierPart, contextSchemaElement, isNew)
		'set the new local context
		set localContext = correspondingProperty.Classifier
		if correspondingProperty.ClassifierSchemaElement is nothing then
			dim newContext
			'make sure the local context exists as SchemaElement
			set newContext = addSchemaElement(localContext, localContext.Name, isNew)
			'set the classifierSchemaElement on the correspondingProperty
			correspondingProperty.ClassifierSchemaElement = newContext
		end if
		'return property
		set createCorrespondingProperty = correspondingProperty
	end function 
	
	private function processIdentifierPart(identifierPart, contextSchemaElement, byRef isNew)
		'get the attribute or association starting from the localContext
		dim correspondingProperty
		set correspondingProperty = contextSchemaElement.getProperty(identifierPart, isNew)
		'return the corresponding property
		if correspondingProperty is nothing and identifierPart <> "content" then
			'check if there is an attribute name "content" and then process the identifierPart on that one
			dim contentProperty 
			set contentProperty = nothing 'initialize
			'turn off error checking
			on error resume next
			set contentProperty = createCorrespondingProperty("content",contextSchemaElement)
			'clear error and turn error checking back on
			if Err.Number <> 0 then
				Err.Clear
			end if
			on error goto 0
			'now try this property
			if not contentProperty is nothing then
				dim contentSchemaElement 
				set contentSchemaElement = contentProperty.ClassifierSchemaElement
				if not contentSchemaElement is nothing then
					set correspondingProperty = contentSchemaElement.getProperty(identifierPart, isNew)
				end if
			end if
			if correspondingProperty is nothing then
				'if still not found then raise error
				Err.Raise vbObjectError + 10, "processIdentifierPart", "Could not find '" & identifierPart & "' in the context of '" & contextSchemaElement.Name & "' with GUID : " & contextSchemaElement.GUID
			end if
		end if
		'return the corresponding property
		set processIdentifierPart = correspondingProperty
	end function
	
	public function tryGetElement(guid, byRef element)
		dim exists
		exists = me.Elements.Exists(guid)
		if exists then
			set element = me.Elements.Item(guid)
		end if
		tryGetElement = exists
	end function
	
	private function getXML()	
		set xmlDOM = CreateObject( "Microsoft.XMLDOM" )
		xmlDOM.validateOnParse = false
		xmlDOM.async = false
		
		dim node 
		set node = xmlDOM.createProcessingInstruction( "xml", "version='1.0'")
		xmlDOM.appendChild node
		
		dim xmlRoot 
		set xmlRoot = xmlDOM.createElement( "message" )
		xmlDOM.appendChild xmlRoot
		
		'add description node
		xmlRoot.appendChild createDescriptionNode(xmlDOM)
		
		'add schema node
		dim xmlSchema 
		set xmlSchema = xmlDOM.createElement("schema")
		'count attribute
		dim xmlcountAtttr 
		set xmlcountAtttr = xmlDOM.createAttribute("count")
		xmlcountAtttr.nodeValue = me.Elements.Count
		xmlSchema.setAttributeNode(xmlcountAtttr)
		'add the schema node to the root
		xmlRoot.appendChild xmlSchema
		'add all the elements to the schema
		dim element 
		for each element in me.Elements.Items
			'don't add elements that are not used
			if element.IsRoot _
				OR element.ReferencingProperties.Count > 0 _
				OR element.Redefines.Count > 0 then
				'add node
				xmlSchema.appendChild createElementNode(xmlDom, element)
			end if
		next
		
		'return xml string
		getXML = xmlDom.xml
	end function
	
	private function createDescriptionNode(xmlDOM)
		dim xmlDescription
		set xmlDescription = xmlDOM.createElement( "description" )
		
		'name attribute
		dim xmlNameAtttr 
		set xmlNameAtttr = xmlDOM.createAttribute("name")
		xmlNameAtttr.nodeValue = me.Context.Name & "_Schema"
		xmlDescription.setAttributeNode(xmlNameAtttr)
		
		'namespace attribute
		dim xmlnamespaceAtttr 
		set xmlnamespaceAtttr = xmlDOM.createAttribute("namespace")
		xmlnamespaceAtttr.nodeValue = ""
		xmlDescription.setAttributeNode(xmlnamespaceAtttr)
		
		'schemaset attribute
		dim xmlschemasetAtttr 
		set xmlschemasetAtttr = xmlDOM.createAttribute("schemaset")
		xmlschemasetAtttr.nodeValue = "ECDM Message Composer"
		xmlDescription.setAttributeNode(xmlschemasetAtttr)
		
		'provider attribute
		dim xmlproviderAtttr 
		set xmlproviderAtttr = xmlDOM.createAttribute("provider")
		xmlproviderAtttr.nodeValue = "ECDM Message Composer"
		xmlDescription.setAttributeNode(xmlproviderAtttr)
		
		'model attribute
		dim xmlmodelAtttr 
		set xmlmodelAtttr = xmlDOM.createAttribute("model")
		xmlmodelAtttr.nodeValue = Repository.ProjectGUID
		xmlDescription.setAttributeNode(xmlmodelAtttr)
		
		'modelURL attribute
		dim xmlmodelURLAtttr 
		set xmlmodelURLAtttr = xmlDOM.createAttribute("modelURL")
		xmlmodelURLAtttr.nodeValue = ""
		xmlDescription.setAttributeNode(xmlmodelURLAtttr)
		
		'version attribute
		dim xmlversionAtttr 
		set xmlversionAtttr = xmlDOM.createAttribute("version")
		xmlversionAtttr.nodeValue = "12.1.1230.1230" 
		xmlDescription.setAttributeNode(xmlversionAtttr)
		
		'xmlns attribute
		dim xmlxmlnsAtttr 
		set xmlxmlnsAtttr = xmlDOM.createAttribute("xmlns")
		xmlxmlnsAtttr.nodeValue = "OCL:"
		xmlDescription.setAttributeNode(xmlxmlnsAtttr)
		
		'type attribute
		dim xmltypeAtttr 
		set xmltypeAtttr = xmlDOM.createAttribute("type")
		xmltypeAtttr.nodeValue = "schema"
		xmlDescription.setAttributeNode(xmltypeAtttr)
		
		'auxiliary node
		dim xmlAuxiliary
		set xmlAuxiliary = xmlDOM.createElement( "auxiliary" )
		
		'xmlns attribute
		dim xmlxmlnsAtttrA 
		set xmlxmlnsAtttrA = xmlDOM.createAttribute("xmlns")
		xmlxmlnsAtttrA.nodeValue = ""
		xmlAuxiliary.setAttributeNode(xmlxmlnsAtttrA)
		'add auxiliary node
		xmlDescription.appendChild xmlAuxiliary
		
		'return node
		set createDescriptionNode = xmlDescription
	end function
	
	private function createElementNode(xmlDOM, element)
		dim xmlClass
		set xmlClass = xmlDOM.createElement( "class" )
		
		'name attribute
		dim xmlNameAtttr 
		set xmlNameAtttr = xmlDOM.createAttribute("name")
		xmlNameAtttr.nodeValue = element.Name
		xmlClass.setAttributeNode(xmlNameAtttr)
		
		'guid attribute
		dim xmlguidAtttr 
		set xmlguidAtttr = xmlDOM.createAttribute("guid")
		xmlguidAtttr.nodeValue = element.GUID
		xmlClass.setAttributeNode(xmlguidAtttr)
		
		'add the properties
		addProperties xmlDom, xmlClass, element
		
		'add redefines if needed
		addRedefines xmlDom, xmlClass, element
		
		'return node
		set createElementNode = xmlClass
	end function
	
	function addProperties (xmlDom, xmlClass, element)
		'create propertiesnode
		dim xmlProperties
		set xmlProperties= xmlDOM.createElement("properties")
		dim schemaProperty
		for each schemaProperty in element.Properties.Items
			xmlProperties.appendChild createPropertyNode (xmlDOM, schemaProperty)
		next
		'add xmlProperties to class node
		xmlClass.appendChild xmlProperties
	end function 
	
	function addRedefines (xmlDom, xmlClass, element)
		'only needed if any redefines exist
		if element.Redefines.Count > 0 then
			'create the redefines node
			dim xmlRedefine
			set xmlRedefine= xmlDOM.createElement("redefines")
			'loop the redefines
			dim redefine
			for each redefine in element.Redefines.Items
				'create the set node
				dim xmlSet
				set xmlSet = xmlDom.createElement("set")
				'add the attribute typename
				dim xmlTypeNameAttr
				set xmlTypeNameAttr = xmlDOM.createAttribute("typename")
				xmlTypeNameAttr.nodeValue = redefine.Name
				xmlSet.setAttributeNode(xmlTypeNameAttr)
				'add the properties for this redefined property
				dim schemaProperty
				for each schemaProperty in redefine.Properties.Items
					xmlSet.appendChild createPropertyNode (xmlDOM, schemaProperty)
				next
				'add the set node to the redefines node
				xmlRedefine.AppendChild xmlSet
			next
			'add the redefines node to the xmlClass node
			xmlClass.AppendChild xmlRedefine
		end if
	end function 

	private function createPropertyNode (xmlDOM, schemaProperty)
		dim xmlProperty
		set xmlProperty = xmlDOM.createElement("property")
		
		'guid attribute
		dim xmlguidAtttr 
		set xmlguidAtttr = xmlDOM.createAttribute("guid")
		xmlguidAtttr.nodeValue = schemaProperty.GUID
		xmlProperty.setAttributeNode(xmlguidAtttr)
		
		'type attribute
		dim xmltypeAtttr 
		set xmltypeAtttr = xmlDOM.createAttribute("type")
		xmltypeAtttr.nodeValue = lcase(schemaProperty.PropertyType)
		xmlProperty.setAttributeNode(xmltypeAtttr)
		
		'check if schemaProperty has a restriction
		if schemaProperty.IsRestricted then
			'type attribute
			dim xmlRestrictedAttr 
			set xmlRestrictedAttr = xmlDOM.createAttribute("restricted")
			xmlRestrictedAttr.nodeValue = "true"
			xmlProperty.setAttributeNode(xmlRestrictedAttr)
			
			'MinOccurs attribute
			dim xmlMinOccursAttr 
			set xmlMinOccursAttr = xmlDOM.createAttribute("minOccurs")
			xmlMinOccursAttr.nodeValue = schemaProperty.minOccurs
			xmlProperty.setAttributeNode(xmlMinOccursAttr)
			
			'maxOccurs attribute
			dim xmlMaxOccursAttr 
			set xmlMaxOccursAttr = xmlDOM.createAttribute("maxOccurs")
			xmlMaxOccursAttr.nodeValue = schemaProperty.maxOccurs
			xmlProperty.setAttributeNode(xmlMaxOccursAttr)
			
			'reDefines attribute
			dim xmlRedefinesAttr 
			set xmlRedefinesAttr = xmlDOM.createAttribute("redefines")
			xmlRedefinesAttr.nodeValue = schemaProperty.redefines
			xmlProperty.setAttributeNode(xmlRedefinesAttr)			

			'byRef attribute -> always 0
			dim xmlByRefAttr 
			set xmlByRefAttr = xmlDOM.createAttribute("byRef")
			xmlByRefAttr.nodeValue = "0"
			xmlProperty.setAttributeNode(xmlByRefAttr)
			
			'inline attribute -> always 0
			dim xmlInlineAttr 
			set xmlInlineAttr = xmlDOM.createAttribute("inline")
			xmlInlineAttr.nodeValue = "0"
			xmlProperty.setAttributeNode(xmlInlineAttr)			
		end if
		
		'return node
		set createPropertyNode = xmlProperty
	end function
	
end Class