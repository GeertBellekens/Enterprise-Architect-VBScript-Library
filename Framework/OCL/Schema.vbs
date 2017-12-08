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
	  if not me.Elements.Exists(me.Context.ElementGUID) then
		addSchemaElement me.Context, me.Context.Name
	  end if
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
			dim source
			'figure out which association end or attribute we are talking about
			source = findSource(ocl.LeftHand)
			if Err.Number <> 0 then
				Repository.WriteOutput outPutName, now() &  " Error processing OCL statement:'" & ocl.Statement & "' ->" & Err.Description, 0
				Err.Clear
			end if
		next
		on error goto 0
	end function
	
	private function addSchemaElement(source, name)
		dim addElement
		'add the context as a SchemaElement
		if not me.Elements.Exists(source.ElementGUID) then
			addElement = true
		else
			dim existingElement
			set existingElement = me.Elements.Item(source.ElementGUID)
			if existingElement.name = name then
				addElement = false
				'return it
				set addSchemaElement = existingElement
			else
				addElement = true
			end if
		end if
		if addElement then
			'create new schema Element
			dim schemaElement
			set schemaElement = new SchemaElement
			schemaElement.Source = source
			schemaElement.Name = name
			'add it to the list
			me.Elements.Add schemaElement.GUID, schemaElement
			'return it
			set addSchemaElement = schemaElement
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
		set contextSchemaElement = addSchemaElement(localContext, localContext.Name)
		dim i
		'start from the second one as the first one will be "self"
		for i = 1 to Ubound(identifierParts)
			identifierPart = identifierParts(i)
			set localContext = processIdentifierPart(identifierPart, contextSchemaElement)
			'make sure the local context exists as SchemaElement
			set contextSchemaElement = addSchemaElement(localContext, localContext.Name)
		next
	end function
	
	private function processIdentifierPart(identifierPart, contextSchemaElement)
		'get the attribute or association starting from the localContext
		dim correspondingProperty
		set correspondingProperty = contextSchemaElement.getProperty(identifierPart)
		'get the new local context
		if not correspondingProperty is nothing then
			set processIdentifierPart = correspondingProperty.Classifier
		else
			Err.Raise vbObjectError + 10, "processIdentifierPart", "Could not find '" & identifierPart & "' in the context of '" & localContext.Name & "'"
		end if
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
			'add node
			xmlSchema.appendChild createElementNode(xmlDom, element)
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
		
		'add propertiesnode
		dim xmlProperties
		set xmlProperties= xmlDOM.createElement("properties")
		dim schemaProperty
		for each schemaProperty in element.Properties.Items
			xmlProperties.appendChild createPropertyNode (xmlDOM, schemaProperty)
		next
				
		'add xmlProperties to class node
		xmlClass.appendChild xmlProperties
		
		'return node
		set createElementNode = xmlClass
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
		xmltypeAtttr.nodeValue = schemaProperty.PropertyType
		xmlProperty.setAttributeNode(xmltypeAtttr)
		
		'return node
		set createPropertyNode = xmlProperty
	end function
	
end Class