'[path=\Projects\Project A\Schema Scripts]
'[group=Schema Scripts]


!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Redirect profile to new source model
' Author: Geert Bellekens
' Purpose: Updates the given profile to point to a new source model (identical to the existing source model)
' Date: 2019-01-29
'
const outPutName = "Redirect profile"

function Main ()
	
	dim artifact as EA.Element
	set artifact = Repository.GetTreeSelectedObject()
	if artifact is Nothing then
		msgbox "Please select a schema artifact"
		exit function
	end if
	if artifact.ObjectType <> otElement then
		msgbox "Please select a schema artifact"
		exit function
	end if
		
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Starting Redirect profile", artifact.ElementID
	
	dim xmlDOM 
	set  xmlDOM = CreateObject( "Microsoft.XMLDOM" )
	'set  xmlDOM = CreateObject( "MSXML2.DOMDocument.4.0" )
	xmlDOM.validateOnParse = false
	xmlDOM.async = false
	'processing instruction
	dim node 
	set node = xmlDOM.createProcessingInstruction( "xml", "version='1.0'")
	xmlDOM.appendChild node
	'get xml schema content
	dim schemaContent
	set schemaContent = getSchemaContent(artifact)
	xmlDom.appendChild schemaContent.documentElement
'	'debug
'	writefile "c:\\temp\\schemaContents.xml", xmlDom.xml
	
	dim successfull
	successfull = redirectProfile(xmlDom)
	
'	'debug
'	writefile "c:\\temp\\schemaContents_cleaned.xml", xmlDom.xml
	if not successfull then
		Repository.WriteOutput outPutName, now() & " Script Cancelled", artifact.ElementID
		exit function
	end if
	'ask for permission to save schema content
	dim userInput
	userinput = MsgBox( "Save Changes to profile?", vbYesNo + vbQuestion, "Save Changes")
	'save the schema content
	if userinput = vbYes then
		dim sqlUpdateSchema
		sqlUpdateSchema = "update t_document set StrContent = N'" & xmlDom.xml & "' where ElementID = '" & artifact.ElementGUID & "'"
		Repository.Execute sqlUpdateSchema
		Repository.AdviseElementChange(artifact.ElementID)
	end if
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Finished Redirect profile", artifact.ElementID
end function

function redirectProfile(xmlDom)
	'default false
	redirectProfile = false
	dim originalSourcePackage as EA.Package
	dim newSourcePackage as EA.Package
	msgbox "Please select the original source package"
	set originalSourcePackage = selectPackage()
	if originalSourcePackage is nothing then
		exit function
	end if
	msgbox "Please select the new source package"
	set newSourcePackage = selectPackage()
	if newSourcePackage is nothing then
		exit function
	end if
	'get the list of old/new guid's
	dim guidDictionary
	set guidDictionary = CreateObject("Scripting.Dictionary")
	tracePackageElements originalSourcePackage, newSourcePackage, guidDictionary
	'get the string from the xmlDom
	dim xmlString 
	xmlString =  xmlDom.xml
	'replace the GUID's
	dim originalGuid
	dim newGUID
	for each originalGuid in guidDictionary.keys
		newGUID = guidDictionary(originalGuid)
		'replace
		xmlString = Replace(xmlString, originalGuid, newGUID)
		'mark successful
		redirectProfile = true
	next
	'write it back to the xmlDom
	xmlDom.LoadXML xmlString
end function



function tracePackageElements(originalPackage, copyPackage, guidDictionary)

	dim originalElement as EA.Element
	dim copyElement as EA.Element
	'find corresponding element
	for each originalElement in originalPackage.Elements
		
		'only process elements that have a name and are Class, Enumeration, Datatype, PrimitiveType
		if len(originalElement.Name) > 0 and _
		  (originalElement.Type = "Class" or originalElement.Type = "Enumeration" _
     	  or originalElement.Type = "DataType" or originalElement.Type = "PrimitiveType" ) then
			'Repository.WriteOutput outputTabName, now() & " Processing " & originalElement.Type & ": " & originalElement.Name ,0
			dim matchFound
			matchFound = false
			for each copyElement in copyPackage.Elements
				if copyElement.Name = originalElement.Name _
				  and copyElement.Type = originalElement.Type then
					'found a match
					traceElements originalElement, copyElement, guidDictionary
					matchFound = true
					exit for
				end if
			next
			if matchFound then
				Repository.WriteOutput outPutName, now() & " Match found for " & originalElement.Type & ": " & originalElement.Name ,0
			else
				Repository.WriteOutput outPutName, now() & " Match NOT found for " & originalElement.Type & ": " & originalElement.Name ,0
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
				tracePackageElements originalSubPackage, copySubpackage, guidDictionary
				exit for
			end if
		next
	next
	'return
	set tracePackageElements = guidDictionary
end function

function traceElements (originalElement,copyElement,guidDictionary)
	'add the original/new guid's of the elements
	guidDictionary.Add originalElement.ElementGUID, copyElement.ElementGUID
	'trace attributes
	traceAttributes originalElement, copyElement, guidDictionary
	'trace associations
	traceAssociations originalElement, copyElement, guidDictionary
end function


function traceAttributes(originalElement,copyElement, guidDictionary)
	dim originalAttribute as EA.Attribute
	dim copyAttribute as EA.Attribute
	for each originalAttribute in originalElement.Attributes
		for each copyAttribute in copyElement.Attributes
			if copyAttribute.Name = originalAttribute.Name then
				'found match, add to dictionary
				guidDictionary.Add originalAttribute.AttributeGUID, copyAttribute.AttributeGUID
				exit for
			end if
		next
	next
end function

function traceAssociations (originalElement,copyElement, guidDictionary)
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
					AND copyConnector.ClientEnd.Role = originalConnector.ClientEnd.Role _
					AND copyConnector.ClientEnd.Aggregation = originalConnector.ClientEnd.Aggregation _
					AND copyConnector.SupplierEnd.Role = originalConnector.SupplierEnd.Role _
					AND copyConnector.SupplierEnd.Aggregation = originalConnector.SupplierEnd.Aggregation then
					'AND copyConnector.ClientEnd.Cardinality = originalConnector.ClientEnd.Cardinality _
					'AND copyConnector.SupplierEnd.Cardinality = originalConnector.SupplierEnd.Cardinality _
					'connector properties match, now check the other ends
					dim originalOtherEnd as EA.Element
					dim copyOtherEnd as EA.Element
					set originalOtherEnd = Repository.GetElementByID(originalConnector.SupplierID)
					set copyOtherEnd = Repository.GetElementByID(copyConnector.SupplierID)
					if copyOtherEnd.Name = originalOtherEnd.Name then
						'found a match, add to dictionary
						guidDictionary.Add originalConnector.ConnectorGUID , copyConnector.ConnectorGUID
						exit for
					end if
				end if
			next
		end if
	next
end function


function addAssociatonsInProfile(xmlDom)
	dim xmlSchema
	set xmlSchema = xmlDOM.selectSingleNode("//schema")
	dim allClassNodes
	'get list of class nodes
	set allClassNodes = getClassNodes(xmlDom)
	dim processedNodes
	set processedNodes = CreateObject("Scripting.Dictionary")
	dim elementGUID
	for each elementGUID in allClassNodes.Keys
		'get the element
		dim element as EA.Element
		set element = Repository.GetElementByGuid(elementGUID)
		if not element is nothing then
			createOrUpdateElementNode xmlDOM, xmlSchema, element, allClassNodes, processedNodes
		end if
	next
end function

function getClassNodes(schemaContent)
	dim classNodeDictionary 
	set classNodeDictionary = CreateObject("Scripting.Dictionary")
	'loop class nodes
	dim classNodes
	set classNodes = schemaContent.SelectNodes("//class")
	dim classNode
	for each classNode in classNodes
		'get guid attribute
		dim classGuid
		classGuid = classNode.GetAttribute("guid")
		'add to Dictionary
		classNodeDictionary.Add classGuid, classNode
	next
	'return dictionary
	set getClassNodes = classNodeDictionary
end function

function getPropertyNodes(schemaContent)
	dim propertyNodeDictionary 
	set propertyNodeDictionary = CreateObject("Scripting.Dictionary")
	'loop property nodes
	dim propertyNodes
	set propertyNodes = schemaContent.SelectNodes("//property")
	dim propertyNode
	for each propertyNode in propertyNodes
		'get guid attribute
		dim propertyuid
		propertyGuid = propertyNode.GetAttribute("guid")
		'add to Dictionary
		propertyNodeDictionary.Add propertyGuid, propertyNode
	next
	'return dictionary
	set getPropertyNodes = propertyNodeDictionary
end function

function getSchemaContent(artifact)
		dim sqlGet, xmlQueryResult
		sqlGet = "select doc.StrContent from t_document doc where doc.ElementID = '" & artifact.ElementGUID & "'"
		xmlQueryResult = Repository.SQLQuery(sqlGet)
		
		xmlQueryResult = replace(xmlQueryResult,"&lt;","<")
		xmlQueryResult = replace(xmlQueryResult,"&gt;",">")
		'Repository.WriteOutput outPutName, "xmlQueryResult: " & xmlQueryResult  , 0
		Dim xDoc 
		Set xDoc = CreateObject("Microsoft.XMLDOM")
		'Set xDoc = CreateObject("Msxml2.DOMDocument")
		'load the resultset in the xml document
		xDoc.LoadXML xmlQueryResult
		dim strContentNode
		for each strContentNode in xDoc.SelectNodes("//message")
			xDoc.LoadXML strContentNode.xml
			exit for
		next
		'return value
		set getSchemaContent = xDoc
end function


function createOrUpdateElementNode(xmlDOM, xmlSchema, element, allNodes, processedNodes)
	'do not continue if already processed
	if processedNodes.Exists(element.ElementGUID) then
		exit function
	end if
	'update log
	Repository.WriteOutput outPutName, "Processing element " & element.Name, element.ElementID
	'add to list of processed nodes
	processedNodes.Add element.ElementGUID, xmlClass
	dim xmlClass
	'get the classNode
	set xmlClass = allNodes(element.ElementGUID)
	'add the element associations
	addElementAssociations xmlClass, xmlDOM, xmlSchema, element, allNodes, processedNodes
end function

function addElementAssociations(xmlClass, xmlDOM, xmlSchema, element, allNodes, processedNodes)
	'get propertiesnode
	dim xmlProperties
	'check if properties node exists
	set xmlProperties = xmlClass.selectSingleNode("./properties")
	if xmlProperties is nothing then
		set xmlProperties= xmlDOM.createElement("properties")
		'add xmlProperties to class node
		xmlClass.appendChild xmlProperties
	end if
	dim propertyNode
	'add associations only if they start at the given element
	dim relation as EA.Connector
	for each relation in element.Connectors
		if (relation.Type = "Association" _
		or relation.Type = "Aggregation" ) _
		and relation.ClientID = element.ElementID then
			'check if the other end is also part of the schema
			dim targetElement as EA.Element
			set targetElement = Repository.GetElementByID(relation.SupplierID)
			if allNodes.Exists(targetElement.ElementGUID) then
				'check if propertyNode exists
				set propertyNode = xmlProperties.selectSingleNode("./property[@guid='" & relation.ConnectorGUID & "']")
				if propertyNode is nothing then
					'update log
					Repository.WriteOutput outPutName, "Adding association '" & relation.name & "' from '" & element.Name & "' to '" &targetElement.Name & "'", element.ElementID
					'add association node
					xmlProperties.appendChild createPropertyNode (xmlDOM, relation.ConnectorGUID, "association")
				end if
			end if
		end if
	next
end function

function addAncestry(xmlClass, xmlDOM, xmlSchema, element, allNodes, processedNodes)
	'not for XSDSimpletypes
'	if element.HasStereotype("XSDsimpleType") then
'		exit function
'	end if
	'loop base elements
	dim sqlGetBaseElements
	sqlGetBaseElements = "select c.End_Object_ID as Object_ID from t_connector c " & _
						" where c.Connector_Type = 'Generalization' " & _
						" and c.Start_Object_ID = " & element.ElementID
	dim baseElements
	set baseElements = getElementsFromQuery(sqlGetBaseElements)
	if baseElements.Count > 0 then
		'composition attribute
		dim xmlCompositionAttr
		set xmlCompositionAttr = xmlDOM.createAttribute("composition")
		xmlCompositionAttr.nodeValue = "inherit"
		xmlClass.setAttributeNode(xmlCompositionAttr)
		'add ancestry node
		dim xmlAncestry
		set xmlAncestry = xmlDOM.createElement("ancestry")
		'loop base elements
		dim baseElement as EA.Element
		for each baseElement in baseElements
			'create ancesterNode
			dim xmlAncestor
			set xmlAncestor = xmlDOM.createElement("ancestor")
			'name attribute
			dim xmlNameAtttr 
			set xmlNameAtttr = xmlDOM.createAttribute("name")
			xmlNameAtttr.nodeValue = baseElement.Name
			xmlAncestor.setAttributeNode(xmlNameAtttr)
			'guid attribute
			dim xmlguidAtttr 
			set xmlguidAtttr = xmlDOM.createAttribute("guid")
			xmlguidAtttr.nodeValue = baseElement.ElementGUID
			xmlAncestor.setAttributeNode(xmlguidAtttr)
			'add to ancestry node
			xmlAncestry.appendChild xmlAncestor
			'create element node for ancestor
			createOrUpdateElementNode xmlDOM, xmlSchema, baseElement, allNodes, processedNodes
		next
		'add to xmlClassNode
		xmlClass.appendChild xmlAncestry
	end if
end function

function createPropertyNode (xmlDOM, guid, propertyType)
	dim xmlProperty
	set xmlProperty = xmlDOM.createElement("property")
	
	'guid attribute
	dim xmlguidAtttr 
	set xmlguidAtttr = xmlDOM.createAttribute("guid")
	xmlguidAtttr.nodeValue = guid
	xmlProperty.setAttributeNode(xmlguidAtttr)
	
	'type attribute
	dim xmltypeAtttr 
	set xmltypeAtttr = xmlDOM.createAttribute("type")
	xmltypeAtttr.nodeValue = propertyType
	xmlProperty.setAttributeNode(xmltypeAtttr)
	
	'return node
	set createPropertyNode = xmlProperty
end function


function createDescriptionNode(xmlDOM, selectedElement)
	dim xmlDescription
	set xmlDescription = xmlDOM.createElement( "description" )
	
	'name attribute
	dim xmlNameAtttr 
	set xmlNameAtttr = xmlDOM.createAttribute("name")
	xmlNameAtttr.nodeValue = selectedElement.Name
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
	xmlversionAtttr.nodeValue = "13.5.1351.1351"
	xmlDescription.setAttributeNode(xmlversionAtttr)
	
	'xmlns attribute
	dim xmlxmlnsAtttr 
	set xmlxmlnsAtttr = xmlDOM.createAttribute("xmlns")
	xmlxmlnsAtttr.nodeValue = "Der:"
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

function writefile(filename, contents)
	dim fileSystemObject
	dim outputFile
		
	set fileSystemObject = CreateObject( "Scripting.FileSystemObject" )
	set outputFile = fileSystemObject.CreateTextFile(filename, true )
	outputFile.Write contents
	outputFile.Close
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

private function createArtifact(ownerPackage)
	'add new artifact in owner package
	dim artifact as EA.Element
	set artifact = ownerPackage.Elements.AddNew(ownerPackage.Name & "_Schema", "Artifact")
	artifact.Update
	'save the Schemacomposer property in the Style settings
	Repository.Execute "update t_object set Style = 'MessageProfile=1;' where ea_guid = '" & artifact.ElementGUID & "'"
	set createArtifact = artifact
end function

'test
main