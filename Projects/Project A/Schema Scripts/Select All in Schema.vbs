'[path=\Projects\Project A\Schema Scripts]
'[group=Schema Scripts]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Select All in schema
' Author: Geert Bellekens
' Purpose: Reads the selected schema, and for each root elements selects all attributes and associations
' Date: 2018-09-24
'
const outPutName = "Select All in Schema"

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
	Repository.WriteOutput outPutName, now() & " Starting Select All in schema profile", artifact.ElementID
	
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
	selectAllInProfile xmlDom
'	'debug
'	writefile "c:\\temp\\schemaContents_cleaned.xml", xmlDom.xml
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
	Repository.WriteOutput outPutName, now() & " Finished Select All in Schema", artifact.ElementID
end function

function selectAllInProfile(xmlDom)
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
	'update the count attribute in the schema node
	dim schemaNode
	set schemaNode = xmlDom.selectSingleNode("//schema")
	'set the count value
	schemaNode.Attributes.getNamedItem("count").Text = allClassNodes.Count
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
	'check if the element does not exist yet
	if not allNodes.Exists(element.ElementGUID) then
		'create the xml Class node
		set xmlClass = xmlDOM.createElement( "class" )
		'add the node to the list
		allNodes.Add element.ElementGUID, xmlClass
		'name attribute
		dim xmlNameAtttr 
		set xmlNameAtttr = xmlDOM.createAttribute("name")
		xmlNameAtttr.nodeValue = element.Name
		xmlClass.setAttributeNode(xmlNameAtttr)
		
		'guid attribute
		dim xmlguidAtttr 
		set xmlguidAtttr = xmlDOM.createAttribute("guid")
		xmlguidAtttr.nodeValue = element.ElementGUID
		xmlClass.setAttributeNode(xmlguidAtttr)
		
		'add the xmlClass node to the schema
		xmlSchema.appendChild xmlClass
	else
		'get the classNode
		set xmlClass = allNodes(element.ElementGUID)
	end if
	'add the element details
	addElementDetails xmlClass, xmlDOM, xmlSchema, element, allNodes, processedNodes

end function

function addElementDetails(xmlClass, xmlDOM, xmlSchema, element, allNodes, processedNodes)
	'ancestry
	addAncestry xmlClass, xmlDOM, xmlSchema, element, allNodes, processedNodes

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
	
	'add attributes
	dim attribute as EA.Attribute
	for each attribute in element.Attributes
		'check if propertyNode exists
		set propertyNode = xmlProperties.selectSingleNode("./property[@guid='" & attribute.AttributeGUID & "']")
		if propertyNode is nothing then
			xmlProperties.appendChild createPropertyNode (xmlDOM, attribute.AttributeGUID, "attribute")
			'add an element node for the type of this attribute
			if attribute.ClassifierID > 0 then
				dim attributeType as EA.Element
				set attributeType = Repository.GetElementByID(attribute.ClassifierID)
				'add the node for the attributeType 
				createOrUpdateElementNode xmlDOM, xmlSchema, attributeType, allNodes, processedNodes
			end if
		end if
	next
	
	'add associations only if they start at the given element
	dim relation as EA.Connector
	for each relation in element.Connectors
		if (relation.Type = "Association" _
		or relation.Type = "Aggregation" ) then
			'check if propertyNode exists
			set propertyNode = xmlProperties.selectSingleNode("./property[@guid='" & relation.ConnectorGUID & "']")
			if propertyNode is nothing then
				'add association node
				xmlProperties.appendChild createPropertyNode (xmlDOM, relation.ConnectorGUID, "association")
				'add element node for the target of the relation
				dim targetElement as EA.Element
				set targetElement = Repository.GetElementByID(relation.SupplierID)
				'add the node for the target
				createOrUpdateElementNode xmlDOM, xmlSchema, targetElement, allNodes, processedNodes
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