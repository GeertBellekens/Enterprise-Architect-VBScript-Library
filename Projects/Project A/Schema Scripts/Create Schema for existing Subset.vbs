'[path=\Projects\Project A\Schema Scripts]
'[group=Schema Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.Util

'
' Script Name: Create Schema for existing Subset
' Author: Geert Bellekens
' Purpose: Create a profile for the currently selected as if the selected package was generated with this schema
' Date: 2019-02-12
'
const outPutName = "Create Schema From Package"

function Main ()
	
		'select subset package
	dim selectedPackage
	set selectedPackage = Repository.GetTreeSelectedPackage()
	dim selectedElement as EA.Element
	set selectedElement = createArtifact(selectedPackage)
	
	if not selectedElement is Nothing AND selectedElement.Type = "Artifact" THEN
		
		'create output tab
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		'set timestamp
		Repository.WriteOutput outPutName, "Starting Create Schema From Package at " & now(), 0
		
		dim xmlDOM 
		set  xmlDOM = CreateObject( "Microsoft.XMLDOM" )
		'set  xmlDOM = CreateObject( "MSXML2.DOMDocument.4.0" )
		xmlDOM.validateOnParse = false
		xmlDOM.async = false
		 
		dim node 
		set node = xmlDOM.createProcessingInstruction( "xml", "version='1.0'")
		xmlDOM.appendChild node
	'
		dim xmlRoot 
		set xmlRoot = xmlDOM.createElement( "message" )
		xmlDOM.appendChild xmlRoot
		'add description node
		xmlRoot.appendChild createDescriptionNode(xmlDOM, selectedElement)
		
		dim packageTree 
		set packageTree = getPackageTree(selectedPackage)
		dim sqlGetClasses
		sqlGetClasses = "select o.Object_ID from t_object o " & _
						"where o.Object_Type in ('Class', 'DataType', 'Enumeration') " & _
						"and o.Package_ID in (" & makePackageIDString(packageTree) & ") order by o.Name"
		dim allElements
		set allElements = getElementsFromQuery(sqlGetClasses)
		'add schema node
		dim xmlSchema 
		set xmlSchema = xmlDOM.createElement("schema")
		'count attribute
		dim xmlcountAtttr 
		set xmlcountAtttr = xmlDOM.createAttribute("count")
		xmlcountAtttr.nodeValue = allElements.Count
		xmlSchema.setAttributeNode(xmlcountAtttr)
		'add the schema node to the root
		xmlRoot.appendChild xmlSchema
		'add all the elements to the schema
		dim element as EA.Element
		for each element in allElements
			'update log
			Repository.WriteOutput outPutName, "Processing element " & element.Name, 0
			'add node
			dim childNode
			set childNode = createElementNode(xmlDom, element)
			if not childNode is nothing then
				xmlSchema.appendChild childNode
			end if
		next
		'crete the t_document row
		createTDocument selectedElement, xmlDOM.xml
		
		'update log
		Repository.WriteOutput outPutName, "Finished Create Schema From Package at " & now(), 0
		
		'writefile "c:\\temp\\schemaContents.xml", slqUpdateSchema
		'main = xmlDOM.xml
	end if
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

function getOriginalElement(element)
	dim originalElement as EA.Element
	set originalElement = nothing
	dim connector as EA.Connector
	'get the element that is connected with a trace and has the same name
	for each connector in element.Connectors
		if connector.Type = "Abstraction" and connector.Stereotype = "trace" then
			set originalElement = Repository.GetElementByID(connector.SupplierID)
			'stop if we foudn the right one
			if originalElement.Name = element.Name then
				exit for
			end if
		end if
	next
	set getOriginalElement =  originalElement
end function

function createElementNode(xmlDOM, element)
	dim xmlClass
	set xmlClass = xmlDOM.createElement( "class" )
	'get original element
	dim originalElement as EA.Element
	set originalElement = getOriginalElement(element)
	if originalElement is nothing then
		'report error
		Repository.WriteOutput outPutName, now() & " ERROR: could not find original element for '" & element.Name & "'" , element.ElementID
		'don't bother to continue
		set createElementNode = nothing
		exit function
	end if
	
	'name attribute
	dim xmlNameAtttr 
	set xmlNameAtttr = xmlDOM.createAttribute("name")
	xmlNameAtttr.nodeValue = originalElement.Name
	xmlClass.setAttributeNode(xmlNameAtttr)
	
	'guid attribute
	dim xmlguidAtttr 
	set xmlguidAtttr = xmlDOM.createAttribute("guid")
	xmlguidAtttr.nodeValue = originalElement.ElementGUID
	xmlClass.setAttributeNode(xmlguidAtttr)
	
	'add propertiesnode
	dim xmlProperties
	set xmlProperties= xmlDOM.createElement("properties")
	
	'add attributes
	dim attribute as EA.Attribute
	for each attribute in element.Attributes
		'get original attribute
		dim originalAttribute as EA.Attribute
		set originalAttribute = getOriginalAttribute(attribute)
		if not originalAttribute is nothing then
			xmlProperties.appendChild createPropertyNode (xmlDOM, originalAttribute.AttributeGUID, "attribute")
		else
			'report error
			Repository.WriteOutput outPutName, now() & " ERROR: could not find original attribute for '" & element.Name & "." &  attribute.Name & "'" , element.ElementID
		end if
	next
	
	'add associations only if they start at the given element
	dim association as EA.Connector
	for each association in element.Connectors
		if (association.Type = "Association" _
		   or association.Type = "Aggregation") _
		    and association.ClientID = element.ElementID then
			dim originalAssociation as EA.Connector
			set originalAssociation = getOriginalAssociation(association)
			if not originalAssociation is nothing then
				xmlProperties.appendChild createPropertyNode (xmlDOM, originalAssociation.ConnectorGUID, "association")
			else
				'report error
				Repository.WriteOutput outPutName, now() & " ERROR: could not find original association for '" & element.Name & "." &  association.Name & "' with guid " & association.ConnectorGUID , element.ElementID
			end if
		end if
	next
	
	'add xmlProperties to class node
	xmlClass.appendChild xmlProperties
	
	'return node
	set createElementNode = xmlClass
end function

function getOriginalAttribute(attribute)
	dim originalAttribute as EA.Attribute
	dim tv as EA.TaggedValue
	set originalAttribute = nothing
	for each tv in attribute.TaggedValues
		if tv.Name = "sourceAttribute" and len(tv.Value) > 0 then
			set originalAttribute = Repository.GetAttributeByGuid(tv.Value)
			exit for
		end if
	next
	'return
	set getOriginalAttribute = originalAttribute
end function

function getOriginalAssociation(association)
	dim originalAssociation as EA.Connector
	dim tv as EA.TaggedValue
	set originalAssociation = nothing
	for each tv in association.TaggedValues
		if tv.Name = "sourceAssociation" and len(tv.Value) > 0 then
			set originalAssociation = Repository.GetConnectorByGuid(tv.Value)
			exit for
		end if
	next
	'return
	set getOriginalAssociation = originalAssociation
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
	xmlversionAtttr.nodeValue = "12.1.1230.1230"
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

'msgbox MyPackageRtfData(3357,"")
function writefile(filename, contents)
	dim fileSystemObject
	dim outputFile
		
	set fileSystemObject = CreateObject( "Scripting.FileSystemObject" )
	set outputFile = fileSystemObject.CreateTextFile(filename, true )
	outputFile.Write contents
	outputFile.Close
end function 

'test
main