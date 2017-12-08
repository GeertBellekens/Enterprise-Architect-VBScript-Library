'[path=\Projects\Project A\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.Util

'
' Script Name: SetPackageToSchema
' Author: Geert Bellekens
' Purpose: Add the contents of the selected package to the selected schema composer artifact
' Date: 2016-06-01
'
const outPutName = "Create Schema From Package"

function Main ()
	
	dim selectedElement as EA.Element
	set selectedElement = Repository.GetTreeSelectedObject()
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
		
		
		'select package
		dim selectedPackage
		set selectedPackage = selectPackage()
		
		if not selectedPackage is nothing then
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
				xmlSchema.appendChild createElementNode(xmlDom, element)
			next
			
			dim sqlUpdateSchema
			sqlUpdateSchema = "update t_document set StrContent = N'" & xmlDOM.xml & "' " & _
							" where ElementType = 'SC_MessageProfile' " & _
							" and ElementID = '" & selectedElement.ElementGUID & "'"
			Repository.Execute sqlUpdateSchema
			'update log
			Repository.WriteOutput outPutName, "Finished Create Schema From Package at " & now(), 0
		end if 
		'writefile "c:\\temp\\schemaContents.xml", slqUpdateSchema
		'main = xmlDOM.xml
	end if
end function

function createElementNode(xmlDOM, element)
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
	xmlguidAtttr.nodeValue = element.ElementGUID
	xmlClass.setAttributeNode(xmlguidAtttr)
	
	'add propertiesnode
	dim xmlProperties
	set xmlProperties= xmlDOM.createElement("properties")
	
	'add attributes
	dim attribute as EA.Attribute
	for each attribute in element.Attributes
		xmlProperties.appendChild createPropertyNode (xmlDOM, attribute.AttributeGUID, "attribute")
	next
	
	'add associations only if they start at the given element
	dim association as EA.Connector
	for each association in element.Connectors
		if (association.Type = "Association" _
		   or association.Type = "Aggregation") _
		    and association.ClientID = element.ElementID then
			xmlProperties.appendChild createPropertyNode (xmlDOM, association.ConnectorGUID, "association")
		end if
	next
	
	'add xmlProperties to class node
	xmlClass.appendChild xmlProperties
	
	'return node
	set createElementNode = xmlClass
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
