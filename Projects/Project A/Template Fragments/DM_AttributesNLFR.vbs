'[path=\Projects\Project A\Template fragments]
'[group=Template fragments]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.Util

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
function MyRtfData (objectID, tagname)

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
	set xmlRoot = xmlDOM.createElement( "EADATA" )
	xmlDOM.appendChild xmlRoot

	dim xmlDataSet
	set xmlDataSet = xmlDOM.createElement( "Dataset_0" )
	xmlRoot.appendChild xmlDataSet
	 
	dim xmlData 
	set xmlData = xmlDOM.createElement( "Data" )
	xmlDataSet.appendChild xmlData
	
	'loop the Attributes
	dim element as EA.Element
	set element = Repository.GetElementByID(objectID)
	dim attribute as EA.Attribute

	if element.Attributes.Count > 0 then
		for each attribute in  element.Attributes
			addRow xmlDOM, xmlData, attribute
		next
		
	else
		'no attributes, add N.A row
		addNotApplicableRow xmlDOM, xmlData
	end if
	MyRtfData = xmlDOM.xml
end function

function addRow(xmlDOM, xmlData, attribute)
	
	dim xmlRow
	set xmlRow = xmlDOM.createElement( "Row" )
	xmlData.appendChild xmlRow
	
	'Attribute name
	dim xmlAttributeName
	set xmlAttributeName = xmlDOM.createElement( "AttributeName" )	
	xmlAttributeName.text = attribute.Name
	xmlRow.appendChild xmlAttributeName

	dim descriptionfull
	descriptionfull = attribute.Notes
	
	'description NL
	dim formattedAttr 
	set formattedAttr = xmlDOM.createAttribute("formatted")
	formattedAttr.nodeValue="1"
	dim xmlDescNL
	set xmlDescNL = xmlDOM.createElement( "DescriptionNL" )	
	xmlDescNL.text = getTagContent(descriptionfull, "NL")
'	xmlDescNL.setAttributeNode(formattedAttr)
	xmlRow.appendChild xmlDescNL
	
	'description FR
	set formattedAttr = xmlDOM.createAttribute("formatted")
	formattedAttr.nodeValue="1"
	dim xmlDescFR
	set xmlDescFR = xmlDOM.createElement( "DescriptionFR" )			
	xmlDescFR.text = getTagContent(descriptionfull, "FR")
'	xmlDescFR.setAttributeNode(formattedAttr)
	xmlRow.appendChild xmlDescFR
	
	'multiplicity
	dim xmlMultiplicity
	set xmlMultiplicity = xmlDOM.createElement( "Multiplicity" )			
	xmlMultiplicity.text = attribute.LowerBound & ".." & attribute.UpperBound
	xmlRow.appendChild xmlMultiplicity
	
	'IsID
	dim xmlIsID
	set xmlIsID = xmlDOM.createElement( "IsID" )
	if attribute.IsID then
		xmlIsID.text = "Y"
	else
		xmlIsID.text = "N"
	end if
	xmlRow.appendChild xmlIsID
	
	'Format
	dim xmlFormat
	set xmlFormat = xmlDOM.createElement( "Format" )			
	xmlFormat.text = attribute.Type
	xmlRow.appendChild xmlFormat
	
	'Alias
	dim xmlAlias
	set xmlAlias = xmlDOM.createElement( "Alias" )			
	xmlAlias.text = attribute.Alias
	xmlRow.appendChild xmlAlias
	
end function

function addNotApplicableRow(xmlDOM, xmlData)
	dim xmlRow
	set xmlRow = xmlDOM.createElement( "Row" )
	xmlData.appendChild xmlRow
	
	'Attribute name
	dim xmlAttributeName
	set xmlAttributeName = xmlDOM.createElement( "AttributeName" )	
	xmlAttributeName.text = "N/A"
	xmlRow.appendChild xmlAttributeName
	
	'description NL
	dim formattedAttr 
	set formattedAttr = xmlDOM.createAttribute("formatted")
	formattedAttr.nodeValue="1"
	dim xmlDescNL
	set xmlDescNL = xmlDOM.createElement( "DescriptionNL" )	
	xmlDescNL.text = "Niet van toepassing"
'	xmlDescNL.setAttributeNode(formattedAttr)
	xmlRow.appendChild xmlDescNL
	
	'description FR
	set formattedAttr = xmlDOM.createAttribute("formatted")
	formattedAttr.nodeValue="1"
	dim xmlDescFR
	set xmlDescFR = xmlDOM.createElement( "DescriptionFR" )			
	xmlDescFR.text = "Sans objet"
'	xmlDescFR.setAttributeNode(formattedAttr)
	xmlRow.appendChild xmlDescFR
	
	'multiplicity
	dim xmlMultiplicity
	set xmlMultiplicity = xmlDOM.createElement( "Multiplicity" )			
	xmlMultiplicity.text = ""
	xmlRow.appendChild xmlMultiplicity
	
	'IsID
	dim xmlIsID
	set xmlIsID = xmlDOM.createElement( "IsID" )
	xmlIsID.text = ""
	xmlRow.appendChild xmlIsID
	
	'Format
	dim xmlFormat
	set xmlFormat = xmlDOM.createElement( "Format" )			
	xmlFormat.text = ""
	xmlRow.appendChild xmlFormat
	
	'Alias
	dim xmlAlias
	set xmlAlias = xmlDOM.createElement( "Alias" )			
	xmlAlias.text = ""
	xmlRow.appendChild xmlAlias
end function


'msgbox MyRtfData(38999, "")
'msgbox MyRtfData(52460, "")