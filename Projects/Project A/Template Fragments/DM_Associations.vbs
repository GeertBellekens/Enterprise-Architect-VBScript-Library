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
function MyRtfData (objectID)

	dim xmlDOM 
	'set  xmlDOM = CreateObject( "Microsoft.XMLDOM" )
	set  xmlDOM = CreateObject( "MSXML2.DOMDocument.4.0" )
	
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
	
	'loop the connectors
	dim element as EA.Element
	set element = Repository.GetElementByID(objectID)
	dim connector as EA.Connector

	if element.Connectors.Count > 0 then
		for each connector in  element.Connectors
			addRow xmlDOM, xmlData, connector
		next
		MyRtfData = xmlDOM.xml
	else
		'no connectors, so return empty string
		MyRtfData = ""
	end if
end function

function addRow(xmlDOM, xmlData, connector)
	
	dim xmlRow
	set xmlRow = xmlDOM.createElement( "Row" )
	xmlData.appendChild xmlRow
	
	'source multiplicity
	dim xmlSMultiplicity
	set xmlSMultiplicity = xmlDOM.createElement( "SMultiplicity" )	
	xmlSMultiplicity.text = connector.ClientEnd.Cardinality
	xmlRow.appendChild xmlSMultiplicity
	
	'target multiplicity
	dim xmlTMultiplicity
	set xmlTMultiplicity = xmlDOM.createElement( "TMultiplicity" )	
	xmlTMultiplicity.text = connector.SupplierEnd.Cardinality
	xmlRow.appendChild xmlTMultiplicity
	
	'source Name
	dim xmlSource
	dim sourceElement as EA.Element
	set sourceElement = Repository.GetElementByID(connector.ClientID)
	set xmlSource = xmlDOM.createElement( "Source" )	
	if not sourceElement is nothing then
		xmlSource.text = sourceElement.Name
	end if
	xmlRow.appendChild xmlSource
	
	'target Name
	dim xmlTarget
	dim targetElement as EA.Element
	set targetElement = Repository.GetElementByID(connector.SupplierID)
	set xmlTarget = xmlDOM.createElement( "Target" )	
	if not targetElement is nothing then
		xmlTarget.text = targetElement.Name
	end if
	xmlRow.appendChild xmlTarget
	
	'ConnectorName
	dim xmlConnectorName
	set xmlConnectorName = xmlDOM.createElement( "ConnectorName" )	
	xmlConnectorName.text = connector.Name
	xmlRow.appendChild xmlConnectorName

	dim descriptionfull
	descriptionfull = connector.Notes
	
	'description NL
	dim formattedAttr 
	set formattedAttr = xmlDOM.createAttribute("formatted")
	formattedAttr.nodeValue="1"
	dim xmlDescNL
	set xmlDescNL = xmlDOM.createElement( "DescriptionNL" )	
	xmlDescNL.text = getTagContent(descriptionfull, "NL")
	xmlDescNL.setAttributeNode(formattedAttr)
	xmlRow.appendChild xmlDescNL
	
	'description FR
	set formattedAttr = xmlDOM.createAttribute("formatted")
	formattedAttr.nodeValue="1"
	dim xmlDescFR
	set xmlDescFR = xmlDOM.createElement( "DescriptionFR" )			
	xmlDescFR.text = getTagContent(descriptionfull, "FR")
	xmlDescFR.setAttributeNode(formattedAttr)
	xmlRow.appendChild xmlDescFR
	
end function

'msgbox MyRtfData(38700, true)