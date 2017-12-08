'[path=\Projects\Project A\Template fragments]
'[group=Template fragments]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.Util

'
' Script Name: MA_SolutionRequirements
' Author: Geert Bellekens
' Purpose: Returns the solution requirements refernenced by elements owned by this M&A Specification
' Date: 2016-05-18
'
function MyRtfData (objectID)

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
	

	'get all BAM Elements linked to an element owned by the M&A element
	dim sqlGetRequirements
	sqlGetRequirements = "select o.Object_ID from t_object o where o.Stereotype like 'BAM%' and  o.parentID =" & objectID
	
	dim requirements
	set requirements = getElementsFromQuery(sqlGetRequirements)
	dim requirement as EA.Element
	'then loop requirements and add the rows
	for each requirement in requirements
		addRow xmlDOM, xmlData, requirement
	next
	
	MyRtfData = xmlDOM.xml
end function

function addRow(xmlDOM, xmlData, requirement)
	
	dim xmlRow
	set xmlRow = xmlDOM.createElement( "Row" )
	xmlData.appendChild xmlRow
	
	'RequirementID
	dim xmlRequirementID
	set xmlRequirementID = xmlDOM.createElement( "RequirementID" )	
	xmlRequirementID.text = requirement.Name
	xmlRow.appendChild xmlRequirementID
	
	dim definitionFull
	definitionFull = getTagContent(requirement.Notes, "definition") 
	'TitleNL
	set formattedAttr = xmlDOM.createAttribute("formatted")
	formattedAttr.nodeValue="1"
	dim xmlTitleNL
	set xmlTitleNL = xmlDOM.createElement( "TitleNL" )	
	xmlTitleNL.text = getTagContent(definitionFull, "NL")
	xmlTitleNL.setAttributeNode(formattedAttr)
	xmlRow.appendChild xmlTitleNL
	
	'TitleFR
	set formattedAttr = xmlDOM.createAttribute("formatted")
	formattedAttr.nodeValue="1"
	dim xmlTitleFR
	set xmlTitleFR = xmlDOM.createElement( "TitleFR" )	
	xmlTitleFR.text = getTagContent(definitionFull, "FR")
	xmlTitleFR.setAttributeNode(formattedAttr)
	xmlRow.appendChild xmlTitleFR
	
	dim descriptionfull
	descriptionfull = getTagContent(requirement.Notes, "description") 
	
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


'msgbox MyRtfData(233339)