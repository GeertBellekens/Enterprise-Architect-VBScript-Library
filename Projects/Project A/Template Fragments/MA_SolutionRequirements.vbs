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
	
	'get the M&A specification element
	dim element as EA.Element
	set element = Repository.GetElementByID(objectID)
	'get all solution requirements linked to an element owned by the M&A element
	dim sqlGetRequirements
	sqlGetRequirements = "select distinct rq.Object_ID                                         " & _
						" from (((t_object o                                                   " & _
						" inner join t_object so on so.ParentID = o.Object_ID)                 " & _
						" inner join t_connector soc on soc.Start_Object_ID = so.Object_ID)    " & _
						" inner join t_object rq on (soc.End_Object_ID = rq.Object_ID          " & _
						" 						and rq.Object_Type = 'Requirement'             " & _
						" 						and rq.Stereotype = 'Solution Requirement'))   " & _
						" where o.ea_guid = '" & element.ElementGUID & "'"
	
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
	
	'TitleNL
	dim xmlTitleNL
	set xmlTitleNL = xmlDOM.createElement( "TitleNL" )	
	xmlTitleNL.text = getTaggedValueValue(requirement, "Title NL")
	xmlRow.appendChild xmlTitleNL
	
	'TitleFR
	dim xmlTitleFR
	set xmlTitleFR = xmlDOM.createElement( "TitleFR" )	
	xmlTitleFR.text = getTaggedValueValue(requirement, "Title FR")
	xmlRow.appendChild xmlTitleFR
	
	dim descriptionfull
	descriptionfull = requirement.Notes
	
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

function addNotApplicableRow(xmlDOM, xmlData)
	dim xmlRow
	set xmlRow = xmlDOM.createElement( "Row" )
	xmlData.appendChild xmlRow
	
	'source multiplicity
	dim xmlSMultiplicity
	set xmlSMultiplicity = xmlDOM.createElement( "SMultiplicity" )	
	xmlSMultiplicity.text = ""
	xmlRow.appendChild xmlSMultiplicity
	
	'target multiplicity
	dim xmlTMultiplicity
	set xmlTMultiplicity = xmlDOM.createElement( "TMultiplicity" )	
	xmlTMultiplicity.text = ""
	xmlRow.appendChild xmlTMultiplicity
	
	'source Name
	dim xmlSource
	set xmlSource = xmlDOM.createElement( "Source" )	
	xmlSource.text = "N/A"
	xmlRow.appendChild xmlSource
	
	'target Name
	dim xmlTarget
	set xmlTarget = xmlDOM.createElement( "Target" )	
	xmlTarget.text = "N/A"
	xmlRow.appendChild xmlTarget
	
	'ConnectorName
	dim xmlConnectorName
	set xmlConnectorName = xmlDOM.createElement( "ConnectorName" )	
	xmlConnectorName.text = "N/A"
	xmlRow.appendChild xmlConnectorName

	'description NL
	dim formattedAttr 
	set formattedAttr = xmlDOM.createAttribute("formatted")
	formattedAttr.nodeValue="1"
	dim xmlDescNL
	set xmlDescNL = xmlDOM.createElement( "DescriptionNL" )	
	xmlDescNL.text = "n.v.t."
	xmlDescNL.setAttributeNode(formattedAttr)
	xmlRow.appendChild xmlDescNL
	
	'description FR
	set formattedAttr = xmlDOM.createAttribute("formatted")
	formattedAttr.nodeValue="1"
	dim xmlDescFR
	set xmlDescFR = xmlDOM.createElement( "DescriptionFR" )			
	xmlDescFR.text = "S.O."
	xmlDescFR.setAttributeNode(formattedAttr)
	xmlRow.appendChild xmlDescFR
		
end function

'msgbox MyRtfData(83755)