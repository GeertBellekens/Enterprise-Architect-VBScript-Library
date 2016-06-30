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
	'msgbox "starting MyRTFdata fo element" & objectID
	'MyRtfData = "<?xml version=""1.0""?><EADATA><Dataset_0><Data><Row><ConstraintName>constraint 1</ConstraintName><DescriptionNL formatted=""1"">Hier staat de nederlandse bescrijving van deze constraint in &lt;b&gt;RichText&lt;/b&gt; (1)</DescriptionNL><DescriptionFR formatted=""1"">ici le français avec éàè characters &lt;b&gt;dôme&lt;/b&gt;</DescriptionFR></Row><Row><ConstraintName>constraint 2</ConstraintName><DescriptionNL formatted=""1"">Hier staat de nederlandse bescrijving van deze constraint in &lt;b&gt;RichText (2)&lt;/b&gt;</DescriptionNL><DescriptionFR formatted=""1"">ici le français avec éàè characters &lt;b&gt;dôme (2)&lt;/b&gt;</DescriptionFR></Row></Data></Dataset_0></EADATA>"
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
	
	'loop the constraints
	dim element as EA.Element
	set element = Repository.GetElementByID(objectID)
	dim constraint as EA.Constraint
	
	if element.Constraints.Count > 0 then
		for each constraint in  element.Constraints
			addRow xmlDOM, xmlData, constraint
		next
	else
		'no constraints, add N.A row
		addNotApplicableRow xmlDOM, xmlData
	end if
	MyRtfData = xmlDOM.xml
end function

function addRow(xmlDOM, xmlData, constraint)
	
	dim xmlRow
	set xmlRow = xmlDOM.createElement( "Row" )
	xmlData.appendChild xmlRow
	
	'constraint names are the same in NL/FR, except when adding an empty row
	dim xmlConstraintNameNL
	set xmlConstraintNameNL = xmlDOM.createElement( "ConstraintNameNL" )	
	xmlConstraintNameNL.text = constraint.Name
	xmlRow.appendChild xmlConstraintNameNL
	
	dim xmlConstraintNameFR
	set xmlConstraintNameFR = xmlDOM.createElement( "ConstraintNameFR" )	
	xmlConstraintNameFR.text = constraint.Name
	xmlRow.appendChild xmlConstraintNameFR
	
	dim descriptionfull
	descriptionfull = constraint.Notes
	
	dim formattedAttr 
	set formattedAttr = xmlDOM.createAttribute("formatted")
	formattedAttr.nodeValue="1"
	dim xmlDescNL
	set xmlDescNL = xmlDOM.createElement( "DescriptionNL" )	

	xmlDescNL.text = getTagContent(descriptionfull, "NL")
	xmlDescNL.setAttributeNode(formattedAttr)
	xmlRow.appendChild xmlDescNL
	
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

	'constraint names are the same in NL/FR, except when adding an empty row
	dim xmlConstraintNameNL
	set xmlConstraintNameNL = xmlDOM.createElement( "ConstraintNameNL" )	
	xmlConstraintNameNL.text = "Niet van toepassing"
	xmlRow.appendChild xmlConstraintNameNL
	
	dim xmlConstraintNameFR
	set xmlConstraintNameFR = xmlDOM.createElement( "ConstraintNameFR" )	
	xmlConstraintNameFR.text = "Sans objet"
	xmlRow.appendChild xmlConstraintNameFR
	
	'add empty tags for DescriptionNL/FR because otherwise the tag name is shown
	dim xmlDescNL
	set xmlDescNL = xmlDOM.createElement( "DescriptionNL" )	
	xmlRow.appendChild xmlDescNL
	
	dim xmlDescFR
	set xmlDescFR = xmlDOM.createElement( "DescriptionFR" )			
	xmlRow.appendChild xmlDescFR
end function


'msgbox MyRtfData(38700, "")