'[path=\Projects\Project A\Template fragments]
'[group=Template fragments]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
function MyRtfData (objectID, tagname)
	
	dim xmlDOM 
	set  xmlDOM = CreateObject( "Microsoft.XMLDOM" )
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
	 
	dim xmlRow
	set xmlRow = xmlDOM.createElement( "Row" )
	xmlData.appendChild xmlRow
	
	dim element as EA.Element
	set element = Repository.GetElementByID(objectID)
	
	dim formattedAttr 
	set formattedAttr = xmlDOM.createAttribute("formatted")
	formattedAttr.nodeValue="1"
	
	dim descriptionfull
	descriptionfull = getTagContent(element.Notes, tagname)
	
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
		
	MyRtfData = xmlDOM.xml
end function

'get the description from the given notes 
'that is the text between <NL> and </NL> or <FR> and </FR>
function getTagContent(notes, tag)
	if tag = "" then
		getTagContent = notes
	else
		getTagContent = ""
		dim startTagPosition
		dim endTagPosition
		startTagPosition = InStr(notes,"&lt;" & tag & "&gt;")
		endTagPosition = InStr(notes,"&lt;/" & tag & "&gt;")
		'Session.Output "notes: " & notes & " startTagPosition: " & startTagPosition & " endTagPosition: " &endTagPosition
		if startTagPosition > 0 and endTagPosition > startTagPosition then
			dim startContent
			startContent = startTagPosition + len(tag) + 8
			dim length 
			length = endTagPosition - startContent
			getTagContent = mid(notes, startContent, length)
		end if
	end if 
end function

'msgbox MyRtfData(18314, "")