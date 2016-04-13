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
function MyPackageRtfData(packageID, tagname)
	dim package as EA.Package
	dim element as EA.Element
	set package = Repository.GetPackageByID(packageID)
	if not package is nothing then
		set element = Repository.GetElementByGuid(package.PackageGUID)
		if not element is nothing then
			MyPackageRtfData = MyRtfData (element.ElementID, tagname)
		end if
	end if 
end function

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

'msgbox MyPackageRtfData(3357,"")
function test
	dim outputString
	dim fileSystemObject
	dim outputFile
	
	outputString =  MyRtfData(62899, "definition")
	
	set fileSystemObject = CreateObject( "Scripting.FileSystemObject" )
	set outputFile = fileSystemObject.CreateTextFile( "c:\\temp\\NLFRtest.xml", true )
	outputFile.Write outputString
	outputFile.Close
end function 

'test