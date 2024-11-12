'[path=\Projects\Project A\Template fragments]
'[group=Template fragments]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: DiagramLink
' Author: Geert Bellekens
' Purpose: get the hyperlink for a diagram
' Date: 2019-12-02
'

function getObjectDiagramLink(objectID, baseUrl)
 dim element
 set element = Repository.GetElementByID(objectID)
 'get diagram
 dim diagram 
 set diagram = getFirstDiagram(element)
 'get the xml data
 getObjectDiagramLink = getRTFData(diagram, baseUrl)
end function

function getPackageDiagramLink (packageID, baseUrl)
 dim package as EA.Package
 set package = Repository.GetPackageByID(packageID)
 'get diagram
 dim diagram 
 set diagram = getFirstDiagram(package)
 'get the xml data
 getPackageDiagramLink = getRTFData(diagram, baseUrl)
end function

function getRTFData(diagram, baseUrl)
 
 dim diagramUrl
 diagramUrl = getDiagramLink(diagram, baseUrl)
 
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
  
 dim formattedAttr 
 set formattedAttr = xmlDOM.createAttribute("formatted")
 formattedAttr.nodeValue="1"
 
 dim xmldiagramLink
 set xmldiagramLink = xmlDOM.createElement( "DiagramLink" ) 

 xmldiagramLink.text = "<a href=""$inet://" & diagramUrl &"""><font color=""#0000ff""><u>" & diagram.Name & "</u></font></a>"
 xmldiagramLink.setAttributeNode(formattedAttr)
 xmlRow.appendChild xmldiagramLink
 
 getRTFData = xmlDOM.xml
end function

function getFirstDiagram(diagramOwner)
 set getFirstDiagram = nothing 'initialize
 'get first diagram
 dim diagram as EA.Diagram
 for each diagram in diagramOwner.Diagrams
  set getFirstDiagram = diagram
  exit function
 next
end function

function getDiagramLink(diagram, baseUrl)
 dim link
 link = baseUrl & "/webea?m=1&o="
 'get diagram GUID
 link = link & diagram.DiagramGUID
 'remove braces
 link = replace(link, "{","")
 link = replace(link, "}","")
 'return 
 getDiagramLink = link
end function

'msgbox MyPackageRtfData(3357,"")
function test
 dim outputString
 dim fileSystemObject
 dim outputFile
 
 'outputString =  getPackageDiagramLink (1783, "http://omnibus2.hampden.local")
 outputString =  getObjectDiagramLink (77915, "http://omnibus2.hampden.local")
 
 set fileSystemObject = CreateObject( "Scripting.FileSystemObject" )
 set outputFile = fileSystemObject.CreateTextFile( "H:\Temp\diagramLink.xml", true )
 outputFile.Write outputString
 outputFile.Close
end function 

'test