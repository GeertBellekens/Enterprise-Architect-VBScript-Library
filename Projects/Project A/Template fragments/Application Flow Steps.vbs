'[path=\Projects\Project A\Template fragments]
'[group=Template fragments]

' Script Name: Application Flow Steps
' Author: Matthias Van der Elst
' Purpose: Lists the steps from the application flow
' Date: 2017-11-09

option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include




'MyRtfData 20037, ""


function MyRtfData(diagramID, tagname)
	dim activityDiagram as EA.Diagram
	set activityDiagram = Repository.GetDiagramByID(diagramID)
	dim sortedDiagramObjects
	set sortedDiagramObjects = sortDiagramObjectsCollection(activityDiagram.DiagramObjects)
	dim diagramObject as EA.DiagramObject
	dim activity as EA.Element
	
	'XML
	dim xmlDOM, xmlRoot, xmlData, xmlDataSet, xmlRow, xmlStep, xmlDesc
	set  xmlDOM = CreateObject( "Microsoft.XMLDOM" )
	xmlDOM.validateOnParse = false
	xmlDOM.async = false
	 
	dim node 
	set node = xmlDOM.createProcessingInstruction( "xml", "version='1.0'")
    xmlDOM.appendChild node

	set xmlRoot = xmlDOM.createElement( "EADATA" )
	xmlDOM.appendChild xmlRoot

	set xmlDataSet = xmlDOM.createElement( "Dataset_0" )
	xmlRoot.appendChild xmlDataSet
	 
	set xmlData = xmlDOM.createElement( "Data" )
	xmlDataSet.appendChild xmlData
	
	for each diagramObject in sortedDiagramObjects
		set activity= Repository.GetElementByID(diagramObject.ElementID)
		if activity.Type = "Activity" then
		
			set xmlRow = xmlDOM.createElement( "Row" )
			xmlData.appendChild xmlRow
			
			set xmlStep = xmlDOM.createElement( "Step" )
			xmlStep.text = activity.Name
			xmlRow.appendChild xmlStep
			
			dim formattedAttr 
			set formattedAttr = xmlDOM.createAttribute("formatted")
			formattedAttr.nodeValue="1"
			set xmlDesc = xmlDOM.createElement( "Description" )
			xmlDesc.text = activity.Notes
			XmlDesc.setAttributeNode(formattedAttr)
			xmlRow.appendChild xmlDesc
			
			'Repository.WriteOutput outPutName, now() & " " & activity.Name & "|" & activity.Notes, 0
		end if
	next
	
	MyRtfData = xmlDOM.xml


end function




function sortDiagramObjectsCollection (diagramObjects)
	dim sortedDiagramObjects 
	dim diagramObject as EA.DiagramObject
	set sortedDiagramObjects = CreateObject("System.Collections.ArrayList")
	for each diagramObject in diagramObjects
		sortedDiagramObjects.Add (diagramObject)
	next
	set sortDiagramObjectsCollection = sortDiagramObjectsArrayList(sortedDiagramObjects)
end function

function sortDiagramObjectsArrayList (diagramObjects)
	dim i
	dim goAgain
	goAgain = false
	dim thisElement as EA.DiagramObject
	dim nextElement as EA.DiagramObject
	for i = 0 to diagramObjects.Count -2 step 1
		set thisElement = diagramObjects(i)
		set nextElement = diagramObjects(i +1)
		if  diagramObjectIsAfterYX(thisElement, nextElement) then
			diagramObjects.RemoveAt(i +1)
			diagramObjects.Insert i, nextElement
			goAgain = true
		end if
	next
	'if we had to swap an element then we go over the list again
	if goAgain then
		set diagramObjects = sortDiagramObjectsArrayList (diagramObjects)
	end if
	'return the sorted list
	set sortDiagramObjectsArrayList = diagramObjects
end function

'returns true if thisElement should come after the nextElement (both diagramObjects)
function diagramObjectIsAfterYX(thisElement, nextElement)
'	dim thisElement as EA.DiagramObject
'	dim nextElement as EA.DiagramObject
	if thisElement.top > nextElement.top then
		diagramObjectIsAfterYX = false
	elseif thisElement.top = nextElement.top then
		if thisElement.left > nextElement.left then
			diagramObjectIsAfterYX = true
		else
			diagramObjectIsAfterYX = false
		end if
	else 
		diagramObjectIsAfterYX = true
	end if
end function