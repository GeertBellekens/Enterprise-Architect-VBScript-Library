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
	'set  xmlDOM = CreateObject( "MSXML2.DOMDocument.4.0" )
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
	 
	dim activities
	set activities = getProcessSpecificActivities(objectID)
	dim activity as EA.Element
	for each activity in activities
		'Add the rows here
		'check the notes of the activty
		if len (activity.Notes) = 0 then
			set activity = getCallingProcess(activity)
		end if
		addRow xmlData, xmlDOM, activity
	next
	MyRtfData = xmlDOM.xml
end function

function getCallingProcess (element)
	set getCallingProcess = element
	if element.Type = "Activity" AND element.Stereotype = "Activity" then
		dim calledActivityTV as EA.TaggedValue
		set calledActivityTV = element.TaggedValues.GetByName("isACalledActivity")
		dim referenceActivityTV as EA.TaggedValue
		set referenceActivityTV = element.TaggedValues.GetByName("calledActivityRef")
		if not calledActivityTV is nothing and not referenceActivityTV is nothing then
			'only do something when the Activity is types a CalledActivity
			'Session.Output "calledActivityTV.Value : " & calledActivityTV.Value 
			'Session.Output "referenceActivityTV.Value :" & referenceActivityTV.Value
			if calledActivityTV.Value = "true" then
				dim calledActivity as EA.Element
				set calledActivity = Repository.GetElementByGuid(referenceActivityTV.Value)
				if not calledActivity is nothing then
					set getCallingProcess = calledActivity
				end if
			end if
		end if
	end if
end function

function getProcessSpecificActivities(objectID)
	dim businessProcess as EA.Element
	set businessProcess = Repository.GetElementByID(objectID)
	dim activities
	set activities = CreateObject("System.Collections.ArrayList")
	if not businessProcess.CompositeDiagram is nothing then
		set activities = getSubProcesses(businessProcess.CompositeDiagram)
	end if
	set getProcessSpecificActivities = activities
end function

function addRow (xmlData, xmlDOM, element)

	dim xmlRow
	set xmlRow = xmlDOM.createElement( "Row" )
	xmlData.appendChild xmlRow
	
	'name
	dim xmlActivityName
	set xmlActivityName = xmlDOM.createElement( "ActivityName" )			
	xmlActivityName.text = element.Name
	xmlRow.appendChild xmlActivityName
	
	'notes
	dim xmlActivityNotes
	set xmlActivityNotes = xmlDOM.createElement("Notes")
	xmlActivityNotes.text = element.Notes
	XmlRow.appendChild xmlActivityNotes
	
	'description
	dim descriptionfull
	descriptionfull = getTagContent(element.Notes, "definition")
	
	dim formattedAttr 
	
	'description NL
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



function getSubProcesses(diagram)
	dim sortedDiagramObjects
	dim sortedSubProcesses
	set sortedSubProcesses = CreateObject("System.Collections.ArrayList")
	set sortedDiagramObjects = sortDiagramObjectsCollection(diagram.DiagramObjects)
	dim diagramObject as EA.DiagramObject
	dim subProcess as EA.Element
	for each diagramObject in sortedDiagramObjects
		set subProcess = Repository.GetElementByID(diagramObject.ElementID)
		if subProcess.Stereotype = "ArchiMate_BusinessProcess" _
			OR subProcess.Stereotype = "Activity" _
			OR (subProcess.Stereotype = "IntermediateEvent" AND len(subProcess.Notes) > 0) _
			OR (subProcess.Stereotype = "StartEvent" AND len(subProcess.Notes) > 0) _
			OR (subProcess.Stereotype = "EndEvent" AND len(subProcess.Notes) > 0) then
			sortedSubProcesses.Add subProcess
		end if
	next
	set getSubProcesses = sortedSubProcesses
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

'returns true if thisElement should come after the nextElement (both diagramObjects)
function diagramObjectIsAfterXY(thisElement, nextElement)
'	dim thisElement as EA.DiagramObject
'	dim nextElement as EA.DiagramObject
	if thisElement.left > nextElement.left then
		diagramObjectIsAfterXY = true
	elseif thisElement.left = nextElement.left then
		if thisElement.top > nextElement.top then
			diagramObjectIsAfterXY = true
		else
			diagramObjectIsAfterXY = false
		end if
	else 
		diagramObjectIsAfterXY = false
	end if
end function

'msgbox MyRtfData(213457)