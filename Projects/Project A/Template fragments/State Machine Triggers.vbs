'[path=\Projects\Project A\Template fragments]
'[group=Template fragments]

' Script Name: State Machine Triggers
' Author: Matthias Van der Elst
' Purpose: Lists the triggers for a specific diagram
' Date: 2017-10-31

option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'const outPutName = "State Machine Diagram"   
'Repository.CreateOutputTab outPutName
'Repository.ClearOutput outPutName
'Repository.EnsureOutputVisible outPutName
dim xmlDOM, xmlRoot, xmlData, xmlDataSet, xmlRow

MyRtfData 5209, ""

function MyRtfData(diagramID, tagname)
	'XML
	set  xmlDOM = CreateObject( "Microsoft.XMLDOM" )
	'set  xmlDOM = CreateObject( "MSXML2.DOMDocument.4.0" )
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
	 
	
	dim stateMachine as EA.Diagram
	set stateMachine = Repository.GetDiagramByID(diagramID)
	
	'Repository.WriteOutput outPutName, now() & " Diagram: '"& stateMachine.Name &"'", 0
	
	dim sortedDiagramObjects
	dim sortedStates
	set sortedStates = CreateObject("System.Collections.ArrayList")
	set sortedDiagramObjects = sortDiagramObjectsCollection(stateMachine.DiagramObjects)
	dim diagramObject as EA.DiagramObject
	dim state as EA.Element
	
	for each diagramObject in sortedDiagramObjects
		set state = Repository.GetElementByID(diagramObject.ElementID)
		if state.Type = "State" or state.Type = "StateNode" then
			sortedStates.Add state
			dim transition as EA.Connector
			for each transition in state.Connectors
				if transition.ClientID = state.ElementID then
					'get the trigger objects
					getTriggers state,transition
				end if
			next
		end if
	next
	
MyRtfData = xmlDOM.xml
end function

function getTriggers(source, transition)
	'transition guid: {13132253-FD87-4056-B3F5-67ED089F957E}
	'Repository.WriteOutput outPutName, now() & " " & source.Name & " | " & transition.Name, 0
	dim target as EA.Element
	set target = Repository.GetElementByID(transition.SupplierID)
	dim sqlGetTriggers
	sqlGetTriggers = "select distinct o.Object_ID " & _ 
					 "from t_xref x " & _
					 "inner join t_object o " & _
					 "on x.Description like '%'+o.ea_guid+'%' " & _
					 "where x.Client = '" & transition.ConnectorGUID & "'"
	
	dim triggers
	set triggers = getElementsFromQuery(sqlGetTriggers)
	dim trigger as EA.Element
	
	'XML
	dim xmlSource, xmlTarget, xmlTrigger, xmlSpecification, xmlGuard 
	
	if triggers.Count = 0 then
		set xmlRow = xmlDOM.createElement( "Row" )
		xmlData.appendChild xmlRow	
		
		set xmlSource = xmlDOM.createElement( "Source" )
		xmlSource.text = source.Name
		xmlRow.appendChild xmlSource
		
		set xmlTarget = xmlDOM.createElement( "Target" )
		xmlTarget.text = target.Name
		xmlRow.appendChild xmlTarget
		
		set xmlTrigger = xmlDOM.createElement( "Trigger" )
		xmlTrigger.text = ""
		xmlRow.appendChild xmlTrigger
		
		set xmlSpecification = xmlDOM.createElement( "Specification" )
		xmlSpecification.text = ""
		xmlRow.appendChild xmlSpecification	
		
		if len(transition.TransitionGuard) > 0 then
			set xmlGuard = xmlDOM.createElement( "Guard" )
			xmlGuard.text = transition.TransitionGuard
			xmlRow.appendChild xmlGuard
		else
			set xmlGuard = xmlDOM.createElement( "Guard" )
			xmlGuard.text = ""
			xmlRow.appendChild xmlGuard
		end if
		
	end if
	
	for each trigger in triggers

		dim sqlGetSpecification
		sqlGetSpecification =  "select x.Description " & _ 
								   "from t_xref x " & _ 
								   "where x.Client = '" & trigger.ElementGUID & "'" & _
								   "and x.Behavior = 'event'"
		dim specc
		specc = getArrayFromQuery(sqlGetSpecification)
		dim spec
		spec = Split(specc(0,0),";", -1, 1)
		dim specification
		specification = Split(spec(1),"Name=",-1,1)
		'Repository.WriteOutput outPutName, now() & " " & source.Name & " | " & target.Name & " | " & trigger.Name & " | " & specification(1) & " | " & transition.TransitionGuard, 0
	
		'XML
		set xmlRow = xmlDOM.createElement( "Row" )
		xmlData.appendChild xmlRow
		
		set xmlSource = xmlDOM.createElement( "Source" )
		xmlSource.text = source.Name
		xmlRow.appendChild xmlSource
		
		set xmlTarget = xmlDOM.createElement( "Target" )
		xmlTarget.text = target.Name
		xmlRow.appendChild xmlTarget
		
		set xmlTrigger = xmlDOM.createElement( "Trigger" )
		xmlTrigger.text = trigger.Name
		xmlRow.appendChild xmlTrigger
		
		if len(specification(1)) > 0 then
			set xmlSpecification = xmlDOM.createElement( "Specification" )
			xmlSpecification.text = specification(1)
			xmlRow.appendChild xmlSpecification
		else 
			set xmlSpecification = xmlDOM.createElement( "Specification" )
			xmlSpecification.text = ""
			xmlRow.appendChild xmlSpecification
		end if

		if len(transition.TransitionGuard) > 0 then
			set xmlGuard = xmlDOM.createElement( "Guard" )
			xmlGuard.text = transition.TransitionGuard
			xmlRow.appendChild xmlGuard
		else
			set xmlGuard = xmlDOM.createElement( "Guard" )
			xmlGuard.text = ""
			xmlRow.appendChild xmlGuard
		end if
	next
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