'[path=\Projects\Project A\Element Group]
'[group=Element Group]
'option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.Util

'
' Script Name: CreateMessageOverviewSequence
' Author: Geert Bellekens
' Purpose: Creates a sequence diagram under the selected businessprocess that shows the sequence of the messages
' from this businessprocess and it's subprocesses
' Date: 31/03/2015
'
'**********EDIT FROM HERE************
'The distance between two lifelines
dim xIncrement
xIncrement = 200
'The width of a lifeline
dim defaultwidth
defaultwidth = 90
'the diagrams name suffix
dim namesuffix
namesuffix = " Message Overview"
'the horizontal space between two levels of boundaries
dim boundaryX
boundaryX = 5
'colors
dim colors
' paars, geel, groen,blauw
colors = Array(16758490,14745599,13434828,16776869)
'**********TO HERE************

dim Yoffset
dim YIncrement

YoffSet = 72
YIncrement = 35

dim lpos
dim mpos
dim rpos
dim lmpos
dim rmpos
lpos = 100
mpos = 400
rpos = 700
lmpos = mpos - 100
rmpos = mpos + 100

sub main
	' get the selected element
	dim process as EA.Element
	set process = Repository.GetTreeSelectedObject
	if process.ObjectType = otElement then
		if process.Type = "Activity" then
			dim userinput
			userinput = MsgBox( "With process boundaries?", vbYesNoCancel + vbQuestion, "Message Overview Diagram")
			if userinput <> vbCancel then
				' make a new diagram -> MessageOverview
				dim overviewDiagram as EA.Diagram
				dim diagramName
				diagramName = process.name & namesuffix
				if userinput = vbYes then
					diagramName = diagramName & " with proces boundaries"
				end if
				set overviewDiagram = getOwnedDiagramByName(process, diagramName)
				if overviewDiagram is nothing then
					set overviewDiagram = process.Diagrams.AddNew(diagramName, "Sequence")
				end if
				overviewDiagram.update
				Repository.AdviseElementChange process.ElementID
				dim messageflows
				set messageflows = getOwnedMessageFlows(process,0, process.ElementID)
				dim messageflow as EA.Connector
				dim sequenceNumber
				sequenceNumber = 1
				dim boundaries
				set boundaries = CreateObject("System.Collections.ArrayList")
				for each messageflow in messageflows
					' Add each message flow tot he MessageOverview diagram
					addMessageFlowToDiagram messageFlow, overviewDiagram, process, sequenceNumber
					if userinput = vbYes then
						addBoundary messageFlow.Alias, overviewDiagram, boundaries,sequenceNumber
					end if
					sequenceNumber = sequenceNumber +1
				next
				Repository.OpenDiagram(overviewDiagram.DiagramID)
				Repository.SaveDiagram(overviewDiagram.DiagramID)
				're-order the lifeLines
				dim totalwidth
				totalwidth = reorderLifeLines(overviewDiagram)
				if userinput = vbYes then
					'resize boundaries
					resizeBoundaries boundaries, totalwidth
				end if
				'set all messages to Asynchronous
				setMessagesAsynchronous (overviewDiagram)
				'reload diagram
				Repository.ReloadDiagram(overviewDiagram.DiagramID)
				'tell the user we are finished
				MsgBox "Finished!"
			end if
		end if
	end if
end sub

main


function resizeBoundaries ( boundaries, totalwidth)
	dim diagramObject as EA.DiagramObject
	for each diagramObject in boundaries
		diagramObject.right = totalwidth - diagramObject.left
		diagramObject.Update
	next
end function

function addBoundary(processIDPath, diagram, boundaries, sequenceNumber)
	dim diagramObject as EA.DiagramObject
	dim processID
	dim processIDs
	processIDs = Split(processIDPath, ".") 
	
	dim level 
	level = Ubound(processIDs) 
	
	if level < 0 then
		processID = processIDPath
		level = 0
	else
		processID = processIDs(level)
	end if
	'debug
	'Session.Output "processIDPath: " & processIDPath & " level: " & level & " processID: " & processID
	'check if the last instance of the boundaries with this level is is the same process
	set diagramObject = getLastBoundaryWithLevel(level, boundaries)
	dim boundary as EA.Element
	dim foundit
	foundit = false
	if not diagramObject is nothing then
		'check if the diagramObject is about the same process
		set boundary = Repository.GetElementByID(diagramObject.ElementID)
		'Session.Output "boundary.Alias: " & boundary.Alias & " processID: " & processID & " boundary.Alias = processID " & (boundary.Alias = processID )
		if boundary.Alias = processIDPath then
			'found it. Elongate the diagramObject
			diagramObject.bottom = (YoffSet + (YIncrement * (sequenceNumber + 1))) * -1
			'save the diagram object
			diagramObject.Update
			foundit = true
		end if
	end if
	'do the "parent" boundaries first
	if level > 0 then
		'remove the last ID form the processIDPath to go one level up
		dim lastDelimiter
		lastDelimiter = InstrRev(processIDPath, ".")
		dim newProcessIDPath 
		newProcessIDPath = left(processIDPath, lastDelimiter -1)
		'make or elongate the parent
		addBoundary newProcessIDPath, diagram, boundaries, sequenceNumber
	end if
	if not foundit = true then
		'get the diagram parent element
		dim diagramOwner as EA.Element
		set diagramOwner = Repository.GetElementByID(diagram.parentID)
		'get the owning process for the message flow
		dim process as EA.Element
		set process = Repository.GetElementByID(processID)
		'create a new boundary
		set boundary = diagramOwner.Elements.AddNew("", "Boundary")
		'set the TreePos so we remember which process is used
		boundary.TreePos = processID
		'set the Alias to the alias of the messageFlow
		boundary.Alias = processIDPath
		'borderstyle
		dim borderstyle
		set borderstyle = boundary.Properties("BorderStyle")
		borderstyle.Value = "Dotted"
		if level > 0 then
			dim colorIndex
			colorIndex = level  MOD (UBound(colors) +1)
			'color
			boundary.SetAppearance 1,0,colors(colorIndex) 'groen
		end if
		'save the boundary
		boundary.Update
		'create a new diagramObject for the boundary
		dim positionString
		positionString =  "l=" & boundaryX * (level +1) & ";r=" & 1000 & ";t=" & YoffSet + (YIncrement * sequenceNumber) & ";b=" & YoffSet + (YIncrement * (sequenceNumber + 1)) & ";"
		'debug
		'Session.Output "positionString for boundary " & process.Name & ": " & positionString
		set diagramObject = diagram.DiagramObjects.AddNew( positionString, "" )
		diagramObject.ElementID = boundary.ElementID
		diagramObject.Sequence = 10 - level
		'save the diagram object
		diagramObject.Update
		'add the diagramObject tot the list of boundaries
		boundaries.Add diagramObject 
		'add the text element
		dim hyperlink as EA.Element
		dim hyperlinkName
		dim compositeDiagramID
		hyperlinkName = "$diagram://"
		if not process.CompositeDiagram is nothing then
			hyperlinkName = hyperlinkName & process.CompositeDiagram.DiagramGUID
			compositeDiagramID = process.CompositeDiagram.DiagramID
		else
			hyperlinkName = hyperlinkName & diagram.DiagramGUID
			compositeDiagramID = diagram.DiagramID
		end if
		set hyperlink = diagramOwner.Elements.AddNew(hyperlinkName, "Text")
		hyperlink.Notes = process.Name
		hyperlink.Update
		'set the link to the composite diagram
		dim hyperlinkSQL
		hyperlinkSQL = "update t_object set PDATA1 = " & compositeDiagramID & " where Object_ID = " & hyperlink.ElementID
		Repository.Execute hyperlinkSQL
		'add the hyperlink to the diagram
		positionstring = "l=" & boundaryX * (level +1) & ";r=" & 500 & ";t=" & YoffSet + (YIncrement * sequenceNumber) & ";b=" &  YoffSet + (YIncrement * sequenceNumber) + 10 & ";"
		dim hyperlinkDiagramObject as EA.DiagramObject
		set hyperlinkDiagramObject = diagram.DiagramObjects.AddNew( positionString, "" )
		hyperlinkDiagramObject.SetStyleEx "HideIcon","1"
		hyperlinkDiagramObject.ElementID = hyperlink.ElementID
		hyperlinkDiagramObject.Update
	end if
end function

function getLastBoundaryWithLevel(level, boundaries)
	dim diagramObject as EA.DiagramObject
	set diagramObject = nothing
	set getLastBoundaryWithLevel = nothing
	dim i
	for i = boundaries.Count -1  to O step -1
		set diagramObject = boundaries(i)
		'debug
		'Session.Output "level: " & level & " diagramObject.left: " & diagramObject.left & " diagramObject.left / boundaryX: " & diagramObject.left / boundaryX
		if (diagramObject.left / boundaryX) = (level + 1) then
			'debug 
			'Session.Output "found one!"
			'found one with the same level
			set getLastBoundaryWithLevel = diagramObject
			exit for
		end if
	next
end function

function setMessagesAsynchronous (diagram)
'There no clean way to do it we we do it with a dirty SQL update
	if (diagram.DiagramID > 0) then
		dim sqlupdate 
		sqlupdate = "update t_connector set PDATA1 = 'Asynchronous' where DiagramID =" & diagram.DiagramID
		Repository.Execute sqlupdate
	end if
end function

function reorderLifeLines(diagram)
	dim diagramObject as EA.DiagramObject

	dim xpos
	xpos = 50
	dim cmsName
	cmsName = "Central Market System"
	dim backendName
	backendName = "DGO-BE System"
	dim backendAdded
	backendAdded = false
	dim orderedDiagramObjects
	Set orderedDiagramObjects = CreateObject("System.Collections.ArrayList")
	dim classifier as EA.Element
	'reorder them in a new arraylist
	for each diagramObject in diagram.DiagramObjects
		set classifier = getDiagramObjectClassifier(diagramObject)
		if not classifier is nothing then
			if classifier.Name = backendName then
				orderedDiagramObjects.Add diagramObject
				backendAdded = true
			elseif classifier.Name = cmsName then
				if not backendAdded then
					orderedDiagramObjects.Add diagramObject
				else
					orderedDiagramObjects.Insert orderedDiagramObjects.Count -1, diagramObject
				end if
			else
				orderedDiagramObjects.Insert 0, diagramObject
			end if
		end if
	next
	'reset their positions
	for each diagramObject in orderedDiagramObjects
		diagramObject.left = xpos
		diagramObject.right = xpos + defaultwidth
		diagramObject.Update
		xpos = xpos + xIncrement
	next
	reorderLifeLines = xpos
end function

function getDiagramObjectClassifier(diagramObject)
	set getDiagramObjectClassifier = nothing
	dim instance as EA.Element
	set instance = Repository.GetElementByID(diagramObject.ElementID)
	if not instance is nothing and instance.ClassifierID > O then
		set getDiagramObjectClassifier = Repository.GetElementByID(instance.ClassifierID)
	end if
end function

function addMessageFlowToDiagram(messageFlow, diagram, process, sequenceNumber)
	'get the start and end element
	dim startElement as EA.Element
	dim startDiagramObject as EA.DiagramObject
	dim endElement as EA.Element
	dim endDiagramObject as EA.DiagramObject
	dim startClassifier as EA.Element
	dim endClassifier as EA.Element
	dim startLifeLine as EA.Element
	dim endLifeLine as EA.Element
	dim sequenceMessage as EA.Connector
	set startElement = Repository.GetElementByID(messageFlow.ClientID)
	set endElement = Repository.GetElementByID(messageFlow.SupplierID)
	if not startElement is nothing AND not endElement is nothing then
		'debug
		'Session.Output "start and endElement found for messageFlow.Name: " & messageflow.Name & "from " & messageflow.ClientID & " to " & messageflow.SupplierID & " id: " & messageflow.ConnectorID
		set startClassifier = getElementClassifier(startElement)
		set endClassifier = getElementClassifier(endElement)
		if not startClassifier is nothing AND not endClassifier is nothing then
			'add message between start and end
			'debug
			'Session.Output "start and endClassifier found for messageFlow.Name: " & messageflow.Name & "from " & messageflow.ClientID & " to " & messageflow.SupplierID & " id: " & messageflow.ConnectorID
			set startLifeLine = getInstanceForClassifier(startClassifier, diagram, process)
			set endLifeLine = getInstanceForClassifier(endClassifier, diagram, process)
			if not startLifeLine is nothing and not endLifeLine is nothing then
				'debug
				'Session.Output "start and endLifeLine found for messageFlow.Name: " & messageflow.Name & "from " & messageflow.ClientID & " to " & messageflow.SupplierID & " id: " & messageflow.ConnectorID
				set sequenceMessage = addSequenceMessage(messageFlow, startLifeLine,endLifeLine,sequenceNumber)
			end if
		end if
	end if
end function



function getInstanceForClassifier(classifier, diagram, process)
	set getInstanceForClassifier = nothing
	dim element as EA.Element
	dim diagramObject as EA.DiagramObject
	for each diagramObject in diagram.DiagramObjects
		'get the element
		set element = Repository.GetElementByID(diagramObject.ElementID)
		if (not element is nothing) and element.ClassifierID = classifier.ElementID then
			set getInstanceForClassifier = element
			exit for
		end if
	next
	'if not already existing then add new one
	if getInstanceForClassifier is nothing then
		set getInstanceForClassifier = addNewLifeline(classifier,process)
		'add it to the diagram
		addElementToDiagram getInstanceForClassifier, diagram, 50, 50 
		'Make sure the diagram knows that there is a new diagramObject
		diagram.DiagramObjects.Refresh
	end if
end function

function addNewLifeline(classifier,process)
	dim lifeLine as EA.Element
	set lifeLine = nothing
	set lifeLine = process.Elements.AddNew("","Object")
	if not lifeLine is nothing then
		lifeLine.ClassifierID = classifier.ElementID
		lifeLine.Update
	end if
	set addNewLifeline = lifeline
end function

function addSequenceMessage(messageFlow, startLifeLine,endLifeLine,sequenceNumber)
	set addSequenceMessage = nothing
	dim sequenceConnector as EA.Connector
	dim messageName
	messageName = ""
	'get the name of the sequence message
	dim messageRefTag as EA.ConnectorTag
	dim messageElement as EA.Element
	'Get the messageRef tagged value
	set messageRefTag = getConnectorTag(messageFlow,"messageRef")
	if not messageRefTag is nothing then
		if len(messageRefTag.Value) > O then
			 set messageElement = Repository.GetElementByGuid(messageRefTag.Value)
			 if not messageElement is nothing then
				messageName = messageElement.Name
			 end if
		end if
	end if
	if len(messageName) = 0 then
		dim intermediateEvent as EA.Element
		set intermediateEvent = Repository.GetElementByID(messageFlow.SupplierID)
		if intermediateEvent.Stereotype <> "IntermediateEvent" then
		set intermediateEvent = Repository.GetElementByID(messageFlow.ClientID)
		end if
		messageName = intermediateEvent.Name & "[MessagRef tag missing!]"
	end if
	'debug
	'messageName = messageFlow.Alias & "." & messageName 
	'add the connector
	set sequenceConnector = startLifeLine.Connectors.AddNew(messageName,"Sequence")
	sequenceConnector.SupplierID = endLifeLine.ElementID
	sequenceConnector.SequenceNo = sequenceNumber
	sequenceConnector.ClientEnd.Constraint = messageFlow.Name
	sequenceConnector.Update
	set addSequenceMessage = sequenceConnector
end function



function getElementClassifier(element)
	'Initialise
	set getElementClassifier = nothing
	dim currentElement as EA.Element
	set currentElement = element
	dim pool as EA.Element
	'intermediate event
	if currentElement.Type = "Event" and currentElement.ParentID > 0 then
		set currentElement = Repository.GetElementByID(currentElement.ParentID)
	end if
	'lane
	if currentElement.Type = "ActivityPartition" and currentElement.Stereotype = "Lane" AND currentElement.ParentID > 0 then
		set currentElement = Repository.GetElementByID(currentElement.ParentID)
	end if
	'Pool
	if currentElement.Type = "ActivityPartition" and currentElement.Stereotype = "Pool" AND currentElement.ClassfierID > 0 then
		set getElementClassifier = Repository.GetElementByID(currentElement.ClassfierID)
	end if
end function

function getXpos(element)
	getXpos = lpos
	if element.Type = "Event" then
		getXpos = mpos
	elseif element.Type = "ActivityPartition" then
		getXpos = lpos
		'DGO-BE- System is the only one that should be on the right side.
		if element.ClassfierID > 0 then
			dim actor
			set actor = Repository.GetElementByID(element.ClassfierID)
			if not actor is nothing AND actor.name = "DGO-BE System" then
				getXpos = rpos
			end if 
		end if
	end if
end function

function addMessageRefToDiagram(messageFlow, diagram, y, x,process)
	dim messageRefTag as EA.ConnectorTag
	dim messageElement as EA.Element
	'Get the messageRef tagged value
	set messageRefTag = getConnectorTag(messageFlow,"messageRef")
	if not messageRefTag is nothing then
		if len(messageRefTag.Value) > O then
			 set messageElement = Repository.GetElementByGuid(messageRefTag.Value)
			 if not messageElement is nothing then
				'add a local object for the message
				dim messageObject as EA.Element
				set messageObject = process.Elements.AddNew("", "Object")
				messageObject.ClassfierID = messageElement.ElementID
				synchronizeElement messageObject
				'add a diagramObject for the local object
				dim diagramObject as EA.DiagramObject
				set diagramObject = addElementToDiagram(messageObject, diagram, y , x)
				setFontOnDiagramObject diagramObject, "Arial Narrow", 12
				diagramObject.Update
			 end if
		end if
	end if
end function

function getConnectorTag(messageFlow, tagName)
	dim connectorTag as EA.ConnectorTag
	set getConnectorTag = nothing
	for each connectorTag in messageFlow.TaggedValues
		if connectorTag.Name = tagName then
			set getConnectorTag = connectorTag
			exit for
		end if
	next
end function

function addElementToDiagram(element, diagram, y, x)
	dim diagramObject as EA.DiagramObject
	dim positionString
	'determine height and width
	dim width 
	dim height
	dim elementType
	dim setVPartition 
	setVPartition = false
	elementType = element.Type
	select case elementType       
		case "Event"
			width = 30
			height = 30
		case "Object"
			width = 40
			height = 25
		case "Activity"
			width = 110
			height = 60
		case "ActivityPartition"
			width = 190
			height = 60
			setVPartition = true
		case else
			'default width and height
			width = 75
			height = 50
	end select
	if diagram.Type = "Sequence" then
		width = 90
		height = 150
	end if
	'to make sure all elements are vertically aligned we subtract half of the width of the x
	x = x - width/2
	'set the position of the diagramObject
	positionString =  "l=" & x & ";r=" & x + width & ";t=" & y & ";b=" & y + height & ";"
	set diagramObject = diagram.DiagramObjects.AddNew( positionString, "" )
	diagramObject.ElementID = element.ElementID
	if setVPartition then
		diagramObject.Style = "VPartition=1"
	end if
	diagramObject.Update
	diagram.DiagramObjects.Refresh
	set addElementToDiagram = diagramObject
end function

function messageFlowIsOnDiagram(messageFlow, diagram)
	dim beginElement as EA.DiagramObject
	dim endElement as EA.DiagramObject
	set beginElement = getDiagramObjectForElementID(messageFlow.ClientID, diagram)
	set endElement = getDiagramObjectForElementID(messageFlow.SupplierID, diagram)
	if not beginElement is nothing and not endElement is nothing then
		messageFlowIsOnDiagram = true
	else
		messageFlowIsOnDiagram = false
	end if
end function

'returns all owned messages of this process and of its subprocesses.
function getOwnedMessageFlows(process, level, processIDPath)
	dim messageflows
	dim messageFlowLinks
	dim messageHeights
	Set messageflows = CreateObject("System.Collections.ArrayList")
	Set messageFlowLinks = CreateObject("System.Collections.ArrayList")
	' find the composite diagam for the selected element
	dim compositeDiagram as EA.Diagram
	set compositeDiagram = process.CompositeDiagram 
	' Make a list of all MessageFlows, ordered by their vertical starting position
	if not compositeDiagram is nothing then
		dim messageFlowLink as EA.DiagramLink
		dim messageFlow as EA.Connector
		for each messageFlowLink in compositeDiagram.DiagramLinks
			if messageFlowLink.IsHidden = false then
				set messageFlow = Repository.GetConnectorByID(messageFlowLink.ConnectorID)
				dim isOnDiagram
				isOnDiagram = messageFlowIsOnDiagram(messageFlow, compositeDiagram)
				if isOnDiagram and messageFlow.Stereotype = "MessageFlow" then
					'ok, found a messageflow, add messageFlow and messageFlowLink
					'abuse the RouteStyle to store the process id
					messageFlow.RouteStyle = process.ElementID
					'abuse the SequenceNo to store the level
					messageFlow.SequenceNo = level
					'abuse the alias field to store the processIDPath
					messageFlow.Alias = processIDPath
					messageFlows.Add messageFlow
					messageFlowLinks.Add messageFlowLink
				end if
			end if
		next
		'sort the messageflows by their location in the diagram
		set messageHeights = sortMessageFlows (messageFlows, messageFlowLinks,compositeDiagram )
		' Make a list of all Activities on the diagram ordered by their vertical position. (if equal from left to right)
		dim sortedActivities
		set sortedActivities = getOrderedActivities(compositeDiagram)
		'Do the same thing for each Activity and add those messageflows to the list
		dim sortedActivity as EA.Element
		dim indexShift
		indexShift = 0
		for each sortedActivity in sortedActivities
			dim ownedMessageFlows
			set ownedMessageFlows = getOwnedMessageFlows(sortedActivity, level +1, processIDPath & "." & sortedActivity.ElementID)
			if ownedMessageFlows.Count > 0 then
				'Check the height of the sorted activity against the sorted messageHeights
				dim activityDiagramObject as EA.DiagramObject
				set activityDiagramObject = getDiagramObjectForElementID(sortedActivity.ElementID, compositeDiagram)
				dim heightindex
				heightindex = getHeightIndex(messageHeights,activityDiagramObject.top)
				'Session.Output "heightindex for " & sortedActivity.Name & "with activityDiagramObject.top: " & activityDiagramObject.top &  " heightindex : " & heightindex
				'insert the messageflows at the heightindex + indexShift
				dim insertIndex
				insertIndex = heightindex + indexShift
				dim j
				for j = ownedMessageFlows.Count -1 to 0 step -1
					'debug
					'Session.Output "inserting: " & ownedMessageFlows(j).Name & " before: " & messageflows(insertIndex).Name
					messageflows.Insert insertIndex, ownedMessageFlows(j)
				next
				'calculate new index shift
				indexShift = IndexShift + ownedMessageFlows.Count
			end if
		next
	end if
	set getOwnedMessageFlows = messageFlows
end function

function getHeightIndex(messageHeights, height)
	dim i 
	getHeightIndex = messageHeights.Count
	for i = 0 to messageHeights.Count -1
		'Session.Output "height: " & height & "messageHeights(i): " & messageHeights(i) 
		if height > messageHeights(i) then
			getHeightIndex = i
			exit for
		end if
	next
end function

function getOrderedActivities(diagram)
	'loop all diagram object
	dim diagramObject as EA.DiagramObject
	dim orderedActivities
	dim orderedDiagramObjects
	set orderedActivities = CreateObject("System.Collections.ArrayList")
	set orderedDiagramObjects = CreateObject("System.Collections.ArrayList")
	for each diagramObject in diagram.DiagramObjects
		dim element as EA.Element
		set element = Repository.GetElementByID(diagramObject.ElementID)
		if not element is nothing and element.Type = "Activity" then
			dim added 
			added = false
			dim i
			for i = 0 to orderedDiagramObjects.Count - 1
				if diagramObject.top = orderedDiagramObjects(i).top then
					'height is equal, check x position
					if diagramObject.left <= orderedDiagramObjects(i).left then
						orderedDiagramObjects.Insert i, diagramObject
						orderedActivities.Insert i, element
						added = true
						exit for
					end if
				elseif diagramObject.top > orderedDiagramObjects(i).top then
					'add before
					orderedDiagramObjects.Insert i, diagramObject
					orderedActivities.Insert i, element
					added = true
					exit for
				end if
			next
			'if not added yet then add it to the back of the list
			if not added then
				orderedDiagramObjects.Add diagramObject
				orderedActivities.Add element
			end if
		end if
	next
	set getOrderedActivities = orderedActivities
end function

function sortMessageFlows (messageFlows, messageFlowLinks,compositeDiagram)
	dim messageFlow as EA.Connector
	dim messageFlowLink as EA.DiagramLink
	dim sortedMessageFlows
	dim sortedMessageFlowLinks
	dim sortedHeights
	Set sortedMessageFlows = CreateObject("System.Collections.ArrayList")
	Set sortedMessageFlowLinks = CreateObject("System.Collections.ArrayList")
	Set sortedHeights = CreateObject("System.Collections.ArrayList")
	dim i
	for i = 0 to messageFlows.Count -1
		set messageFlow = messageFlows(i)
		set messageFlowLink = messageFlowLinks(i)
		dim height
		height = getStartingHeight(messageFlow, messageFlowLink,compositeDiagram)
		dim added
		added = false
		'loop the already sorted elements
		dim j
		for j = 0 to sortedMessageFlows.Count -1
			dim sortedHeight
			sortedHeight = sortedHeights(j)
			if sortedHeight <= height then
				sortedMessageFlows.Insert j, messageFlow 
				sortedMessageFlowLinks.Insert j, messageFlowLink
				sortedHeights.Insert j, height
				added = true
				exit for
			end if
		next
		'if it is the first element then just add it
		if not added then
			sortedMessageFlows.Add messageFlow
			sortedMessageFlowLinks.Add messageFlowLink
			sortedHeights.Add height
		end if
		
	next
	set messageFlows = sortedMessageFlows
	set messageFlowLinks = sortedMessageFlowLinks
	set sortMessageFlows = sortedHeights
end function


function getStartingHeight(connector, diagramLink, diagram)
	'check start element
	dim startElement as EA.Element
	dim elementID
	set startElement = Repository.GetElementByID(connector.ClientID)
	elementID = startElement.ElementID
	if startElement.Type = "ActivityPartition" then
		'check end element
		dim endElement as EA.Element
		set endElement = Repository.GetElementByID(connector.SupplierID)
		if endElement.Type <> "ActivityPartition" then
			elementID = endElement.ElementID
		end if
	end if
	dim diagramObject as EA.DiagramObject
	set diagramObject = getDiagramObjectForElementID(ElementID, diagram)
	if not diagramObject is nothing then
		getStartingHeight = diagramObject.top
	else
		getStartingHeight = 0
	end if
end function

function getDiagramObjectForElementID(elementID, diagram)
	set getDiagramObjectForElementID = nothing
	dim diagramObject as EA.DiagramObject
	for each diagramObject in diagram.DiagramObjects
		if diagramObject.ElementID = elementID then
			set getDiagramObjectForElementID = diagramObject
			exit for
		end if
	next
end function



function setLinkStyles(overviewDiagram)
	dim diagramLink as EA.DiagramLink
	for each diagramLink in overviewDiagram.DiagramLinks
		dim styleParts
		styleParts = Split (diagramLink.Style, ";") 
		dim i
		dim stylepart
		dim modepart 
		modepart = "Mode=2"
		dim modeSet
		modeSet = false
		dim treepart
		treepart = "TREE=OR"
		dim treeSet
		treeSet = false
		for i = 0 to Ubound(styleParts) -1
			stylepart = styleParts(i)
			'Session.Output "stylepart: " & stylepart & ", i: " & i & ", Instr(stylepart,Mode=): " & Instr(stylepart,"Mode=")
			if Instr(stylepart,"Mode=") >= 0 then
				styleParts(i) = modepart
				modeSet = true
'			elseif Instr(stylepart,"TREE=") >= 0 then
'				styleParts(i) = treepart
'				treeSet = true
			end if
		next
		diagramLink.Style = join(styleParts,";")
		if not modeSet then
			if len(diagramLink.Style) > 0 then
				diagramLink.Style = modepart & ";"& diagramLink.Style
			else 
				diagramLink.Style = modepart & ";"
			end if
		end if
'		if not treeSet then
'			diagramLink.Style = diagramLink.Style & ";" & treepart
'		end if
		diagramLink.Update
	next
end function

function getOwnedDiagramByName(element, diagramName)
	set getOwnedDiagramByName = nothing
	dim diagram as EA.Diagram
	for each diagram in element.Diagrams
		if diagram.Name = diagramName then
			set getOwnedDiagramByName = diagram
			exit for
		end if
	next
end function