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
	 
	dim messages
	set messages = getProcessSpecificMessages(objectID)
	dim message as EA.Connector
	for each message in messages
		addRow xmlData, xmlDOM, message
	next
	if messages.Count = 0 then
		'if no messages found then we add N.v.t / S.O as output
		addNotApplicableRow xmlDOM, xmlData
	end if
	MyRtfData = xmlDOM.xml
end function


function getProcessSpecificMessages(objectID)
	dim businessProcess as EA.Element
	set businessProcess = Repository.GetElementByID(objectID)
	dim messages
	set messages = CreateObject("System.Collections.ArrayList")
	if not businessProcess.CompositeDiagram is nothing then
		set messages = getMessages(businessProcess.CompositeDiagram)
	end if
	set getProcessSpecificMessages = messages
end function

function addRow (xmlData, xmlDOM, messageFlow)
	
	dim messageElement as EA.Element
	set messageElement = getMessageFromMessageFlow(messageFlow)
	if not messageElement is nothing then
		dim technicalMessage as EA.Element
		set technicalMessage = getTechnicalMessageFromMessage(messageElement)
		
		dim xmlRow
		set xmlRow = xmlDOM.createElement( "Row" )
		xmlData.appendChild xmlRow

		dim xmlFISName
		set xmlFISName = xmlDOM.createElement( "FISName" )			
		xmlFISName.text = messageElement.Name
		xmlRow.appendChild xmlFISName
		
		dim xmlMessageName
		set xmlMessageName = xmlDOM.createElement( "MessageName" )	
		if not technicalMessage is nothing then		
			xmlMessageName.text = technicalMessage.Name
		else
			xmlMessageName.text = ""
		end if
		xmlRow.appendChild xmlMessageName
		
		'description
		dim notes
		if len(messageFlow.Notes) > 0 then
			notes = messageFlow.Notes
		elseif len(messageElement.Notes) > 0 then
			notes = messageElement.Notes
		elseif not technicalMessage is nothing then
			notes = technicalMessage.Notes
		end if
		
		dim descriptionfull
		descriptionfull = getTagContent(notes, "")
		
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
	end if
end function

function addNotApplicableRow(xmlDOM, xmlData)
	dim xmlRow
	set xmlRow = xmlDOM.createElement( "Row" )
	xmlData.appendChild xmlRow

	dim xmlFISName
	set xmlFISName = xmlDOM.createElement( "FISName" )			
	xmlFISName.text = "n.v.t. - S.O."
	xmlRow.appendChild xmlFISName
	
	dim xmlMessageName
	set xmlMessageName = xmlDOM.createElement( "MessageName" )	
	xmlMessageName.text = "n.v.t. - S.O."
	xmlRow.appendChild xmlMessageName

	dim formattedAttr 
	
	'description NL

	dim xmlDescNL
	set xmlDescNL = xmlDOM.createElement( "DescriptionNL" )	
	xmlDescNL.text = "n.v.t."
	xmlRow.appendChild xmlDescNL
	
	'description FR
	
	dim xmlDescFR
	set xmlDescFR = xmlDOM.createElement( "DescriptionFR" )			
	xmlDescFR.text = "S.O."
	xmlRow.appendChild xmlDescFR
	
end function

function getTechnicalMessageFromMessage(messageElement)
	set getTechnicalMessageFromMessage = nothing
	dim connector as EA.Connector
	'Session.Output "MessageElement.Name: " & messageElement.Name & " Connectors.Count: " & messageElement.Connectors.Count
	for each connector in messageElement.Connectors
		'Session.Output "connector.Type = " & connector.Type
		'check if it is a realization
		if connector.Type = "Realization" OR  connector.Type = "Realisation" then
			dim technicalElement as EA.Element
			set technicalElement = Repository.GetElementByID(connector.ClientID)
			if technicalElement.Stereotype = "Message" then
				set getTechnicalMessageFromMessage = technicalElement
				exit for
			end if
		end if
	next
end function


function getMessages(diagram)
	dim sortedDiagramObjects
	dim sortedMessages
	set sortedMessages = CreateObject("System.Collections.ArrayList")
	dim messageFlows
	Set messageFlows = CreateObject("System.Collections.ArrayList")
	dim messageFlowLinks
	Set messageFlowLinks = CreateObject("System.Collections.ArrayList")
	dim diagramObject as EA.DiagramObject
	dim message as EA.Element
	dim messageFlowLink as EA.DiagramLink
	dim messageFlow as EA.Connector
	for each messageFlowLink in diagram.DiagramLinks
		if messageFlowLink.IsHidden = false then
			set messageFlow = Repository.GetConnectorByID(messageFlowLink.ConnectorID)
			if messageFlow.Stereotype = "MessageFlow" then
				'ok, found a messageflow, add messageFlow and messageFlowLink
				messageFlows.Add messageFlow
				messageFlowLinks.Add messageFlowLink
			end if
		end if
	next
	'sort the messageFlows
	sortMessageFlows messageFlows, messageFlowLinks,diagram
	'return them
	set getMessages = messageFlows
end function


function getMessageFromMessageFlow(messageFlow)
	set getMessageFromMessageFlow = nothing
	dim messageRefTag as EA.ConnectorTag
	'Get the messageRef tagged value
	set messageRefTag = getConnectorTag(messageFlow,"messageRef")
	if not messageRefTag is nothing then
		if len(messageRefTag.Value) > 0 then
			 set getMessageFromMessageFlow = Repository.GetElementByGuid(messageRefTag.Value)
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
			if sortedHeight < height then
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

function test
	dim outputString
	dim fileSystemObject
	dim outputFile
	
	outputString = MyRtfData(9721)
	
	set fileSystemObject = CreateObject( "Scripting.FileSystemObject" )
	set outputFile = fileSystemObject.CreateTextFile( "c:\\temp\\processSpecificMessages.txt", true )
	outputFile.Write outputString
	outputFile.Close
	
end function

'test