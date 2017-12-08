'[path=\Projects\Project A\Diagram Group]
'[group=Diagram Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.DocGenUtil
'
' Script Name: Information Model Document
' Author: Geert Bellekens
' Purpose: Create the virtual document for the Information Models based on the selected Message objects.
' Use: Continued use for creating Information Model documents for MIG-DGO.
' Date: 08/05/2015
'

dim IMPackageGUID
IMPackageGUID = "{74E1FDC4-6027-4401-BCDC-2EE41092E0DD}"



dim DiagramTemplate, MessageTemplate_part1, MessageTemplate_part2

DiagramTemplate = "PackageDiagram IM"
MessageTemplate_part1 = "Atrias IM message part 1"
MessageTemplate_part2 = "Atrias IM message part 2"


'test function
sub test()
	dim currentDiagram as EA.Diagram
	'set currentDiagram = Repository.GetContextObject
	set currentDiagram = Repository.GetDiagramByGuid("{560A6A84-CE91-4fba-BDEC-82B3DF265391}")
	dim selectedDiagramObject as EA.DiagramObject
	dim selectedElement as EA.Element
	dim sortedObjects
	set sortedObjects = sortDiagramObjectsCollection (currentDiagram.DiagramObjects)
	dim selectedMessages
	set selectedMessages = CreateObject("System.Collections.ArrayList")
	'loop the sorted diagram objects
	for each selectedDiagramObject in sortedObjects
		set selectedElement = Repository.GetElementByID ( selectedDiagramObject.ElementID)
		if selectedElement.Stereotype = "Message" then
			selectedMessages.Add selectedElement
		end if
	next
	'ask user for document name
	dim documentName
	documentName = InputBox("Please enter a name for this IM document", "Document Name", "UMIG DGO - IM - [XX] - 05 - [BusinessDomain] v[X.X]")
	
	createIMdocument selectedMessages, documentName
	Msgbox "Finished!"
end sub

'test

'
' Diagram Script main function
'
sub OnDiagramScript()

	' Get a reference to the current diagram
	dim currentDiagram as EA.Diagram
	set currentDiagram = Repository.GetCurrentDiagram()

	if not currentDiagram is nothing then
		' Get a reference to any selected connector/objects
		dim selectedConnector as EA.Connector
		dim selectedObjects as EA.Collection
		set selectedConnector = currentDiagram.SelectedConnector
		set selectedObjects = currentDiagram.SelectedObjects
		
		if not selectedConnector is nothing then
			' A connector is selected
		elseif selectedObjects.Count > 0 then
			dim selectedDiagramObject as EA.DiagramObject
			dim selectedElement as EA.Element
			dim selectedMessages
			set selectedMessages = CreateObject("System.Collections.ArrayList")
			'sort the diagram objects
			dim sortedObjects
			set sortedObjects = sortDiagramObjectsCollection (selectedObjects)
			'loop the sorted diagram objects
			for each selectedDiagramObject in sortedObjects
				set selectedElement = Repository.GetElementByID ( selectedDiagramObject.ElementID)
				if selectedElement.Stereotype = "Message" then
					selectedMessages.Add selectedElement
				end if
			next
			'ask user for document name
			dim documentName
			documentName = InputBox("Please enter a name for this IM document", "Document Name", "UMIG DGO - IM - [XX] - 05 - [BusinessDomain] v[X.X]")
			
			createIMdocument selectedMessages, documentName
		else
			' Nothing is selected
		end if
		Msgbox "Finished!"
	else
		Session.Prompt "This script requires a diagram to be visible", promptOK
	end if
end sub

OnDiagramScript

'create a process book for the given business processes with the given document name
function createIMdocument(selectedMessages, documentName)
	dim message as EA.Element
	'first create a master document
	dim masterDocument as EA.Package
	set masterDocument = addMasterDocument (IMPackageGUID, documentName)
	dim i
	i = 1
	for each message in selectedMessages
		'add model document for template Atrias element
		addModelDocumentForElement masterDocument,message, i, MessageTemplate_part1
		i = i + 1
		'add model document for template diagram if exists composite diagram
		if not message.CompositeDiagram is nothing then
			addModelDocumentForDiagram masterDocument,message.CompositeDiagram , i, DiagramTemplate
			i = i + 1
		end if
		'add part 2
		'get the xsd root for the message element
		dim xsdRoot as EA.Element
		set xsdRoot = getXSDRootForMessage(message)
		if not xsdRoot is nothing then
			'add the model document
			addModelDocumentForElement masterDocument,xsdRoot, i, MessageTemplate_part2
			i = i + 1
		end if
	next
	
	'reload the package to schow the correct order
	Repository.RefreshModelView(masterDocument.PackageID)
end function

function getXSDRootForMessage(message)
	dim connector as EA.Connector
	dim xsdRoot as EA.Element
	set xsdRoot = nothing
	for each connector in message.Connectors
		if connector.SupplierID = message.ElementID _
		AND (connector.Type = "Realisation" OR connector.Type = "Realization") then
			dim connectedElement as EA.Element
			set connectedElement = Repository.GetElementByID(connector.ClientID)
			if connectedElement.Stereotype = "XSDtopLevelElement" then
				set xsdRoot = connectedElement
				exit for
			end if
		end if
	next
	set getXSDRootForMessage = xsdRoot
end function

function addModelDocumentForElement(masterDocument,documentedElement, treepos, template_in)
	dim modelDocElement as EA.Element
	dim elementName 
	dim template
	template = template_in
	elementName = documentedElement.Name & " element"
	'if the documentedElement contains a linked document then we take that instead of the content of the notes
'	dim linkedDocument
'	linkedDocument = documentedElement.GetLinkedDocument()
'	if len(linkedDocument) > 0  AND template = MessageTemplate_part1 then
'		template = template + " LD"
'		'Session.Output "Business process: "  & documentedElement.name & " len(linkedDocument): " &len(linkedDocument)
'	end if
	addModelDocument masterDocument, template,elementName, documentedElement.ElementGUID, treepos
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
		if  diagramObjectIsAfterXY(thisElement, nextElement) then
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
			diagramObjectIsAfterXY = false
		else
			diagramObjectIsAfterXY = true
		end if
	else 
		diagramObjectIsAfterXY = false
	end if
end function

function containsElement(list, elementID)
	dim element as EA.Element
	containsElement = false
	for each element in list
		if element.ElementID = elementID then
			containsElement = true
			exit for
		end if
	next
end function