'[path=\Projects\Project A\Diagram Group]
'[group=Diagram Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.DocGenUtil

'
' Script Name: Process Book
' Author: Geert Bellekens
' Purpose: Create the virtual document for a process book based on the selected processes
' Date: 08/05/2015
'

dim processBooksPackageGUID
processBooksPackageGUID = "{C771D6FE-5233-4a62-857F-3711AE1FFB0A}"

dim BPO_Template, BPTemplate 

BPO_Template = "BR_BPO"
BPTemplate = "BR_BPMD"


'test function
sub test()
	dim currentDiagram as EA.Diagram
	'set currentDiagram = Repository.GetContextObject
	set currentDiagram = Repository.GetDiagramByGuid("{806AC569-C91C-4327-9C0B-E9EA7C67A655}")
	dim selectedDiagramObject as EA.DiagramObject
	dim selectedElement as EA.Element
	dim sortedObjects
	set sortedObjects = sortDiagramObjectsCollection (currentDiagram.DiagramObjects)
	dim selectedBusinessProcesses
	set selectedBusinessProcesses = CreateObject("System.Collections.ArrayList")
	'loop the sorted diagram objects
	for each selectedDiagramObject in sortedObjects
		set selectedElement = Repository.GetElementByID ( selectedDiagramObject.ElementID)
		if selectedElement.Stereotype = "ArchiMate_BusinessProcess" then
			selectedBusinessProcesses.Add selectedElement
		end if
	next
	'ask user for document name
	dim documentName
	documentName = InputBox("Please enter the name for this BR document", "Document Name", "MIG-DGO 6.0 - BR - XX - NN - XYZ v N.N")
	
	createProcessBook selectedBusinessProcesses, documentName
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
			dim selectedBusinessProcesses
			set selectedBusinessProcesses = CreateObject("System.Collections.ArrayList")
			'sort the diagram objects
			dim sortedObjects
			set sortedObjects = sortDiagramObjectsCollection (selectedObjects)
			'loop the sorted diagram objects
			for each selectedDiagramObject in sortedObjects
				set selectedElement = Repository.GetElementByID ( selectedDiagramObject.ElementID)
				if selectedElement.Stereotype = "ArchiMate_BusinessProcess" then
					selectedBusinessProcesses.Add selectedElement
				end if
			next
			'ask user for document name
			dim documentName
			documentName = InputBox("Please enter the name for this process book", "Document Name", "MIG-DGO-PB-XX-NN-CMS Process Book xxx")
			
			createProcessBook selectedBusinessProcesses, documentName
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
function createProcessBook(selectedBusinessProcesses, documentName)
	dim businessProcess as EA.Element
	'first create a master document
	dim masterDocument as EA.Package
	set masterDocument = addMasterDocument (processBooksPackageGUID, documentName)
	dim i
	i = 1
	dim subProcesses
	set subProcesses = CreateObject("System.Collections.ArrayList")
	for each businessProcess in selectedBusinessProcesses
		'add model document for template Atrias element
		addModelDocumentForElement masterDocument,businessProcess, i, BPO_Template
		i = i + 1
		'add model document for template diagram if exists composite diagram
		if not businessProcess.CompositeDiagram is nothing then
			'get the list of all sub-processes on the composite diagram for later processing
			subProcesses.AddRange getSubProcesses(businessProcess.CompositeDiagram, subProcesses)
		end if
	next
	dim subProcess as EA.Element
	for each subProcess in subProcesses
		'add model document for template Atrias element
		addModelDocumentForElement masterDocument,subProcess, i, BPTemplate
		i = i + 1
	next
	'reload the package to schow the correct order
	Repository.RefreshModelView(masterDocument.PackageID)
end function

function getSubProcesses(diagram, existingSubProcesses)
	dim sortedDiagramObjects
	dim sortedSubProcesses
	set sortedSubProcesses = CreateObject("System.Collections.ArrayList")
	set sortedDiagramObjects = sortDiagramObjectsCollection(diagram.DiagramObjects)
	dim diagramObject as EA.DiagramObject
	dim subProcess as EA.Element
	for each diagramObject in sortedDiagramObjects
		if not containsElement(existingSubProcesses, diagramObject.ElementID) then
			set subProcess = Repository.GetElementByID(diagramObject.ElementID)
			if subProcess.Stereotype = "ArchiMate_BusinessProcess" _
				OR subProcess.Stereotype = "Activity" _
				OR subProcess.Stereotype = "BusinessProcess"then
				sortedSubProcesses.Add subProcess
			end if
		end if
	next
	set getSubProcesses = sortedSubProcesses
end function


function addModelDocumentForElement(masterDocument,businessProcess, treepos, template_in)
	dim modelDocElement as EA.Element
	dim elementName 
	dim template
	template = template_in
	elementName = businessProcess.Name & " element"
	'if the businessprocess contains a linked document then we take that instead of the content of the notes
	dim linkedDocument
	linkedDocument = businessProcess.GetLinkedDocument()
	if len(linkedDocument) > 0 then
		template = template + " Linked Document"
		'Session.Output "Business process: "  & businessProcess.name & " len(linkedDocument): " &len(linkedDocument)
	end if
	addModelDocument masterDocument, template,elementName, businessProcess.ElementGUID, treepos
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