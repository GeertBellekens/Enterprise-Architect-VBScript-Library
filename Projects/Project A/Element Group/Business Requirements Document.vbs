'[path=\Projects\Project A\Element Group]
'[group=Element Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.DocGenUtil
!INC Atrias Scripts.Util

'
' Script Name: Process Book
' Author: Geert Bellekens
' Purpose: Create the virtual document for a process book based on the selected processes
' Date: 08/05/2015
'

dim processBooksPackageGUID
processBooksPackageGUID = "{C771D6FE-5233-4a62-857F-3711AE1FFB0A}"
dim businessProcessesPackageGUID
businessProcessesPackageGUID = "{7EAA1987-6FB1-427f-8BA1-2610ED339905}"
dim reusableSubProcessesPackageGUID
reusableSubProcessesPackageGUID = "{5D830EDF-0470-4d41-9358-93C2EB410521}"

dim BPO_Template, BPTemplate 

BPO_Template = "BR_BPO"
BPTemplate = "BR_BPMD"

'test

sub main()
	'get the selected element
	dim domainGrouping as EA.Element
	set domainGrouping = Repository.GetContextObject()
	dim selectionOK
	selectionOK = false
	if domainGrouping.ObjectType = otElement then
		if domainGrouping.Stereotype = "ArchiMate_Grouping" then
			selectionOK = true
			'ask user for document version
			dim documentVersion
			documentVersion = InputBox("Please enter the version for this document", "Document Version", "x.y")
			'create the processbook
			createProcessBook domainGrouping, documentVersion
			Msgbox "Finished!"
		end if
	end if
	if selectionOK = false then
		Session.Prompt "Please select the Archimate Grouping that represents the domain", promptOK
	end if
end sub

main

'create a process book for the given business processes with the given document name
function createProcessBook(domainGrouping, documentVersion)
	'get the BPO processes from the domain grouping
	dim BPOs
	set BPOs = CreateObject("System.Collections.ArrayList")
	dim BPO as EA.Element
	for each BPO in domainGrouping.Elements
		if BPO.Stereotype = "ArchiMate_BusinessProcess" then
			BPOs.Add BPO
		end if
	next
	'get document info
	dim masterDocumentName,documentAlias,documentName,documentTitle,documentStatus
	documentAlias = "UMIG DGO" 
	documentName = domainGrouping.Name
	documentTitle = "UMIG DGO - BR - " & domainGrouping.Alias & " - 02 - " & domainGrouping.Name
	documentStatus = "Voor implementatie / Pour implémentation"
	masterDocumentName = documentTitle & " v" & documentVersion
	'first create a master document
	dim masterDocument as EA.Package
	set masterDocument = addMasterDocumentWithDetailTags (processBooksPackageGUID,masterDocumentName,documentAlias,documentName,documentTitle,documentVersion,documentStatus)
	dim i
	i = 1
	'sort the business processes
	'TODO
	for each BPO in BPOs
		'add model document for template Atrias element
		addModelDocumentForElement masterDocument,BPO, i, BPO_Template
		i = i + 1
	next
	'get the business processes
	dim businessProcesses
	set businessProcesses = getProcesses(businessProcessesPackageGUID, domainGrouping.Name)
	dim businessProcess as EA.Element
	for each businessProcess in businessProcesses
		'add model document for busines process
		addModelDocumentForElement masterDocument,businessProcess, i, BPTemplate
		i = i + 1
	next
	'get the reusable subprocesses
	dim reusableSubProcesses
	set reusableSubProcesses = getProcesses(reusableSubProcessesPackageGUID, domainGrouping.Name)
	dim reusableSubProcess as EA.Element
	for each reusableSubProcess in reusableSubProcesses
		'add model document for reusable sub-process
		addModelDocumentForElement masterDocument,reusableSubProcess, i, BPTemplate
		i = i + 1
	next
	'reload the package to schow the correct order
	Repository.RefreshModelView(masterDocument.PackageID)
end function

function getProcesses(parentPackageGUID, domainName)
	dim sqlGetProcesses
	sqlGetProcesses = "select o.Object_ID from ((((t_object o " & _
					" inner join t_package p on o.Package_ID = p.Package_ID) " & _
					" inner join t_package pp on pp.Package_ID = p.Parent_ID) " & _
					" inner join t_package ppp on ppp.Package_ID = pp.Parent_ID) " & _
					" inner join t_diagram d on d.ParentID = o.Object_ID) " & _
					" where o.Object_Type = 'Activity' " & _
					" and o.Stereotype in ('BusinessProcess', 'Activity') " & _
					" and ppp.ea_guid = '" & parentPackageGUID & "' " & _
					" and pp.Name = '" & domainName & "' " & _
					" order by o.Name " 
	set getProcesses = getElementsFromQuery(sqlGetProcesses)
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