'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]

!INC Local Scripts.EAConstants-VBScript
!INC Bellekens DocGen.DocGenHelpers
!INC Bellekens DocGen.Util
!INC Wrappers.Include

'******************************************************************************
' Script Name: Genereer Proceskaart - Algemeen	
' Author: Alain Van Goethem
' Purpose: Creatie van een virtueel document (ProcesKaart, of kort PK) in EA op basis van een Proces-element
' Date: 2019-08
'******************************************************************************
dim WC
const PK_Template_History = "PK_DocumentHistory"
const PK_Template_BusinessProcess = "PK_BusinessProcess"
const PK_Template_OverviewDiagram = "PK_OverviewDiagram"
const PK_Template_ConceptualProcessOwner = "PK_ConceptualProcessOwner"
const PK_Template_ProcessUniversum = "PK_ProcessUniversum"
const PK_Template_ConceptualProcesses = "PK_ConceptualProcesses"
const PK_Template_ConceptualProcessDiagramViaElement = "PK_ConceptualProcessDiagramViaElement"
const PK_Template_ConceptualProcessDiagramViaPackage = "PK_ConceptualProcessDiagramViaPackage"
const PK_Template_LinksToTurtleDiagram = "PK_LinksToTurtleDiagram"
const PK_Template_TurtleDiagram = "PK_TurtleDiagram"
const PK_Template_TurtleDiagramElements = "PK_TurtleDiagramElements"
const PK_Template_Diagram = "PK_BPMNDiagram"
const PK_Template_DiagramElements = "PK_BPMNDiagramElements"
const PK_Template_DiagramRelatedElements = "PK_BPMNDiagramRelatedElements"
const PK_Template_RisksAndOpportunities = "PK_RisksAndOpportunities"
const PK_Template_Subprocesses = "PK_Subprocesses"
const PK_Template_SubprocessDiagram = "PK_SubprocessDiagram"
const PK_Template_SubprocessDiagramElements = "PK_SubProcessDiagramElements"
const PK_Template_StageParts = "PK_StageParts"
const PK_Template_OpenIssues = "PK_OpenIssues"

'*** update GUID to match the GUID of the 'Documents' package of this model ***
const DocumentsPackageGUID = "{0D62104A-981E-434f-B094-F3EE46ABDA05}"
'***************************************************************************************************

function createBPMasterDocument(currentElement)
	'define variables
	dim elementName	
	dim documentName
	dim documentVersion
	dim documentTitle
	dim documentAlias
	dim documentStatus

	'prepare master document
	Session.output "Preparing master document..."
	set createBPMasterDocument = nothing
	elementName = currentElement.Name
	
	'ask user for document version
	'documentVersion = InputBox("Gelieve de versie van het document in te vullen", "Proceskaart - Document versie", "x.y.z")
	'if documentVersion <> "" then
	'	documentName = "Proceskaart - " & elementName & " - v. " & documentVersion
	'else
		documentName = "Proceskaart - " & elementName
	'end if
	documentVersion = ""
	documentAlias = documentName
	documentTitle = documentName
	documentStatus = "Draft"

	'Remove older version of the master document
	'removeMasterDocumentDuplicates DocumentsPackageGUID, documentTitle

	'create master document
	Session.output "Creating master document..."
	dim masterDocument as EA.Package
	set masterDocument = addMasterDocumentWithDetailTags (DocumentsPackageGUID,documentName,documentAlias,documentName,documentTitle,documentVersion,documentStatus)
	'set masterDocument = addMasterDocumentWithDetails (DocumentsPackageGUID,documentName,documentVersion,documentAlias)
	set createBPMasterDocument = masterDocument

end function

function createBPDocument(element)
	'initialize wildcard character
	WC = getWC
	'dim element as EA.Element
	dim conceptualbusinessprocesses
	
	'check if selected element is a BusinessProcess
	if element.Stereotype = "BusinessProcess" Then
	
		'first create a master document
		dim masterDocument as EA.Package
		set masterDocument = createBPMasterDocument(element)
		if not masterDocument is nothing then
			Session.output "Master document created"
			dim i
			i = 0
			
			'get Package containing the BusinessProcess object
			dim bpPackage as EA.Package
			set bpPackage = Repository.GetPackageByID(element.PackageID)
			
			'add model document with History (from LinkedDocument of BusinessProcess)
			Session.output "Adding Document History..."
			addModelDocumentForPackage masterDocument,bpPackage, "Document History", i, PK_Template_History
			i = i + 1
			
			'add model document with info of Business Process
			Session.output "Adding Business Process..."
			addModelDocumentForPackage masterDocument,bpPackage, "Business Process Info", i, PK_Template_BusinessProcess
			i = i + 1
			
			'get related Conceptual Business Processes (from Process Universe  / PUR)
			set conceptualbusinessprocesses = getConceptualBusinessProcesses(element)
			if conceptualbusinessprocesses.Count > 0 then
				'add Process owner details of related conceptual Business Process (Assumption: exactly 1 result)
				dim relatedBusinessProcess
				Session.output "Adding Process Owner details..."
				for each relatedBusinessProcess in conceptualbusinessprocesses
					Session.Output "Found Process: " + relatedBusinessProcess.Name
					addModelDocumentWithSearch masterDocument, PK_Template_ConceptualProcessOwner,relatedBusinessProcess.Name, relatedBusinessProcess.ElementGUID, i, "ZDG_ElementByGUID"
					'addModelDocumentForElement masterDocument,relatedBusinessProcess, i, PK_Template_ConceptualProcessOwner
					i = i + 1
				next 
			end if		
			
	'		'get Package containing the Overview diagram
	'		dim OverviewPackage as EA.Package
	'		set OverviewPackage = Repository.GetPackageByID(bpPackage.ParentID)
	'		
	'		'get Overview diagram
	'		dim overviewDiagram as EA.Diagram
	'		set overviewDiagram = OverviewPackage.Diagrams.GetAt(0) 
	'		if not overviewDiagram is nothing Then			
	'			'add model document for diagram
	'			Session.output "Adding overview diagram..."
	'			addModelDocumentForPackage masterDocument,OverviewPackage, "Overview diagram", i, PK_Template_OverviewDiagram
	'			i = i + 1
	'		else
	'			Session.Output "Error - overview diagram not found"
	'		end if	
			
			'add model document with Title, link (if any) to "Turtle diagram" 
			Session.output "Adding Turtle Diagram info..."
			addModelDocumentForPackage masterDocument,bpPackage, "Turtle diagram link", i, PK_Template_LinksToTurtleDiagram
			i = i + 1

			'Search Turtle diagram  (to avoid error from GetByName when diagram not present)
			dim tmpdiagrams
			set tmpdiagrams = bpPackage.Diagrams
			dim tmpdiagram as EA.Element
			dim turtleDiagram as EA.Diagram
			'Session.output tmpdiagrams.Count
			if tmpdiagrams.Count>0 then
				'Session.Output "diagrams found"
				for each tmpdiagram in tmpdiagrams
					'Session.Output tmpdiagram.Name
					if tmpdiagram.Name="Schildpad" then
						set turtleDiagram = bpPackage.Diagrams.GetByName("Schildpad")
						
						'add model document for diagram
						Session.output "Adding Turtle diagram..."
						addModelDocumentForPackage masterDocument,bpPackage, "Turtle diagram", i, PK_Template_TurtleDiagram
						i = i + 1
				
						' add Turtle diagram Elements
						Session.output "Adding Turtle diagram elements..."
						addModelDocumentForPackage masterDocument,bpPackage, "Turtule Diagram Elements", i, PK_Template_TurtleDiagramElements		
						i = i + 1
					end if
				next
			end if			
				
			'add related Business Processes (that are outside of diagram)  & related CIM-objects (with CRUD)
			Session.output "Adding Related Elements..."
			addModelDocumentForPackage masterDocument,bpPackage, "Related Elements", i, PK_Template_DiagramRelatedElements
			i = i + 1
			
			'add Risks and Opportunities (from LinkedDocument of Parent Package)
			'Session.output "Adding Risks and Opportunities..."
			'addModelDocumentForPackage masterDocument,bpPackage, "Risks And Opportunities", i, PK_Template_RisksAndOpportunities
			'i = i + 1
			
			'add title for Process universe 
			Session.output "Adding Process universe..."
			addModelDocumentForPackage masterDocument, bpPackage, "Process universe", i, PK_Template_ProcessUniversum
			i = i + 1
			'add related Conceptual Business Processes from the ProcessUniverse (ArchiMate)
			addConceptualBusinessProcesses masterdocument, conceptualbusinessprocesses, bpPackage, element, i
		
			'********************************************************* MAIN PROCESS **************************************************************
			'get BPMN diagram under BusinessProcess object
			dim diagram as EA.Diagram
			set diagram = element.Diagrams.GetAt(0) 
			if not diagram is nothing Then			
				'add model document for diagram
				Session.output "Adding BPMN diagram..."
				addModelDocumentForPackage masterDocument,bpPackage, "BPMN diagram", i, PK_Template_Diagram
				i = i + 1
				
				' add BPMN Elements
				Session.output "Adding BPMN diagram elements..."
				'addModelDocumentForPackage masterDocument,bpPackage, "BPMN diagram elements", i, PK_Template_DiagramElements	
				addModelDocumentWithSearch masterDocument, PK_Template_DiagramElements ,element.Name + "- elements", element.ElementGUID, i, "ZDG_ElementByGUID"
				i = i + 1
			else
				Session.Output "Error - diagram not found"
			end if	
			
			'********************************************************* SUBPROCESSES **************************************************************
			' add title Subprocesses
			Session.output "Adding Subprocesses..."
			addModelDocumentForPackage masterDocument,bpPackage, "Subprocesses", i, PK_Template_Subprocesses		
			i = i + 1
			
			'get the subprocesses
			dim subProcesses
			set subProcesses = getSubProcesses(bpPackage, diagram)
			dim subProcess as EA.Element
			for each subProcess in subProcesses
				Session.output "Adding Subprocess... " + subProcess.Name
				if subProcess.Stereotype = "Activity" _
				  or right(subProcess.Stereotype, 4) = "Task" then
					'get BPMN diagram under SubProcess object
					dim subdiagram as EA.Diagram
					set subdiagram = subProcess.CompositeDiagram
					if not subdiagram is nothing Then			
						'add model document for subprocess diagram
						Session.output "Adding BPMN Subprocess diagram..."
						addModelDocumentWithSearch masterDocument, PK_Template_SubprocessDiagram,subProcess.Name, subProcess.ElementGUID, i, "ZDG_ElementByGUID"
						'addModelDocumentForElement masterDocument,subProcess, i, PK_Template_SubprocessDiagram
						i = i + 1
						
						' add BPMN subprocess Elements
						Session.output "Adding BPMN Subprocess diagram elements..."
						'addModelDocumentForPackage masterDocument,bpPackage, "BPMN subprocess diagram elements", i, PK_Template_SubprocessDiagramElements		
						addModelDocumentWithSearch masterDocument, PK_Template_SubprocessDiagramElements ,subProcess.Name + "- elements", subProcess.ElementGUID, i, "ZDG_ElementByGUID"
						i = i + 1	
					else
						Session.Output "Error - diagram not found"
					end if	
				else
					'process stage
					addModelDocumentWithSearch masterDocument, PK_Template_StageParts ,subProcess.Name + "- tasks", subProcess.ElementGUID, i, "ZDG_ElementByGUID"
					i = i + 1
				end if
			next
			
			'********************************************************* OPEN ISSUES **************************************************************
			'add model document with open issues
			Session.output "Adding Open issues..."
			addModelDocumentForPackage masterDocument,bpPackage, "Open issues", i, PK_Template_OpenIssues
			i = i + 1
			
			'finished, refresh model view to make sure the order is reflected in the model.
			Session.output "Reloading master document..."
			Repository.ReloadPackage(masterDocument.PackageID)
		else
			Session.output "Error - Master document not created"
		end if
	else
		Session.output "Wrong type of element selected"
		MsgBox("Gelieve een element van type BusinessProcess te selecteren om een proceskaart te genereren.")
	end if
end function

function getConceptualBusinessProcesses(element)
	dim ArchiMateProcesses
	dim sql	
	'find related ArchiMate_BusinessProcesses for BusinessProcess via Trace-relationship
	Session.Output "Searching conceptual Business Processes..."
	sql = "select o.Object_ID  "& _
			" from ((t_object o "& _
			" inner join t_connector c on (o.Object_ID = c.End_Object_ID)) "& _
			" inner join t_object o2 on (o2.Object_ID = c.Start_Object_ID)) "& _
			" where o.Stereotype = 'ArchiMate_BusinessProcess' "& _
			" and o2.Stereotype = 'BusinessProcess' "& _
			" and c.Stereotype = 'trace' "& _
			" and o2.ea_guid = '" & element.ElementGUID & "'" & _
			" order by o.Name"	
	set ArchiMateProcesses = getElementsFromQuery(sql)
	set getConceptualBusinessProcesses = ArchiMateProcesses
end function

function addConceptualBusinessProcesses(masterDocument,ArchiMateProcesses,bpPackage,element,i)
	dim ArchiMateProcess as EA.Element
	dim compositeDiagram as EA.Diagram
	
	if ArchiMateProcesses.Count > 0 then
		Session.Output "Adding conceptual Business Processes..."
		'add related ArchiMate BusinessProcesses from ProcessUniversum
		for each ArchiMateProcess in ArchiMateProcesses
			' add model document per Element
			Session.Output ArchiMateProcess.Name
			'addModelDocumentForElement masterDocument, ArchiMateProcess, i, PK_Template_ConceptualProcesses
			addModelDocumentWithSearch masterDocument, PK_Template_ConceptualProcesses, ArchiMateProcess.Name, ArchiMateProcess.ElementGUID, i, "ZDG_ElementByGUID"
			i = i + 1
			
			set compositeDiagram = ArchiMateProcess.CompositeDiagram
			if not compositeDiagram is nothing then
				'Add different template depending on parent of found composite diagram
				if compositeDiagram.ParentID = 0 then
					'get Package containing the Composite diagram
					dim CompositeDiagramPackage as EA.Package
					set CompositeDiagramPackage = Repository.GetPackageByID(compositeDiagram.PackageID)
					addModelDocumentForPackage masterDocument,CompositeDiagramPackage, "CompositeDiagram", i, PK_Template_ConceptualProcessDiagramViaPackage
				else
					addModelDocumentForElement masterDocument, ArchiMateProcess, i, PK_Template_ConceptualProcessDiagramViaElement
				end if
				i = i + 1
			end if
		next
	else
		Session.Output "No conceptual Business Processes found."
	end if

	addConceptualBusinessProcesses = i
end function

function addElements(masterDocument, elements, i)
	'Session.Output "# of elements : " & element.count
	if elements.count > 0 then	
		dim element as EA.Element
		for each element in elements
			' add model document per Element
			addModelDocumentForElement masterDocument, element, i, PK_Template_Element
			i = i + 1
		next	
	else
		Session.Output "No elements added."
	end if
	'return the new i
	addElements = i
end function

function addModelDocumentForElement(masterDocument,element, treepos, template_in)
	dim modelDocElement as EA.Element
	dim elementName 
	dim template
	template = template_in
	elementName = element.Name & " element"
	'if the element contains a linked document then we take that instead of the content of the notes
	dim linkedDocument
	linkedDocument = element.GetLinkedDocument()
	if len(linkedDocument) > 0 then
		template = template + " Linked Document"
	end if
	addModelDocument masterDocument, template,elementName, element.ElementGUID, treepos
end function

function getSubProcesses(parentpackage, diagram)
	dim sortedDiagramObjects
	set sortedDiagramObjects = sortDiagramObjectsCollection(diagram.DiagramObjects)
	dim sortedSubProcesses
	set sortedSubProcesses = CreateObject("System.Collections.ArrayList")
	dim sortedSubDiagramObjects

	dim diagramObject as EA.DiagramObject
	dim subdiagramObject as EA.DiagramObject
	dim subProcess as EA.Element
	dim compositeDiagram as EA.Diagram
	'Check each Activity object if it contains a composite diagram 
	'and add to Subprocesses if it does
	for each diagramObject in sortedDiagramObjects
		set subProcess = Repository.GetElementByID(diagramObject.ElementID)
		if subProcess.Stereotype = "Stage" then
			Session.Output "Found stage lvl1: " + subProcess.Name
			sortedSubProcesses.Add subProcess
		elseif subProcess.Stereotype = "Activity" _
		    or right(subProcess.Stereotype, 4) = "Task" then
			set compositeDiagram = subProcess.CompositeDiagram
			
			'debug
			dim hasSubDiagram
			if compositeDiagram is nothing then
				hasSubDiagram = false
			else
				hasSubDiagram = true
			end if
			Session.Output "Processing " & subProcess & " " & subProcess.Name & " has composite diagram: " & hasSubDiagram
			
			if not compositeDiagram is nothing then
				Session.Output "Found subprocess lvl1: " & subProcess.Name
				sortedSubProcesses.Add subProcess
				
				'Search for 2nd level of subprocesses inside the subprocess
				set sortedSubDiagramObjects = sortDiagramObjectsCollection(compositeDiagram.DiagramObjects)
				for each subdiagramObject in sortedSubDiagramObjects
					set subProcess = Repository.GetElementByID(subdiagramObject.ElementID) 
					if subProcess.Stereotype = "Activity" _
						or right(subProcess.Stereotype,4) = "Task" then
						set compositeDiagram = subProcess.CompositeDiagram
						if not compositeDiagram is nothing then
							Session.Output "Found subprocess lvl2: " + subProcess.Name
							sortedSubProcesses.Add subProcess
						end if
					end if
				next
			end if
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
