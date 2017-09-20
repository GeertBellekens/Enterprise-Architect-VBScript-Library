'[path=\Projects\Project Bellekens\Bellekens Docgen]
'[group=Bellekens DocGen]
!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
'

' Script Name: UseCaseDocuemnt
' Author: Geert Bellekens
' Purpose: Create the virtual document for a Use Case Document based on the given diagram
' Date: 11/11/2015
'

dim useCaseDocumentsPackageGUID

function createUseCaseDocument( diagram, documentsPackageGUID)
	
	useCaseDocumentsPackageGUID = documentsPackageGUID
	'first create a master document
	dim masterDocument as EA.Package
	set masterDocument = makeUseCaseMasterDocument(diagram)
	if not masterDocument is nothing then
		dim i
		i = 0
		'use case diagram part 1
		addModelDocumentForDiagram masterDocument,diagram, i, "UCD_Use Case Diagram"
		i = i + 1
		'add Actors
		dim diagramPackage as EA.Package
		set diagramPackage = Repository.GetPackageByID(diagram.PackageID)
		addModelDocumentForPackage masterDocument, diagramPackage, diagram.Name & " Actors", i, "UCD_Actors"
		i = i + 1
		' We only want to report the use cases that are shown within the scope boundary on this diagram
		'get the boundary diagram object in the diagram
		dim boundaries
		set boundaries = getDiagramObjects(diagram,"Boundary")
		Session.Output boundaries.Count
		'get the use cases
		dim usecases		
		if boundaries.Count > 0 then
			set usecases = getElementsFromDiagramInBoundary(diagram, "UseCase",boundaries(0))
			Session.Output "boundary found"
		else
			set usecases = getElementsFromDiagram(diagram, "UseCase")
		end if
		
		'sort use cases alphabetically
		set usecases = sortElementsByName(usecases)
		
		'add the use cases
		i = addUseCases(masterDocument, usecases, i)
		
		Repository.RefreshModelView(masterDocument.PackageID)
		'select the created master document in the project browser
		Repository.ShowInProjectView(masterDocument)
	end if
end function


function makeUseCaseMasterDocument(currentDiagram)
	dim documentTitle
	dim documentVersion
	dim documentName
	dim documentAlias
	dim masterDocumentName
	dim documentStatus
	set makeUseCaseMasterDocument = nothing
	'we should ask the user for a version
	documentVersion = ""
	documentVersion = InputBox("Please enter the version of this document", "Document version", "x.y.z" )
	'to make sure document version is filled in
	if documentVersion <> "" then
		'OK, we have a version, continue
		documentName = "UCD - " & currentDiagram.Name & " v. " & documentVersion
		'fill in the master document details
		documentAlias = currentDiagram.Name
		documentTitle = documentName
		documentAlias = documentName
		masterDocumentName = documentName
		documentStatus = "Draft"
		'then actually create the master document
		dim masterDocument as EA.Package
		set masterDocument = addMasterDocumentWithDetailTags (useCaseDocumentsPackageGUID,masterDocumentName,documentAlias,documentName,documentTitle,documentVersion,documentStatus)
		'set masterDocument = addMasterDocumentWithDetails(useCaseDocumentsPackageGUID, documentName,documentVersion,diagramName)
		set makeUseCaseMasterDocument = masterDocument
	end if
end function

'add the use cases to the document
function addUseCases(masterDocument, usecases, i)
	dim usecase as EA.Element
	for each usecase in usecases
		'use case part 1
		addModelDocument masterDocument, "UCD_Use Case details part1", "UC " & usecase.Name & " Part 1", usecase.ElementGUID, i
		i = i + 1
		
		'get the nested scenario diagram
		dim activity as EA.Element
		set activity = getActivityForUsecase(usecase)
		
		'add scenario diagram
		if not activity is nothing then
		addModelDocument masterDocument, "UCD_Use Case Scenarios Diagram", "UC " & usecase.Name & " Scenarios diagram", activity.ElementGUID, i
			i = i + 1
		end if
		
		'use case part 2
		addModelDocument masterDocument, "UCD_Use Case details part2","UC " &  usecase.Name & " Part 2", usecase.ElementGUID, i
		i = i + 1
		
	next
	'return the new i
	addUseCases = i
end function

function getActivityForUsecase(usecase)
	set getActivityForUsecase = getNestedDiagramOnwerForElement(usecase, "Activity")
end function

function getInteractionForUseCase(usecase)
	set getInteractionForUseCase = getNestedDiagramOnwerForElement(usecase, "Interaction")
end function

function getNestedDiagramOnwerForElement(element, elementType)
	dim diagramOnwer as EA.Element
	set diagramOnwer = nothing
	dim nestedElement as EA.Element
	for each nestedElement in element.Elements
		if nestedElement.Type = elementType and nestedElement.Diagrams.Count > 0 then
			set diagramOnwer = nestedElement
			exit for
		end if
	next
	set getNestedDiagramOnwerForElement = diagramOnwer
end function


'sort the elements in the given ArrayList of EA.Elements by their name 
function sortElementsByName (elements)
	dim i
	dim goAgain
	goAgain = false
	dim thisElement as EA.Element
	dim nextElement as EA.Element
	for i = 0 to elements.Count -2 step 1
		set thisElement = elements(i)
		set nextElement = elements(i +1)
		if  elementIsAfter(thisElement, nextElement) then
			elements.RemoveAt(i +1)
			elements.Insert i, nextElement
			goAgain = true
		end if
	next
	'if we had to swap an element then we go over the list again
	if goAgain then
		set elements = sortElementsByName (elements)
	end if
	'return the sorted list
	set sortElementsByName = elements
end function

'check if the name of the next element is bigger then the name of the first element
function elementIsAfter (thisElement, nextElement)
	dim compareResult
	compareResult = StrComp(thisElement.Name, nextElement.Name,1)
	if compareResult > 0 then
		elementIsAfter = True
	else
		elementIsAfter = False
	end if
end function