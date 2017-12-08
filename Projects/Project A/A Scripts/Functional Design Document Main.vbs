'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.DocGenUtil
!INC Atrias Scripts.Util
'
'
' Script Name: Functional Analysis Document
' Author: Geert Bellekens
' Purpose: Create the virtual document for a functional Analysis document based on the given diagram
' Date: 2016-12-16
'

'
dim WC


function makeFAMasterDocument(currentDiagram)

	'the diagram name should be in the format [Domain Abbr] - [UCMD] - [Functional Name]
	'the master document should have following properties:
	' - DocumentTitle -> to go into the tagged value RTFName [Functional Name]
	' - Version -> version of the master document
	' - DocumentName - name of the Master document
	' - Domain name -> Alias 
	
	'we should ask the user for a version
	dim documentTitle
	dim documentVersion
	dim documentName
	dim domainName
	dim diagramName
	dim functionalName
	set makeFAMasterDocument = nothing
	diagramName = currentDiagram.Name
	'get the abbreviation from the diagram name
	dim nameparts
	dim abbreviation
	nameparts = Split (diagramName, "-")
	if Ubound(nameparts) > 0 then
		abbreviation = Trim(nameparts(0))
		functionalName = Trim(nameparts(Ubound(nameparts)))
		Session.Output "functionalName: " & functionalName
	else
		abbreviation = diagramName
	end if
	'get the domain name based onthe abbreviation
	domainName = getFullDomainName(abbreviation)
	if domainName <> "" then
		'to make sure document version is filled in
		documentVersion = ""
		documentVersion = InputBox("Please enter the version of this document", "Document version", "x.y.z" )
		if documentVersion <> "" then
			'get document info
			dim masterDocumentName,documentAlias,documentStatus
			documentAlias = getApplication(currentDiagram)
			documentName =  domainName & " - " & functionalName 
			documentTitle = documentAlias & " - FD - " & documentName
			documentStatus = "For consultation"
			masterDocumentName = documentAlias & " - FD - " & abbreviation & " - " & functionalName & " v." & documentVersion
			'Remove older version of the master document
			removeMasterDocumentDuplicates FDDocumentsPackageGUID, documentAlias & " - FD - " & abbreviation & " - " & functionalName
			'then create a new master document
			dim masterDocument as EA.Package
			set masterDocument = addMasterDocumentWithDetailTags (FDDocumentsPackageGUID,masterDocumentName,documentAlias,documentName,documentTitle,documentVersion,documentStatus)
			set makeFAMasterDocument = masterDocument
		end if
	end if	
end function

function getFullDomainName(abbreviation)
	'default "Unknown"
	getFullDomainName = "Unknown Domain"
	dim getDomainNameSQL
	getDomainNameSQL = "select o.Object_ID from t_object o " & _
						" where o.Stereotype = 'Archimate_Grouping' " & _
						" and o.Alias = '" & abbreviation & "'"
    dim domainElements
	set domainElements = getElementsFromQuery(getDomainNameSQL)
	if domainElements.Count > 0 then
		dim domainElement as EA.Element
		set domainElement = domainElements(0)
		getFullDomainName = domainElement.Name
	else
		dim domainName
		domainName = InputBox("Could not find domain name for abbreviation '" & abbreviation & "'" & vbNewLine & "Please enter the full domain name" _
										, "Enter full domain name", "Domain Name" )
		if len(domainName) > 0 then
			getFullDomainName = domainName
		end if
	end if
end function

function createFADocument( diagram)
	'initialize wildcard character
	WC = getWC
	'dim diagram as EA.Diagram
	'first create a master document
	dim masterDocument as EA.Package
	set masterDocument = makeFAMasterDocument(diagram)
	if not masterDocument is nothing then
		dim i
		i = 0
		
		'get the boundary diagram object in the diagram
		dim boundaries
		set boundaries = getDiagramObjects(diagram,"Boundary")
		'get the use cases 
		dim usecases		
		if boundaries.Count > 0 then
			set usecases = getElementsFromDiagramInBoundary(diagram, "UseCase",boundaries(0))
			Session.Output "boundary found"
		else
			set usecases = getElementsFromDiagram(diagram, "UseCase")
		end if
		
		'sort use cases alphabetically -> At this point, the changelog table should be created
		set usecases = sortElementsByName(usecases)
		
		Session.Output "usecases.Count : " & usecases.Count
		dim ucPackage as EA.Package
		set ucPackage = Repository.GetPackageByID(diagram.PackageID)
			
		'add user changelog
		addModelDocumentForPackage masterDocument,ucPackage, diagram.name & " - Changelog", i, "FD_Changelog"
		i = i + 1


		'use case diagram part 1
		addModelDocumentForDiagram masterDocument,diagram, i, "FD_Use Case Diagram"
		i = i + 1
		
		'get the Business Process Activities
'		dim bpas
'		set bpas = getBusinessProcessActivitiesForUseCases(usecases)
'		Session.Output "bpas.count : " & bpas.Count
'		
'		'get the Business processes
'		dim bpmds
'		set bpmds = getBusinessProcessesForActivities(bpas)
'		Session.Output "bpmds.count : " & bpmds.Count
'		
'		'get the Business Process Overviews
'		dim bpos
'		set bpos = getBPOsForBPMDs(bpmds)
'		Session.Output "bpos.count: " & bpos.Count
'		
'		'get the Application Functions
'		dim apfs
'		set apfs = getAplicationFunctionsForUseCases(usecases)
		
		'add the applicationFunction diagrams
		'dim apfDiagrams
		'set apfDiagrams = addApplicationFunctionDiagrams(masterDocument,apfs)
		
		'get the data stores
	'	dim dataStores
	'	set dataStores = getDataStoresForApplicationFunctions(apfs)
		
		'make the process map diagram
'		dim processsMapDiagram
'		set processsMapDiagram = addProcessmapDiagram(masterDocument, diagram, bpos)
'		'add the diagram tot the document
'		addModelDocumentForDiagram masterDocument,processsMapDiagram, i, "FD_PackageDiagram"
'		i = i + 1
		
		'add the title "Related Business Process Overview Elements"
'		addModelDocument masterDocument, "FD_Related BPO elements title","Related Business Process Overview Elements", "", i
'		i = i + 1
			
		'add the composite diagrams for the bpos to the document
'		i = addCompositeDiagrams(masterDocument, bpos, i)
		
		'add the title "Related Business Process Overview Elements"
'		addModelDocument masterDocument, "FD_Related Business Processes title","Related Business Processes", "", i
'		i = i + 1
			
		'add the composite diagrams for the bpos to the document
'		i = addCompositeDiagrams(masterDocument, bpmds, i)

		'add the matrix businessprocesses vs requirements
'		dim diagramPackage as EA.Package
'		set diagramPackage = Repository.GetPackageByID(diagram.PackageID)
'		addModelDocumentForPackage masterDocument, diagramPackage, "BusinessActivities x Requirements", i, "FD_BusinessActivities x Requirements"
'		i = i + 1
		
		'add Actors
		dim diagramPackage as EA.Package
		set diagramPackage = Repository.GetPackageByID(diagram.PackageID)
		addModelDocumentForPackage masterDocument, diagramPackage, diagram.Name & " Actors", i, "FD_Actors"
		i = i + 1
		
		'add use case diagram part 2
'		addModelDocumentForDiagram masterDocument,diagram, i, "FD_Use Case Diagram Part2"
'		i = i + 1
		
		'add the use cases
		i = addUseCases(masterDocument, usecases, i)
		
		'add the Related Class diagrams title
	'	addModelDocument masterDocument, "Related Class Diagrams title","Related Class Diagrams title", "", i
	'	i = i + 1
		
		'add the data stores
	'	i = addDataStores(masterDocument, dataStores, i)
		
		'finished, refresh model view to make sure the order is reflected in the model.
		Repository.RefreshModelView(masterDocument.PackageID)
	end if
end function


function addDataStores(masterDocument, dataStores, i)
	dim dataStore as EA.Element
	for each dataStore in dataStores
		addModelDocument masterDocument, "FD_DataObject", "DataStore -" & dataStore.Name , dataStore.ElementGUID, i
		i = i + 1
	next
	addDataStores = i
end function

function addUseCases(masterDocument, usecases, i)
	dim usecase as EA.Element
	for each usecase in usecases
		'use case part 1
		addModelDocument masterDocument, "FD_Use Case details part1", usecase.Name & " Part 1", usecase.ElementGUID, i
		i = i + 1
		
		'get the nested Activity diagram
		dim activity as EA.Element
		set activity = getActivityForUsecase(usecase)
		
		'add activity diagram
		if not activity is nothing then
		addModelDocument masterDocument, "FD_Use Case Activity Diagram", usecase.Name & " Activity diagram", activity.ElementGUID, i
			i = i + 1
		end if
		
		'use case part 2
		addModelDocument masterDocument, "FD_Use Case details part2", usecase.Name & " Part 2", usecase.ElementGUID, i
		i = i + 1
		
		'get the nested Sequence diagram
		dim interAction as EA.Element
		set interAction = getInteractionForUseCase(usecase)
		
		'add sequence diagram
		if not interAction is nothing then
		addModelDocument masterDocument, "FD_Use Case Sequence Diagram", usecase.Name & " Sequence diagram", interAction.ElementGUID, i
			i = i + 1
		end if
		
		'add user Rules
		addModelDocumentWithSearch masterDocument, "FD_Rules", usecase.Name & " Rules", usecase.ElementGUID, i, "ZDG_RulesByUseCaseGUID"
		i = i + 1
		
		'add user interface details
		addModelDocumentWithSearch masterDocument, "FD_User Interface details", usecase.Name & " User Interfaces", usecase.ElementGUID, i, "ZDG_ApplicationInterfaceByUseCaseGUID"
		i = i + 1
		
		'add traceability
		addModelDocument masterDocument, "FD_User Case details Traceability", usecase.Name & " Traceability", usecase.ElementGUID, i
		i = i + 1
		
		'use case part 3
'		addModelDocument masterDocument, "FD_Use Case details part3", usecase.Name  & " Part 3", usecase.ElementGUID, i
'		i = i + 1
		
'		'add the applicaiton function diagrams (should be only one per use case)
'		dim applicationFunctions
'		dim useCaseCollection
'		set useCaseCollection = CreateObject("System.Collections.ArrayList")
'		useCaseCollection.Add usecase
'		set applicationFunctions = getAplicationFunctionsForUseCases(useCaseCollection)
'		dim applicationFunction as EA.Element
'		for each applicationFunction in applicationFunctions
'			'get the diagram from the dictorary
'			dim apfDiagram as EA.Diagram
'			set apfDiagram = apfDiagrams(applicationFunction.ElementID)
'			'add the diagram to the document
'			addModelDocumentForDiagram masterDocument,apfDiagram, i, "FD_PackageDiagram"
'			i = i + 1
'		next
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

'adds the composite diagrams of all the elements int he given list
function addCompositeDiagrams(masterDocument, elements, i)
	dim element as EA.Element
	for each element in elements
		dim compositeDiagram
		set compositeDiagram = element.CompositeDiagram
		if not compositeDiagram is nothing then
				'add the diagram tot the document
				addModelDocumentForDiagram masterDocument,compositeDiagram, i, "FD_PackageDiagram"
				i = i + 1
		end if
	next
	'return i
	addCompositeDiagrams = i
end function

function addApplicationFunctionDiagrams(masterDocument,apfs)
	dim applicationFunction as EA.Element
	'Dictonary to keep the diagrams
	dim apfDiagrams
	set apfDiagrams = CreateObject("Scripting.Dictionary")
	'create a package and diagram for each application function (usually only one)
	for each applicationFunction in apfs
		'add package
		dim apfDiagramPackage as EA.Package
		set apfDiagramPackage = masterDocument.Packages.AddNew("ApplicationFunction - " & applicationFunction.Name,"")
		apfDiagramPackage.Update
		
		'add diagram
		dim apfDiagram as EA.Diagram
		set apfDiagram = apfDiagramPackage.Diagrams.AddNew(apfDiagramPackage.Name, "Analysis")
		apfDiagram.Update
		
		'add the Application Function tot he diagram
		addElementToDiagram applicationFunction, apfDiagram, 50, 100
		
		'add the diagram to the dictionary
		apfDiagrams.Add applicationFunction.ElementID, apfDiagram
	next
	set addApplicationFunctionDiagrams = apfDiagrams
end function

function addProcessmapDiagram(masterDocument, diagram, bpos)
	'add package
	dim processMapPackage as EA.Package
	set processMapPackage = masterDocument.Packages.AddNew(diagram.Name & " Process Map","")
	processMapPackage.Update
	
	'add diagram
	dim processMapDiagram as EA.Diagram
	set processMapDiagram = processMapPackage.Diagrams.AddNew(processMapPackage.Name, "Analysis")
	processMapDiagram.Update
	
	dim bpoGroupings
	set bpoGroupings = sortBPOandGetGroupings(bpos)
	Session.Output "bpoGroupings.Count: " & bpoGroupings.Count
	'add elements on diagram
	dim x
	x = 100
	dim y 
	Y = 50
	dim bpo
	dim groupingID
	groupingID = 0
	for each bpo in bpos
		if groupingID > 0 and groupingID <> bpo.ParentID then
			'we go to a new grouping
			' add the current grouping to the diagram
			addBPOGrouping processMapDiagram,bpoGroupings,groupingID, x, y
			'move down the y
			y = y + 150
		end if
		groupingID = bpo.ParentID
		'add the bpo to the diagram
		addElementToDiagram bpo, processMapDiagram, y, x
		x = x + 200
	next
	'add the last grouping here
	if (groupingID > 0) then
		addBPOGrouping processMapDiagram,bpoGroupings,groupingID, x, y
	end if
		
	'return diagram
	set addProcessmapDiagram = processMapDiagram
end function

function addBPOGrouping(processMapDiagram,bpoGroupings,groupingID, x, y)
	dim grouping
	set grouping = bpoGroupings(groupingID)
	'add the grouping to the diagram
	dim groupingDiagramObject as EA.DiagramObject
	set groupingDiagramObject = addElementToDiagram(grouping, processMapDiagram, y - 30 , 50)
	groupingDiagramObject.right = x - 75
	groupingDiagramObject.left = 15
	groupingDiagramObject.bottom = (y * -1) - 70
	Session.Output "groupingDiagramObject.right: " & groupingDiagramObject.right & "x + 50: " & x + 50
	groupingDiagramObject.Update
end function			

function sortBPOandGetGroupings(bpos)
	dim bpoGroupings
	set bpoGroupings = CreateObject("Scripting.Dictionary")
	dim bpo as EA.Element
	'first get the bpogroupings
	for each bpo in bpos
		if not bpoGroupings.Exists(bpo.ParentID) then
			bpoGroupings.Add bpo.ParentID, Repository.GetElementByID(bpo.ParentID)
			Session.Output "adding grouping: " & bpo.ParentID
		end if
	next
	'then start sorting
	set bpos = sortBPOsWithGroupings (bpos, bpoGroupings)
	set sortBPOandGetGroupings = bpoGroupings
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

'check if the name of the next elemnt is bigger then the name of the first element
function elementIsAfter (thisElement, nextElement)
	dim compareResult
	compareResult = StrComp(thisElement.Name, nextElement.Name,1)
	if compareResult > 0 then
		elementIsAfter = True
	else
		elementIsAfter = False
	end if
end function

function sortBPOsWithGroupings (bpos, bpoGroupings)
	dim i
	dim goAgain
	goAgain = false
	dim thisElement as EA.Element
	dim nextElement as EA.Element
	for i = 0 to bpos.Count -2 step 1
		set thisElement = bpos(i)
		set nextElement = bpos(i +1)
		if  bpoIsAfter(thisElement, nextElement, bpoGroupings) then
			bpos.RemoveAt(i +1)
			bpos.Insert i, nextElement
			goAgain = true
		end if
	next
	'if we had to swap an element then we go over the list again
	if goAgain then
		set bpos = sortBPOsWithGroupings (bpos, bpoGroupings)
	end if
	'return the sorted list
	set sortBPOsWithGroupings = bpos
end function

function bpoIsAfter(thisElement, nextElement, bpoGroupings) 
	'first check if they have the same parent
	if thisElement.ParentID = nextElement.ParentID then
		'same parent so compare names
		if StrComp(thisElement.Name,nextElement.Name,1) > 0 then
			bpoIsAfter = true
		else 
			bpoIsAfter = false
		end if
	else
		'no different parents. order by parent
		dim bpoGrouping as EA.Element
		for each bpoGrouping in bpoGroupings
			if bpoGrouping.ElementID = thisElement.ParentID then
				bpoIsAfter = false
				exit for
			elseif bpoGrouping.ElementID = nextElement.ParentID then
				bpoIsAfter = true
				exit for
			end if
		next
	end if
end function


'returns the applicaton functions that are linked with the given use cases
function getAplicationFunctionsForUseCases(usecases)
	dim usecaseIDstring
	useCaseIDString = makeIDString(usecases)
	dim sqlSelect 
	sqlSelect = 	"select distinct af.Object_ID from (( t_object uc  "&_
					"inner join [t_connector] ucaf on uc.[Object_ID] = ucaf.[End_Object_ID]) "&_
					"inner join t_object af on af.Object_ID = ucaf.[Start_Object_ID]) "&_
					"where  "&_
					"uc.[Object_ID] in (" & usecaseIDstring & ") "&_
					"and af.[Object_Type] = 'Activity' and af.[Stereotype] = 'Archimate_ApplicationFunction'"
	
	set getAplicationFunctionsForUseCases = getElementsFromQuery(sqlSelect)
end function

function getDataStoresForApplicationFunctions(apfs)
	dim apfIDstring
	apfIDstring = makeIDString(apfs)
	dim sqlSelect 
	sqlSelect = 	"select distinct ds.Object_ID from (( t_object af   "&_
					"inner join [t_connector] afds on af.Object_ID = afds.[Start_Object_ID]) "&_
					"inner join t_object ds on ds.[Object_ID] = afds.[End_Object_ID] )  "&_
					"where  "&_
					"af.[Object_ID] in (" & apfIDstring & ") "&_
					"and ds.[Object_Type] = 'Class' and ds.[Stereotype] = 'Archimate_DataObject'"
	
	set getDataStoresForApplicationFunctions = getElementsFromQuery(sqlSelect)
end function

function getBusinessProcessActivitiesForUseCases(usecases)
	dim usecaseIDstring
	useCaseIDString = makeIDString(usecases)
	dim sqlSelect 
	sqlSelect = "select distinct bpa.Object_ID from (((( t_object uc "&_ 
					" inner join t_connector ucrq on ucrq.[Start_Object_ID] = uc.[Object_ID]) "&_ 
					" inner join t_object rq on ucrq.[End_Object_ID] = rq.[Object_ID] ) "&_ 
					" inner join [t_connector] rqbpa on rq.[Object_ID] = rqbpa.[Start_Object_ID])  "&_ 
					" inner join t_object bpa on bpa.Object_ID = rqbpa.[End_Object_ID])  "&_ 
					"where  "&_
					"uc.[Object_ID] in (" & usecaseIDstring & ") "&_
					"and bpa.[Object_Type] = 'Activity' and bpa.[Stereotype] = 'Activity' "
	
	set getBusinessProcessActivitiesForUseCases = getElementsFromQuery(sqlSelect)
end function

function getBusinessProcessesForActivities(activities)
	dim bpaIDstring
	bpaIDString = makeIDString(activities)
	dim sqlSelect 
	sqlSelect = "select distinct bp.Object_ID from ((((t_object bpa  "&_
				"inner join [t_objectproperties] bpatv on bpatv.VALUE like bpa.[ea_guid]) "&_
				"inner join t_object bpai on bpai.[Object_ID] = bpatv.[Object_ID]) "&_
				"inner join t_diagramObjects bpaido on bpaido.[Object_ID] = bpai.Object_ID) "&_
				"inner join t_object bp on bp.pdata1 like bpaido.Diagram_ID) "&_
				"where bpa.[Object_ID] in (" & bpaIDString & ") "&_
				"and bpa.[Object_Type] = 'Activity' and bpa.[Stereotype] = 'Activity'  "&_
				"and bpatv.Property = 'calledActivityRef' "&_
				"and bp.[Object_Type] = 'Activity' and bp.[Stereotype] = 'Archimate_BusinessProcess' and bp.name like '"& WC &"BPMD"& WC &"' "
	set getBusinessProcessesForActivities = getElementsFromQuery(sqlSelect)
end function 

function getBPOsForBPMDs(businessprocesses)
	dim bpmdIDstring
	bpmdIDstring = makeIDString(businessprocesses)
	dim sqlSelect
	sqlSelect = "select distinct bpo.Object_ID from ((t_object bp  "&_
				"inner join t_diagramobjects bpdo on bp.object_id = bpdo.object_id) "&_
				"inner join t_object bpo on bpo.pdata1 like bpdo.Diagram_ID) "&_
				"where bp.[Object_ID] in (" & bpmdIDstring & ") "&_
				"and bpo.[Object_Type] = 'Activity' and bpo.[Stereotype] = 'Archimate_BusinessProcess' and bpo.name like '"& WC &"BPO"& WC &"' "
	set getBPOsForBPMDs = getElementsFromQuery(sqlSelect)
end function