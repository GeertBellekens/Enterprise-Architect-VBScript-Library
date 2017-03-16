'[path=\Projects\Project M\Metallo Modelling Standards]
'[group=Metallo Modelling Standards]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
'
' Script Name: Business Process Document Main
' Author: Geert Bellekens
' Purpose: Create the virtual document for the Business Process Document based on the selected Archimate Process
' Date: 2017-02-16
'
const outputName = "Create Business Process Document"

sub createNewBusinessProcessDocument(businessProcessDocumentsPackageGUID, rootBusinessProcess)
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'inform user we are starting
	Repository.WriteOutput outPutName,now() & " Starting Create Business Process Document"  , 0
	'validate input
	if validateInput(rootBusinessProcess) then
		'get the virtual documents package from the user
		dim virtualDocumentPackage as EA.Package
		set virtualDocumentPackage = getVirtualDocumentPackage(businessProcessDocumentsPackageGUID)
		if not virtualDocumentPackage is nothing then
			'ask the user the language
			dim documentLanguage
			documentLanguage = getUserSelectedLanguage()
			if len(documentLanguage) > 0 then
				'we can now start the actual creation of the document
				'try locking the Virtual Document package
				if isRequireUserLockEnabled() then
					if not virtualDocumentPackage.ApplyUserLock() then
						msgbox "Please apply user lock to the virtual document package",vbOKOnly+vbExclamation,"Virtual Document Pakage not locked!"
						exit sub
					end if
				end if
				createBusinessProcessDocument virtualDocumentPackage, rootBusinessProcess, documentLanguage
			end if
		end if
	end if
	'Reload completely
	Repository.RefreshModelView 0
	'inform user the document is finished
	Repository.WriteOutput outPutName,now() & " Finished Create Business Process Document"  , 0
end sub

function getUserSelectedLanguage()
	dim messageBoxResult
	messageBoxResult = msgbox("Generate the document in English?" & vbNewline & "Press 'No' for Dutch",vbYesNoCancel+vbQuestion, "Select Document Language")
	select case messageBoxResult
		case vbYes
			getUserSelectedLanguage = "EN"
		case vbNo
			getUserSelectedLanguage = "NL"
		case vbCancel
			getUserSelectedLanguage = ""
	end select
end function

'create the actual document
function createBusinessProcessDocument(virtualDocumentPackage,rootBusinessProcess,language)
	dim masterDocumentName,documentAlias,documentName,documentTitle,documentVersion,documentStatus
	documentVersion = rootBusinessProcess.Version
	documentName = rootBusinessProcess.Name
	documentAlias = "Business Process Document"
	documentTitle = documentName
	documentStatus = rootBusinessProcess.Status
	masterDocumentName = rootBusinessProcess.Name & " v." & documentVersion & "_" & language
	'delete previous version if it exists
	deletePreviousVersion virtualDocumentPackage, masterDocumentName
	'create the master document
	dim virtualPackageGUID
	virtualPackageGUID = virtualDocumentPackage.PackageGUID
	dim masterDocument as EA.Package
	set masterDocument = addMasterDocumentWithDetailTags(virtualPackageGUID,masterDocumentName,documentAlias,documentName,documentTitle,documentVersion,documentStatus)
	'add the root business process to the document
	addE2EProcessToDocument masterDocument, rootBusinessProcess,language
end function 

function addE2EProcessToDocument(masterDocument, rootBusinessProcess, language)
	'set the counter
	dim i
	i = 0
	'start with the rootBusinessProcess BP_E2E Archimate Process_EN
	addModelDocument masterDocument, "BP_E2E Archimate Process_" & language, rootBusinessProcess.Name & " Element", rootBusinessProcess.ElementGUID, i
	i = i + 1
	'get the composite diagram
	dim rootDiagram as EA.Diagram
	set rootDiagram = rootBusinessProcess.CompositeDiagram
	if not rootDiagram is nothing then
		' case of NL then we should create the dutch diagram and add a section to the document
		if language <> "EN" then
			dim dutchDiagram
			set dutchDiagram = createDutchDiagram(rootDiagram,masterDocument)
			if not dutchDiagram is nothing then
				addModelDocumentForDiagram masterDocument,dutchDiagram, i, "BP_PackageDiagram"
				i = i + 1
			end if
		end if
		'get the Archimate Business Processes shown on the composite diagram
		dim businessProcesses
		dim sqlGetBusinessProcesses
		sqlGetBusinessProcesses = 	"select act.Object_ID from (t_object act                          " & _
									" inner join t_diagramobjects do on do.Object_ID = act.Object_ID) " & _
									" where do.Diagram_ID = " & rootDiagram.DiagramID & "             " & _
									" and act.Object_Type = 'Activity'                                " & _
									" and act.Stereotype = 'ArchiMate_BusinessProcess'                " & _
									" order by do.RectLeft, do.RectTop                                "
		set businessProcesses = getElementsFromQuery(sqlGetBusinessProcesses)
		'loop the business processes and add them to the document
		dim businessprocess as EA.Element
		for each businessprocess in businessProcesses
			i = addBusinessProcessToDocument(masterDocument,businessprocess,i)
		next
	end if
end function

function createDutchDiagram(diagram,masterDocument)
	dim dutchDiagram as EA.Diagram
	set dutchDiagram = nothing
	'get the diagram package
	dim NlDiagramPackage as EA.Package
	set NlDiagramPackage = getOrCreateDiagramPackage(masterDocument)
	'add a package for the diagram
	dim diagrampackage
	set diagrampackage = NlDiagramPackage.Packages.AddNew(diagram.Name, "")
	diagrampackage.Update
	'add a copy of the diagram with he "use alias if available" enabled
	'the only way to copy a diagram is to clone the package containing the diagram and then remove everything except for the diagram
	set dutchDiagram = copyDiagram(diagram, diagrampackage)
	'tell the user if it went wrong
	if dutchDiagram is nothing then
		Repository.WriteOutput outPutName,now() & " ERROR: Failed to create Dutch diagram for '" & diagram & "'"  , 0
	else
		'set the "use alias switch
		if instr(diagram.ExtendedStyle, "UseAlias=0") > 0 then
			diagram.ExtendedStyle = replace(diagram.ExtendedStyle,"UseAlias=0","UseAlias=1")
		else
			diagram.ExtendedStyle = diagram.ExtendedStyle & "UseAlias=1;"
		end if
		diagram.Update
	end if
	'return
	set createDutchDiagram = dutchDiagram
end function



function getOrCreateDiagramPackage(masterDocument)
	dim diagramPackage as EA.Package
	'initialize at nothing
	set diagramPackage = nothing
	dim currentPackage as EA.Package
	for each currentPackage in masterDocument.Packages
		if currentPackage.Name = "Diagrams_NL" then
			set diagramPackage = currentPackage
		end if
	next
	'check if diagramPackage was found, if not create it
	if diagramPackage is nothing then
		set diagramPackage = masterDocument.Packages.AddNew("Diagrams_NL","")
		diagramPackage.Update
	end if
	'return the diagramPackage
	set getOrCreateDiagramPackage = diagramPackage
end function

function addBusinessProcessToDocument(masterDocument,businessprocess,i)
	'add the part for the element
	addModelDocument masterDocument, "BP_Archimate Process_EN", businessprocess.Name & " Element", businessprocess.ElementGUID, i
	i = i + 1
	'add the part for the diagram
	dim compositeDiagram as EA.Diagram
	set compositeDiagram = businessprocess.CompositeDiagram
	if not compositeDiagram is nothing then
		addModelDocumentForDiagram masterDocument,compositeDiagram, i, "BP_PackageDiagram"
		i = i + 1
		'then add all the workprocesses on the diagram
		dim workProcesses
		dim sqlGetworkProcesses
		sqlGetworkProcesses = 	"select act.Object_ID from (t_object act                         " & _
								" inner join t_diagramobjects do on do.Object_ID = act.Object_ID) " & _
								" where do.Diagram_ID = " & compositeDiagram.DiagramID & "        " & _
								" and act.Object_Type = 'Activity'                                " & _
								" and act.Stereotype = 'Activity'             				      " & _
								" order by do.RectLeft, do.RectTop                                "
		set workProcesses = getElementsFromQuery(sqlGetworkProcesses)
		'loop the workprocesses and add them to the document
		dim workProcess
		for each workProcess in workProcesses
			i = addWorkProcessToDocument(masterDocument,workProcess,i)
		next
	end if
	'return position
	addBusinessProcessToDocument = i
end function

function addWorkProcessToDocument(masterDocument,workProcess,i)
	'add the part for the element
	addModelDocument masterDocument, "BP_BPMN Process_EN", workProcess.Name & " Element", workProcess.ElementGUID, i
	i = i + 1
	'add the part for the diagram
	dim compositeDiagram as EA.Diagram
	set compositeDiagram = workProcess.CompositeDiagram
	if not compositeDiagram is nothing then
		addModelDocumentForDiagram masterDocument,compositeDiagram, i, "BP_WP Diagram Details_EN"
		i = i + 1
	end if 
	addWorkProcessToDocument = i
end function

'Delete the previous version if it exists
function deletePreviousVersion(virtualDocumentPackage, masterDocumentName)
	dim i
	for i = 0 to virtualDocumentPackage.Packages.Count -1
		dim currentPackage as EA.Package
		set currentPackage = virtualDocumentPackage.Packages.GetAt(i)
		if currentPackage.Name =  masterDocumentName AND currentPackage.Element.Stereotype = "master document" then
			if currentPackage.ApplyUserLockRecursive(true, true, true) then
				virtualDocumentPackage.Packages.DeleteAt i,false
			else
				Repository.WriteOutput outPutName,now() & " WARNING! Previous version of virtual document could not be deleted!"  , currentPackage.Element.ElementID
			end if
			exit for
		end if
	next
end function

function getVirtualDocumentPackage(businessProcessDocumentsPackageGUID)
	msgbox "Please select the package to create the virtual document in", vbOKOnly,"Select Virtual Documents Package"
	'let the user select the package but propose the given package if it exists
	dim virtualDocumentElementID 
	virtualDocumentElementID = Repository.InvokeConstructPicker("IncludedTypes=Package;Selection=" & businessProcessDocumentsPackageGUID & ";")
	if virtualDocumentElementID > 0 then
		dim packageElement as EA.Element
		set packageElement = Repository.GetElementByID(virtualDocumentElementID)
		set getVirtualDocumentPackage = Repository.GetPackageByGuid(packageElement.ElementGUID)
	else
		set getVirtualDocumentPackage = nothing
	end if
end function

function validateInput(rootBusinessProcess)
	'check the root business process
	dim rootProcessValid
	rootProcessValid = false
	if not rootBusinessProcess is nothing then
		if rootBusinessProcess.ObjectType = otElement then
			if rootBusinessProcess.Type = "Activity" AND rootBusinessProcess.Stereotype = "ArchiMate_BusinessProcess" then
				rootProcessValid = true
			end if
		end if
	end if
	'inform the user in case the rootprocess is not valid
	if not rootProcessValid then
		msgbox "Please select an Archimate Business Process to start the document creation",vbOKOnly+vbExclamation,"Invalid Element Selection!"
	end if
	validateInput = rootProcessValid
end function