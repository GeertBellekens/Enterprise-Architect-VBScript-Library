'[path=\Projects\Project AR\Diagram Group]
'[group=Diagram Group]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
'
' Script Name: Create Technical Design Document
' Author: Geert Bellekens
' Purpose: Create the virtual document for the Technical Design Document based on the selected Diagram
' Date: 2017-07-06
'
const outputName = "Create Technical Design Document"
const technicalDesignDocumentsPackageGUID = "{09E9857A-2DE1-4c69-89E5-567C78CE0B56}"

sub main
	dim selectedDiagram as EA.Diagram
	set selectedDiagram = Repository.GetCurrentDiagram()
	createTechnicalDesignDocumentMain selectedDiagram
end sub

sub test
	dim selectedDiagram as EA.Diagram
	set selectedDiagram = Repository.GetDiagramByGuid("{92626968-948B-4d86-B63B-78CC9736026A}")
	createTechnicalDesignDocumentMain selectedDiagram
end sub

sub createTechnicalDesignDocumentMain(selectedDiagram)
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'inform user we are starting
	Repository.WriteOutput outPutName,now() & " Starting creation virtual document"  , 0
	'get the package for the virtual document
	dim virtualDocumentPackage as EA.Package
	dim masterDocument as EA.Package
	set virtualDocumentPackage = getVirtualDocumentPackage(technicalDesignDocumentsPackageGUID)
	if not virtualDocumentPackage is nothing then
		if not selectedDiagram is nothing then
			'we can now start the actual creation of the document
			set masterDocument = createTechnicalDesignDocument(virtualDocumentPackage, selectedDiagram)
		end if
	end if

	'Reload completely
	Repository.RefreshModelView 0
	'inform user the document is finished
	Repository.WriteOutput outPutName,now() & " Finished creating creation virtual document, please select the virtual document and press F8 to generate the document"  , 0
	Repository.WriteOutput outPutName,now() & " Please select the virtual document and press F8 to generate the document"  , masterDocument.Element.ElementID
	'select the virtual document
	Repository.ShowInProjectView masterDocument
end sub

'create the actual document
function createTechnicalDesignDocument(virtualDocumentPackage, selectedDiagram)
	
	dim masterDocumentName,documentAlias,documentName,documentTitle,documentVersion,documentStatus
	'ask the user for the name of the document
	documentName = InputBox("Please enter the name for this document", "Document Name", selectedDiagram.Name & " - (project nummer/program nummer)")
	documentVersion = "0.0"
	documentAlias = ""
	documentTitle = ""
	documentStatus = "Vertrouwelijk"
	masterDocumentName = documentName
	'only continue if the name is filled in
	if len(documentName) > 0 then
		'try locking the Virtual Document package
		if isRequireUserLockEnabled() then
			if not virtualDocumentPackage.ApplyUserLock() then
				msgbox "Please apply user lock to the virtual document package",vbOKOnly+vbExclamation,"Virtual Document Pakage not locked!"
				exit function
			end if
		end if
		'delete previous version if it exists
		deletePreviousVersion virtualDocumentPackage, masterDocumentName
		'create the master document
		dim virtualPackageGUID
		virtualPackageGUID = virtualDocumentPackage.PackageGUID
		dim masterDocument as EA.Package
		set masterDocument = addMasterDocumentWithDetailTags(virtualPackageGUID,masterDocumentName,documentAlias,documentName,documentTitle,documentVersion,documentStatus)
		'create the model documents
		createTechnicalDesignDetails masterDocument, selectedDiagram
	end if
	'return the master document
	set createTechnicalDesignDocument = masterDocument
end function 

function createTechnicalDesignDetails(masterDocument, selectedDiagram)
	'set the counter
	dim i
	i = 0
	dim selectedDiagramPackage as EA.Package
	set selectedDiagramPackage  = Repository.GetPackageByID(selectedDiagram.PackageID)
	'Start with the prefix document
	addModelDocumentForPackage masterDocument,selectedDiagramPackage,selectedDiagramPackage.Name & " Prefix", i, "TD_Linked Document"
	i = i + 1
	'add the manual logical Flow document
	addModelDocumentForDiagram masterDocument,selectedDiagram, i, "TD_Logical Data Flow"
	i = i + 1
	'get the environment package (DVL, ACC, PRD, ...)
	dim environmentPackage as EA.Package
	set environmentPackage = getEnvironmentPackage(selectedDiagram)
	if not environmentPackage is nothing then
		'add the network document if available
		i = addNetworkDocument(masterDocument,environmentPackage, i)
		'add the Samenwerking document if available
		i = addSamenwerkingDocument(masterDocument,environmentPackage, i)
		'add all servers
		i = addServers(masterDocument,selectedDiagram,i)
		'add load balancer template
		addModelDocumentForPackage masterDocument,selectedDiagramPackage,"Load Balancer", i, "TD_Load Balancer"
		i = i + 1
		'add databases
		i = addDatabases(masterDocument,selectedDiagram,i)
		'add Message Queuing
		addModelDocumentForPackage masterDocument,selectedDiagramPackage,"Message Queuing", i, "TD_Message Queuing"
		'add File Transfer
		i = addFiletransfer(masterDocument,environmentPackage, i)
		'add network
		i = addNetwork(masterDocument,environmentPackage, i)
		'add assumptions and concerns
		addModelDocumentForPackage masterDocument,selectedDiagramPackage,"Assumptions and Concerns", i, "TD_Assumption and Concerns"
	end if
end function

function addNetwork(masterDocument,environmentPackage, i)
	dim networkPackage as EA.Package
	'loop the pakages to find the network package
	for each networkPackage in environmentPackage.Packages
		if  InStr(LCase(networkPackage.Name),"vpn") > 0 then
			'found the network package, add the model document
			addModelDocumentForPackage masterDocument,networkPackage,networkPackage.Name, i, "TD_Network"
			i = i + 1
			exit for
		end if
	next
	'return the new i
	addNetwork = i
end function

function addFiletransfer(masterDocument,environmentPackage, i)
	dim fileTransferPackage as EA.Package
	'loop the pakages to find the network package
	for each fileTransferPackage in environmentPackage.Packages
		if  InStr(LCase(fileTransferPackage.Name),"mft") > 0 then
			'found the network package, add the model document
			addModelDocumentForPackage masterDocument,fileTransferPackage,fileTransferPackage.Name, i, "TD_File Transfer"
			i = i + 1
			exit for
		end if
	next
	'return the new i
	addFiletransfer = i
end function

function addDatabases(masterDocument,selectedDiagram,i)
	dim sqlGetApplications
	sqlGetApplications = "select o.[Object_ID]                                                               " & _
						" from ((([t_diagramobjects] do                                                      " & _
						" inner join [t_object] o on (o.[Object_ID] = do.[Object_ID]                         " & _
						"                            and o.[Stereotype] = 'ArchiMate_ApplicationComponent')) " & _
						" inner join [t_connector] c on c.[Start_Object_ID] = o.[Object_ID])                 " & _
						" inner join [t_object] odb on (odb.[Object_ID] = c.[End_Object_ID]                  " & _
						"                              and odb.[Stereotype] = 'ArchiMate_DataObject'))       " & _
						" where do.[Diagram_ID] = " & selectedDiagram.DiagramID
	dim applications
	set applications = getElementsFromQuery(sqlGetApplications)
	dim appliation
	for each application in applications
		addModelDocument masterDocument, "TD_Databases", application.Name & " Databases", application.ElementGUID, i
		i = i +1
	next
	'return i
	addDatabases = i
end function

function addServers(masterDocument,selectedDiagram,i)
	'find all servers on the diagram based on their position
	dim sqlGetserversOnDiagram
	sqlGetserversOnDiagram = "select o.[Object_ID]                                               "& _
							" from ([t_diagramobjects] do                                        "& _
							" inner join [t_object] o on (o.[Object_ID] = do.[Object_ID]         "& _
							"                            and o.[Stereotype] = 'ArchiMate_Node')) "& _
							" where do.[Diagram_ID] = " & selectedDiagram.DiagramID & "          "& _
							" order by do.[RectTop] desc, do.[RectLeft]                          "
	dim servers
	set servers = getElementsFromQuery(sqlGetserversOnDiagram)
	dim server as EA.Element
	for each server in servers
		i = addServer(masterDocument,server,selectedDiagram,i)
	next
	'return i
	addServers = i
end function

function addServer(masterDocument,server,selectedDiagram,i)
	'add server template
	addModelDocument masterDocument, "TD_Server", server.Name & " Server", server.ElementGUID, i
	i = i +1
	'add container template if needed
	dim sqlGetcontainers
	sqlGetcontainers = "select o.[Object_ID]                                                           " & _
						" from (([t_diagramobjects] do                                                 " & _
						" inner join [t_object] o on (o.[Object_ID] = do.[Object_ID]                   " & _
						"                            and o.[Stereotype] = 'ArchiMate_SystemSoftware')) " & _
						" inner join [t_connector] c on c.[Start_Object_ID] = o.[Object_ID])           " & _
						" where do.[Diagram_ID] = " & selectedDiagram.DiagramID & "                    " & _
						" and c.[End_Object_ID] = " & server.ElementID
	dim containers
	set containers = getElementsFromQuery(sqlGetcontainers)
	dim container as EA.Element
	for each container in containers
		addModelDocument masterDocument, "TD_Container", container.Name & " Container", container.ElementGUID, i
		i = i +1
	next
	'add shared folder template if needed
	dim sqlGetSharedFolder
	sqlGetSharedFolder = "select o.[Object_ID]                                                            " & _
						" from (([t_diagramobjects] do                                                    " & _
						" inner join [t_object] o on (o.[Object_ID] = do.[Object_ID]                      " & _
						"                            and o.[Stereotype] = 'ArchiMate_TechnologyFunction'))" & _
						" inner join [t_connector] c on c.[End_Object_ID] = o.[Object_ID])                " & _  
						" where do.[Diagram_ID] = " & selectedDiagram.DiagramID & "                       " & _
						" and c.[Start_Object_ID] = " & server.ElementID
	dim sharedFolders
	set sharedFolders = getElementsFromQuery(sqlGetSharedFolder)
	dim sharedFolder as EA.Element
	for each sharedFolder in sharedFolders
		addModelDocument masterDocument, "TD_Technology Function", sharedFolder.Name, sharedFolder.ElementGUID, i
		i = i +1
	next
	'return i
	addServer = i
end function

function getEnvironmentPackage(selectedDiagram)
	'get the parent package of the selected diagram
	dim parentPackage as EA.Package
	set parentPackage = Repository.GetPackageByID(selectedDiagram.PackageID)
	if not parentPackage is nothing then
		dim environmentPackage as EA.Package
		set environmentPackage = Repository.GetPackageByID(parentPackage.ParentID)
		if not environmentPackage is nothing then
			set getEnvironmentPackage = environmentPackage
		else
			set getEnvironmentPackage = null
		end if
	end if
end function

function addSamenwerkingDocument(masterDocument,environmentPackage, i)
	dim samenwerkingPackage as EA.Package
	'loop the pakages to find the network package
	for each samenwerkingPackage in environmentPackage.Packages
		if  InStr(LCase(samenwerkingPackage.Name),"interactions") > 0 then
			'found the network package, add the model document
			addModelDocumentForPackage masterDocument,samenwerkingPackage,samenwerkingPackage.Name, i, "TD_Samenwerking"
			i = i + 1
			exit for
		end if
	next
	'return the new i
	addSamenwerkingDocument = i
end function

function addNetworkDocument(masterDocument,environmentPackage, i)
	dim deploymentPackage as EA.Package
	set deploymentPackage = Repository.GetPackageByID(environmentPackage.ParentID)
	if not deploymentPackage is nothing then
		dim networkPackage as EA.Package
		'loop the pakages to find the network package
		for each networkPackage in deploymentPackage.Packages
			if LCase(left(networkPackage.Name,4)) = "netw" then
				'found the network package, add the model document
				addModelDocumentForPackage masterDocument,networkPackage,"Physical Data Flow Diagram", i, "TD_Physical Data Flow"
				i = i + 1
				exit for
			end if
		next
	end if
	'return the new i
	addNetworkDocument = i
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
			if isRequireUserLockEnabled() then
				if currentPackage.ApplyUserLockRecursive(true, true, true) then
					virtualDocumentPackage.Packages.DeleteAt i,false
				else
					Repository.WriteOutput outPutName,now() & " WARNING! Previous version of virtual document could not be deleted!"  , currentPackage.Element.ElementID
				end if
			else
				virtualDocumentPackage.Packages.DeleteAt i,false
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

'call the functionality
main
'test