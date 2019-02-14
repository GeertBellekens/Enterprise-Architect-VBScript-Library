'[path=\Projects\Project MZ\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Utils.Include

'
' Script Name: Migrate from Signavio
' Author: Geert Bellekens
' Purpose: Performs a number of corrections needed after importing a Signavio model
' - Import the .bpmn files in the selected file directory
' - Strip .bpmn from package name
' - Remove Collaboration model and move the contents under the business process
' - Set the name of the diagrams to their owner
' - Set the diagrams as composite of their process/subprocess object
' - Move the lanes into the pools (add pools if needed)
' - Adjust the colors to the standard from the project template package
' Date: 2019-02-06
'
const outPutName = "Migrate from Signavio"
dim templateElements

function Main ()
	'fill the template elemnts dictionary
	getTemplateElements()
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage()
	if not selectedPackage is Nothing then
		'create output tab
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		'inform user
		Repository.WriteOutput outPutName, now() & " Starting Migrate from Signavio for package '" & selectedPackage.Name & "'" , selectedPackage.Element.ElementID
		'import the .bpmn files
		importBPMNFiles selectedPackage
		'switch back to our output 
		Repository.EnsureOutputVisible outPutName
		'refresh the subPackages
		selectedPackage.Packages.Refresh
		'do the actual work
		processPackage selectedPackage
		'inform user
		Repository.WriteOutput outPutName, now() & " Finished Migrate from Signavio for package '" & selectedPackage.Name & "'" , selectedPackage.Element.ElementID
		'refresh
		Repository.RefreshModelView 0
	end if
end function

function importBPMNFiles(package)
	dim selectedFolder
	set selectedFolder = new FileSystemFolder
	'let the user select a folder
	set selectedFolder = selectedFolder.getUserSelectedFolder("")
	if not selectedfolder is nothing then
		'start importing this folder
		importBPMNFolder selectedfolder, package
	end if
end function

function importBPMNFolder(folder, package)
	'create package first
	dim bpmnPackage as EA.Package
	'make sure the package is locked
	package.ApplyUserLock
	set bpmnPackage = package.Packages.AddNew(folder.Name, "")
	bpmnPackage.Update
	'import files in the folder
	dim file
	for each file in folder.TextFiles
		if lcase(file.Extension) = "bpmn" then
			'inform user
			Repository.EnsureOutputVisible outPutName
			Repository.WriteOutput outPutName, now() & " Importing file  '" & file.FullPath & "'" , package.Element.ElementID
			dim project as EA.Project
			set project = Repository.GetProjectInterface
			project.ImportPackageXMI project.GUIDtoXML(bpmnPackage.PackageGUID), file.FullPath , 1, 0
		end if
	next
	'then process subfolders
	dim subFolder
	for each subFolder in folder.SubFolders
		importBPMNFolder subFolder, bpmnPackage
	next
end function

function processPackage(package)
	dim businessProcesses
	set businessProcesses = getBusinessProcesses(package)
	'remove the .bpmn suffix from package name
	removeBpmnSuffix package
	'move diagram packages to the appropriate business process object
	movePackagediagrams package, businessProcesses
	'Remove Collaboration Model
	removeCollaborationModel package, businessProcesses
	'Add missing pools
	addMissingPools package
	'format diagrams
	formatDiagrams package
	'set composite diagram
	setCompositediagrams package
	'process subPackages
	dim subPackage as EA.Package
	for each subPackage in package.Packages
		processPackage subPackage
	next
	'reload package
	Repository.ReloadPackage(package.PackageID)
end function

function addMissingPools(package)
	'in case a business process has any lanes that are not owned by a pool, we add the pool, move the lane underneath and add the pool to the diagram
	'inform user
	Repository.WriteOutput outPutName, now() & " Adding missing pools for '" & package.Name & "'" , package.Element.ElementID
	dim sqlGetLanes
	sqlGetlanes = "select l.Object_ID from (t_object o                 " & _
					" inner join t_object l on (l.ParentID = o.Object_ID  " & _
					" 						and l.Stereotype = 'Lane'))   " & _
					" where o.Object_Type = 'Activity'                    " & _
					" and o.Stereotype = 'BusinessProcess'                " & _
					" and o.Package_ID = " & package.PackageID
	dim lanes
	set lanes = getElementsFromQuery(sqlGetlanes)
	dim lane as EA.Element
	for each lane in lanes
		'get the business process owning the lane
		dim businessProcess as EA.Element
		set businessProcess = Repository.GetElementByID(lane.ParentID)
		'add a pool tot the businessProcess
		dim pool as EA.Element
		set pool = businessProcess.Elements.AddNew(businessProcess.Name, "BPMN2.0::Pool")
		pool.update
		'set the processRef tagged value
		dim tv as EA.TaggedValue
		for each tv in pool.TaggedValues
			if lcase(tv.Name) = "processref" then
				tv.Value = businessProcess.ElementGUID
				tv.Update
				exit for
			end if
		next
		'move the lane underneatch the pool
		lane.ParentID = pool.ElementID
		lane.Update
		'put the pool on the diagram
		dim diagram as EA.Diagram
		for each diagram in businessProcess.Diagrams
			addElementToDiagram pool, diagram, 20, 20
		next
	next
end function

function setCompositediagrams(package)
	'inform user
	Repository.WriteOutput outPutName, now() & " Setting composite diagrams in package '" & package.Name & "'" , package.Element.ElementID
	'get all activities that have a diagram
	dim sqlGetCompositeActivities
	sqlGetCompositeActivities  = "select o.Object_ID from t_object o                  " & _
								" inner join t_diagram d on d.ParentID = o.Object_ID " & _
								" where o.Object_Type = 'Activity' " & _
								" and o.Package_ID = " & package.PackageID
	dim compositeActivities
	set compositeActivities = getElementsFromQuery(sqlGetCompositeActivities)
	dim compositeActivity as EA.Element
	'loop the activities
	for each compositeActivity in compositeActivities
		dim diagram as EA.Diagram
		for each diagram in compositeActivity.Diagrams
			'set composite
			setCompositeDiagram compositeActivity, diagram
			'rename diagram if needed
			if diagram.Name <> compositeActivity.Name then
				diagram.Name = compositeActivity.Name
				diagram.Update
			end if
			'exit after one
			exit for
		next
	next
end function

function formatDiagrams(package)
	'inform user
	Repository.WriteOutput outPutName, now() & " Formatting diagrams in package '" & package.Name & "'" , package.Element.ElementID
	dim sqlGgetDiagrams
	sqlGgetDiagrams = "select d.Diagram_ID from t_diagram d where d.Package_ID = " & package.PackageID
	dim diagrams
	set diagrams = getDiagramsFromQuery(sqlGgetDiagrams)
	dim diagram
	for each diagram in diagrams
		formatDiagram diagram
	next
end function

function movePackagediagrams(package, businessProcesses)
	dim diagram as EA.Diagram
	for each diagram in package.Diagrams
		dim parentFound
		parentFound = false
		dim businessProcess as EA.Element
		for each businessProcess  in businessProcesses
		'if we find an diagramObject that is owned by the business process we move it
			dim diagramObject as EA.DiagramObject
			dim element as EA.Element
			for each diagramObject in diagram.DiagramObjects
				set element = Repository.GetElementByID(diagramObject.ElementID)
				if element.ParentID = businessProcess.ElementID then
					parentFound = true
					exit for
				end if
			next
			''move the diagram
			if parentFound then
				moveBusinessProcessDiagram diagram, businessProcess
				exit for
			end if
		next
	next
end function

function getBusinessProcesses(package)
	'get the business process objects
	dim businessProcesses
	dim sqlGetBusinessProcesses
	sqlGetBusinessProcesses = "select o.Object_ID from t_object o      " & _
								" where                                " & _
								" o.Stereotype = 'BusinessProcess'     " & _
								" and o.Object_Type = 'Activity'       " & _
								" and o.Package_ID = " & package.PackageID
	set businessProcesses = getElementsFromQuery(sqlGetBusinessProcesses)
	'return
	set getBusinessProcesses = businessProcesses
end function

function moveBusinessProcessDiagram(diagram, businessProcess)
	dim diagramMoved
	diagramMoved = false
	'inform user
	Repository.WriteOutput outPutName, now() & " Moving Diagram to '" & businessProcess.Name & "'", businessProcess.ElementID
	'only if there are no diagrams underneatch the business process
	businessProcess.Diagrams.Refresh
	if businessProcess.Diagrams.Count = 0 then
		'move the diagram
		diagram.ParentID = businessProcess.ElementID
		diagram.Name = businessProcess.Name
		diagram.Update
		'set flag
		diagramMoved = true
		'set the diagram as composite
		setCompositeDiagram businessProcess, diagram
		'set the composite diagram for all callActivity Tasks
		dim sqlCalledactivityTasks
		sqlCalledactivityTasks = "select distinct o.Object_ID from t_object o                    " & _
								" inner join t_objectproperties tv on tv.Object_ID = o.Object_ID " & _
								" where                                                          " & _
								" tv.Property = 'calledActivityRef'                              " & _
								" and tv.Value = '" & businessProcess.ElementGUID & "'           "
		dim calledActivityTasks
		set calledActivityTasks = getElementsFromQuery(sqlCalledactivityTasks)
		'set composite 
		dim calledActivityTask as EA.Element
		for each calledActivityTask in calledActivityTasks
			setCompositeDiagram calledActivityTask, diagram
		next
	end if
	'return
	moveBusinessProcessDiagram = diagramMoved
end function



function formatDiagram(diagram)
	dim diagramObject as EA.DiagramObject
	for each diagramObject in diagram.DiagramObjects
		dim element as EA.Element
		set element = Repository.GetElementByID(diagramObject.ElementID)
		'get the corresponding diagramObject from the list of template elements
		dim templateDiagramObject
		set templateDiagramObject = getCorrespondingTemplateElement(element)
		if not templateDiagramObject is nothing then
			'set the diagramObject colors to the one from the template
			diagramObject.BackgroundColor = templateDiagramObject.BackgroundColor
			diagramObject.BorderColor = templateDiagramObject.BorderColor
			diagramObject.FontColor = templateDiagramObject.FontColor
			diagramObject.Style = templateDiagramObject.Style
			diagramObject.Update
		end if
	next

end function 

function getCorrespondingTemplateElement(element)
	dim keyElement as EA.Element
	dim templateDiagramObject as EA.DiagramObject
	'initialize
	set templateDiagramObject = nothing
	for each keyElement in templateElements.Keys
		if keyElement.Type = element.Type _
		  and keyElement.Stereotype = element.Stereotype then
			set templateDiagramObject = templateElements(keyElement)
			exit for
		end if
	next
	'return
	set getCorrespondingTemplateElement = templateDiagramObject
end function

function getTemplateElements()
	'get all diagramobjects in the template packages in a dictionary with the element as key.
	set templateElements = CreateObject("Scripting.Dictionary")
	dim templatePackage as EA.Package
	set templatePackage = getTemplatePackage()
	if not templatePackage is nothing then
		dim diagram as EA.Diagram
		for each diagram in templatePackage.Diagrams
			dim diagramObject as EA.DiagramObject
			for each diagramObject in diagram.DiagramObjects
				dim element 
				set element = Repository.GetElementByID(diagramObject.ElementID)
				set templateElements(element) = diagramObject
			next
		next
	end if
end function

function getTemplatePackage()
	'initialize at nothing
	set getTemplatePackage = nothing
	dim sqlGetPackageObject 
	sqlGetPackageObject = "select o.Object_ID from ((t_package p " & _
							" inner join usys_system syst on (syst.Property = 'TemplatePkg' " & _
							"  								and syst.Value like p.Package_ID)) " & _
							" inner join t_object o on o.ea_guid = p.ea_guid) "
    dim packageObjectCollection
	set packageObjectCollection = Repository.GetElementSet(sqlGetPackageObject, 2)					
	dim packageObject as EA.Element
	dim templatePackage as EA.Package
	for each packageObject in packageObjectCollection
		set templatePackage = Repository.GetPackageByGuid(packageObject.ElementGUID)
		set getTemplatePackage = templatePackage
		exit for
	next
end function

function getCollaborationModels(package)
	'find the collaboration model objects
	dim collaborationModels
	dim sqlGetCollaborationModels
	sqlGetCollaborationModels = "select o.Object_ID from t_object o    " & _
								" where                                " & _
								" o.Stereotype = 'CollaborationModel'  " & _
								" and o.Object_Type = 'Activity'       " & _
								" and o.Package_ID = " & package.PackageID
	set collaborationModels = getElementsFromQuery(sqlGetCollaborationModels)
	'return
	set getCollaborationModels = collaborationModels
end function

function removeCollaborationModel(package, businessProcesses)
	'get the collaborationModels
	dim collaborationModels
	set collaborationModels = getCollaborationModels(package)
	'loop through collaborationModels (should only be one)
	dim collaborationModel as EA.Element
	for each collaborationModel in collaborationModels
		dim collaborationDiagram as EA.Diagram
		set collaborationDiagram = nothing
		if collaborationModel.Diagrams.Count > 0 then
			set collaborationDiagram = collaborationModel.Diagrams.GetAt(0)
		end if
		dim diagramMoved
		diagramMoved = false
		'loop through business processes
		dim businessProcess as EA.Element
		for each businessProcess in businessProcesses
			'move the diagram to the business process
			if not collaborationDiagram is nothing then
				if not diagramMoved then
					'move the diagram
					diagramMoved = moveBusinessProcessDiagram(collaborationDiagram, businessProcess)
				else
					'copy the diagram
					'inform user
					Repository.WriteOutput outPutName, now() & " Copying Diagram to '" & businessProcess.Name & "'", businessProcess.ElementID
					dim copiedDiagram as EA.Diagram
					set copiedDiagram = copyDiagram(collaborationDiagram, businessProcess)
					'save the copied diagram
					copiedDiagram.Update
					'reload the diagram (or else the formatting doesn't work)
					set copiedDiagram = Repository.GetDiagramByID(copiedDiagram.DiagramID)
					'set it as business process diagram
					moveBusinessProcessDiagram copiedDiagram, businessProcess
				end if
			end if
			'move all pools except if the pool has a filled-in processRef tag. 
			'In that case only move it if the process is referenced
			'inform user
			Repository.WriteOutput outPutName, now() & " Moving pools to '" & businessProcess.Name & "'", businessProcess.ElementID
			dim pool as EA.Element
			for each pool in collaborationModel.Elements
				dim movePool
				movePool = true
				dim isMainPool
				isMainPool = false
				'figure out if we need to leave the "pool" element
				if lcase(pool.Stereotype) = "pool" then
					dim tag as EA.TaggedValue
					for each tag in pool.TaggedValues
						if lcase(tag.Name) = "processref" then
							if len(tag.Value) > 0 then
								if tag.Value = businessProcess.ElementGUID then
									isMainPool = true
								else
									movePool = false
									exit for
								end if
							end if
						end if
					next
				end if
				'move the pool, if allowed
				if movePool then
					pool.ParentID = businessProcess.ElementID
					pool.Update
					'check if there are any lanes in this busines process and move those underneath the pool
					if isMainPool then
						dim lane as EA.Element
						for each lane in businessProcess.Elements
							if lcase(lane.Stereotype) = "lane" then
								lane.ParentID = pool.ElementID
								lane.Update
							end if
						next
					end if
				end if
			next
		next
		'refresh to make sure we are looking at the correct collections
		collaborationModel.Elements.Refresh
		collaborationModel.Diagrams.Refresh
		'check if everything has been moved
		if collaborationModel.Elements.Count = 0 and collaborationModel.Diagrams.Count = 0 then
			'inform user
			Repository.WriteOutput outPutName, now() & " Deleting Collaboration in '" & package.Name & "'" , package.Element.ElementID
			'delete the collaborationModel
			deleteElement collaborationModel, package
		else
			Repository.WriteOutput outPutName, now() & " ERROR Collaboration in '" & package.Name & "' could not be removed because it still contains items" , collaborationModel.ElementID
		end if
	next
end function

function deleteElement (elementToDelete, owner)
	dim i
	dim element as EA.Element
	for i = owner.Elements.Count -1 to i = 0 step -1
		set element = owner.Elements.GetAt(i)
		if element.ElementGUID = elementToDelete.ElementGUID then
			'same element, delete it
			owner.Elements.DeleteAt i, false
			exit for
		end if
	next
end function


function removeBpmnSuffix(package)
	if Right(package.Name, 5) = ".bpmn" then
		package.Name = left(package.Name, len(package.Name) - 5)
		package.Update
	end if
end function

'Execute
main