'[path=\Projects\Project E\Documents]
'[group=Documents]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
'
'
' Script Name: Architecture Blueprint
' Author: Geert Bellekens
' Purpose: Generate the Architecture Blueprint document
' Date: 2024-04-12
'


dim ArchitectureBlueprintDocument
set ArchitectureBlueprintDocument = new Document
ArchitectureBlueprintDocument.Name = "Architecture Blueprint"
ArchitectureBlueprintDocument.Description = "A Word document containing the Architecture Blueprint of a project"
ArchitectureBlueprintDocument.ValidQuery = "select top 1 p.Package_ID from t_package p where p.Name like '%Capability Map%' and p.Package_ID in (#Branch#)"

sub GenerateArchitectureBlueprintDocument
	Repository.WriteOutput outPutName, now() & " Starting generation of " & ArchitectureBlueprintDocument.Name , 0
	'get selected package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	if package is nothing then
		msgbox "please select a package before running this script"
		exit sub
	end if
	'do the actual work
	createArchitectureBlueprint package
	Repository.WriteOutput outPutName, now() & " Finished generation of " & ArchitectureBlueprintDocument.Name , 0
end sub
'GenerateArchitectureBlueprintDocument 'test

dim packageTreeIDString

function createArchitectureBlueprint(package)
	'set the global packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	'get document package
	dim documentPackage as EA.Package
	set documentPackage = getDocumentPackage(package)
	'create master document
	dim masterDocument as EA.Package
	set masterDocument = createArchitectureBlueprintMasterDocument(documentPackage, package)
	if masterDocument is nothing then
		Repository.WriteOutput outPutName, now() & " ERROR: Can't create master document!" , 0
		exit function
	end if
	'counter
	dim i
	i = 0
	'Introduction
	i = addArchitectureBlueprintIntroduction(documentPackage, masterDocument, i)
	'Capability Map
	i = addCapabilityMap(package, masterDocument, i)
	'Business Information concepts
	i = addInformationConcepts(package, masterDocument, i)
	'Capabilities managing CBICS + Data Domains
	i = addDataSection(package, masterDocument, i)
	'Capability Stories
	i = addCapabilityStories(package, masterDocument, i)
	'Solution Architecture
	i = addSolutionArchitecture(package, masterDocument, i)
	'reload model
	Repository.ReloadPackage documentPackage.PackageID
	'generate document
	generateMasterDocument masterDocument
end function

function addSolutionArchitecture(package, masterDocument, i)
	dim solutionArchitecturePackage as EA.Package
	set solutionArchitecturePackage = getPackageInContextPackage("Solution architecture")
	if solutionArchitecturePackage  is nothing then
		exit function
	end if
	'create the model document
	dim solutionArchitectureModelDocument as EA.Element
	set solutionArchitectureModelDocument = addEmptyModelDocument(masterDocument, "Solution Architecture", i, "ABP_SolutionArchitecture")
	'add all packages to the model document
	addPackagesToModelDocument solutionArchitectureModelDocument, solutionArchitecturePackage
	i =  i + 1
	'return i
	addSolutionArchitecture = i
end function

function addCapabilityStories(package, masterDocument, i)
	dim storiesPackage as EA.Package
	set storiesPackage = getPackageInContextPackage("Capability Stories")
	if storiesPackage  is nothing then
		exit function
	end if
	'add modeld document for stories package
	addModelDocumentForPackage masterDocument, storiesPackage, "Capability Stories", i, "ABP_Story"
	i =  i + 1
	'return i
	addCapabilityStories = i
end function

function getPackageInContextPackage(packageNamePattern)
	set getPackageInContextPackage = nothing
	dim sqlGetData
	sqlGetData = "select top 1 p.Package_ID from t_package p        " & vbNewLine & _
				 " where p.Name like '" & packageNamePattern & "'   " & vbNewLine & _
				 " and p.Package_ID in (" & packageTreeIDString & ")"
	dim packages
	set packages = getPackagesFromQuery(sqlGetData)
	if not packages.Count > 0 then
		Repository.WriteOutput outPutName, now() & " ERROR: Can't find package with name '" & packageNamePattern & "'!" , 0
		exit function
	end if
	'return
	set getPackageInContextPackage = packages(0)		
end function


function addDataSection(package, masterDocument, i)
	dim dataDomainsPackage as EA.Package
	set dataDomainsPackage = getPackageInContextPackage("Data Domains")
	if dataDomainsPackage is nothing then
		exit function
	end if
	'add model document for capabilities managing CBICS
	addModelDocumentForPackage masterDocument,dataDomainsPackage, "Capabilities Managing CBICS", i, "ABP_CBICS and Capabilities"
	i =  i + 1
	'add model document for Data Domains
	addModelDocumentForPackage masterDocument,dataDomainsPackage, "Data Domains", i, "ABP_Data Domains"
	i =  i + 1
	'return i
	addDataSection = i
end function


function addInformationConcepts(package, masterDocument, i)
	dim informationPackage as EA.Package
	set informationPackage = getPackageInContextPackage("%Information Architecture%")
	if informationPackage is nothing then
		exit function
	end if
	'create the model document
	dim informationModelDocument as EA.Element
	set informationModelDocument = addEmptyModelDocument(masterDocument, "CBIM", i, "ABP_BIM")
	i =  i + 1
	'add all packages to the model document
	addPackagesToModelDocument informationModelDocument, informationPackage
	'return i
	addInformationConcepts = i
end function

function addPackagesToModelDocument(modelDocElement, package)
	dim subPackage as EA.Package
	for each subPackage in package.Packages
		if subPackage.Diagrams.Count > 0 then
			'add the package to the model document, only if it has at least one diagram
			addPackageToModelDocument modelDocElement, subPackage
		end if
		'process it's subpackages
		addPackagesToModelDocument modelDocElement, subPackage
	next
end function


function addCapabilityMap(package, masterDocument, i)
	'get capabilityMap package
	dim capabilityMapPackage as EA.Package
	set capabilityMapPackage = getCapabilityMapPackage(package)
	'add model document
	addModelDocumentForPackage masterDocument,capabilityMapPackage, "Capability Map", i, "ABP_CapabilityMap"
	i = i + 1
	i = addL2CapabilitiesFromCapabilityMap(capabilityMapPackage,masterDocument, i)
	'return i
	addCapabilityMap = i
end function

function addL2CapabilitiesFromCapabilityMap(capabilityMapPackage, masterDocument, i)
	dim sqlGetData
	sqlGetData = "select distinct o.Object_ID, do.RectTop, do.RectLeft                                " & vbNewLine & _
				" from t_object o                                                                    " & vbNewLine & _
				"  inner join t_objectproperties tv on tv.Object_ID = o.Object_ID                    " & vbNewLine & _
				"  								and tv.Property = 'Level'                            " & vbNewLine & _
				"  								and tv.Value = 'L2'                                  " & vbNewLine & _
				"  inner join t_diagramobjects do on do.Object_ID = o.Object_ID                      " & vbNewLine & _
				"  inner join t_diagram d on d.Diagram_ID = do.Diagram_ID                            " & vbNewLine & _
				" 						and d.Diagram_ID =                                           " & vbNewLine & _
				" 						(select top 1 d2.Diagram_ID from t_diagram d2                " & vbNewLine & _
				" 						where d2.Package_ID = "& capabilityMapPackage.PackageID &"   " & vbNewLine & _
				" 						and exists (select * from t_diagramobjects do2               " & vbNewLine & _
				" 									where do2.Diagram_ID = d2.Diagram_ID             " & vbNewLine & _
				" 									and do2.Object_ID = o.Object_ID)                 " & vbNewLine & _
				" 						order by d2.TPos, d2.Name)                                   " & vbNewLine & _
				"  where o.stereotype = 'ArchiMate_Capability'                                       " & vbNewLine & _
				"  and d.Package_ID = "& capabilityMapPackage.PackageID &"                           " & vbNewLine & _
				"  order by do.RectTop desc, do.RectLeft                                             "
	dim l2Capabilities
	set l2Capabilities = getElementsFromQuery(sqlGetData)
	dim l2Capability as EA.Element
	for each l2Capability in l2Capabilities
		'add modeldocument
		addModelDocument masterDocument, "ABP_L2Capability",l2Capability.Name, l2Capability.ElementGUID, i
		i = i + 1
	next
	'return i
	addL2CapabilitiesFromCapabilityMap = i
end function

function getCapabilityMapPackage(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	dim sqlGetData
	sqlGetData = replace(ArchitectureBlueprintDocument.ValidQuery, "#Branch#", packageTreeIDString)
	dim packages
	set packages = getPackagesFromQuery(sqlGetData)
	if not packages.Count > 0 then
		Repository.WriteOutput outPutName, now() & " ERROR: Can't find package with name 'Business Capability Map'!" , 0
		exit function
	end if
	'return
	set getCapabilityMapPackage = packages(0)
end function

function addArchitectureBlueprintIntroduction(documentPackage, masterDocument, i)
	'get intro document artifact
	dim documentArtifact as EA.Element
	set documentArtifact = getIntroDocumentArtifact(documentPackage, masterDocument)
	'add modeldocument
	addModelDocument masterDocument, "ABP_LinkedDocument",documentArtifact.Name, documentArtifact.ElementGUID, i
	i = i + 1
	'return i
	addArchitectureBlueprintIntroduction = i
end function

function getIntroDocumentArtifact(documentPackage, masterDocument)
	dim documentArtifact as EA.Element
	dim sqlGetData
	sqlGetData = "select top 1 o.Object_ID from t_object o              " & vbNewLine & _
				" where o.Object_Type = 'Artifact'                      " & vbNewLine & _
				" and o.Stereotype = 'Document'                         " & vbNewLine & _
				" and o.Package_ID = "& masterDocument.ParentID & "    " & vbNewLine & _
				" and o.Name = '"& masterDocument.Name &" Introduction' "
	dim result
	set result = getElementsFromQuery(sqlGetData)
	if result.Count > 0 then
		set documentArtifact = result(0)
	else
		'create the introduction artifact if not found already
		set documentArtifact = documentPackage.Elements.AddNew(masterDocument.Name & " Introduction", "Artifact")
		documentArtifact.StereotypeEx = "StandardProfileL2::Document"
		documentArtifact.Update
	end if
	'return
	set getIntroDocumentArtifact = documentArtifact
end function

function createModelDocumentsForL0Capability(l0Capability, masterDocument, i)
	'Capability overview
	dim capabilityPackage as EA.Package
	set capabilityPackage  = Repository.GetPackageByID(l0Capability.PackageID)
	'add modeldocument
	addModelDocumentForPackage masterDocument,capabilityPackage, l0Capability.Name, i, "BB_L0 Capability"
	i = i + 1
	'Ambitions
	dim ambitions
	set ambitions = getAmbitionsForL0capability(l0Capability)
	dim ambition
	for each ambition in ambitions
		addModelDocument masterDocument, "BB_Ambition", "Ambition -" & ambition.Name , ambition.ElementGUID, i
		i = i + 1
	next
	'L1 Capability details
	dim L1Capabilities
	set L1Capabilities = getL1Capabilities(l0Capability)
	dim L1Capability as EA.Element
	for each L1Capability in L1Capabilities
		dim l1Package as EA.Package
		set l1Package = Repository.GetPackageByID(L1Capability.PackageID)
		addModelDocumentForPackage masterDocument,l1Package, "Capability -" & L1Capability.Name, i, "BB_Capability_Details"
		i = i + 1
	next
	
	'return i
	createModelDocumentsForL0Capability = i
end function

function getL1Capabilities(l0Capability)
	dim sqlGetData
	sqlGetData = "select l1.Object_ID                                                        " & vbNewLine & _
				" from t_object l1                                                          " & vbNewLine & _
				" inner join t_connector l1l0 on l1l0.End_Object_ID = l1.Object_ID          " & vbNewLine & _
				" 							and l1l0.Stereotype = 'ArchiMate_Composition'   " & vbNewLine & _
				" where l1.Stereotype = 'ArchiMate_Capability'                              " & vbNewLine & _
				" and l1L0.Start_Object_ID  = " & l0Capability.ElementID & "                " & vbNewLine & _
				" order by l1.Name                                                          "
	dim result
	set result = getElementsFromQuery(sqlGetData)
	'return
	set getL1Capabilities = result
end function

function getAmbitionsForL0capability(l0Capability)
	dim sqlGetData
	sqlGetData = "select distinct o.Object_ID                                               " & vbNewLine & _
				" from t_object o                                                           " & vbNewLine & _
				" inner join t_connector rz on rz.End_Object_ID = o.Object_ID               " & vbNewLine & _
				" 							and rz.Stereotype = 'ArchiMate_Realization'     " & vbNewLine & _
				" inner join t_object g on g.Object_ID = rz.Start_Object_ID                 " & vbNewLine & _
				" 							and g.Stereotype = 'Elia_Gap'                   " & vbNewLine & _
				" inner join t_connector gc on gc.Start_Object_ID = g.Object_ID             " & vbNewLine & _
				" 							and gc.Stereotype = 'ArchiMate_Association'     " & vbNewLine & _
				" inner join t_object l2 on  l2.Object_ID = gc.End_Object_ID                " & vbNewLine & _
				" 							and l2.Stereotype = 'ArchiMate_Capability'      " & vbNewLine & _
				" inner join t_connector l2l1 on l2l1.End_Object_ID = l2.Object_ID          " & vbNewLine & _
				" 							and l2l1.Stereotype = 'ArchiMate_Composition'   " & vbNewLine & _
				" inner join t_object l1 on	l1.Object_ID = l2l1.Start_Object_ID             " & vbNewLine & _
				" 							and l1.Stereotype = 'ArchiMate_Capability'      " & vbNewLine & _
				" inner join t_connector l1l0 on l1l0.End_Object_ID = l1.Object_ID          " & vbNewLine & _
				" 							and l1l0.Stereotype = 'ArchiMate_Composition'   " & vbNewLine & _
				" where o.Stereotype = 'ELia_Ambition'                                      " & vbNewLine & _
				" and l1L0.Start_Object_ID = " & l0Capability.ElementID & "                 " & vbNewLine & _
				" order by 1                                                                "
	dim result
	set result = getElementsFromQuery(sqlGetData)
	'return
	set getAmbitionsForL0capability = result
end function

function getL0Capabilities(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	dim sqlGetData
	sqlGetData = "select * from t_object o                                          " & vbNewLine & _
				" inner join t_objectproperties tv on tv.Object_ID = o.Object_ID   " & vbNewLine & _
				" 									and tv.Property = 'Level'      " & vbNewLine & _
				" 									and tv.Value = 'L0'            " & vbNewLine & _
				" where o.Stereotype = 'ArchiMate_Capability'                      " & vbNewLine & _
				" and o.Package_ID in (" & packageTreeIDString & ")                "
	dim result
	set result = getElementsFromQuery(sqlGetData)
	'return
	set getL0Capabilities = result
end function

function createArchitectureBlueprintMasterDocument(documentPackage, package)
	dim documentTitle
	dim documentVersion
	dim documentName
	dim masterDocumentName
	dim documentAlias
	dim documentStatus
	
	'get document info
	documentTitle = "Architecture Blueprint " & package.Name
	documentAlias = documentTitle
	documentName =  documentTitle
	documentStatus = package.Element.Status
	documentVersion = package.Version
	masterDocumentName = documentTitle & " v." & documentVersion
	'Remove older version of the master document
	removeMasterDocumentDuplicates documentPackage.PackageGUID, documentTitle
	'then create a new master document
	dim masterDocument as EA.Package
	set masterDocument = addMasterDocumentWithDetailTags (documentPackage.PackageGUID, masterDocumentName, documentAlias, documentName, documentTitle, documentVersion, documentStatus)
	'return
	set createArchitectureBlueprintMasterDocument = masterDocument
end function

function getDocumentPackage(package)
	dim subPackage as EA.Package
	for each subPackage in package.Packages
		if lcase(subPackage.Name) = "document" then
			'return and exit
			set getDocumentPackage = subPackage
			exit function
		end if
	next
	'create the new document package
	set subPackage = package.Packages.AddNew("Document", "")
	subPackage.Update
	'return
	set getDocumentPackage = subPackage
end function

'main