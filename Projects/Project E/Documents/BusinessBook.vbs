'[path=\Projects\Project E\Documents]
'[group=Documents]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
'
'
' Script Name: Generate Business Book
' Author: Geert Bellekens
' Purpose: Create the virtual document for the Business Book based on the selected pacakge
' Date: 2016-12-16
'

'

'const outPutName = "Export Data Model"
const businessBookDocumentsPackageGUID = "{059FE4EE-E32E-43b6-AE66-15DBB0AD1298}"
const capabilityPackageGUID = "{414F318A-079B-4e61-9558-FD65F92322E7}"

dim BusinessBookDocument
set BusinessBookDocument = new Document
BusinessBookDocument.Name = "BusinessBook"
BusinessBookDocument.Description = "A Word document with all the details of all capabilities in this package branch"
BusinessBookDocument.ValidQuery = "select top 1 o.Object_ID from t_object o where o.Stereotype = 'ArchiMate_Capability' and o.Package_ID in (#Branch#)"

sub BusinessBook
	'get selected package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	if package is nothing then
		exit sub
	end if
	'do the actual work
	createBusinessBook package
end sub

function createBusinessBook(package)

	'create master document
	dim masterDocument as EA.Package
	set masterDocument = createBusinessBookMasterDocument(package)
	if masterDocument is nothing then
		Repository.WriteOutput outPutName, now() & " ERROR: Can't create master document!" , 0
		exit function
	end if
	'counter
	dim i
	i = 0
	'Introduction (top level diagram)
	i = addBusinessBookIntroduction(masterDocument, i)
	'get all L0 capabilities
	dim l0Capabilities
	set l0Capabilities = getL0Capabilities(package)
	dim l0Capability as EA.Element
	for each l0Capability in l0Capabilities
		i = createModelDocumentsForL0Capability(l0Capability, masterDocument, i)
	next
	generateMasterDocument masterDocument
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

function addBusinessBookIntroduction(masterDocument, i)
	'get introduction package
	dim capabilityPackage
	set capabilityPackage = Repository.GetPackageByGuid(capabilityPackageGUID)
	'add modeldocument
	addModelDocumentForPackage masterDocument,capabilityPackage, "Introduction", i, "BB_Overview"
	i = i + 1
	'return i
	addBusinessBookIntroduction = i
end function

function createBusinessBookMasterDocument(package)
	dim documentTitle
	dim documentVersion
	dim documentName
	dim masterDocumentName
	dim documentAlias
	dim documentStatus
	
	'get document info
	documentTitle = "Business Book " & package.Name
	documentAlias = documentTitle
	documentName =  documentTitle
	documentStatus = package.Element.Status
	documentVersion = package.Version
	masterDocumentName = documentTitle & " v." & documentVersion
	'Remove older version of the master document
	removeMasterDocumentDuplicates businessBookDocumentsPackageGUID, documentTitle
	'then create a new master document
	dim masterDocument as EA.Package
	set masterDocument = addMasterDocumentWithDetailTags (businessBookDocumentsPackageGUID,masterDocumentName,documentAlias,documentName,documentTitle,documentVersion,documentStatus)
	'return
	set createBusinessBookMasterDocument = masterDocument
end function

'main