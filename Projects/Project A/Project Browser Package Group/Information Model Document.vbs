'[path=\Projects\Project A\Project Browser Package Group]
'[group=Project Browser Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

' Script Name: Information Model Document
' Author: Geert Bellekens
' Purpose: Generate the virtual document in EA for a Information Model Document based on a package containing MIG6 message definitions («DOCLibrary» packages)
' This script is supposed to be executed on a package representing a domain.
' Date: 2018-08-17

'*************configuration*******************
'***update GUID to match the GUID of the "Documents" packages of this model***
dim documentsPackageGUID
documentsPackageGUID = "{762AFB36-8B7E-46d7-A2C2-28C3E7700319}"
'*************configuration*******************

const outPutName = "Information Model Document"

sub main()
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get the selected package
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage
	if selectedPackage is nothing then
		'no package selected
		Msgbox "Please select a package before executing this script!"
		exit sub
	end if
	'report progress
	Repository.WriteOutput outPutName, now() & " Starting generating Information Model Document for '" & selectedPackage.Name & "'" , selectedPackage.Element.ElementID
	
	'create the actual virtual document
	createInformationModelDocument(selectedPackage)
	'report progress
	Repository.WriteOutput outPutName, now() & " Finished generating Information Model Document for '" & selectedPackage.Name & "'" , selectedPackage.Element.ElementID
end sub

function createInformationModelDocument(selectedPackage)
	'get the documents package
	dim documentsPackage as EA.Package
	set documentsPackage = Repository.GetPackageByGuid(documentsPackageGUID)
	if documentsPackage is nothing then
		msgbox "Could not find documents package with GUID " & documentsPackageGUID
		exit function
	end if
	'create the virtual document
	'get document info
	dim masterDocumentName,documentAlias,documentName,documentTitle,documentStatus,documentVersion
	documentAlias = "UMIG" 
	documentName = selectedPackage.Name
	documentTitle = "UMIG - IM - " & selectedPackage.Alias & " - 05 - " & selectedPackage.Name & " Exchanged Information"
	documentStatus = "For Implementation"
	documentVersion = selectedPackage.Version
	masterDocumentName = documentTitle & " v" & selectedPackage.Version
	'Remove older version of the master document
	removeMasterDocumentDuplicates documentsPackageGUID, masterDocumentName
	'first create a master document
	dim masterDocument as EA.Package
	set masterDocument = addMasterDocumentWithDetailTags (documentsPackageGUID,masterDocumentName,documentAlias,documentName,documentTitle,documentVersion,documentStatus)
	dim i
	i = 1
	'loop subdomains
	dim subDomainPackage as EA.Package
	for each subDomainPackage  in selectedPackage.Packages 
		'create the part for the subdomain
		i = createDocumentForSubdomain(masterDocument, subDomainPackage, i)
	next
	'add FIS message table document
	addModelDocumentForPackage masterDocument, selectedPackage, "FIS - Message reference", i, "BRIM_FIS Message reference"
	'finished, refresh model view to make sure the order is reflected in the model.
	Repository.RefreshModelView(masterDocument.PackageID)
	'select the masterDocument in the project browser
	Repository.ShowInProjectView masterDocument
end function

function createDocumentForSubdomain(masterDocument, subDomainPackage, i)
	'create the subdomain model document
	addModelDocumentForPackage masterDocument,subDomainPackage, "Subdomain " & subDomainPackage.Name, i, "BRIM_PackageHeader"
	'addModelDocumentForPackage(masterDocument,package,name, treepos, template)
	i = i + 1
	'loop «Doclibrary» packages
	dim docLibraryPackage as EA.Package
	for each docLibraryPackage in subDomainPackage.Packages
		Repository.WriteOutput outPutName, now() & " Processing '" & docLibraryPackage.Name & "'" , docLibraryPackage.Element.ElementID
		'add part 1
		addModelDocumentForPackage masterDocument,docLibraryPackage, docLibraryPackage.Name & " part1", i, "BRIM_Doclibrary part1"
		i = i + 1
		'add part 2
		addModelDocumentForPackage masterDocument,docLibraryPackage, docLibraryPackage.Name & " part2", i, "BRIM_Doclibrary part2"
		i = i + 1
	next
	'return counter
	createDocumentForSubdomain = i
end function

main