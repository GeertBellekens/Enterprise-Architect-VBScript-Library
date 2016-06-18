'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.DocGenUtil
!INC Atrias Scripts.Util
'

'
' Script Name: Monitoring and Alerting Document
' Author: Geert Bellekens
' Purpose: Create the virtual document for the M&A document
' Date: 2016-06-15
'
dim documentsPackageGUID, businessProcessesPackageGUID, subProcessesPackageGUID
'*************configuration*******************
documentsPackageGUID = "{A15738BC-3B18-46be-8357-2190FC05436F}"
businessProcessesPackageGUID = "{7EAA1987-6FB1-427f-8BA1-2610ED339905}"
subProcessesPackageGUID = "{5D830EDF-0470-4d41-9358-93C2EB410521}"
'*************configuration*******************

const outPutName = "Create M&A document"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'report start of process
	Repository.WriteOutput outPutName, "Starting creation of M&A document at " & now(), 0
	'create document
	createMandADocument
	'report end of process
	Repository.WriteOutput outPutName, "Finished creation of M&A document at " & now(), 0
end sub

function createMandADocument()
	'ask user for document name
	dim documentName
	documentName = InputBox("Please enter the name for this document", "Document Name", "MIG-DGO-M&A Business Requirements v X.Y")
	'first create a master document
	dim masterDocument as EA.Package
	set masterDocument = addMasterDocument (documentsPackageGUID, documentName)
	dim i
	i = 1
	
	'get the processes package
	dim businessProcessesPackage as EA.Package
	set businessProcessesPackage = Repository.GetPackageByGuid(businessProcessesPackageGUID)
	'loop all process packages
	dim domainPackage as EA.Package
	for each domainPackage in businessProcessesPackage.Packages
		'add domain title
		addModelDocumentForPackage masterDocument,domainPackage,domainPackage.Name, i, "MA_Domain Title"
		i = i + 1
		'report processes
		i = reportOwnedProcesses(masterDocument, domainPackage, i)
		'get the subProcesses package for this domain
		dim subProcessPackage
		set subProcessPackage = getCorrespondingSubProcessPackage(domainPackage)
		if not subProcessPackage is nothing then
			'report subProcesses
			i = reportOwnedProcesses(masterDocument, subProcessPackage, i)
		end if
	next
	'reload the package to show the actual order
	Repository.RefreshModelView masterDocument.PackageID
end function

function reportOwnedProcesses(masterDocument, domainPackage, i)
		dim processes
		set processes = getMAndAProcesses(domainPackage)
		dim process as EA.Element
		for each process in processes
			'add model documents for process
			i = addProcessDocuments(masterDocument, process, i)
		next
	reportOwnedProcesses = i
end function

function getCorrespondingSubProcessPackage(domainPackage)
	'initialize
	set getCorrespondingSubProcessPackage = nothing
	dim subProcessesPackage as EA.Package
	set subProcessesPackage = Repository.GetPackageByGuid(subProcessesPackageGUID)
	dim subProcessPackage as EA.Package
	'find the subpackage with the same name
	for each subProcessPackage in subProcessesPackage.Packages
		if subProcessPackage.Name = domainPackage.Name then
			'found it
			set getCorrespondingSubProcessPackage = subProcessPackage
			exit for
		end if
	next
end function

function addProcessDocuments(masterDocument, process, i)
	'add section for the Business Process
	addModelDocument masterDocument, "MA_Business Process",process.Name, process.ElementGUID, i
	i = i + 1
	'add sections for each of the BAM specifications
	dim BAMspecifications
	set BAMspecifications = getBAMSpecifications(process)
	dim BAMSpecification as EA.Element
	for each BAMSpecification in BAMSpecifications
		'add section for the BAM Specification
		addModelDocument masterDocument, "MA_Specifications",BAMSpecification.Name, BAMSpecification.ElementGUID, i
		i = i + 1
	next
	addProcessDocuments = i
end function

function getMAndAProcesses(domainPackage)
	dim sqlGetProcesses
	sqlGetProcesses = "select bp.Object_ID from ((t_object bp " & _
				" inner join t_package bpp on bp.Package_ID = bpp.Package_ID) " & _
				" inner join t_diagram d on d.ParentID = bp.Object_ID) " & _
				" where  " & _
				" bp.Stereotype in ('Activity', 'BusinessProcess') " & _
				" and exists (select bam.Object_ID from t_object bam " & _
				" where bam.Package_ID = bp.Package_ID " & _
				" and bam.Stereotype = 'BAM_Specification') " & _
				" and bpp.Parent_ID =" & domainPackage.PackageID & _
				" order by bp.Name"
    set getMAndAProcesses = getElementsFromQuery(sqlGetProcesses)
end function

function getBAMSpecifications(process)
	dim sqlGetBamSpecifications
	sqlGetBamSpecifications = "select bam.Object_ID " & _
							" from t_object bam " & _
							" where bam.Stereotype = 'BAM_Specification' " & _
							" and bam.Package_ID = " & process.PackageID & _
							" order by bam.name"
	set getBAMSpecifications = getElementsFromQuery(sqlGetBamSpecifications)
end function


main