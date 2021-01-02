'[path=\Projects\Project S\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Generate Documents
' Author: Geert Bellekens
' Purpose: Generate the documents for each individual model document here
' Date: 2020-05-28

const outPutName = "Generate Documents"

sub main

	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get selected package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage
	'exit if not selected
	if package is nothing then
		msgbox "Please select a package before running this script"
		exit sub
	end if
	'inform user
	Repository.WriteOutput outPutName, now() & " Starting Generate documents for '" & package.Name & "'" , 0
	'do the actual work
	generateDocuments package
	'inform user
	Repository.WriteOutput outPutName, now() & " Finished Generate documents for '" & package.Name & "'" , 0
end sub

function generateDocuments(package)
	'get user selected folder
	dim selectedFolder
	set selectedFolder = New FileSystemFolder
	'set selectedFolder = selectedFolder.getUserSelectedFolder("C:\Temp\")
	set selectedFolder = selectedFolder.getUserSelectedFolder("")
	if not selectedFolder is nothing then
		dim modelDocument as EA.Element
		for each modelDocument in package.Elements
			if modelDocument.Stereotype = "model document" then
				generateDocument modelDocument, selectedFolder
			end if
		next
	end if
end function

function generateDocument(modelDocument, selectedFolder)
	'inform user
	Repository.WriteOutput outPutName, now() & " Generating '" & modelDocument.Name & "'" , 0
	'get document generator
	dim docgen as EA.DocumentGenerator
	set docgen = Repository.CreateDocumentGenerator()
	'get template
	dim templateName
	templateName = getTemplate(modelDocument)
	'create new document
	docgen.NewDocument("")
	'add contents
	dim attribute as EA.Attribute
	for each attribute in modelDocument.Attributes
		dim packageElement as EA.Element
		set packageElement = Repository.GetElementByID(attribute.ClassifierID)
		dim package as EA.Package
		set package = Repository.GetPackageByGuid(packageElement.ElementGUID)
		if not package is nothing then
			dim success
			success = docgen.DocumentPackage(package.PackageID, 0, templateName)
			'report error if any
			if not success then
				Repository.WriteOutput outPutName, now() & " ERROR: '" & docgen.GetLastError & "'" , 0
			end if
		end if
	next
	'save document
	dim path
	path = selectedFolder.FullPath & "\" & modelDocument.Name & ".docx"
	dim saveSuccess
	saveSuccess = docgen.SaveDocument(path, dtDOCX)
	'report error if any
	if not saveSuccess then
		Repository.WriteOutput outPutName, now() & " ERROR: '" & docgen.GetLastError & "'" , 0
	end if
end function

function getTemplate(modelDocument)
	dim templateName
	templateName = ""
	dim tag as EA.TaggedValue
	for each tag in modelDocument.TaggedValues
		if tag.Name = "RTFTemplate" then
			templateName = tag.Value
		end if
	next
	'return
	getTemplate = templateName
end function

main