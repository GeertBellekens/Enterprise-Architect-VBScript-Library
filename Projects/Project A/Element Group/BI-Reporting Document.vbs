'[path=\Projects\Project A\Element Group]
'[group=Element Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.DocGenUtil

' Script Name: Generate BI-Reporting Document
' Author: Alain Van Goethem
' Purpose: Generate the virtual document in EA for a BI-Reporting document based on the selected BI-Reporting element
' Date: 18/01/2018

'*************configuration*******************
'***update GUID to match the GUID of the "Documents" packages of this model***
dim documentsPackageGUID
documentsPackageGUID = "{CE66E294-F6ED-4292-B9AF-CD2EFCF473C5}"
'*************configuration*******************

const outPutName = "Generate BI-Reporting document"

sub OnProjectBrowserElementScript()
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	Repository.WriteOutput outPutName, now() & " " & "*** Start script *** " & outputName, 0
	
	' Get the selected element
	dim selectedElement as EA.Element
	set selectedElement = Repository.GetContextObject()
	dim eStereoType 'Element stereotype
	eStereoType = selectedElement.Stereotype
	
	Select Case eStereoType 
		Case "BI-REP", "BI-Dataset"
			'create virtual document
			createReportingDocument(selectedElement)
			
		Case Else
			Repository.WriteOutput outPutName, "The selected element is not of stereotype BI-REP or BI-Dataset", 0
			Session.Prompt "The selected element is not of stereotype BI-REP or BI-Dataset", promptOK
	End Select
	
	Repository.WriteOutput outPutName, now() & " " & "*** End script ***", 0
end sub

sub createReportingDocument(selectedElement)
	dim masterDocument as EA.Package
	set masterDocument = createMasterDocument(selectedElement)
	if masterDocument is nothing then
		'something went wrong or the user cancelled
		exit sub
	end if
	
	'Model document counter
	dim i
	i = 1
	
	dim eStereoType 'Element stereotype
	eStereoType = selectedElement.Stereotype
	
	'get parent package of selectedElement
	dim parentpackage as EA.Package
	set parentpackage = Repository.GetPackageByID(selectedElement.PackageID)
	
	Select Case eStereoType 
		Case "BI-Dataset"			
			'add model document using template
			addModelDocumentForPackage masterDocument,parentpackage, selectedElement.name & " - " & " Main", i, "BI_Dataset Main Template"
			'addModelDocument masterDocument, "BI_Dataset Main Template", documentName & " - " & " Main", parentpackage.PackageGUID, i
			Repository.WriteOutput outPutName, now() & " " & eStereotype & ": " & "Added Model Document: " & selectedElement.name & " - " & " Main", 0
			i = i + 1
		Case "BI-REP"
			'add model document using template
			addModelDocumentForPackage masterDocument,parentpackage, selectedElement.name & " - " & " Main", i, "BI_Report Main Template"
			'addModelDocument masterDocument, "BI_Report Main Template", documentName & " - " & " Main", parentpackage.PackageGUID, i
			Repository.WriteOutput outPutName, now() & " " & eStereotype & ": " & "Added Model Document: " & selectedElement.name & " - " & " Main", 0
			i = i + 1			
		Case Else
			Repository.WriteOutput outPutName, now() & " " & eStereoType & " - " & "Stereotype not recognized", 0
			'do nothing			
	End Select
	
	'end of process
	Repository.WriteOutput outPutName, now() & " " & "End creation of master document", 0
	
	'finished, refresh model view to make sure the order is reflected in the model.
	Repository.RefreshModelView(0)
	'select the masterDocument in the project browser
	Repository.ShowInProjectView masterDocument
	'let the user know
	msgbox "Finished!" & vbNewline & vbNewline & "Please select the virtual document and press F8 to generate the document"
end sub


function createMasterDocument(selectedElement)
	'start of process
	Repository.WriteOutput outPutName, now() & " " & "Preparing creation of master document...", 0
	
	'ask user for document version
	dim documentVersion
	documentVersion = ""
	documentVersion = InputBox("Please enter the version for this document", "Document Version", "x.y.z")
	if documentVersion <> "" then
		Repository.WriteOutput outPutName, now() & " " & "Document version ok", 0
		'get document info
		dim masterDocumentName,documentAlias,documentName,documentTitle,documentStatus
		'determine the alias
		if selectedElement.Stereotype = "BI-REP" then
			documentAlias = "Report"
		else
			documentAlias = "Dataset"
		end if
		documentName = selectedElement.Name
		documentTitle = documentAlias & " - " & documentName
		documentStatus = "Voor implementatie / Pour implémentation"
		masterDocumentName = documentTitle & " - v. " & documentVersion
		
		'Remove older version of the master document
		removeMasterDocumentDuplicates documentsPackageGUID, documentTitle
		
		'then create a new master document
		Repository.WriteOutput outPutName, now() & " " & "Creation of master document...", 0
		dim masterDocument as EA.Package
		set masterDocument = addMasterDocumentWithDetailTags (documentsPackageGUID,masterDocumentName,documentAlias,documentName,documentTitle,documentVersion,documentStatus)
		set createMasterDocument = masterDocument
	else
		'return nothing
		set createMasterDocument = nothing
	end if
end function

OnProjectBrowserElementScript