'[path=\Projects\Project A\Element Group]
'[group=Element Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.DocGenUtil

' Script Name: Generate CMS Actors Document
' Author: Kristof Smeyers
' Purpose: Generate the virtual document in EA for a document with CMS Actors
' Date: 18/01/2018

'*************configuration*******************
'***update GUID to match the GUID of the "Documents" packages of this model***
dim documentsPackageGUID
documentsPackageGUID = "{7C6C8274-3C3C-426c-967E-5D39F50E2934}"
'*************configuration*******************

const outPutName = "Generate CMS Actor"

sub OnProjectBrowserElementScript()
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	Repository.WriteOutput outPutName, now() & " " & "*** Start script *** " & outputName, 0
	
	' Get the selected element
	dim selectedElement as EA.Element
	set selectedElement = Repository.GetContextObject()
	dim eObject_Type 'Element Object_Type
	eObject_Type = selectedElement.ObjectType
	
	Select Case eObject_Type 
		Case "Actor"
			'create virtual document
			createMasterDocument(selectedElement)
			
		Case Else
			Repository.WriteOutput outPutName, "The selected element is not an Actor", 0
			Session.Prompt "The selected element is not an Actor", promptOK
	End Select
	
	Repository.WriteOutput outPutName, now() & " " & "*** End script ***", 0
end sub

sub createMasterDocument(selectedElement)
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
		documentAlias = selectedElement.Name
		documentName = selectedElement.Name
		documentTitle = "Report - " & documentName
		documentStatus = "Voor implementatie / Pour implémentation"
		masterDocumentName = documentTitle & " v" & documentVersion
		
		'Remove older version of the master document
		'removeMasterDocumentDuplicates FDDocumentsPackageGUID, documentAlias & " - FD - " & abbreviation & " - " & functionalName
		'-> to correct
		
		'then create a new master document
		Repository.WriteOutput outPutName, now() & " " & "Creation of master document...", 0
		dim masterDocument as EA.Package
		set masterDocument = addMasterDocumentWithDetailTags (documentsPackageGUID,masterDocumentName,documentAlias,documentName,documentTitle,documentVersion,documentStatus)
		'set createMasterDocument = masterDocument
	end if
	
	' Add different Model documents to the Master document
	dim i
	i = 1
	
	dim eObject_Type 'Element Object_Type
	eObject_Type = selectedElement.ObjectType
	
	'get parent package of selectedElement
	dim parentpackage as EA.Package
	Repository.WriteOutput outPutName, now() & ": " & "Parent Package: " & selectedElement.PackageID, 0
	set parentpackage = Repository.GetPackageByID(selectedElement.PackageID)
	
	Select Case eObject_Type 
		Case "Actor"			
			'add model document using template
			addModelDocument masterDocument, "CMS Actors", documentName & " - " & " Main", parentpackage.PackageGUID, i
			Repository.WriteOutput outPutName, now() & " " & eObject_Type & ": " & "Added Model Document: " & documentName, 0
			i = i + 1		
		Case Else
			Repository.WriteOutput outPutName, now() & " " & eObject_Type & " - " & "Object_Type not recognized", 0
			'do nothing			
	End Select
	
	'finished, refresh model view to make sure the order is reflected in the model.
	Repository.RefreshModelView(masterDocument.PackageID)
	
	'end of process
	Repository.WriteOutput outPutName, now() & " " & "End creation of master document", 0
end sub


OnProjectBrowserElementScript