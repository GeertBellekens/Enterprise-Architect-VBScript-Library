'[path=\Projects\Project A\Element Group]
'[group=Element Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Atrias Scripts.DocGenUtil

' Script Name: Generate BI-Reporting FD
' Author: Matthias Van der Elst
' Purpose: Generate FD for the requested BI-Reporting
' Date: 03/04/2017

dim documentsPackageGUID
'*************configuration*******************
documentsPackageGUID = "{CE66E294-F6ED-4292-B9AF-CD2EFCF473C5}"
'*************configuration*******************

const outPutName = "Create BI & Reporting document"

' Project Browser Script main function
sub OnProjectBrowserElementScript()
	' Get the selected element
	dim selectedElement as EA.Element
	set selectedElement = Repository.GetContextObject()
	dim eStereoType 'Element stereotype
	eStereoType = selectedElement.Stereotype
	if eStereoType = "REP_BI-Group" then
		'create output tab
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		'report start of process
		Repository.WriteOutput outPutName, "Starting creation of BI & Reporting document at " & now(), 0
		'create document
		GenerateFD(selectedElement)
		'report end of process
		Repository.WriteOutput outPutName, "Finished creation of BI & Reporting document at " & now(), 0
	else
		Session.Prompt "The selected element is not of the correct type", promptOK
	 end if
	
end sub

sub GenerateFD(selectedElement)
	'ask user for document version
	dim documentVersion
	documentVersion = InputBox("Please enter the version for this document", "Document Version", "x.y")
	'get document info
	dim masterDocumentName,documentAlias,documentName,documentTitle,documentStatus
	documentAlias = selectedElement.Name
	documentName = selectedElement.Name
	documentTitle = "BI and Reporting - " & documentName
	documentStatus = "Voor implementatie / Pour implémentation"
	masterDocumentName = documentTitle & " v" & documentVersion
	'first create a master document
	dim masterDocument as EA.Package
	set masterDocument = addMasterDocumentWithDetailTags (documentsPackageGUID,masterDocumentName,documentAlias,documentName,documentTitle,documentVersion,documentStatus)
	dim i
	i = 1
	
	'add the BI Group
	documentName = selectedElement.Name & " - BI-Group"
	addModelDocument masterDocument, "REP_BI-Group",documentName, selectedElement.ElementGUID, i	
	i = i + 1
	
	'add the BI Report
	documentName = selectedElement.Name & " - BI Report"
	dim reports
	set reports = CreateObject("System.Collections.ArrayList")
	dim sqlGetReport
	sqlGetReport = 	"select rep.object_id " & _
					"from ((t_object bi " & _
					"inner join t_connector con " & _
					"on bi.object_id = con.End_Object_ID) " & _
					"inner join t_object rep " & _
					"on con.Start_Object_ID = rep.Object_ID) " & _
					"where bi.Object_ID = '" & selectedElement.ElementID & "' "
					
	set reports = getElementsFromQuery(sqlGetReport)
	
	dim report as EA.Element
	for each report in reports
		addModelDocument masterDocument, "REP_BI Report",documentName, report.ElementGUID, i	
		i = i + 1
	next
	
	
	'finished, refresh model view to make sure the order is reflected in the model.
	Repository.RefreshModelView(masterDocument.PackageID)
	
end sub

OnProjectBrowserElementScript