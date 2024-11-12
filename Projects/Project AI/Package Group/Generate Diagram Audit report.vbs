'[path=\Projects\Project AI\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Generate Diagram Audit Report
' Author: Geert Bellekens
' Purpose: Generate Diagram Audit Report
' Date: 2024-06-17
'
'name of the output tab
const outPutName = "Generate Diagram Audit Report"
const auditStartDate = "2024-01-01"
const auditEndDate = "2024-07-18"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'set timestamp
	Repository.WriteOutput outPutName, now() &  " Starting " & outPutName , 0
	'Actually do the thing
	generateDiagramAuditReport
	'set timestamp
	Repository.WriteOutput outPutName, now() &  " Finished " & outPutName , 0
end sub


function generateDiagramAuditReport()
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()

	dim diagram as EA.Diagram
	dim docgen as EA.DocumentGenerator
	set docgen = Repository.CreateDocumentGenerator()
	
	'get the packageID's of all packages under this branch
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	docgen.NewDocument("diagram audit")
	'set the project constants
	docgen.SetProjectConstant "DateFrom", auditStartDate
	docgen.SetProjectConstant "DateTo", auditEndDate
	docgen.SetProjectConstant "Subject", package.Name
	'set the cover page
	docgen.InsertCoverPageDocument("DAU_Coverpage")
	'New diagrams title
	docgen.DocumentPackage package.PackageID, 0, "DAU_New Diagrams Title"
	dim diagrams
	set diagrams = getRecentNewDiagrams(packageTreeIDString)
	for each diagram in diagrams
		docgen.DocumentDiagram diagram.DiagramID, 1, "DAU_Diagram Audit Main"
	next
	'Modified diagrams title
	docgen.DocumentPackage package.PackageID, 0, "DAU_Modified Diagrams Title"
	dim modifiedDiagrams
	set modifiedDiagrams = getRecentModifiedDiagrams(packageTreeIDString)
	for each diagram in modifiedDiagrams
		docgen.DocumentDiagram diagram.DiagramID, 1, "DAU_Diagram Audit Main"
	next
	
	'get the all deleted diagrams part
	docgen.DocumentPackage package.PackageID, 1, "DAU_All Deleted Diagrams"
	'save the document
	dim fileName
	fileName = getUserSelectedDocumentFileName()
	docgen.SaveDocument fileName, dtDOCX
	'open the file
	CreateObject("WScript.Shell").Run("""" & fileName & """")
end function

function getUserSelectedDocumentFileName()
	dim filename
	filename = Repository.GetProjectInterface().GetFileNameDialog ("", "Word documents |*.docx", 1, 2 ,"", 1) 
	getUserSelectedDocumentFileName = fileName
end function

function getRecentNewDiagrams(packageTreeIDString)
	'find the recent diagrams
	dim sqlGetData
	sqlGetData = "select d.Diagram_ID                                   " & vbNewLine & _
				" from t_diagram d                                      " & vbNewLine & _
				" where d.CreatedDate between '" & auditStartDate & "'  " & vbNewLine & _
				"                       and '" & auditEndDate &   "'    " & vbNewLine & _
				" and d.Package_ID in (" & packageTreeIDString & ")  "
	dim result
	set result = getDiagramsFromQuery(sqlGetData)
	'return
	set getRecentNewDiagrams = result
end function

function getRecentModifiedDiagrams(packageTreeIDString)
	'find the recent diagrams
	dim sqlGetData
	sqlGetData = "select d.Diagram_ID                                   " & vbNewLine & _
				" from t_diagram d                                      " & vbNewLine & _
				" where d.ModifiedDate between '" & auditStartDate & "' " & vbNewLine & _
				"                       and '" & auditEndDate &   "'    " & vbNewLine & _
				" and not d.CreatedDate between '" & auditStartDate & "'" & vbNewLine & _
				"                       and '" & auditEndDate &   "'    " & vbNewLine & _				
				" and d.Package_ID in (" & packageTreeIDString & ")  "
	dim result
	set result = getDiagramsFromQuery(sqlGetData)
	'return
	set getRecentModifiedDiagrams = result
end function

main