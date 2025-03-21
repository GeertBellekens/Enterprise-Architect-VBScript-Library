'[path=\Projects\Project E\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Documents.Include

'
' Script Name: Generate Documentation
' Author: Geert Bellekens
' Purpose: Export the Project Dependencies under this package, and the links to gaps, projects and capabilities
' Date: 2023-06-16

const outPutName = "Generate Documentation"

sub main
	'reset output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'report progress
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName, 0
	'do the actual work
	generateDocumentation()
	'report progress
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName, 0
end sub

function generateDocumentation()
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage
	generateDocumentsForPackage(package)
end function

main
