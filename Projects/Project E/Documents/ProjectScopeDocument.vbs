'[path=\Projects\Project E\Documents]
'[group=Documents]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
'
'
' Script Name: Project Scope document
' Author: Geert Bellekens
' Purpose: Generate the project scope document for the selected package
' Date: 2024-05-17
'


dim scopeDocument
set scopeDocument = new Document
scopeDocument.Name = "Project Scope"
scopeDocument.Description = "A Word document that contains the scope of the projects in this package in terms of Requirements, Gaps, Ambitions and Capabilities"
scopeDocument.ValidQuery = "select top 1 o.Object_ID from t_object o where o.Stereotype = 'Elia_Project' and o.Package_ID in (#Branch#)"

sub GenerateScopeDocument
	'get selected package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	if package is nothing then
		exit sub
	end if
	dim project as EA.Project
	set project  = Repository.GetProjectInterface()
	dim fileName
	fileName = getUserSelectedDocumentFileName()
	project.RunReport package.PackageGUID, "PRJ_Scope", fileName
	CreateObject("WScript.Shell").Run("""" & fileName & """")
end sub