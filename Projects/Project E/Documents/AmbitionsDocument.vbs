'[path=\Projects\Project E\Documents]
'[group=Documents]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
'
'
' Script Name: AmbitionsDocument
' Author: Geert Bellekens
' Purpose: Create the virtual document for the Business Book based on the selected pacakge
' Date: 2016-12-16
'


dim AmbitionsDocument
set AmbitionsDocument = new Document
AmbitionsDocument.Name = "Ambitions"
AmbitionsDocument.Description = "A Word document with the details of the ambitions in a table"
AmbitionsDocument.ValidQuery = "select top 1 o.Object_ID from t_object o where o.Stereotype = 'Elia_Ambition' and o.Package_ID in (#Branch#)"

sub GenerateAmbitionsDocument
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
	project.RunReport package.PackageGUID, "AMB_AmbitionsMain", fileName
	CreateObject("WScript.Shell").Run("""" & fileName & """")
end sub