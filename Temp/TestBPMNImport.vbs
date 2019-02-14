'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
	dim selectedPackage as EA.Package
	set selectedPackage  = Repository.GetTreeSelectedPackage
	dim project as EA.Project
	set project = Repository.GetProjectInterface
	project.ImportPackageXMI project.GUIDtoXML(selectedPackage.PackageGUID), "C:\Temp\test 5.bpmn", 1, 0
end sub

main