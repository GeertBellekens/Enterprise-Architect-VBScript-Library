'[path=\Projects\Project E\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Generate Documentation
' Author: Geert Bellekens
' Purpose: Generate the documentation for this package if it is a master document
' Date: 2025-03-17

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
	'check the stereotype fo the package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage
	if not lcase(package.Element.Stereotype) = "master document" then
		Repository.WriteOutput outPutName, now() & " No «master document» package selected - exiting script", 0
		exit function
	end if
	'ask user if they want to save a snapshot of the model
	dim exportToXmi
	dim userIsSure
	userIsSure = Msgbox("Also create snapshot of the model?", vbYesNo+vbQuestion, "Create Model Snapshot?")
	if userIsSure = vbYes then
		exportToXmi = true
	else
		exportToXmi = false
	end if
	'export document
	dim project as EA.Project
	set project  = Repository.GetProjectInterface()
	dim fileName
	fileName = getUserSelectedDocumentFileName()
	project.RunReport package.PackageGUID, "", fileName
	if exportToXmi then
		'export to xmi and commit to git
		exportControlledRootPackages
		'TODO commit to git
		dim shell 
		set shell  = CreateObject("WScript.Shell")
		shell.Run "cmd.exe /K " & "git status", 1, False  
	end if
end function

function test
	dim shell 
	set shell  = CreateObject("WScript.Shell")
	shell.Run "cmd.exe /K " & "git status", 1, False  
end function

test
'main
