'[path=\Projects\Project E\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Generate Documentation
' Author: Geert Bellekens
' Purpose: Generate the documentation for this package if it is a master document. 
' This will also generate xmi files for all controlled root packages, and commit/add and push them to git.
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
		dim paths
		set paths = exportControlledRootPackages()
		'get the directories from the paths
		dim directories
		set directories = getDirectories(paths)
		dim commitMessage
		commitMessage = inputBox("Please enter a commit message")
		if len(commitMessage) > 0 then
			'Add, commit and push in each directory
			dim shell 
			set shell  = CreateObject("WScript.Shell")
			dim directory
			for each directory in directories
				shell.Run "cmd.exe /K cd """ & directory & """ & git add . & git commit -am """ & commitMessage & """ & git push" , 1, False  
			next
		end if		
	end if
end function

function getDirectories(paths)
	dim directories
	set directories = CreateObject("Scripting.Dictionary")
	dim path
	for each path in paths
		'find last backslash
		dim endPos
		endPos = instrrev(path, "\")
		dim directory
		directory = left(path, endPos -1)
		if not directories.Exists(directory) then
			directories.Add directory, directory
		end if
	next
	'return
	set getDirectories = directories
end function

main
