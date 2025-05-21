'[path=\Projects\Project K\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include


'
' Script Name: Create Application Model Structure
' Author: Geert Bellekens
' Purpose: Create a copy of the Application model structure and rename where needed to application name
' Date: 2020-03-27

const outPutName = "Create Application Model Structure"
const applicationTemplatePackage = "{2A9CE59E-3173-4319-8B93-BB6E35062298}"
const toReName = "[Application Name]"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Starting Create Application Model Structure", 0
	'get selected package
	dim package as EA.Package

	set package = Repository.GetTreeSelectedPackage()
	'exit if not selected
	if package is nothing then
	 msgbox "Please select a package before running this script"
	 exit sub
	end if
	'start work
	createApplicationModelStructure package
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Finished Create Application Model Structure", 0
end sub

function createApplicationModelStructure(package)
	'ask for new name
	dim newName
	newName = ""
	newName = InputBox("Please enter the name of the new applicaiton", "Application")
	'exit if no name entered
	if len(newName) = 0 then
	exit function
	end if
	dim templatePackage as EA.Package
	set templatePackage    = Repository.GetPackageByGuid(applicationTemplatePackage)
	Repository.WriteOutput outPutName, now() & " Copying template package", 0
	'copy template package
	dim clonedPackage as EA.Package
	set clonedPackage = templatePackage.Clone()
	'move to target package
	clonedPackage.ParentID = package.PackageID
	clonedPackage.Update
	Repository.WriteOutput outPutName, now() & " Renaming [Application Name]", 0
	'get current author name
	dim author
	author = getCurrentAuthorName(clonedPackage)
	'rename
	rename newName, clonedPackage, author
	'select in project browser
	Repository.ShowInProjectView clonedPackage
end function

function getCurrentAuthorName(package)
	dim temp as EA.Element
	set temp = package.Elements.AddNew("temp", "Boundary")
	dim author
	author = temp.Author
	getCurrentAuthorName = author
end function

function rename(newName, item, author)
	'process item itself
	renameItem newName, item
	setAuthor item, author
	'process diagrams
	dim diagram as EA.Package
	for each diagram in item.Diagrams
		renameItem newName, diagram
		setAuthor diagram, author
	next
	'process owned elements
	dim element as EA.Element
	for each element in item.Elements
		rename newName, element, author
		setAuthor element, author
	next
	'process subpackages
	if item.ObjectType = otPackage then
		dim subPackage as EA.Package
		for each subPackage in item.Packages
			rename newName, subPackage, author
		next
	end if
end function

function setAuthor(item, author)
	if item.ObjectType = otPackage then
		setAuthor item.Element, author
	else
		item.Author = author
		item.update
	end if
end function

function renameItem (newName, item)
	dim renamedName
	renamedName = replace(item.Name, toReName, newName)
	if renamedName <> item.Name then
		item.name = renamedName
		item.Update
	end if
end function

main