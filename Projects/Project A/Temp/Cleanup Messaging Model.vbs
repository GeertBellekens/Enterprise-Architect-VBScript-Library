'[path=\Projects\Project A\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Cleanup Messaging Model
' Author: Geert Bellekens
' Purpose: moves the «MA» and «invEnvelope» elements to a separate folder in the messaging model 
' moves the schema file into the «DOCLibrary» package and moves «DOCLibrary» package to the Messages package.
' each package will remain in its original structure
' Date: 2018-06-14
'

const MessagesRootGUID = "{0EEBD3FD-9B2D-4de0-9252-5C8F142194AC}"
const outPutName = "Cleanup Messaging Model"

sub main
	'get the Messages and Messaging root packages
	dim messagesRootPackage
	set messagesRootPackage = Repository.GetPackageByGuid(MessagesRootGUID)
	
	'get the selected package
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage()
	if selectedPackage is nothing then
		msgbox "Please select a packages in the project browser before starting this script."
		exit sub
	end if
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'report progress
	Repository.WriteOutput outPutName, now() & " Cleanup started from '"& selectedPackage.Name &"'", 0
	
	'get all the MA's in the selected package and subpackages
	dim sqlGetDocLibraries
	sqlGetDocLibraries = "select o.Object_ID from t_object o " & _
					"where o.Stereotype = 'DOCLibrary' " & _
					"and o.Package_ID in (" & getPackageTreeIDString(selectedPackage) & ")"
	
	dim docLibraries
	set docLibraries = getElementsFromQuery(sqlGetDocLibraries) 'all doclibraries
	'create the package dictionaries
	dim messagesPackages 
	set messagesPackages = CreateObject("Scripting.Dictionary")
	'add the the messagesroot package with the parent ID of the selected package
	messagesPackages.Add selectedPackage.ParentID, messagesRootPackage
	'loop doclibraries and process them individually
	dim docLibraryElement as EA.Element
	for each docLibraryElement in docLibraries
		'tell the user what we are doing
		Repository.WriteOutput outPutName, now() & " Processing Package '"& docLibraryElement.Name &"'", 0
		processDocLibrary docLibraryElement, messagesRootPackage,  messagesPackages
	next
	'report progress
	Repository.WriteOutput outPutName, now() & " Cleanup finished for '"& selectedPackage.Name &"'", 0
end sub

function processDocLibrary(docLibraryElement, messagesRootPackage, messagesPackages)
	'get the docLibrary package
	dim docLibraryPackage
	set docLibraryPackage = Repository.GetPackageByGuid(docLibraryElement.ElementGUID)
	'move the schema profile element(s) into the docLibraryPackage
	moveProfileElements docLibraryPackage
	
	'start from the grand parent package for the messages
	dim parentPackage as EA.Package
	set parentPackage = Repository.GetPackageByID(docLibraryElement.PackageID)
	dim grandParentPackage as EA.Package
	set grandParentPackage = Repository.GetPackageByID(parentPackage.ParentID)
	'create the package structure going up to the root package
	dim messagePackage
	set messagePackage = addPackageToCreatedPackages(grandParentPackage,  messagesPackages)
	'move the docLibrary package to the messagePackage
	docLibraryPackage.ParentID = messagePackage.PackageID
	docLibraryPackage.update
end function

function addPackageToCreatedPackages(package, messagesPackages)
	'check if this package exists
	if not messagesPackages.Exists(package.PackageID) then
		dim newParentPackage as EA.Package
		if not messagesPackages.Exists(package.ParentID) then
			dim parentPackage as EA.Package
			set parentPackage = Repository.GetPackageByID(package.parentID)
			'go up
			set newParentPackage = addPackageToCreatedPackages(parentPackage,  messagesPackages)
		else
			'parent package exists
			set newParentPackage = messagesPackages(package.ParentID)
		end if
		'create new package under new parent package
		dim newPackage as EA.Package
		set newPackage = newParentPackage.Packages.AddNew(package.Name, "")
		newPackage.Update
		'add newPackage to dictionary 
		messagesPackages.add package.PackageID, newPackage
		'return 
		set addPackageToCreatedPackages = newPackage
	else
		'already exists, return it.
		set addPackageToCreatedPackages = messagesPackages.Item(package.PackageID)
	end if
end function

function moveProfileElements(docLibraryPackage)
	dim sqlGetProfileElements
	sqlGetProfileElements = "select o.Object_ID from t_object o                     " & _
							" inner join t_package p on p.Package_ID = o.Package_ID " & _
							" where p.Package_ID = " & docLibraryPackage.ParentID & " " & _
							" and o.Object_Type = 'Artifact'                        "
	dim profileElements
	set profileElements = getElementsFromQuery(sqlGetProfileElements)
	dim profileElement as EA.Element
	for each profileElement in profileElements
		'move profile element into docLibraryPackage
		profileElement.PackageID = docLibraryPackage.PackageID
		profileElement.Update
	next
end function

main