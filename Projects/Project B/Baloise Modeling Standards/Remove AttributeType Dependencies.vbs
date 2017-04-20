'[path=\Projects\Project B\Baloise Modeling Standards]
'[group=Baloise Modeling Standards]

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Remove Attribute Type Dependencies
' Author: Geert Bellekens
' Purpose: This script will remove all attribute type dependencies created by the EA Message Composer
'    script can be executed as custom script after generating a message, or as a diagram group script.
		
' Date: '2017-04-20
'
sub RemoveAttributeDependenciesFromCurrentDiagram()
	' get the message package
	dim currentDiagram as EA.Diagram
	set currentDiagram = Repository.GetCurrentDiagram
	if not currentDiagram is nothing then
		dim messagePackage as EA.Package
		set messagePackage = Repository.GetPackageByID(currentDiagram.PackageID)
		' ask the user for confirmation before updating the tags just in case something went wrong with the generation of the subset and we have the wrong diagram
		dim response
		response = msgbox("Remove all attribute type dependencies in '" & messagePackage.Name & "'?", vbYesNo+vbQuestion, "Remove attribute type dependencies?")
		if response = vbYes then
			dim packageIDString
			packageIDString = getPackageTreeIDString(messagePackage)
			dim sqlDeleteDependencies
			sqlDeleteDependencies = "delete c                                                   " & _
									" from ((t_connector c                                       " & _
									" inner join t_object so on c.Start_Object_ID = so.Object_ID)" & _
									" inner join t_attribute a on (a.Object_ID = so.Object_ID    " & _
									" 						and a.Name = c.Name                  " & _
									" 						and a.Classifier = c.End_Object_ID)) " & _
									" where c.Connector_Type = 'dependency'                      " & _
									" and so.Package_ID in (" & packageIDString & ")             "
			Repository.Execute sqlDeleteDependencies
			'reload the diagram
			dim saveDiagram
			saveDiagram = msgbox("Save the current diagram '" & currentDiagram.Name & "' before reloading?", vbYesNo+vbQuestion, "Save Diagram?")
			if saveDiagram = vbYes then
				Repository.SaveDiagram currentDiagram.DiagramID
			end if
			Repository.ReloadDiagram currentDiagram.DiagramID
		end if
	end if
	
end sub

'get the package id string of the given package tree
function getPackageTreeIDString(package)
	'initialize at "0"
	getPackageTreeIDString = "0"
	dim packageTree
	dim currentPackage as EA.Package
	if not package is nothing then
		'get the whole tree of the selected package
		set packageTree = getPackageTree(package)
		' get the id string of the tree
		getPackageTreeIDString = makePackageIDString(packageTree)
	end if 
end function

'make an id string out of the package ID of the given packages
function makePackageIDString(packages)
	dim package as EA.Package
	dim idString
	idString = ""
	dim addComma 
	addComma = false
	for each package in packages
		if addComma then
			idString = idString & ","
		else
			addComma = true
		end if
		idString = idString & package.PackageID
	next 
	'if there are no packages then we return "0"
	if packages.Count = 0 then
		idString = "0"
	end if
	'return idString
	makePackageIDString = idString
end function

'returns an ArrayList of the given package and all its subpackages recursively
function getPackageTree(package)
	dim packageList
	set packageList = CreateObject("System.Collections.ArrayList")
	addPackagesToList package, packageList
	set getPackageTree = packageList
end function

'add the given package and all subPackges to the list (recursively
function addPackagesToList(package, packageList)
	dim subPackage as EA.Package
	'add the package itself
	packageList.Add package
	'add subpackages
	for each subPackage in package.Packages
		addPackagesToList subPackage, packageList
	next
end function

RemoveAttributeDependenciesFromCurrentDiagram