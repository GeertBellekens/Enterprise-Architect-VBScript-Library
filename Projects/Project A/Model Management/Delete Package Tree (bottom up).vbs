'[path=\Projects\Project A\Model Management]
'[group=Model Management]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Delete Package Tree (bottom up)
' Author: Geert Bellekens
' Purpose: Delete a whole package tree bottom up
' Date: 2019-04-03
'
const outPutName = "Delete package tree"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get selected package
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage
	if not selectedPackage is nothing then
		'ask for confirmation
		dim userIsSure
		userIsSure = Msgbox("Are you sure you want to delete the package '" &selectedPackage.Name & "' and all its subPackages? '", vbYesNo+vbExclamation, "Delete Package " &selectedPackage.Name & "?")
		if userIsSure = vbYes then
			Repository.WriteOutput outPutName, now() & " Starting delete package tree for package '"& selectedPackage.Name &"'", 0
			'delete the package using it's parent
			deletePackageTree selectedPackage, nothing
			'refresh
			Repository.RefreshModelView 0
			'let user know
			Repository.WriteOutput outPutName, now() & " Finished delete package tree for package '"& selectedPackage.Name &"'", 0
		end if
	end if
end sub

function deletePackageTree(package, parentPackage)
	'first delete subPackages
	dim subPackage as EA.Package
	for each subPackage in package.Packages
		deletePackageTree subPackage, package
	next
	'then delete owned elements
	dim j
	for j = package.Elements.Count -1 to 0 step -1
		package.Elements.DeleteAt j , false
	next
	'then get parent package if needed
	if parentPackage is nothing _
	  and package.ParentID > 0 then
		set parentPackage = Repository.GetPackageByID(package.ParentID)
	end if
	'stop if parentPackage still not found
	if parentPackage is nothing then
		exit function
	end if
	'then delete this package using it's parent package
	dim i
	i = 0
	dim tempPackage as EA.Package
	for each tempPackage in parentPackage.Packages
		if tempPackage.PackageID = package.PackageID then
			Repository.WriteOutput outPutName, now() & " Deleting package '"& package.Name &"'", 0
			parentPackage.Packages.DeleteAt i, false
			parentPackage.Packages.Refresh
			exit for
		end if
		'up counter
		i = i + 1
	next
end function

main