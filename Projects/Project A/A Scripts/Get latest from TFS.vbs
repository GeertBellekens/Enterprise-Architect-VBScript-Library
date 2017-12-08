'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Get latest from TFS
' Author: Geert Bellekens
' Purpose: Does GetLatest on all version controlled packages in the tree.
' Date: 2017-10-19
'

const outPutName = "Get latest from TFS"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get the selected package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	'let the user know we started
	Repository.WriteOutput outPutName, now() & " Starting get latest from TFS for package '"& package.Name &"'", 0
	'ask the user if he is sure
	dim userIsSure
	userIsSure = Msgbox("Do you really want to get latest from TFS for package '" &package.Name & "' ?", vbYesNo+vbQuestion, "Get latest from TFS?")
	if userIsSure = vbYes then
		'actually do the getlatest
		getLatestFromTFS package
		'do a reconcile to make sure all relations are there
		Repository.WriteOutput outPutName, now() & " Reconciling...", 0
		Repository.ScanXMIAndReconcile
		'make sure the output is visible again
		Repository.EnsureOutputVisible outPutName
	end if
	'let the user know it is finished
	Repository.WriteOutput outPutName, now() & " Finished get latest from TFS for package '"& package.Name &"'", 0
end sub

function getLatestFromTFS(package)
	'first process this package
	if package.IsVersionControlled then
		Repository.WriteOutput outPutName, now() & " Getting latest version for package '"& package.Name &"'", 0
		dim isReadWrite
		if isRequireUserLockEnabled() then
			isReadWrite = package.ApplyUserLockRecursive(true, true, true)
		else
			isReadWrite = true
		end if
		'then add this package to version control
		if not isReadWrite then
			Repository.WriteOutput outPutName, now() & " ERROR: Cannot proceed as package '"& package.Name &"' cannot be locked", package.Element.ElementID
			exit function
		end if
		package.VersionControlGetLatest(true) 'true for force import
	end if
	'make sure we have an up-to set of subPackages
	dim newPackage as EA.Package
	set newPackage = Repository.GetPackageByGuid(package.PackageGUID)
	'then process subPackages
	dim subPackage
	for each subPackage in newPackage.Packages
		getLatestFromTFS subPackage
	next
end function

main