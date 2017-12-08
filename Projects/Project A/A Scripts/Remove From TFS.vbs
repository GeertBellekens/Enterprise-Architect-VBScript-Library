'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Remove from TFS
' Author: Geert Bellekens
' Purpose: Removes the whole package tree from TFS
' Date: 2017-10-19
'

const outPutName = "Remove from TFS"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get the selected package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	'let the user know we started
	Repository.WriteOutput outPutName, now() & " Starting Remove from TFS for package '"& package.Name &"'", 0
	'ask the user if he is sure
	dim userIsSure
	userIsSure = Msgbox("Do you really want to remove package '" &package.Name & "' from TFS?", vbYesNo+vbQuestion, "Remove package from TFS?")
	if userIsSure = vbYes then
		'Actually remove the packages from version control
		removeFromVersionControl package
	end if
	'let the user know it is finished
	Repository.WriteOutput outPutName, now() & " Finished Remove from TFS for package '"& package.Name &"'", 0
end sub

function removeFromVersionControl(package)
	'first process this package
	if package.IsVersionControlled then
		Repository.WriteOutput outPutName, now() & " Removing package '"& package.Name &"' from TFS", 0
		dim vcStatus
		'check version control status for this package
		vcStatus = package.VersionControlGetStatus
		if vcStatus = csCheckedIn then
			'then remove this package from version control
			package.VersionControlRemove
		else
			'tell the user  we couln't remove this package
			Repository.WriteOutput outPutName, now() & " Error removing package '"& package.Name &"' from TFS", 0
		end if
	end if
	'then process subPackages
	dim subPackage
	for each subPackage in package.Packages
		removeFromVersionControl subPackage
	next
end function

main