'[path=\Projects\Project AP\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Delete Backup Packages
' Author: Geert Bellekens
' Purpose: Remove the backup packages from the selected package
' Date: 2022-05-27
'
const outPutName = "Delete Backup Packages"

function Main ()
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage()
	if selectedPackage is nothing then
		exit function
	end if
	'inform user
	Repository.WriteOutput outPutName, now() & " Starting Delete Backup Packages for package '" & selectedPackage.name & "'", 0
	'do the actual work
	deleteBackupPackages selectedPackage
	'inform user
	Repository.WriteOutput outPutName, now() & " Finished Delete Backup Packages for package '" & selectedPackage.name & "'", 0
		
end function

function deleteBackupPackages(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	dim sqlGetData
	sqlGetData = "select p.Package_ID from t_package p                                  " & vbNewLine & _
				" where p.Name like '%_backup'                                          " & vbNewLine & _
				" and not exists (select o.Object_ID from t_object o                    " & vbNewLine & _
				" 				where o.Package_ID = p.Package_ID                       " & vbNewLine & _
				" 				and o.Object_Type not in ('Text', 'Boundary','Note'))   " & vbNewLine & _
				" and p.Package_ID in (" & packageTreeIDString & ")                     " & vbNewLine & _
				" order by p.Parent_ID                                                   "
	dim results
	set results = getPackagesFromQuery(sqlGetData)
	Repository.WriteOutput outPutName, now() & " Found " & results.Count & " packages to delete", 0
	if results.Count = 0 then
		'no packages found
		exit function
	end if
	'make sure the user is sure:
	dim userIsSure
	userIsSure = Msgbox("Delete " & results.Count & " backup packages in package '" & package.Name & "'?", vbYesNo+vbQuestion, "Delete Backup Packages?")
	if userIsSure = vbYes then
		dim backupPackage as EA.Package
		dim i
		i = 0
		dim parentPackage as EA.Package
		set parentPackage = package
		for each backupPackage in results
			i = i + 1
			'inform user
			Repository.WriteOutput outPutName, now() & " Deleting package " & i & " of " & results.Count & ": '" & backupPackage.Name & "'", 0
			if parentPackage.PackageID <> backupPackage.ParentID then
				set parentPackage = Repository.GetPackageByID(backupPackage.ParentID)
			end if
			deletePackageFromParent parentPackage, backupPackage
		next
	end if
	'reload the package
	Repository.ReloadPackage package.PackageID
end function

function deletePackageFromParent(parentPackage, package)
	dim i
	dim currentPackage as EA.Package
	for i = parentPackage.Packages.Count -1 to 0 step -1
		set currentPackage = parentPackage.Packages.GetAt(i)
		if currentPackage.PackageID = package.PackageID then
			parentPackage.Packages.DeleteAt i, true 
			exit function
		end if 
	next
end function

main