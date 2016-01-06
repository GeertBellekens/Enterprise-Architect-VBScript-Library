'[path=\Projects\Project A\Project Browser Package Group]
'[group=Project Browser Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: ExportToXPDL
' Author: Geert Bellekens
' Purpose: Exports each package to xpdl
' Date: 08/09/2015
'
sub main
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	
	'get the folder from the user
	dim folder
    set folder = new FileSystemFolder
	set folder = folder.getUserSelectedFolder("")
	'export
	if not folder is nothing then
		exportToXPDL package, folder
	end if
end sub

function exportToXPDL(package, folder)

	'only process if the package has nu subpackages
	if package.Packages.Count = 0 then
		dim projectInterface as EA.Project
		set projectInterface = Repository.GetProjectInterface()
		projectInterface.ExportPackageXMIEx package.PackageGUID, xmiXPDL22,1,0,0,0,folder.FullPath & "\" & package.Name & ".xpdl",epExcludeEAExtensions
	else
		dim subPackage as EA.Package
		for each subPackage in package.Packages
			exportToXPDL subPackage, folder
		next
	end  if
end function


main