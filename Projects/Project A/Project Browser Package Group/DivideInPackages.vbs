option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: DivideInPackages
' Author: Geert Bellekens
' Purpose: Puts each BusinessProcess or SubProcess/Activity in each own package. To be run from a package in the project browser
' Date: 08/09/2015
'
sub main
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	dim process as EA.Element
	for each process in package.Elements
		if process.Stereotype = "BusinessProcess" _
		or process.Stereotype = "Activity" then
			dim subPackage as EA.Package
			set subPackage = package.Packages.AddNew(process.Name,"Package")
			'msgbox subPackage.Name
			subPackage.Update
			process.PackageID = subPackage.PackageID
			process.Update
		end if
	next
	Repository.RefreshModelView package.PackageID
	msgbox "finished!"
end sub

main