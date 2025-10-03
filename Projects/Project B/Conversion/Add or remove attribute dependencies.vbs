'[path=\Projects\Project B\Conversion]
'[group=Conversion]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Conversion.Transformation utils

'
' Script Name: Transform to JSON
' Author: Geert Bellekens
' Purpose: Transforms the current package into a new package, transforming the stereotypes
' Date: 2025-01-16
'

const outPutName = "Add or remove attribute dependencies"


sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get the selected package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	dim addDependencies
	addDependencies = Msgbox("Do you want to:" & vbNewLine & "-Add (YES) " & vbNewLine & "-Remove (NO) " & vbNewLine & "Attribute dependencies in package '" &package.Name & "'?", vbYesNoCancel+vbQuestion, "Add/Remove attribute dependencies")
	if addDependencies = vbYes then
		'let the user know we started
		Repository.WriteOutput outPutName, now() & " Starting adding attribute dependencies for package '" & package.Name &"'", 0
		'do the actual work
		addAttributeDependencies packageTreeIDString
		'let the user know it is finished
		Repository.WriteOutput outPutName, now() & " Finished adding attribute dependencies for package '"& package.Name &"'", 0
	elseif addDependencies = vbNo then
		'let the user know we started
		Repository.WriteOutput outPutName, now() & " Starting removing attribute dependencies for package '" & package.Name &"'", 0
		'do the actual work
		removeAttributeDependencies packageTreeIDString
		'let the user know it is finished
		Repository.WriteOutput outPutName, now() & " Finished removing attribute dependencies for package '"& package.Name &"'", 0
	end if
end sub

main