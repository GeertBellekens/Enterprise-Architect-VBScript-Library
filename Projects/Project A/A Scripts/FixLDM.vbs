'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
dim response
		response = Msgbox("This script will switch alias and name for all enumeration values" & vbnewLine & " Are you sure?", vbYesNo+vbExclamation, "Fix LDM")
		if response = vbYes then
			'create output tab
			Repository.CreateOutputTab "FixLDM"
			Repository.ClearOutput "FixLDM"
			Repository.EnsureOutputVisible "FixLDM"
	
			'get the selected package
			dim package as EA.Package 
			set package = Repository.GetTreeSelectedPackage()
			'switch alias and name for all attributes
			switchAttrAliasAndName package
			'tell the user we are done
			Repository.WriteOutput "FixLDM","Finished'",0
		end if
end sub

main

function switchAttrAliasAndName(package)
	dim element as EA.Element
	dim subPackage as EA.Package
	'first loop owned elements
	for each element in package.Elements
		switchAttrAliasAndNameOnElement element
	next
	'then do the owned packages
	for each subPackage in package.Packages
		switchAttrAliasAndName subPackage
	next
end function

function switchAttrAliasAndNameOnElement(element)
	dim attribute as EA.Attribute
	dim tempAlias
	dim subElement as EA.Element
	'log progress
	Repository.WriteOutput "FixLDM","Processing: " & element.Name,0
	'first do all owned attributes
	for each attribute in element.Attributes
		tempAlias = attribute.Alias
		attribute.Alias = attribute.Name
		attribute.Name = tempAlias
		attribute.Update
	next
	'then do the owned element
	for each subElement in element.Elements
		switchAttrAliasAndNameOnElement subElement
	next
end function