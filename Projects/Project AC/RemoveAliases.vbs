'[path=\Projects\Project AC]
'[group=Acerta Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
	dim selectedPackage as EA.Package
	 set selectedPackage  = Repository.GetTreeSelectedPackage
	 if not selectedPackage is nothing then
		removeAliases selectedPackage
	 end if
end sub

function removeAliases(package)
	dim element as EA.Element
	for each element in package.Elements
		element.Alias = ""
		element.Update
		'attributes
		dim attribute as EA.Attribute
		for each attribute in element.Attributes
			attribute.Alias = ""
			attribute.Update
		next
	next
end function

main