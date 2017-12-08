'[path=\Projects\Project A\Project Browser Package Group]
'[group=Project Browser Package Group]

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Add Attribute Dependencies
' Author: Geert Bellekens
' Purpose: This script will remove all attribute type dependencies created by the EA Message Composer
'    script can be executed on a package an will remove the dependencies from the selected package and all sub packages.
		
' Date: '2017-04-20
'
'name of the output tab
const outPutName = "Add Attribute Type Dependencies"

sub main

	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage()
	
	if not messagePackage is nothing then
		'tell the user we are starting
		Repository.WriteOutput outPutName, now() & " Starting adding Attribute Type Dependencies for '" & selectedPackage.Name & "'", selectedPackage.Element.ElementID
		'do the actual work
		addAttributeDependencies(package)
		'tell the user we have finished
		Repository.WriteOutput outPutName, now() & " Finished adding Attribute Type Dependencies for '" & selectedPackage.Name & "'", selectedPackage.Element.ElementID
	endif 
	
end sub

function addAttributeDependencies(package)
	dim element as EA.Element
	'loop elements
	for each element in package.Elements
		dim attribute as EA.Attribute
		'loop attributes
		for each attribute in element.Attributes
			if attribute.ClassifierID > 0 then
				'add dependency
				dim typeDependency as EA.Connector
				set typeDependency = element.Connectors.AddNew("","Dependency")
				typeDependency.SupplierID = attribute.ClassifierID
				typeDependency.Update
				if len(attribute.LowerBound) > 0 AND len (attribute.UpperBound) > 0 then
					typeDependency.SupplierEnd.Cardinality = attribute.LowerBound & .. & attribute.UpperBound
					typeDependency.SupplierEnd.Update;
				end if
			end if
		next
	next
	'go one level deeper
	dim subPackage as EA.Package
	for each subPackage in package.Packages
		
	next
	
end function



main