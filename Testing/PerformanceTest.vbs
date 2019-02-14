'[group=Testing]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
'
' Script Name: PerformanceTest
' Author: Geert Bellekens
' Purpose: Test the performance difference between different options
' Date: 
'
const outPutName = "PerformanceTest"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get selected package
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage
	'start measuring
	dim startTimeStamp
	dim endTimeStamp
	'now using a query
	startTimeStamp = Timer()
	printAttributesQuery selectedPackage
	endTimeStamp = Timer()
	Repository.WriteOutput outPutName, "Print attributes with query finished in " &  endTimeStamp - startTimeStamp & " s", 0
	'print attributes using EA.Collections
	startTimeStamp = Timer()
	printAttributesCollection selectedPackage
	endTimeStamp = Timer()
	Repository.WriteOutput outPutName, "Print attributes with collection finished in " &  endTimeStamp - startTimeStamp & " s", 0

end sub

function printAttributesCollection(package)
	dim element as EA.Element
	for each element in package.Elements
		dim attribute as EA.Attribute
		for each attribute in element.Attributes
			Session.Output "Attribute: " & element.Name & "." & attribute.Name & " : " & attribute.Type
		next
	next
	'loop subPackages
	dim subPackage
	for each subPackage in package.Packages
		printAttributesCollection subPackage
	next
end function

function printAttributesQuery(package)
	dim element as EA.Element
	for each element in package.Elements
		dim getAttributeQuery
		getAttributesQuery = "select a.ID from t_attribute a where a.Object_ID = " & element.ElementID
		dim attributes
		dim attribute as EA.Attribute
		set attributes = getattributesFromQuery(getAttributesQuery)
		for each attribute in attributes
			Session.Output "Attribute: " & element.Name & "." & attribute.Name & " : " & attribute.Type
		next
	next
	'loop subPackages
	dim subPackage
	for each subPackage in package.Packages
		printAttributesQuery subPackage
	next
end function

main