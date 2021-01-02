'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
	'get selected package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage
	'create element
	dim element as EA.Element
	set element = package.Elements.AddNew("newClass", "Class")
	element.update
	'create attribute
	dim attr as EA.Attribute
	set attr = element.Attributes.AddNew("newAttribute","string")
	attr.Update
	'create attributeConstraint
	dim attributeConstraint as EA.AttributeConstraint
	set attributeConstraint = attr.Constraints.AddNew("Constraint1", "Restriction")
	attributeConstraint.Update
	'refresh constraints collection
	attr.Constraints.Refresh
	'loop through attribute constraints again
	dim tempConstraint as EA.AttributeConstraint
	for each tempConstraint in attr.Constraints
		Session.Output tempConstraint.Name
	next
	'again make new one
	set attributeConstraint = attr.Constraints.AddNew("Constraint1", "Restriction")
	attributeConstraint.Update
	'refresh constraints collection
	attr.Constraints.Refresh
	'loop through attribute constraints again
	for each tempConstraint in attr.Constraints
		Session.Output tempConstraint.Name
		tempConstraint.Update
	next
	
	'get element again
	set element = Repository.GetElementByGuid(element.ElementGUID)
	'loop through attributes
	for each attr in element.Attributes
		'add a new constraint
		set attributeConstraint = attr.Constraints.AddNew("Second name 'with' single quotes", "Restriction")
		attributeConstraint.Update
	next
	
end sub

main