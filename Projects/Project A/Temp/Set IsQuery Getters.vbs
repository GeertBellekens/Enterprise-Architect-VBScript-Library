'[path=\Projects\Project A\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Set IsQuery Getter
' Author: Geert Bellekens
' Purpose: sets the IsQuery property to true for all operations in the selected package with stereotype "get"
' Date: 2016-01-22
'
sub main
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage()
	dim element as EA.Element
	dim operation as EA.Method
	for each element in selectedPackage.Elements
		for each operation in element.Methods
			if operation.Stereotype = "get" then
				operation.IsQuery = true
				operation.Update
			end if
		next
	next
end sub

main