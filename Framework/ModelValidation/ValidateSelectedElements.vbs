'[path=\Framework\ModelValidation]
'[group=Search Scripts]

!INC Local Scripts.EAConstants-VBScript
!INC ModelValidation.Include
!INC Wrappers.Include

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
	dim selectedElement as EA.Element
	dim selectedElements as EA.Collection
	dim treeSelectedType
	dim selectedPackage
	dim modelValidator
	set modelValidator = new ModelValidator
	dim ValidationResults, validationResult
	dim options
	set options = nothing
	'get the elements to validate
	dim itemsToValidate 
	set itemsToValidate = CreateObject("System.Collections.ArrayList")
	
	treeSelectedType = Repository.GetTreeSelectedItemType()
	select case treeSelectedType
		case otElement
			set selectedElements = Repository.GetTreeSelectedElements
			dim element as EA.Element
			for each element in selectedElements
				itemsToValidate.Add element
				itemsToValidate.AddRange getElementsFromElement(element)
			next
		case otPackage
			set itemsToValidate = getElementsFrompackage(Repository.GetTreeSelectedPackage)
		case else
			msgbox "This scripts works only if you select either a package or one or more elements int the project browser"
	end select 

	set ValidationResults = modelValidator.Validate(itemsToValidate, false, false, options)
	if ValidationResults.Count > 0 then
		dim searchresults 
		set searchresults = new SearchResults
		searchResults.Name = "Validation Result"
		'add headers
		searchResults.Fields = ValidationResults(0).Headers
		
		'add results
		for each validationResult in ValidationResults
			searchResults.Results.Add validationResult.getResultFields()
		next
		'show the results
		searchResults.Show
	end if

end sub

function getElementsFromPackage(package)
	'loop elements and add to a collection
	dim elements
	set elements  = CreateObject("System.Collections.ArrayList")
	dim element as EA.Element
	dim subPackage as EA.Package
	for each element in package.Elements
		elements.Add element
		elements.AddRange getElementsFromElement(element)
	next
	'loop subpackages
	for each subPackage in package.Packages
		elements.AddRange getElementsFromPackage(subPackage)
	next
	'return
	set getElementsFromPackage = elements
end function

function getElementsFromElement(element)
	dim subElement as EA.Element
	dim elements
	set elements  = CreateObject("System.Collections.ArrayList")
	for each subElement in element.Elements
		elements.Add subElement
		elements.AddRange getElementsFromElement(subElement)
	next
	'return
	set getElementsFromElement = elements
end function

main