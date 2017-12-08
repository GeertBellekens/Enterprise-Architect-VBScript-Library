'[path=\Framework\ModelValidation]
'[group=Project Browser Group]

!INC Local Scripts.EAConstants-VBScript
!INC ModelValidation.Include
!INC Wrappers.Include
!INC Atrias Scripts.Util

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
dim outputTabName
outputTabName = "ModelValidation"

sub main
	Repository.CreateOutputTab outputTabName
	Repository.ClearOutput outputTabName
	Repository.EnsureOutputVisible outputTabName
	'tell the user we are starting
	Repository.WriteOutput outputTabName, now() & ": Starting Model Validation",0
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
	'tell the user what we are doing
	Repository.WriteOutput outputTabName, now() & ": Getting list of elements to validate",0
	treeSelectedType = Repository.GetTreeSelectedItemType()
	select case treeSelectedType
		case otElement
			set selectedElements = Repository.GetTreeSelectedElements
			dim element as EA.Element
			for each element in selectedElements
				itemsToValidate.Add element
				Repository.WriteOutput outputTabName, "Adding " & getItemName(element) & " to list to validate",0
				itemsToValidate.AddRange getElementsFromElement(element)
			next
		case otPackage
			set itemsToValidate = getElementsFrompackage(Repository.GetTreeSelectedPackage)
		case else
			msgbox "This scripts works only if you select either a package or one or more elements int the project browser"
	end select 
	'starting validation
	Repository.WriteOutput outputTabName, "Start validating " & itemsToValidate.Count & " items",0
	set ValidationResults = modelValidator.Validate(itemsToValidate, false, false, options, outputTabName)
	Repository.WriteOutput outputTabName, "Processing Validation Results",0
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
		Repository.WriteOutput outputTabName, "Results processed, starting showing results",0
		'show the results
		searchResults.Show
	end if
	Repository.WriteOutput outputTabName, now() & ": Finished Model Validation",0

end sub

function getElementsFromPackage(package)
'	'loop elements and add to a collection
'	dim elements
'	set elements  = CreateObject("System.Collections.ArrayList")
'	dim element as EA.Element
'	dim subPackage as EA.Package
'	for each element in package.Elements
'		elements.Add element
'		Repository.WriteOutput outputTabName, "Adding " & getItemName(element) & " to list to validate",0
'		elements.AddRange getElementsFromElement(element)
'	next
'	'loop subpackages
'	for each subPackage in package.Packages
'		elements.AddRange getElementsFromPackage(subPackage)
'	next
'	'return
'	set getElementsFromPackage = elements
	dim packageList
	set packageList = getPackageTree(package)
	
	dim sqlGetElements
	sqlGetElements = "select o.Object_ID from t_object o where o.Package_ID in (" & makePackageIDString(packageList) & ")"
	set getElementsFromPackage = getElementsFromQuery(sqlGetElements)
end function

function getElementsFromElement(element)
	dim subElement as EA.Element
	dim elements
	set elements  = CreateObject("System.Collections.ArrayList")
	for each subElement in element.Elements
		elements.Add subElement
		Repository.WriteOutput outputTabName, "Adding " & getItemName(subElement) & " to list to validate",0
		elements.AddRange getElementsFromElement(subElement)
	next
	'return
	set getElementsFromElement = elements
end function

main