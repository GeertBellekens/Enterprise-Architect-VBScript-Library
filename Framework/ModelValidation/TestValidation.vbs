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
	set selectedElements = Repository.GetTreeSelectedElements
	dim modelValidator
	set modelValidator = new ModelValidator
	dim ValidationResults, validationResult
	dim options
	set options = nothing
	set ValidationResults = modelValidator.Validate(selectedElements, false, false,options)
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
	
'	searchResults = "<ReportViewData UID="""">" & _
'					"	<Fields>" & _
'					"		<Field name=""CLASSGUID""/>" & _
'					"		<Field name=""CLASSTYPE""/>" & _
'					"		<Field name=""Name""/>" & _
'					"	</Fields>" & _
'					"	<Rows>" & _
'					"		<Row>" & _
'					"			<Field name=""CLASSGUID"" value=""{68FFEBD2-B90B-4b55-A697-962FA0768AA5}""/>" & _
'					"			<Field name=""CLASSTYPE"" value=""UseCase""/>" & _
'					"			<Field name=""Name"" value=""UC - ME - 007 - Process iExV ""/>" & _
'					"		</Row>" & _
'					"	</Rows>" & _
'					"</ReportViewData>"
'	Repository.RunModelSearch "searchName","searchTerm","searchOptions",searchResults
end sub

main