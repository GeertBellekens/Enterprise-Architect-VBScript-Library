'[path=\Projects\Project A\Search Scripts]
'[group=Search Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.Util

'
' This code has been included from the default Search Script template.
' If you wish to modify this template, it is located in the Config\Script Templates
' directory of your EA install path.   
'
' Script Name:
' Author:
' Purpose:
' Date:
'

' TODO 1: Define your search specification:
' The columns that will apear in the Model Search window
dim SEARCH_SPECIFICATION 
SEARCH_SPECIFICATION = "<ReportViewData>" &_
							"<Fields>" &_
								"<Field name=""CLASSGUID""/>" &_
								"<Field name=""CLASSTYPE"" />" &_
								"<Field name=""Element Name"" />" &_
								"<Field name=""Comments"" />" &_
							"</Fields>" &_
							"<Rows/>" &_
						"</ReportViewData>"

'
' Search Script main function
' 
sub OnSearchScript()
	'get the search term
	dim searchTerm
	searchTerm = InputBox( "Please enter term to search for", "search term" )
	'get the linked documents
	dim allLinkedDocuments
	set allLinkedDocuments = getAllLinkedDocuments()
	
	
	' Create a DOM object to represent the search tree
	dim xmlDOM
	set xmlDOM = CreateObject( "MSXML2.DOMDocument.4.0" )
	xmlDOM.validateOnParse = false
	xmlDOM.async = false
	
	' Load the search template
	if xmlDOM.loadXML( SEARCH_SPECIFICATION ) = true then
	
		dim rowsNode
		set rowsNode = xmlDOM.selectSingleNode( "//ReportViewData//Rows" )
	
		' TODO 2: Gather the required data from the repository
		' This template adds a result row for a bogus class to the search document
		AddRow xmlDOM, rowsNode, "{2917209A-D3E0-4de7-8AED-C7D7F059D96F}", "ResultClass", _
			"Here are some comments about this class!"
		
		' Fill the Model Search window with the results
		Repository.RunModelSearch "", "", "", xmlDOM.xml
		
	else
		Session.Prompt "Failed to load search xml", promptOK
	end if
end sub	

'
' TODO 3: Modify this function signature to include all information required for the search
' results. Entire objects (such as elements, attributes, operations etc) may be passed in.
'
' Adds an entry to the xml row node 'rowsNode'
'
sub AddRow( xmlDOM, rowsNode, elementGUID, elementName, comments )

	' Create a Row node
	dim row
	set row = xmlDOM.createElement( "Row" )
	
	' Add the Model Search row data to the DOM
	AddField xmlDOM, row, "CLASSGUID", elementGUID
	AddField xmlDOM, row, "CLASSTYPE", "Class"
	AddField xmlDOM, row, "Name", elementName
	AddField xmlDOM, row, "Comments", comments
	
	' Append the newly created row node to the rows node
	rowsNode.appendChild( row )

end sub

'
' Adds an Element to the DOM called Field which makes up the Row data for the Model Search window.
' <Field name "" value ""/>
'
sub AddField( xmlDOM, row, name, value )

	dim fieldNode
	set fieldNode = xmlDOM.createElement( "Field" )
	
	' Create first attribute for the name
	dim nameAttribute
	set nameAttribute = xmlDOM.createAttribute( "name" )
	nameAttribute.value = name
	fieldNode.attributes.setNamedItem( nameAttribute )
	
	' Create second attribute for the value
	dim valueAttribute 
	set valueAttribute = xmlDOM.createAttribute( "value" )
	valueAttribute.value = value
	fieldNode.attributes.setNamedItem( valueAttribute )
	
	' Append the fieldNode
	row.appendChild( fieldNode )

end sub

'returns a dictionary of all elements that have a linked document as key and the text of the linked document as value.
function getAllLinkedDocuments
	dim queryString
	queryString =	"select o.object_ID from t_document d " & _
					" inner join t_object o on d.ElementID = o.ea_guid " & _
					" where d.ElementType = 'ModelDocument'"
	dim elementsWithLinkedDocument
	set elementsWithLinkedDocument = getElementsFromQuery(queryString)
	Session.Output "number of elements with a linked document: " & elementsWithLinkedDocument.count
	dim linkedDocumentsDictionary
	set linkedDocumentsDictionary = CreateObject("Scripting.Dictionary")
	dim element as EA.Element
	'loop the elements and add element and its linked document to the dictionary
	for each element in elementsWithLinkedDocument
		dim linkedDocumentText
		linkedDocumentText = getLinkedDocumentContent(element, "TXT")
		linkedDocumentsDictionary.Add element, linkedDocumentText
	next
	set getAllLinkedDocuments = linkedDocumentsDictionary
end function


OnSearchScript()