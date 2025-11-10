'[path=\Projects\Project IFB\Template fragments]
'[group=Template fragments]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: UsedGlossaryTermsTranslated
' Author: Geert Bellekens
' Purpose: Returns all GlossaryTerms used in the translation of the process and its nested elements
' Date: 2025-11-05
'

function MyRtfData (ObjectID, language)
	dim xmlQueryResult
	xmlQueryResult = Repository.SQLQuery(sqlGetData)
	dim xmlDOM 
	set xmlDOM = CreateObject("MSXML2.DOMDocument")
	dim businessProcess as EA.Element
	set businessProcess = Repository.GetElementByID(objectID)
	dim usedGlossaryterms
	set usedGlossaryterms = getUsedGlossaryTerms(businessProcess, language)
	addGlossaryTermsToEAData xmlDom, usedGlossaryterms
	'debug
	set fileSystemObject = CreateObject( "Scripting.FileSystemObject" )
	set outputFile = fileSystemObject.CreateTextFile( "c:\temp\glossaryterms.xml", true )
	outputFile.Write xmlDOM.xml
	outputFile.Close
	MyRtfData = xmlDOM.xml
end function

function addGlossaryTermsToEAData(xmlDom, usedGlossaryterms)
	xmlDOM.validateOnParse = false
	xmlDOM.async = false
	dim node 
	set node = xmlDOM.createProcessingInstruction( "xml", "version='1.0'")
	xmlDOM.appendChild node
	dim xmlRoot 
	set xmlRoot = xmlDOM.createElement( "EADATA" )
	xmlDOM.appendChild xmlRoot
	dim xmlDataSet
	set xmlDataSet = xmlDOM.createElement( "Dataset_0" )
	xmlRoot.appendChild xmlDataSet
	dim xmlData 
	set xmlData = xmlDOM.createElement( "Data" )
	xmlDataSet.appendChild xmlData
	'add rows
	dim termList
	for each termList in usedGlossaryterms.Items
		dim termRow
		for each termRow in termList
			dim xmlRow
			set xmlRow = xmlDOM.createElement( "Row" )
			xmlData.appendChild xmlRow
			dim xmlTerm
			set xmlTerm = xmlDOM.createElement("Term")
			xmlTerm.Text = termRow(0)
			xmlRow.appendChild xmlTerm
			dim xmlCategory
			set xmlCategory = xmlDOM.createElement("Category")
			xmlCategory.Text = termRow(1)
			xmlRow.appendChild xmlCategory
			dim xmlNotes
			set xmlNotes = xmlDOM.createElement("Notes")
			dim formattedAttr 
			set formattedAttr = xmlDOM.createAttribute("formatted")
			formattedAttr.nodeValue="1"
			xmlNotes.setAttributeNode(formattedAttr)
			xmlNotes.Text = termRow(2)
			xmlRow.appendChild xmlNotes
		next
	next
end function

function getUsedGlossaryTerms(element, language)
	dim usedGlossaryTerms
	set usedGlossaryTerms = CreateObject("Scripting.Dictionary")
	'get dictionary of all glossaryterms
	dim glossaryDictionary
	set glossaryDictionary = getGlossaryDictionary()
	'get full translated string
	dim fullTranslatedString
	fullTranslatedString = getFullTranslatedString(element, language)
	'build regular expression
	dim regexp
	set regexp = CreateObject("VBScript.RegExp")
	regexp.Global = True
	regexp.IgnoreCase = True
	regexp.Pattern = "\b(" & Join(glossaryDictionary.Keys, "|") & ")\b"
	dim matches
	' Execute once on the entire text
	set matches = regexp.Execute(fullTranslatedString)
	dim match
	for each match in matches
		dim term
		term = lcase(match)
		if not usedGlossaryTerms.Exists(term) and glossaryDictionary.Exists(term) then
			usedGlossaryTerms.Add term, glossaryDictionary(term)
		end if
	next
	'return used glossary
	set getUsedGlossaryTerms = usedGlossaryTerms
end function



function getGlossaryDictionary()
	dim glossaryDictionary
	set glossaryDictionary = CreateObject("Scripting.Dictionary")
	dim sqlGetData
	sqlGetData = "select o.Name as Term,  d.Name as Type, o.Note as [Meaning], 'True' as [ModelItem]    " & vbNewLine & _
				" from t_object o                                                                      " & vbNewLine & _
				" inner join t_diagramobjects do on do.Object_ID = o.Object_ID                         " & vbNewLine & _
				" inner join t_diagram d on d.Diagram_ID = do.Diagram_ID                               " & vbNewLine & _
				"       and d.StyleEx like '%MDGDgm=Glossary Item Lists::GlossaryItemList;%'           " & vbNewLine & _
				" where len(o.Name) > 0                                                                " & vbNewLine & _
				" union                                                                                " & vbNewLine & _
				" select o.Name as Term,  p.Name as Type, o.Note as [Meaning], 'True' as [ModelItem]   " & vbNewLine & _
				" from t_object o                                                                      " & vbNewLine & _
				" inner join t_package p on p.Package_ID = o.Package_ID                                " & vbNewLine & _
				" where o.Stereotype = 'GlossaryEntry'                                                 " & vbNewLine & _
				" and len(o.Name) > 0                                                                  " & vbNewLine & _
				" union                                                                                " & vbNewLine & _
				" select g.Term, g.Type, g.Meaning, 'False' as [ModelItem]                             " & vbNewLine & _
				" from t_glossary g                                                                    " & vbNewLine & _
				" where len(g.Term) > 0                                                                " & vbNewLine & _
				" order by 1, 2, 3                                                                     "
	Dim results
	Set results = getArrayListFromQuery(sqlGetData)
	'loop over results
	dim row
	for each row in results
		dim term
		term = lcase(row(0))
		dim entry
		if glossaryDictionary.Exists(term) then
			'term already in dictionary, row can be added to the existing list
			set entry = glossaryDictionary(term)
		else
			'create new list 
			set entry = CreateObject("System.Collections.ArrayList")
			'add list to dictionary
			glossaryDictionary.Add term, entry
		end if
		'then add the row
		entry.Add row
	next
	'debug
	Session.Output "glossaryDictionary.Count: " & glossaryDictionary.Count
	'return dictionary
	set getGlossaryDictionary = glossaryDictionary
end function

function getFullTranslatedString(element, language)
	'get all related elements
	dim relatedElements
	set relatedElements = getRelatedElements(element)
	'build string with name and notes in the requested language
	dim fullString
	fullString = ""
	'loop elements
	dim relatedElement as EA.Element
	for each relatedElement in relatedElements
		fullString = fullString & relatedElement.GetTXName(language, 0) & " " & relatedElement.GetTXNote(language, 0) & " "
	next
	'return
	getFullTranslatedString = fullString
end function

function getRelatedElements(element)
	dim sqlGetData
	sqlGetData = "select  o.Object_ID from t_object bp                                                                                                 " & vbNewLine & _
				" left join t_object o1 on o1.ParentID = bp.Object_ID                                                                                 " & vbNewLine & _
				" left join t_object o2 on o2.ParentID = o1.Object_ID                                                                                 " & vbNewLine & _
				" left join t_object o3 on o3.ParentID = o2.Object_ID                                                                                 " & vbNewLine & _
				" inner join t_object o on o.Object_ID in (bp.Object_ID, o1.Object_ID, o2.Object_ID, o3.Object_ID)                                    " & vbNewLine & _
				" where bp.Object_ID = " & element.ElementID & "                                                                                      " & vbNewLine & _
				" and o.Object_Type not in ('Text', 'Note', 'Artifact')                                                                               " & vbNewLine & _
				" union                                                                                                                               " & vbNewLine & _
				" select oo.Object_ID from t_object bp                                                                                                " & vbNewLine & _
				" left join t_object o1 on o1.ParentID = bp.Object_ID                                                                                 " & vbNewLine & _
				" left join t_object o2 on o2.ParentID = o1.Object_ID                                                                                 " & vbNewLine & _
				" left join t_object o3 on o3.ParentID = o2.Object_ID                                                                                 " & vbNewLine & _
				" inner join t_object o on o.Object_ID in (bp.Object_ID, o1.Object_ID, o2.Object_ID, o3.Object_ID)                                    " & vbNewLine & _
				" inner join t_objectproperties tv on tv.Object_ID = o.Object_ID                                                                      " & vbNewLine & _
				" inner join t_object oo on oo.ea_guid = tv.Value                                                                                     " & vbNewLine & _
				" where bp.Object_ID = " & element.ElementID & "                                                                                      " & vbNewLine & _
				" and o.Object_Type not in ('Text', 'Note', 'Artifact')                                                                               " & vbNewLine & _
				" union                                                                                                                               " & vbNewLine & _
				" select oo.Object_ID from t_object bp                                                                                                " & vbNewLine & _
				" left join t_object o1 on o1.ParentID = bp.Object_ID                                                                                 " & vbNewLine & _
				" left join t_object o2 on o2.ParentID = o1.Object_ID                                                                                 " & vbNewLine & _
				" left join t_object o3 on o3.ParentID = o2.Object_ID                                                                                 " & vbNewLine & _
				" inner join t_object o on o.Object_ID in (bp.Object_ID, o1.Object_ID, o2.Object_ID, o3.Object_ID)                                    " & vbNewLine & _
				" inner join t_connector c on o.Object_ID in (c.End_Object_ID, c.Start_Object_ID)                                                     " & vbNewLine & _
				" 						and c.Stereotype = 'trace'                                                                                    " & vbNewLine & _
				" inner join t_object oo on oo.Object_ID in (c.End_Object_ID, c.Start_Object_ID)                                                      " & vbNewLine & _
				" 						and oo.Object_ID <> o.Object_ID                                                                               " & vbNewLine & _
				" 						and oo.Stereotype in ('Risk', 'ArchiMate_ApplicationComponent','ArchiMate_Location', 'ArchiMate_Equipment')   " & vbNewLine & _
				" where bp.Object_ID = " & element.ElementID & "                                                                                      " & vbNewLine & _
				" and o.Object_Type not in ('Text', 'Note', 'Artifact')                                                                               "
	dim result
	set result = getElementsFromQuery(sqlGetData)
	'return
	set getRelatedElements = result
end function




function addFormattedAttribute(fieldNode, xmlDOM)
	dim formattedAttr
	set formattedAttr = xmlDOM.createAttribute("formatted")
	formattedAttr.nodeValue="1"
	fieldNode.setAttributeNode(formattedAttr)
end function



'msgbox MyPackageRtfData(3357,"")
function test
	MyRtfData 139544, "nl"
end function 

test