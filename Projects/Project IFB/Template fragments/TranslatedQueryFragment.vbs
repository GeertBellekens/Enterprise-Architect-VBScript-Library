'[path=\Projects\Project IFB\Template fragments]
'[group=Template fragments]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: TranslatedQueryFragment
' Author: Geert Bellekens
' Purpose: Returns the translated values for data in the query. It looks for columns that end with Name, Alias or Notes-Formatted and checks if there is a corresponding column with extra suffix _guid
' if there is then it will replace the contents of the column with the translated name/alias/notes of the item corresponding to the guid.
' if there are multiple guids, then the each will be replaced with the translation of the name/alias/notes.
' Date: 2025-09-03
'

function MyRtfData (sqlGetData, language)
	dim xmlQueryResult
	xmlQueryResult = Repository.SQLQuery(sqlGetData)
	dim xmlDOM 
	set xmlDOM = CreateObject("MSXML2.DOMDocument")
	if not xmlDOM.LoadXML(xmlQueryResult) Then  
		MyRtfData = ""
		exit function
	end if
	dim itemsCache
	set itemsCache = CreateObject("Scripting.Dictionary")
	'select the rows
	dim rowList
	Set rowList = xmlDOM.SelectNodes("//Row")
	dim rowNode 
	for each rowNode in rowList
		processRow rowNode, language, xmlDOM, itemsCache
	next
	'debug
'	set fileSystemObject = CreateObject( "Scripting.FileSystemObject" )
'	set outputFile = fileSystemObject.CreateTextFile( "c:\temp\NLFRtest.xml", true )
'	outputFile.Write outputString
'	outputFile.Close
	MyRtfData = xmlDOM.xml
end function


function processRow(rowNode, language, xmlDOM, itemsCache)
	dim fieldsDictionary
	set fieldsDictionary = createDictionaryFromRowNode(rowNode)
	'translate the fields in the dictionary
	translateFields fieldsDictionary, language, itemsCache
	'replace the values in the xml by the translated values in the dictoary
	dim fieldNode
	For Each fieldNode In rowNode.ChildNodes
		fieldNode.Text = fieldsDictionary(fieldnode.nodeName)
		if lcase(right(fieldnode.nodeName, len("formatted"))) = "formatted" then
			addFormattedAttribute fieldNode, xmlDOM
		end if
	next
end function

function addFormattedAttribute(fieldNode, xmlDOM)
	dim formattedAttr
	set formattedAttr = xmlDOM.createAttribute("formatted")
	formattedAttr.nodeValue="1"
	fieldNode.setAttributeNode(formattedAttr)
end function

function translateFields(fieldsDictionary, language, itemsCache)
	dim fieldName
	for each fieldName in fieldsDictionary.Keys
		'check if there is a _guid equivalent for this field
		if fieldsDictionary.Exists(fieldName & "_guid") then
			dim guidContent
			guidContent = fieldsDictionary(fieldName & "_guid")
			'TODO: process multiple guids in one field
			'add the item to the cache if needed
			dim item
			if not itemsCache.Exists(guidContent) then
				itemsCache.Add guidContent, getItemFromGUID(guidContent)
			end if
			'get the item from the cache dictionary
			set item = itemsCache(guidContent)
			dim translation
			'Check whether its name, alias or notes field
			if lcase(right(fieldName, len("name"))) = "name" then
				translation = item.GetTXName(language, 0)
			elseif lcase(right(fieldName, len("alias"))) = "alias" then
				translation = item.GetTXAlias(language, 0)
			elseif lcase(right(fieldName, len("note"))) = "note" _
			  or lcase(right(fieldName, len("notes"))) = "notes" _ 
			  or lcase(right(fieldName, len("notes-formatted"))) = "notes-formatted" _ 
			  or lcase(right(fieldName, len("note-formatted"))) = "note-formatted" then
				translation = item.GetTXNote(language, 0)
			end if
			'set the translation value
			if len(translation) > 0 then
				fieldsDictionary(fieldName) = translation
			end if
		end if
	next
end function

function createDictionaryFromRowNode(rowNode)
	dim fieldsDictionary
	set fieldsDictionary = CreateObject("Scripting.Dictionary")
	dim fieldNode
	For Each fieldNode In rowNode.ChildNodes
		if not fieldsDictionary.Exists(fieldnode.nodeName) then
			fieldsDictionary.Add fieldnode.nodeName, fieldNode.Text
		end if
	next
	'return
	set createDictionaryFromRowNode = fieldsDictionary
end function

'msgbox MyPackageRtfData(3357,"")
function test
	dim sqlGetData
	sqlGetData = "select o.name, o.ea_guid as name_guid, o.note, o.ea_guid as note_guid, o.alias, o.ea_guid as alias_guid    " & vbNewLine & _
				" from t_object o                                                                                           " & vbNewLine & _
				" inner join t_package p on p.Package_ID = o.Package_ID                                                     " & vbNewLine & _
				" where p.ea_guid = '{B1E39191-9D6E-943C-B4AE-4DFB1ED0D488}'                                                " & vbNewLine & _
				" and o.Stereotype = 'Activity'                                                                             "
	dim outputString
	outputString =  MyRtfData(sqlGetdata, "nl")

	set fileSystemObject = CreateObject( "Scripting.FileSystemObject" )
	set outputFile = fileSystemObject.CreateTextFile( "c:\temp\NLFRtest.xml", true )
	outputFile.Write outputString
	outputFile.Close
end function 

'test