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
			if not item is nothing then
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
	sqlGetData = "select bpa.name, bpa.ea_guid as name_guid                                                                     " & vbNewLine & _
				" ,isnull(bpa.note, ca.Note) as [note-formatted], isnull(bpa.ea_guid, ca.ea_guid) as [note-formatted_guid]     " & vbNewLine & _
				" from t_object bp                                                                                             " & vbNewLine & _
				" left join t_object pl on pl.ParentID = bp.Object_ID                                                          " & vbNewLine & _
				" 						and pl.Stereotype = 'Pool'                                                             " & vbNewLine & _
				" left join t_object ln on ln.ParentID in (pl.Object_ID, bp.Object_ID)                                         " & vbNewLine & _
				" 						and ln.Stereotype = 'Lane'                                                             " & vbNewLine & _
				" inner join t_object bpa on bpa.ParentID in (ln.Object_ID, pl.Object_ID, bp.Object_ID)                        " & vbNewLine & _
				" 						and (bpa.Stereotype = 'Activity'                                                       " & vbNewLine & _
				" 						or bpa.Stereotype = 'StartEvent' and bpa.Note is not null)                             " & vbNewLine & _
				" left join t_objectproperties tv on tv.Object_ID = bpa.Object_ID                                              " & vbNewLine & _
				" 					and tv.Property = 'calledActivityRef'                                                      " & vbNewLine & _
				" left join t_object ca on ca.ea_guid = tv.Value                                                               " & vbNewLine & _
				" inner join t_diagramobjects do on do.Object_ID = bpa.Object_ID                                               " & vbNewLine & _
				" inner join t_diagram d on d.Diagram_ID = do.Diagram_ID                                                       " & vbNewLine & _
				" 					and d.ParentID = bp.Object_ID                                                              " & vbNewLine & _
				" where bp.Object_ID = 139544                                                                                  " & vbNewLine & _
				" order by case when bpa.Stereotype = 'StartEvent' then 1 else 2 end, do.RectLeft, do.RectTop desc, bpa.Name   "
	dim outputString
	outputString =  MyRTFData ("select bpa.name as activityname                                                                           " & vbNewLine & _
" , bpa.ea_guid as activityname_guid                                                                       " & vbNewLine & _
" , r.rolename                                                                                             " & vbNewLine & _
" , r.ea_guid as rolename_guid                                                                             " & vbNewLine & _
" , r.RASCI                                                                                                " & vbNewLine & _
" , CASE WHEN r.RASCI = 'R' THEN r.RASCI ELSE NULL END as R                                                " & vbNewLine & _
" , CASE WHEN r.RASCI = 'A' THEN r.RASCI ELSE NULL END as A                                                " & vbNewLine & _
" , CASE WHEN r.RASCI = 'S' THEN r.RASCI ELSE NULL END as S                                                " & vbNewLine & _
" , CASE WHEN r.RASCI = 'C' THEN r.RASCI ELSE NULL END as C                                                " & vbNewLine & _
" , CASE WHEN r.RASCI = 'I' THEN r.RASCI ELSE NULL END as I                                                " & vbNewLine & _
" from t_object bp                                                                                         " & vbNewLine & _
" left join t_object pl on pl.ParentID = bp.Object_ID                                                      " & vbNewLine & _
" 						and pl.Stereotype = 'Pool'                                                         " & vbNewLine & _
" left join t_object ln on ln.ParentID in (pl.Object_ID, bp.Object_ID)                                     " & vbNewLine & _
" 						and ln.Stereotype = 'Lane'                                                         " & vbNewLine & _
" inner join t_object bpa on bpa.ParentID in (ln.Object_ID, pl.Object_ID, bp.Object_ID)                    " & vbNewLine & _
" 						and bpa.Stereotype = 'Activity'                                                    " & vbNewLine & _
" left join t_objectproperties tv on tv.Object_ID = bpa.Object_ID                                          " & vbNewLine & _
" 					and tv.Property = 'calledActivityRef'                                                  " & vbNewLine & _
" left join t_object ca on ca.ea_guid = tv.Value                                                           " & vbNewLine & _
" left join                                                                                                " & vbNewLine & _
" 		(select c.Start_Object_ID, ro.Name as RoleName, ro.ea_guid, tv.VALUE as RASCI from t_connector c   " & vbNewLine & _
" 		inner join t_object ro on ro.Object_ID = c.End_Object_ID                                           " & vbNewLine & _
" 								and ro.Stereotype = 'PartnerRole'                                          " & vbNewLine & _
" 		left join t_connectortag tv on tv.ElementID = c.Connector_ID                                       " & vbNewLine & _
" 						and tv.Property = 'RASCI'                                                          " & vbNewLine & _
" 		where c.Stereotype = 'trace') r on r.Start_Object_ID = bpa.Object_ID                               " & vbNewLine & _
" inner join t_diagramobjects do on do.Object_ID = bpa.Object_ID                                            " & vbNewLine & _
" inner join t_diagram d on d.Diagram_ID = do.Diagram_ID                                                    " & vbNewLine & _
" 					and d.ParentID = bp.Object_ID                                                          " & vbNewLine & _
" where bp.ea_guid = '{8777C211-98CA-413e-ADDD-BFAC97BFA09E}'                                                                          " & vbNewLine & _
" order by do.RectLeft, do.RectTop desc, bpa.Name                                                          " & vbNewLine & _
" , CASE WHEN r.RASCI = 'R' THEN 1                                                                         " & vbNewLine & _
" 		WHEN r.RASCI = 'A' THEN 2                                                                          " & vbNewLine & _
" 		WHEN r.RASCI = 'S' THEN 3                                                                          " & vbNewLine & _
" 		WHEN r.RASCI = 'C' THEN 4                                                                          " & vbNewLine & _
" 		WHEN r.RASCI = 'I' THEN 5 END                                                                      ", "nl")

	set fileSystemObject = CreateObject( "Scripting.FileSystemObject" )
	set outputFile = fileSystemObject.CreateTextFile( "c:\temp\NLFRtest.xml", true )
	outputFile.Write outputString
	outputFile.Close
end function 

'test