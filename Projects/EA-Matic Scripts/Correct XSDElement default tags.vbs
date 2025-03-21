'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
option explicit
'
' Script Name: 
' Author: Geert Bellekens
' Purpose: Set correct default tagged values for attribute with stereotype XSDelement (doesn't work if we set the steroetype later)
' 
' Date: 2022-05-20
' EA-Matic

function EA_OnPostNewAttribute(Info)
	dim attributeID
	attributeID = Info.Get("AttributeID")
	
	setXSDElementDefaults attributeID
	
end function

function setXSDElementDefaults(attributeID)
	dim attribute as EA.Attribute
	set attribute = Repository.GetAttributeByID(attributeID)
	dim defaultAttribute as EA.Attribute
	set defaultAttribute = getDefaultXSDElement()
	if defaultAttribute is nothing then
		exit function
	end if
	dim defaultTag as EA.AttributeTag
	for each defaultTag in defaultAttribute.TaggedValues
		dim tag as EA.AttributeTag
		for each tag in attribute.TaggedValues
			if lcase(tag.Name) = lcase(defaultTag.Name) then
'				'Debug
'				Session.Output "tagName: " & tag.Name & " tagValue: " & tag.Value & " defaultValue: " & defaultTag.Value
				if tag.Value <> defaultTag.Value then
					tag.Value = defaultTag.Value
					tag.Update
				end if
			end if
		next
	next
end function

function getDefaultXSDElement()
	set getDefaultXSDElement = nothing
	dim sqlGetData
	sqlGetData = "select a.ID from t_attribute a                            " & vbNewLine & _
				" inner join t_object o on o.Object_ID = a.Object_ID       " & vbNewLine & _
				" inner join usys_system u on u.Property = 'TemplatePkg'   " & vbNewLine & _
				" 						and u.Value = o.Package_ID         " & vbNewLine & _
				" where a.Stereotype = 'XSDelement'                        "
	dim results
	set results = getAttributesFromQuery(sqlGetData)
	dim attribute
	for each attribute in results
		set getDefaultXSDElement = attribute
		exit for
	next
end function

'function test
'	dim attribute as EA.Attribute
'	set attribute = Repository.GetTreeSelectedObject()
'	setXSDElementDefaults attribute.AttributeID
'end function
'
'test

function getAttributesFromQuery(sqlQuery)
	dim xmlResult
	xmlResult = Repository.SQLQuery(sqlQuery)
	dim attributeIDs
	attributeIDs = convertQueryResultToArray(xmlResult)
	dim attributes 
	set attributes = CreateObject("System.Collections.ArrayList")
	dim attributeID
	dim attribute as EA.Attribute
	for each attributeID in attributeIDs
		if attributeID > 0 then
			set attribute = Repository.GetAttributeByID(attributeID)
			if not attribute is nothing then
				attributes.Add(attribute)
			end if
		end if
	next
	set getattributesFromQuery = attributes
end function

'converts the query results from Repository.SQLQuery from xml format to a two dimensional array of strings
Public Function convertQueryResultToArray(xmlQueryResult)
    Dim arrayCreated
    Dim i 
    i = 0
    Dim j 
    j = 0
    Dim result()
    Dim xDoc 
    Set xDoc = CreateObject( "MSXML2.DOMDocument" )
    'load the resultset in the xml document
    If xDoc.LoadXML(xmlQueryResult) Then        
		'select the rows
		Dim rowList
		Set rowList = xDoc.SelectNodes("//Row")

		Dim rowNode 
		Dim fieldNode
		arrayCreated = False
		'loop rows and find fields
		For Each rowNode In rowList
			j = 0
			If (rowNode.HasChildNodes) Then
				'redim array (only once)
				If Not arrayCreated Then
					ReDim result(rowList.Length, rowNode.ChildNodes.Length)
					arrayCreated = True
				End If
				For Each fieldNode In rowNode.ChildNodes
					'write f
					result(i, j) = fieldNode.Text
					j = j + 1
				Next
			End If
			i = i + 1
		Next
	end if
    convertQueryResultToArray = result
End Function