
!INC Local Scripts.EAConstants-VBScript

'Author: Geert Bellekens
'Date: 2015-12-07

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

