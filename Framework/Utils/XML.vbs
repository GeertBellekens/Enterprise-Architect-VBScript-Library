'[path=\Framework\Utils]
'[group=Utils]
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

public Function sanitizeXMLString(invalidString)
	Dim tmp, i 
	tmp = invalidString
	'first replace ampersand
	tmp = Replace(tmp, chr(38), "&amp;") 
	'then the other special characters
	For i = 160 to 255
		tmp = Replace(tmp, chr(i), "&#" & i & ";")
	Next
	'and then the special characters
	tmp = Replace(tmp, chr(34), "&quot;")
	tmp = Replace(tmp, chr(39), "&apos;")
	tmp = Replace(tmp, chr(60), "&lt;")
	tmp = Replace(tmp, chr(62), "&gt;")
	'tmp = Replace(tmp, chr(32), "&nbsp;")
	sanitizeXMLString = tmp
end function

'sub test
'	dim test
'	test = sanitizeXMLString("invali""d'strèiçng<&>")
'	Session.Output "sanitized: " & test
'end sub
'test

'returns the string encoded int he base64 zipped contents of the given xml string
public function decodeBase64zippedXML(xmlString,nodeName)
		Dim xDoc 
		Set xDoc = CreateObject( "MSXML2.DOMDocument" )
		decodeBase64zippedXML = ""
		'get the stereotype
		'load the resultset in the xml document
		If xDoc.LoadXML(xmlString) Then    
			dim contentsNode 
			set contentsNode = xDoc.SelectSingleNode("//" & nodeName)
			if not contentsNode is nothing then
				dim contentsDecoded
				contentsDecoded = contentsNode.nodeTypedValue
				'save as temp zip file
				dim tempZipFile
				set tempZipFile = new BinaryFile
				tempZipFile.FullPath = replace(getTempFilename, ".tmp",".zip")
				tempZipFile.Contents = contentsDecoded
				tempZipFile.Save
				'unzip 
				dim tempFolderPath
				tempfolderPath = unzip(tempZipFile.FullPath)
				'get the text file 
				dim tempFolder
				set tempFolder = new FileSystemFolder
				tempFolder.FullPath = tempfolderPath
				dim contentsFile
				For each contentsFile in tempfolder.TextFiles
					decodeBase64zippedXML = contentsFile.Contents
					'there should be only one file
					exit for
				next
				'delete the temp folder and temp file name
				tempfolder.Delete
				tempZipFile.Delete
			end if
		end if
end function