'[path=\Framework\Tools\Script Management]
'[group=Script Management]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
'
' Script Name: ExportAllShapeScripts
' Author: Geert Bellekens
' Purpose: Export all shapescripts to files on the file system
' Date: 2016-07-18
'
sub main
	'get all attributes with name _image that have shapescript in the default field and a parent with stereotype «stereotype»
	dim sqlGetShapescriptAttributes
	sqlGetShapescriptAttributes = "select a.ID from (t_attribute a " & _
								  " inner join t_object o on (o.Object_ID = a.Object_ID " & _
								  "						and o.Stereotype = 'stereotype')) " & _
								  " where a.Name = '_image' " & _
								  " and a.[Default] like '<Image type=""EAShapeScript" & getWC & "'"
	dim shapeScriptAttributes
	set shapeScriptAttributes = getAttributesByQuery(sqlGetShapescriptAttributes)
	dim shapeScriptAttribute as EA.Attribute
	'get the user selected folder
	dim selectedFolder
	set selectedFolder = new FileSystemFolder
	set selectedFolder = selectedFolder.getUserSelectedFolder("")
	'loop the shape script attributes
	for each shapeScriptAttribute in shapeScriptAttributes
		Dim xDoc 
		Set xDoc = CreateObject( "MSXML2.DOMDocument.3.0" )
		dim stereotype as EA.Element
		set stereotype = Repository.GetElementByID(shapeScriptAttribute.ParentID)
		dim profile as EA.Package
		set profile = Repository.GetPackageByID(stereotype.PackageID)
		'load the resultset in the xml document
		If xDoc.LoadXML(shapeScriptAttribute.Default) Then    
			dim imageNode 
			set imageNode = xDoc.SelectSingleNode("//Image")
			dim shapeScriptEncoded
			shapeScriptEncoded = imageNode.text
			dim shapeScriptDecoded
			shapeScriptDecoded = imageNode.nodeTypedValue
			'save as temp zip file
			dim tempZipFile
			set tempZipFile = new BinaryFile
			tempZipFile.FullPath = replace(getTempFilename, ".tmp",".zip")
			tempZipFile.Contents = shapeScriptDecoded
			tempZipFile.Save
			'unzip 
			dim tempFolderPath
			tempfolderPath = unzip(tempZipFile.FullPath)
			'get the text file 
			dim tempFolder
			set tempFolder = new FileSystemFolder
			tempFolder.FullPath = tempfolderPath
			dim scriptFile
			For each scriptfile in tempfolder.TextFiles
				'save the script
				scriptFile.FullPath = selectedFolder.FullPath & "\" & profile.Name & "\" & stereotype.Name & ".shapeScript"
				scriptFile.Save
			next
			'delete the temp folder and temp file name
			tempfolder.Delete
			tempZipFile.Delete
		end if
	next
end sub

'Function SaveBinaryData(FileName, ByteArray)
'	Const adTypeBinary = 1
'	Const adSaveCreateOverWrite = 2
'	'Create Stream object
'	Dim BinaryStream
'	Set BinaryStream = CreateObject("ADODB.Stream")
'	'Specify stream type – we want To save binary data.
'	BinaryStream.Type = adTypeBinary
'	'Open the stream And write binary data To the object
'	BinaryStream.Open
'	BinaryStream.Write ByteArray
'	'Save binary data To disk
'	BinaryStream.SaveToFile FileName, adSaveCreateOverWrite
'End Function



main