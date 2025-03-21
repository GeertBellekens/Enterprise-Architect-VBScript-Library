'[path=\Framework\Wrappers\]
'[group=Wrappers]


!INC Utils.Include
!INC Local Scripts.EAConstants-VBScript


'
' Script Name: Shapescript
' Author: Geert Bellekens
' Purpose: Reusable methods to save shapescripts from MDG files
' Date: 2024-08-16

function exportShapescriptsFromMDGFile(mdgFilePath, outputFolderPath)
	'get the MDG file
	dim mdgFile
	set mdgFile = new TextFile
	if len(mdgFilePath) = 0 then
		if not mdgFile.UserSelect("C:\Temp","MDG Files (*.xml)|*.xml") then
			exit function
		end if
	else
		mdgFile.FullPath = mdgFilePath
		mdgFile.LoadContents
	end if
	'get the output folder
	dim selectedFolder
	set selectedFolder = new FileSystemFolder
	if len(outputFolderPath) = 0 then
		set selectedFolder = selectedFolder.getUserSelectedFolder("")
	else
		selectedFolder.FullPath = outputFolderPath
	end if
	'read the mdg file xml
	Dim xDoc 
	Set xDoc = CreateObject( "MSXML2.DOMDocument" )
	If xDoc.LoadXML(mdgFile.Contents) Then    
		Session.Output "mdg file read"
		'get the profile nodes
		dim profiles
		set profiles = xDoc.SelectNodes("//UMLProfile")
		dim profile
		for each profile in profiles
			'get the documentation tag
			dim documentationTag
			set documentationTag = profile.SelectSingleNode("./Documentation")
			'get the profile name
			dim profileName
			profileName = documentationTag.Attributes.GetNamedItem("name").Text
			'loop the stereotype nodes nodes
			dim stereotypeNodes
			set stereotypeNodes = profile.SelectNodes("./Content/Stereotypes/Stereotype")
			dim stereotypeNode
			for each stereotypeNode in stereotypeNodes	
				'get the name fo the stereotype
				dim stereotypeName
				stereotypeName = stereotypeNode.Attributes.GetNamedItem("name").Text
				'get the shapescript
				dim shapeScript
				shapeScript = decodeBase64zippedXML(stereotypeNode.xml,"Image")
				if len(shapeScript) > 0 then
					dim scriptFile
					set scriptFile = New TextFile
					scriptfile.Contents = shapeScript
					'save the script
					scriptFile.FullPath = selectedFolder.FullPath & "\" & profileName & "\" & stereotypeName & ".shapeScript"
					scriptFile.Save
				end if
			next
		next
	end if

end function