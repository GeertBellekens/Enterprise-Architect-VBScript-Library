'[path=\Framework\Tools\Script Management]
'[group=Script Management]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
'
' Script Name: ExtractShapescriptsFromMDG
' Author: Geert Bellekens
' Purpose: Export all shapescripts to files on the file system
' Date: 2016-06-28
'
sub main
	'get the MDG file
	dim mdgFile
	set mdgFile = new TextFile
	if mdgFile.UserSelect("C:\Temp","MDG Files (*.xml)|*.xml") then
		'get the user selected folder
		dim selectedFolder
		set selectedFolder = new FileSystemFolder
		set selectedFolder = selectedFolder.getUserSelectedFolder("")
		'read the mdg file xml
		Dim xDoc 
		Set xDoc = CreateObject( "MSXML2.DOMDocument" )
		If xDoc.LoadXML(mdgFile.Contents) Then    
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
						'debug info
						Session.Output "saving script: " & scriptFile.FullPath
					end if
				next

			next
		end if
	end if
	
end sub


main