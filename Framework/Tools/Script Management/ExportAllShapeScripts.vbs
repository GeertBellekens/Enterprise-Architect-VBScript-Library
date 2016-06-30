'[path=\Framework\Tools\Script Management]
'[group=Script Management]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
'
' Script Name: ExportAllShapeScripts
' Author: Geert Bellekens
' Purpose: Export all shapescripts to files on the file system
' Date: 2016-06-18
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
		'get the stereotype
		dim stereotype as EA.Element
		set stereotype = Repository.GetElementByID(shapeScriptAttribute.ParentID)
		dim profile as EA.Package
		set profile = findProfilePackage(stereotype)
		'load the resultset in the xml document
		dim shapeScript
		shapeScript = decodeBase64zippedXML(shapeScriptAttribute.Default,"Image")
		if len(shapeScript) > 0 then
			dim scriptFile
			set scriptFile = New TextFile
			scriptfile.Contents = shapeScript
			'save the script
			scriptFile.FullPath = selectedFolder.FullPath & "\" & profile.Name & "\" & stereotype.Name & ".shapeScript"
			scriptFile.Save
			'debug info
			Session.Output "saving script: " & scriptFile.FullPath
		end if
	next
end sub

'finds the owning package with stereotype Profile.
'if not found returns the owning package of the attribute
function findProfilePackage(stereotype)
	dim profile as EA.Package
	set profile = Repository.GetPackageByID(stereotype.PackageID)
	if profile.StereotypeEx <> "profile" then
		dim parentProfile as EA.Package
		set parentProfile = getParentProfilePackage(profile)
		if not parentProfile is nothing then
			set profile = parentProfile
		end if
	end if
	set findProfilePackage = profile
end Function

'recurse up the tree until the profile package is found
function getParentProfilePackage(profile)
	dim parentProfile as EA.Package
	set parentProfile = nothing
	if profile.ParentID > 0 then
		set parentProfile = Repository.GetPackageByID(profile.ParentID)
		if parentProfile.StereotypeEx <> "profile" then
			set parentProfile = getParentProfilePackage(parentProfile)
		end if
	end if
	set getParentProfilePackage = parentProfile
end function

main