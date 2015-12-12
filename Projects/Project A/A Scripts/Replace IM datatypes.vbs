'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Replace IM datatypes
' Author: Geert Bellekens
' Purpose: Replace the datatype references to the IM to local datatype references
' Date: 2015-11-02
'

dim outputTabName
outputTabName = "Replace IM datatypes"

sub main
		
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	dim response
	response = Msgbox("Replace IM datatype references in package """ & package.Name & """?", vbYesNo+vbQuestion, "Replace IM datatypes")
	'only do something if the user clicked "Yes"
	if response = vbYes then

		Repository.CreateOutputTab outputTabName
		Repository.ClearOutput outputTabName
		Repository.EnsureOutputVisible outputTabName
		dim IMPackage as EA.Package
		dim IMPackageGUID 
		IMPackageGUID = "{DC6C38F5-3043-46be-8A55-52AEAC84BAED}"
		set IMPackage = Repository.GetPackageByGuid(IMPackageGUID)
		dim IMElementsDictionary
		Set IMElementsDictionary = CreateObject("Scripting.Dictionary")
		Repository.WriteOutput outputTabName, "Loading IM datatypes...",0
		'put all IM elements in a dictionary with the ID as key
		addClassesToIdDictionary IMPackage,IMElementsDictionary
		dim mainElementsDictionary
		Set mainElementsDictionary = CreateObject("Scripting.Dictionary")
		Repository.WriteOutput outputTabName, "Loading selected package datatypes...",0
		'domain model elements are stored in a dictionary by name
		addClassesToDictionary package, mainElementsDictionary
		
		dim classElement as EA.Element
		dim attribute as EA.Attribute
		'loop the domain model elements
		for each classElement in mainElementsDictionary.Items
			Repository.WriteOutput outputTabName, "Processing " & classElement.Name ,0
			for each attribute in classElement.Attributes
				if IMElementsDictionary.Exists(attribute.ClassifierID) then
					dim attributeType
					attributeType = IMElementsDictionary(attribute.ClassifierID).Name
					if mainElementsDictionary.Exists(attributeType) then
						'if the attribute references a type from the IM, and the equivalent exists in the main dictionary then we use the element from the main dictionary
						Repository.WriteOutput outputTabName, "replacing type " & attribute.Type & " for attribute " & classElement.Name & "." & attribute.Name,0
						attribute.ClassifierID = mainElementsDictionary(attributeType).ElementID
						attribute.Update
					end if
				end if	
			next 'Attribute			
		next 'Element
		msgbox "finished!"
	end if
end sub

main


function addClassesToDictionary(package, dictionary)
	dim classElement as EA.Element
	dim subpackage as EA.Package
	'process owned elements
	for each classElement in package.Elements
		if (classElement.Type = "Class" OR classElement.Type = "Enumeration" OR classElement.Type = "DataType" ) _
			AND len(classElement.Name) > 0 _ 
			AND not dictionary.Exists(classElement.Name) then
			Repository.WriteOutput outputTabName, "Loading element: " & classElement.Name ,0
			dictionary.Add classElement.Name,  classElement
		end if
	next
	'process subpackages
	for each subpackage in package.Packages
		addClassesToDictionary subpackage, dictionary
	next
end function

function addClassesToIdDictionary(package, dictionary)
	dim classElement as EA.Element
	dim subpackage as EA.Package
	'process owned elements
	for each classElement in package.Elements
		if (classElement.Type = "Class" OR classElement.Type = "Enumeration" OR classElement.Type = "DataType" ) _
			AND len(classElement.Name) > 0 _ 
			AND not dictionary.Exists(classElement.Name) then
			Repository.WriteOutput outputTabName, "Loading element: " & classElement.Name ,0
			dictionary.Add classElement.ElementID,  classElement
		end if
	next
	'process subpackages
	for each subpackage in package.Packages
		addClassesToIdDictionary subpackage, dictionary
	next
end function