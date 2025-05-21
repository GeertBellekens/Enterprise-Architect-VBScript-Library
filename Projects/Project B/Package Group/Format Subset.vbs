'[path=\Projects\Project B\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Format Subset
' Author: Geert Bellekens
' Purpose: Formats the subset to make it similar to the original. Creates subPackages and diagrams to mirror the original package
' Date: 2019-01-03
'
const outPutName = "Format Subset"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get selected package
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage
	if selectedPackage is nothing then 
		exit sub
	end if
	'get the other package
	MsgBox "Please select the source package"
	dim sourcePackage as EA.Package
	set sourcePackage = selectPackage()
	if not sourcePackage is nothing then
		dim response
		response = Msgbox("Do you want to format the subset '" & selectedPackage.Name & _
		"' according to the source package '" & sourcePackage.Name & "'" , vbYesNo + vbQuestion, "Format subset")
		If response = vbYes Then
			'set timestamp
			Repository.WriteOutput outPutName, now() & " Start Format Subset for '" & selectedPackage.Name &  "'", 0
			'actually do the work
			formatSubset selectedPackage, sourcePackage
			'set timestamp
			Repository.WriteOutput outPutName, now() & " Finished Format Subset for '" & selectedPackage.Name &  "'", 0
		end if
	end if
end sub

function formatSubset(package, sourcePackage)
	'inform user of progress
	Repository.WriteOutput outPutName, now() & " Getting subset elements", 0
	'make a list of all elements in this package
	dim elements
	set elements = getElementsList(package)
	'inform user of progress
	Repository.WriteOutput outPutName, now() & " Getting source elements", 0
	'make a list of all elements in the sourcePackage
	dim sourceElements
	set sourceElements = getElementsList(sourcePackage)
	'inform user of progress
	Repository.WriteOutput outPutName, now() & " Matching subset elements with source elements", 0
	'make a list of all elements in the elementsList that have an equivalent int the source package list
	dim correspondingElements
	set correspondingElements = getCorrespondingList(elements, sourceElements)
	dim packages
	set packages = CreateObject("Scripting.Dictionary")
	'put the root packages in the dictionary
	packages.Add sourcePackage.PackageID, package
	dim element as EA.Element
	dim sourceID
	dim sourceElement as EA.Element
	for each sourceID in correspondingElements.Keys
		set element = correspondingElements(sourceID)
		set sourceElement = sourceElements(sourceID)
		processElement element, sourceElement, package, sourcePackage, packages
	next
	'get diagram translations
	dim translations
	set translations = getTranslations(package)
	'process diagrams
	processDiagrams packages, correspondingElements, translations
	'reload package
	Repository.RefreshModelView package.PackageID
end function

function getTranslations(package)
	dim translationsDictionary
	set translationsDictionary = CreateObject("Scripting.Dictionary")
	dim translationsTag as EA.TaggedValue
	for each translationsTag in package.Element.TaggedValues
		if lcase(translationsTag.Name) = "translations" then
			dim keyValuesString
			keyValuesString = translationsTag.Value
			set translationsDictionary = getKeyValuePairs(keyValuesString)
		end if
	next
	'return
	set getTranslations = translationsDictionary
end function

function processDiagrams(packages, correspondingElements, translations)
	dim currentSourcePackage as EA.Package
	dim currentTargetPackage as EA.Package
	dim sourcePackageID
	for each sourcePackageID in packages.Keys
		set currentSourcePackage = Repository.GetPackageByID(sourcePackageID)
		set currentTargetPackage = packages(sourcePackageID)
		'loop diagrams in current source packages
		dim sourceDiagram as EA.Diagram
		for each sourceDiagram in currentSourcePackage.Diagrams
			'process diagram
			processDiagram sourceDiagram, currentTargetPackage, correspondingElements, translations
		next
	next
end function

function translateDiagramName(name, translations)
	dim original
	dim translation
	dim returnName
	returnName = name
	for each original in translations.keys
		translation = translations.Item(original)
		returnName = replace(returnName, original, translation)
	next
	'return
	translateDiagramName = returnName
end function

function processDiagram(sourceDiagram, currentTargetPackage, correspondingElements, translations)
	Repository.WriteOutput outPutName, now() & " Processing diagram '" & currentTargetPackage.Name  & "." & sourceDiagram.Name &  "'", 0
	'find or create corresponding diagram
	dim targetDiagram as EA.Diagram
	set targetDiagram = nothing
	dim tempDiagram as EA.Diagram
	'get the translated diagram name
	dim translatedDiagramName
	translatedDiagramName = translateDiagramName(sourceDiagram.Name, translations)
	'search corresponding diagram
	for each tempDiagram in currentTargetPackage.Diagrams
		if tempDiagram.Name = translatedDiagramName _
		  AND tempDiagram.Type = sourceDiagram.Type then
			set targetDiagram = tempDiagram
			exit for
		end if
	next
	'create if not found
	if targetDiagram is nothing then
		set targetDiagram = currentTargetPackage.Diagrams.AddNew(translatedDiagramName, sourceDiagram.Type)
		targetDiagram.Update
	end if
	'set diagram properties
	targetDiagram.ExtendedStyle = sourceDiagram.ExtendedStyle
	targetDiagram.Orientation = sourceDiagram.Orientation
	targetDiagram.StyleEx = sourceDiagram.StyleEx
	targetDiagram.Update
	'process diagram objects
	dim sourceDiagramObject as EA.DiagramObject
	for each sourceDiagramObject in sourceDiagram.DiagramObjects
		dim targetElement as EA.Element
		if correspondingElements.Exists(sourceDiagramObject.ElementID) then
			set targetElement = correspondingElements(sourceDiagramObject.ElementID)
			'find the corresponding diagramObject
			dim targetDiagramObject as EA.DiagramObject
			set targetDiagramObject = nothing
			dim tempDiagramObject as EA.DiagramObject
			for each tempDiagramObject in targetDiagram.DiagramObjects
				if tempDiagramObject.ElementID = targetElement.ElementID then
					set targetDiagramObject = tempDiagramObject
				end if
			next
			'if not found then create it
			if targetDiagramObject is nothing then
				set targetDiagramObject = targetDiagram.DiagramObjects.AddNew("l=10;r=50;t=10;b=50;","")
				targetDiagramObject.ElementID = targetElement.ElementID
				targetDiagramObject.Update
				'set size and location
				targetDiagramObject.left = sourceDiagramObject.left
				targetDiagramObject.right = sourceDiagramObject.right
				targetDiagramObject.top = sourceDiagramObject.top
				targetDiagramObject.bottom = sourceDiagramObject.bottom
				targetDiagramObject.update
			end if
		end if
	next
end function

function processElement(element, sourceElement, package, sourcePackage, packages)
	Repository.WriteOutput outPutName, now() & " Processing element '" & sourceElement.Name &  "'", 0
	'make sure to create the package structure for the package
	dim targetPackage as EA.Package
	'get the corresponding target package
	set targetPackage = createPackageStructure (sourceElement.PackageID, package, sourcePackage, packages)
	'move the element tot the target package
	if element.PackageID <> targetPackage.PackageID or _
	  element.TreePos <> sourceElement.TreePos then
		'set package 
		element.PackageID = targetPackage.PackageID 
		'set tree pos
		element.TreePos = sourceElement.TreePos
		element.update
	end if
end function

function createPackageStructure (sourcePackageID, package, sourcePackage, packages)
	if not packages.Exists(sourcePackageID) then
		dim currentSourcePackage as EA.Package
		set currentSourcePackage = Repository.GetPackageByID(sourcePackageID)
		'get the parent target package
		dim parentPackage as EA.Package
		set parentPackage = createPackageStructure (currentSourcePackage.ParentID, package, sourcePackage, packages)
		'create new target package
		dim newPackage as EA.Package
		set newPackage = nothing
		'check if it already exists
		dim tempPackage as EA.Package
		for each tempPackage in parentPackage.Packages
			if tempPackage.Name = currentSourcePackage.Name then
				set newPackage = tempPackage
				exit for
			end if
		next
		'create new if not found
		if newPackage is nothing then
			set newPackage = parentPackage.Packages.AddNew(currentSourcePackage.Name, "")
			newPackage.TreePos = currentSourcePackage.TreePos
			newPackage.Notes = currentSourcePackage.Notes
			newPackage.Update
		end if
		'add the target package to the dictionary
		packages.Add sourcePackageID, newPackage
	end if
	'return the corresponding target package
	set createPackageStructure = packages(sourcePackageID)
end function


function getCorrespondingList(elements, sourceElements)
	dim element as EA.Element
	dim result
	set result = CreateObject("Scripting.Dictionary")
	for each element in elements.Items
		dim sqlGetData
		sqlGetData = "select distinct o.Object_ID from t_objectproperties tv    " & vbNewLine & _
					" inner join t_object o on o.ea_guid = tv.Value            " & vbNewLine & _
					" where tv.Property = 'sourceElement'                      " & vbNewLine & _
					" and tv.Object_ID = " & element.ElementID
		dim results
		set results = getVerticalArrayListFromQuery(sqlGetData)
		if results.Count > 0 then
			dim correspondingID
			correspondingID = Clng(results(0)(0))
			if sourceElements.Exists(correspondingID) then
				dim sourceElement as EA.Element
				set sourceElement = sourceElements(correspondingID)
				result.Add correspondingID, element
			end if
		end if
	next
	'return
	set getCorrespondingList = result
end function 

function getElementsList(package)
	dim sqlGetElements
	sqlGetElements = "select o.Object_ID  from t_object o where  o.Package_ID in (" & getPackageTreeIDString(package) & ")"
	dim elements
	set elements = getElementsFromQuery(sqlGetElements)
	'make a dictionary with elementID as key
	dim result
	set result = CreateObject("Scripting.Dictionary")
	dim element as EA.Element
	for each element in elements
		if not result.Exists(element.ElementID) then
			result.Add element.ElementID, element
		end if
	next
	set getElementsList = result
end function




main