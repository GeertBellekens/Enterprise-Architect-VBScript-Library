option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Link to Domain Model
' Author: Geert Bellekens
' Purpose: Link elements with classes in the domain model based on their name.
' Date: 15/11/2015
'


sub main
	'let the user select the domain model package
	dim response
	response = Msgbox ("Please select the domain model package", vbOKCancel+vbQuestion, "Selecte domain model")
	if response = vbOK then
		
		dim domainModelPackage as EA.Package
		set domainModelPackage = selectPackage()
		if not domainModelPackage is nothing then
			Repository.CreateOutputTab "Link to Domain Model"
			Repository.ClearOutput "Link to Domain Model"
			Repository.EnsureOutputVisible "Link to Domain Model"
			'link selected elements to the domain model elements
			linkSelectionToDomainModel(domainModelPackage)
			'tell the user we are done
			Repository.WriteOutput "Link to Domain Model", "Finished!", 0
		end if
	end if
end sub

'this is the actual call to the main function
main

function linkSelectionToDomainModel(domainModelPackage)
	'first get the pattern from all the classes in the domain model
	dim dictionary
	Set dictionary = CreateObject("Scripting.Dictionary")
	
	'create domain model dictionary
	'tell the user what we are doing
	Repository.WriteOutput "Link to Domain Model", "Creating domain model dictionary", 0
	addToClassDictionary domainModelPackage.PackageGUID, dictionary
	
	'tell the user what we are doing
	Repository.WriteOutput "Link to Domain Model", "Interpreting dictionary", 0
	' and prepare the regex object
	dim pattern
	'create the pattern based on the names in the dictionary
	pattern = createRegexPattern(dictionary)
	Dim regExp  
	Set regExp = CreateObject("VBScript.RegExp")
	regExp.Global = True   
	regExp.IgnoreCase = False
	regExp.Pattern = pattern
	
	' Get the type of element selected in the Project Browser
	dim treeSelectedType
	treeSelectedType = Repository.GetTreeSelectedItemType()
	' Either process the selected Element, or process all elements in the selected package
	select case treeSelectedType
		case otElement
			' Code for when an element is selected
			dim selectedElements as EA.Collection
			set selectedElements = Repository.GetTreeSelectedElements()
			'link the selected elements with the 
			linkDomainClassesWithElements dictionary,regExp,selectedElements 
		case otPackage
			' Code for when a package is selected
			dim selectedPackage as EA.Package
			set selectedpackage = Repository.GetTreeSelectedObject()
			'link use domain classes with the elements in the selected package
			linkDomainClassesWithElementsInPackage dictionary, regExp,selectedPackage
		case else
			' Error message
			Session.Prompt "You have to select Elements or a Package", promptOK
	end select
end function

'this function will get all elements in the given package and subpackages recursively and link them to the domain classes
function linkDomainClassesWithElementsInPackage(dictionary,regExp,selectedPackage)
	dim packageList 
	set packageList = getPackageTree(selectedPackage)
	dim packageIDString
	packageIDString = makePackageIDString(packageList)
	dim getElementsSQL
	getElementsSQL = "select o.Object_ID from t_object o where o.Package_ID in (" & packageIDString & ")"
	dim usecases
	set usecases = getElementsFromQuery(getElementsSQL)
	linkDomainClassesWithElements dictionary,regExp,usecases
end function


function linkDomainClassesWithElements(dictionary,regExp,elements)
	dim element as EA.Element
	'loop de elements
	for each element in elements
		'tell the user what we are doing
		Repository.WriteOutput "Link to Domain Model", "Linking element: " & element.Name, 0
		'first remove all automatic traces
		removeAllAutomaticTraces element
		'match based on notes and linked document
		dim elementText
		'get full text (name +notes + linked document + scenario names + scenario notes)
		elementText = element.Name
		elementText = elementText & vbNewLine & Repository.GetFormatFromField("TXT",element.Notes)
		elementText = elementText & vbNewLine & getLinkedDocumentContent(element, "TXT")
		elementText = elementText & vbNewLine & getTextFromScenarios(element)
		dim matches
		set matches = regExp.Execute(elementText)
		'for each match create a «trace» link with the element
		linkMatchesWithelement matches, element, dictionary
		'link based on text in scenariosteps
		dim scenario as EA.Scenario
		'get all dependencies left
		dim dependencies
		set dependencies = getDependencies(element)
		'loop scenarios
		for each scenario in element.Scenarios
			dim scenarioStep as EA.ScenarioStep
			for each scenarioStep in scenario.Steps
				'first remove any additional terms in the uses field
				scenarioStep.Uses = removeAddionalUses(dependencies, scenarioStep.Uses)
				set matches = regExp.Execute(scenarioStep.Name)
				dim classesToMatch 
				set classesToMatch = getClassesToMatchDictionary(matches, dictionary)
				dim classToMatch as EA.Element
				for each classToMatch in classesToMatch.Items
					if not instr(scenarioStep.Uses,classToMatch.Name) > 0 then
						scenarioStep.Uses = scenarioStep.Uses & " " & classToMatch.Name
					end if
					'create the dependency between the use case and the domain model class
					linkElementsWithAutomaticTrace element, classToMatch
				next
				'save scenario step
				scenarioStep.Update
				scenario.Update
			next
		next
	next
end function

function linkMatchesWithelement(matches, element, dictionary)
	dim classesToMatch
	'get the classes to match
	Set classesToMatch = getClassesToMatchDictionary(matches,dictionary)
	dim classToMatch as EA.Element
	'actually link the classes
	for each classToMatch in classesToMatch.Items
		linkElementsWithAutomaticTrace element, classToMatch
	next
end function


'get the text from the scenarios name and notes
function getTextFromScenarios(element)
	dim scenario as EA.Scenario
	dim scenarioText
	scenarioText = "" 
	for each scenario in element.Scenarios
		scenarioText = scenarioText & vbNewLine & scenario.Name
		scenarioText = scenarioText & vbNewLine & Repository.GetFormatFromField("TXT",scenario.Notes)
	next
	getTextFromScenarios = scenarioText
end function

function removeAddionalUses(dependencies, uses)
	dim dependency
	dim filteredUses
	filteredUses = ""
	if len(uses) > 0 then
		for each dependency in dependencies.Keys
			if Instr(uses,dependency) > 0 then
				if len(filteredUses) > 0 then
					filteredUses = filteredUses & " " & dependency
				else
					filteredUses = dependency
				end if
			end if
		next
	end if
	removeAddionalUses = filteredUses
end function

'returns a dictionary of elements with the name as key and the element as value.
function getDependencies(element)
	dim getDependencySQL
	getDependencySQL =  "select dep.Object_ID from ( t_object dep " & _
						" inner join t_connector con on con.End_Object_ID = dep.Object_ID)   " & _
						" where con.Connector_Type = 'Dependency'  " & _
						" and con.Start_Object_ID = " & element.ElementID   
	set getDependencies = getElementDictionaryFromQuery(getDependencySQL)
end function

function removeAllAutomaticTraces(element)
		dim i
		dim connector as EA.Connector
		'remove all the traces to domain model classes
		for i = element.Connectors.Count -1 to 0 step -1
			set connector = element.Connectors.GetAt(i)
			if connector.Alias = "automatic" and connector.Stereotype = "trace" then
				element.Connectors.DeleteAt i,false 
			end if
		next
end function

function getClassesToMatchDictionary(matches, allClassesDictionary)
	dim match
	dim classesToMatch
	dim className
	Set classesToMatch = CreateObject("Scripting.Dictionary")
	'create list of elements to link
	For each match in matches
		if not allClassesDictionary.Exists(match.Value) then
			'strip the last 's'
			className = left(match.Value, len(match.Value) -1)
		else
			className = match.Value
		end if
		if not classesToMatch.Exists(className) then
			classesToMatch.Add className, allClassesDictionary(className)
		end if
	next
	set getClassesToMatchDictionary = classesToMatch
end function

'Create a «trace» relation between source and target with "automatic" as alias
function linkElementsWithAutomaticTrace(sourceElement, targetElement)
	dim linkExists 
	linkExists = false
	'first make sure there isn't already a trace relation between the two
	dim existingConnector as EA.Connector
	'make sure we are using the up-to-date connectors collection
	sourceElement.Connectors.Refresh
	for each existingConnector in sourceElement.Connectors
		if existingConnector.SupplierID = targetElement.ElementID _
		   and existingConnector.Stereotype = "trace" then
		   linkExists = true
		   exit for
		end if
	next
	if not linkExists then
		'tell the user what we are doing
		Repository.WriteOutput "Link to Domain Model", "Adding trace between " &sourceElement.Name & " and " & targetElement.Name, 0
		dim trace as EA.Connector
		set trace = sourceElement.Connectors.AddNew("","trace")
		trace.Alias = "automatic"
		trace.SupplierID = targetElement.ElementID
		trace.Update
	end if
end function

function addToClassDictionary(PackageGUID, dictionary)
	dim package as EA.Package
	set package = Repository.GetPackageByGuid(PackageGUID)
	
	'get the classes in the dictionary (recursively
	addClassesToDictionary package, dictionary
end function

function addClassesToDictionary(package, dictionary)
	dim classElement as EA.Element
	dim subpackage as EA.Package
	'process owned elements
	for each classElement in package.Elements
		if classElement.Type = "Class" AND len(classElement.Name) > 0 AND not dictionary.Exists(classElement.Name) then
			dictionary.Add classElement.Name,  classElement
		end if
	next
	'process subpackages
	for each subpackage in package.Packages
		addClassesToDictionary subpackage, dictionary
	next
end function


'Create a reges pattern like this "\b(name1|name2|name3)\b" based on the 
function createRegexPattern(dictionary)
	Dim patternString
	dim className
	'add begin
	patternString = "\b("
	dim addPipe
	addPipe = FALSE
	for each className in dictionary.Keys
			if addPipe then
				patternString = patternString & "|"
			else
				addPipe = True
			end if
			patternString = patternString & className
	next
	'add end
	patternString = patternString & ")s?\b"
	'return pattern
	createRegexPattern = patternString
end function

'returns an ArrayList of the given package and all its subpackages recursively
function getPackageTree(package)
	dim packageList
	set packageList = CreateObject("System.Collections.ArrayList")
	addPackagesToList package, packageList
	set getPackageTree = packageList
end function

'make an id string out of the package ID of the given packages
function makePackageIDString(packages)
	dim package as EA.Package
	dim idString
	idString = ""
	dim addComma 
	addComma = false
	for each package in packages
		if addComma then
			idString = idString & ","
		else
			addComma = true
		end if
		idString = idString & package.PackageID
	next 
	'if there are no packages then we return "0"
	if packages.Count = 0 then
		idString = "0"
	end if
	'return idString
	makePackageIDString = idString
end function

'returns an ArrayList with the elements accordin tot he ObjectID's in the given query
function getElementsFromQuery(sqlQuery)
	dim elements 
	set elements = Repository.GetElementSet(sqlQuery,2)
	dim result
	set result = CreateObject("System.Collections.ArrayList")
	dim element
	for each element in elements
		result.Add Element
	next
	set getElementsFromQuery = result
end function

'returns a dictionary of all elements in the query with their name as key, and the element as value.
'for elements with the same name only one will be returned
function getElementDictionaryFromQuery(sqlQuery)
	dim elements 
	set elements = Repository.GetElementSet(sqlQuery,2)
	dim result
	set result = CreateObject("Scripting.Dictionary")
	dim element
	for each element in elements
		if not result.Exists(element.Name) then
		result.Add element.Name, element
		end if
	next
	set getElementDictionaryFromQuery = result
end function

'gets the content of the linked document in the given format (TXT, RTF or EA)
function getLinkedDocumentContent(element, format)
	dim linkedDocumentRTF
	dim linkedDocumentEA
	dim linkedDocumentPlainText
	linkedDocumentRTF = element.GetLinkedDocument()
	if format = "RTF" then
		getLinkedDocumentContent = linkedDocumentRTF
	else
		linkedDocumentEA = Repository.GetFieldFromFormat("RTF",linkedDocumentRTF)
		if format = "EA" then
			getLinkedDocumentContent = linkedDocumentEA
		else
			linkedDocumentPlainText = Repository.GetFormatFromField("TXT",linkedDocumentEA)
			getLinkedDocumentContent = linkedDocumentPlainText
		end if
	end if
end function

'let the user select a package
function selectPackage()
	dim documentPackageElementID 		
	documentPackageElementID = Repository.InvokeConstructPicker("IncludedTypes=Package") 
	if documentPackageElementID > 0 then
		dim packageElement as EA.Element
		set packageElement = Repository.GetElementByID(documentPackageElementID)
		dim package as EA.Package
		set package = Repository.GetPackageByGuid(packageElement.ElementGUID)
	else
		set package = nothing
	end if 
	set selectPackage = package
end function

'add the given package and all subPackges to the list (recursively
function addPackagesToList(package, packageList)
	dim subPackage as EA.Package
	'add the package itself
	packageList.Add package
	'add subpackages
	for each subPackage in package.Packages
		addPackagesToList subPackage, packageList
	next
end function