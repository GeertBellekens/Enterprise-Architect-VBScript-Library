option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.Util

'
' Script Name: Link with Logical Data Model
' Author: Geert Bellekens
' Purpose: Links uses cases and rules to elements Logical Data Model entities they use in their scenariosteps for the use case and in the notes/linked documents for the Business Rules
' Date: 28/09/2015
'


sub main

	'reference to the domeain model package
	dim domainModelPackageGUID 
	domainModelPackageGUID = "{F7AD826C-6BD6-4513-82B7-E71737A50DB0}"
'	dim extendedDomainModelPackageGUID 
'	extendedDomainModelPackageGUID = "{967ED68D-A6D0-45ea-BBDE-F87E2BE34CE0}"
	
	'create output tab
	Repository.CreateOutputTab "Link to Logical Data Model"
	Repository.ClearOutput "Link to Logical Data Model"
	Repository.EnsureOutputVisible "Link to Logical Data Model"
	
	
	'first get the pattern from all the classes in the Logical Data Model
	dim dictionary
	Set dictionary = CreateObject("Scripting.Dictionary")
	'Logical Data Model
	Repository.WriteOutput "Link to Logical Data Model", "Creating dictionary from Logical Data Model", 0
	addToClassDictionary domainModelPackageGUID, dictionary
	'extended Logical Data Model
	'addToClassDictionary extendedDomainModelPackageGUID, dictionary
	
	' and prepare the regex object
	dim pattern
	'create the pattern based on the names in the dictionary
	Repository.WriteOutput "Link to Logical Data Model", "Creating regular expression", 0
	pattern = createRegexPattern(dictionary)
	Dim regExp  
	Set regExp = CreateObject("VBScript.RegExp")
	regExp.Global = True   
	regExp.IgnoreCase = False
	regExp.Pattern = pattern
	
	' Get the type of element selected in the Project Browser
	dim treeSelectedType
	treeSelectedType = Repository.GetTreeSelectedItemType()
	
	select case treeSelectedType
	
		case otElement
			' Code for when an element is selected
			dim selectedElements as EA.Collection
			set selectedElements = Repository.GetTreeSelectedElements()
			'process use cases
			dim usecases
			set usecases = getUseCasesFromEACollection(selectedElements)
			linkDomainClassesWithUseCases dictionary,regExp,usecases 
			'process business rules
			dim businessRules
			set businessRules = getBusinessRulesFromEACollection(selectedElements)
			Session.Output "business rules found: " & businessRules.Count
			linkDomainClassesWithBusinessRules dictionary,regExp, businessRules
		case otPackage
			' Code for when a package is selected
			dim selectedPackage as EA.Package
			set selectedpackage = Repository.GetTreeSelectedObject()
			'link use domain classes with use cases under the selected package
			linkDomainClassesWithUseCasesInPackage dictionary, regExp,selectedPackage
			'link the domain classes with the business rules under the selected package
			linkDomainClassesWithBusinessRulesInPackage dictionary, regExp,selectedPackage		
		case else
			' Error message
			Repository.WriteOutput "Link to Logical Data Model", "Error: wrong type selected. You need to select a package or one or more elements", 0
			
	end select

	Repository.WriteOutput "Link to Logical Data Model", "Finished!", 0
end sub

function getUseCasesFromEACollection(selectedElements)
	dim usecases 
	set usecases = CreateObject("System.Collections.ArrayList")
	dim element as EA.Element
	for each element in selectedElements
		if element.Type = "UseCase" then
			usecases.Add element
		end if
	next
	set getUseCasesFromEACollection = usecases
end function

function getBusinessRulesFromEACollection(selectedElements)
	dim businessRules 
	set businessRules = CreateObject("System.Collections.ArrayList")
	dim element as EA.Element
	for each element in selectedElements
		if element.Type = "Activity" and element.Stereotype = "Business Rule" then
			businessRules.Add element
		end if
	next
	set getBusinessRulesFromEACollection = businessRules
end function

function linkDomainClassesWithUseCasesInPackage(dictionary,regExp,selectedPackage)
	dim packageList 
	set packageList = getPackageTree(selectedPackage)
	dim packageIDString
	packageIDString = makePackageIDString(packageList)
	dim getElementsSQL
	getElementsSQL = "select uc.Object_ID from t_object uc where uc.Object_Type = 'UseCase' and uc.Package_ID in (" & packageIDString & ")"
	dim usecases
	set usecases = getElementsFromQuery(getElementsSQL)
	linkDomainClassesWithUseCases dictionary,regExp,usecases
end function

function linkDomainClassesWithUseCases(dictionary,regExp,usecases)
	Session.Output usecases.Count & " use cases found"
	dim usecase as EA.Element
	
	'loop de use cases
	for each usecase in usecases
		Repository.WriteOutput "Link to Logical Data Model", "Processing use case: " & usecase.Name, 0
		'first remove all automatic traces
		removeAllAutomaticTraces usecase
		'get all dependencies left
		dim dependencies
		set dependencies = getDependencies(usecase)
		dim scenario as EA.Scenario
		'loop scenarios
		for each scenario in usecase.Scenarios
			dim scenarioStep as EA.ScenarioStep
			for each scenarioStep in scenario.Steps
				'first remove any additional terms in the uses field
				scenarioStep.Uses = removeAddionalUses(dependencies, scenarioStep.Uses)
				dim matches
				set matches = regExp.Execute(scenarioStep.Name)
				dim classesToMatch 
				set classesToMatch = getClassesToMatchDictionary(matches, dictionary)
				dim classToMatch as EA.Element
				for each classToMatch in classesToMatch.Items
					Session.Output "scenarioStep.Uses before " & scenarioStep.Uses
					if not instr(scenarioStep.Uses,"LDM-" & classToMatch.Name) > 0 then
						if len(scenarioStep.Uses) > 0 then 
							'add a space if needed
							scenarioStep.Uses = scenarioStep.Uses & " "
						end if
						'add the name of the class
						scenarioStep.Uses = scenarioStep.Uses & "LDM-" & classToMatch.Name
					end if
					'create the dependency between the use case and the Logical Data Model class
					linkElementsWithAutomaticTrace usecase, classToMatch
					Session.Output "adding link between " & usecase.Name & " and Logical Data Model class " & classToMatch.Name & " because of step " & scenario.Name & "." & scenarioStep.Name
				next
				'save scenario step
				scenarioStep.Update
				scenario.Update
			next
		next
	next
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

function linkDomainClassesWithBusinessRulesInPackage(dictionary,regExp,selectedPackage)
	'get a list of all business rules in the selected package
	dim packageList 
	set packageList = getPackageTree(selectedPackage)
	dim packageIDString
	packageIDString = makePackageIDString(packageList)
	dim getElementsSQL
	getElementsSQL = "select r.Object_ID from t_object r where r.stereotype = 'Business Rule' and r.Package_ID in (" & packageIDString & ")"
	dim businessRules
	set businessRules = getElementsFromQuery(getElementsSQL)
	linkDomainClassesWithBusinessRules dictionary,regExp, businessRules
end function

function linkDomainClassesWithBusinessRules(dictionary,regExp, businessRules)
	Session.Output businessRules.Count &" business rules found"
	dim businessRule as EA.Element
	dim connector as EA.Connector
	dim i
	for each businessRule in BusinessRules
		Repository.WriteOutput "Link to Logical Data Model", "Processing Business Rule: " & businessRule.Name, 0
		'first remove all automatic trace elements
		removeAllAutomaticTraces(businessRule)
		dim ruleText
		'get full text (notes + linked document)
		ruleText = businessRule.Name
		ruleText = ruleText & vbNewLine & Repository.GetFormatFromField("TXT",businessRule.Notes)
		ruleText = ruleText & vbNewLine & getLinkedDocumentContent(businessRule, "TXT")
		dim matches
		set matches = regExp.Execute(ruleText)
		'for each match create a <<trace link>> with business rule
		linkMatchesWithBusinessRule matches, businessRule, dictionary
	next
	
end function

function removeAllAutomaticTraces(element)
		dim i
		dim connector as EA.Connector
		'remove all the traces to Logical Data Model classes
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

function linkMatchesWithBusinessRule(matches, businessRule, dictionary)
	dim classesToMatch
	'get the classes to match
	Set classesToMatch = getClassesToMatchDictionary(matches,dictionary)
	dim classToMatch as EA.Element
	'actually link the classes
	for each classToMatch in classesToMatch.Items
		linkElementsWithAutomaticTrace businessRule, classToMatch
	next
end function

function linkElementsWithAutomaticTrace(sourceElement, TargetElement)
	dim trace as EA.Connector
	set trace = sourceElement.Connectors.AddNew("","trace")
	trace.Alias = "automatic"
	trace.SupplierID = TargetElement.ElementID
	trace.Update
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
			Repository.WriteOutput "Link to Logical Data Model", "Adding class: " & classElement.Name, 0
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



main
