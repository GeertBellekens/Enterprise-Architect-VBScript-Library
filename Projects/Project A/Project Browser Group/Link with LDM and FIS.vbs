'[path=\Projects\Project A\Project Browser Group]
'[group=Project Browser Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Link with Logical Data Model
' Author: Geert Bellekens
' Purpose: Links uses cases and rules to elements Logical Data Model entities they use in their scenariosteps for the use case and in the notes/linked documents for the Business Rules
' Date: 28/09/2015
'
'name of the output tab
const outPutName = "Link to LDM and FIS"
const reportingRootGUID = "{DDED9219-D9BB-4a8f-80A2-CDA8405C6FE1}"

sub main
	
	'figure out if we are in the reporting model
	dim selectedPackage
	set selectedPackage = Repository.GetTreeSelectedPackage
	dim reporting 
	reporting = isReporting(selectedPackage)
	
	'reference to the domain model package
	dim domainModelPackageGUID 
	dim FISPackageGUID
	if reporting then
		'the parent package of the LDM DW
		domainModelPackageGUID = "{B460AE6A-19A2-42ed-BD43-0221BA94B3B9}"
	else
		domainModelPackageGUID = "{8A528D7F-D23B-4a85-B89A-15F5B41CE384}"
		FISPackageGUID = "{A4B198D1-FF9B-4375-8B9B-0015096DE9AD}"
	end if
		
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'set timestamp
	Repository.WriteOutput outPutName, "Starting link to LDM and FIS at " & now(), 0
	
	'first get the pattern from all the classes in the Logical Data Model
	dim dictionary
	Set dictionary = CreateObject("Scripting.Dictionary")
	'Logical Data Model
	Repository.WriteOutput outPutName, "Creating dictionary from Logical Data Model", 0
	addToClassDictionary domainModelPackageGUID, dictionary
	'extended Logical Data Model
	'addToClassDictionary extendedDomainModelPackageGUID, dictionary
	
	'FISSES
	if not reporting then
		addToClassDictionary FISPackageGUID, dictionary
	end if
	
	' and prepare the regex object
	dim pattern
	'create the pattern based on the names in the dictionary
	Repository.WriteOutput outPutName, "Creating regular expression", 0
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
			linkDomainClassesWithBusinessRules dictionary,regExp, businessRules
			'process reporting elements
			dim reportingElements
			set reportingElements = getReportingElementsEACollection(selectedElements)
			linkDomainClassesWithBusinessRules dictionary,regExp, reportingElements
		case otPackage
			' Code for when a package is selected
			dim packageIDString 
			packageIDString = getPackageTreeIDString(selectedPackage)
			'link use domain classes with use cases under the selected package
			linkDomainClassesWithUseCasesInPackage dictionary, regExp, packageIDString
			'link the domain classes with the business rules under the selected package
			linkDomainClassesWithBusinessRulesInPackage dictionary, regExp, packageIDString	
			'link the reporting elemnets in the selected package
			linkDomainClassesWithReportingElementsInPackage dictionary, regExp, packageIDString	
		case else
			' Error message
			Repository.WriteOutput outPutName, "Error: wrong type selected. You need to select a package or one or more elements", 0
			
	end select
	Repository.WriteOutput outPutName, "Finished link to LDM and FIS at " & now(), 0
end sub

function isReporting(selectedPackage)
	'we check if the selected packag is part of the reporting model by going up in package until we reach the top level reporting package.
	if selectedPackage.PackageGUID = reportingRootGUID then
		isReporting = true
	else
		if selectedPackage.ParentID > 0 then
			'get the parent package
			dim parentPackage
			set parentPackage = Repository.GetPackageByID(selectedPackage.ParentID)
			'go up
			isReporting = isReporting(parentPackage)
		else
			isReporting = false
		end if
	end if
end function

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

function getReportingElementsEACollection(selectedElements)
	dim reportingElements 
	set reportingElements = CreateObject("System.Collections.ArrayList")
	dim element as EA.Element
	for each element in selectedElements
		if element.Stereotype = "BI-REP" or element.Stereotype = "BI-Dataset" then
			reportingElements.Add element
		end if
	next
	set getReportingElementsEACollection = reportingElements
end function

function linkDomainClassesWithUseCasesInPackage(dictionary,regExp,packageIDString)
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
		Repository.WriteOutput outPutName, "Processing use case: " & usecase.Name, 0
		'get all dependencies left
		dim dependencies
		set dependencies = getDependencies(usecase)
		
		'first remove all automatic traces
		removeAllAutomaticTraces usecase
		
		dim useCaseText
		'add the notes the text to verify
		useCaseText = usecase.Notes & vbNewLine
		'Loop the constraints
		dim constraint as EA.Constraint
		for each constraint in usecase.Constraints
			'add text (name and notes) of the constraint to the useCaseText
			useCaseText = useCaseText & constraint.Name & vbNewLine & constraint.Notes & vbNewLine
		next
		'loop scenarios
		dim scenario as EA.Scenario
		for each scenario in usecase.Scenarios
			dim scenarioStep as EA.ScenarioStep
			for each scenarioStep in scenario.Steps
				'add the scenario text to the use case text				
				useCaseText = useCaseText & scenarioStep.Name & vbNewline
			next
		next
		'check for matches
		dim matches
		set matches = regExp.Execute(useCaseText)
		dim classesToMatch 
		set classesToMatch = getClassesToMatchDictionary(matches, dictionary)
		dim classIDToMatch
		for each classIDToMatch in classesToMatch.Items
			'create the dependency between the use case and the Logical Data Model class
			linkElementsWithAutomaticTrace usecase, classIDToMatch
		next
	next
end function

function addUsesToScenarioStep (classToMatch, scenarioStep, prefix)
	if not instr(scenarioStep.Uses,prefix & "-" & classToMatch.Name) > 0 then
		if len(scenarioStep.Uses) > 0 then 
			'add a space if needed
			scenarioStep.Uses = scenarioStep.Uses & " "
		end if
		'add the name of the class
		scenarioStep.Uses = scenarioStep.Uses & prefix & "-"  & classToMatch.Name
	end if
end function

function removeAddionalUses( dependencies, uses, dictionary)
	dim refName
	dim dependency
	if Instr(uses,"LDM-") > 0 _
		or Instr(uses, "FIS-") > 0 then
			'first loop all dependencies
			for each dependency in dependencies
				'remove LDM-<name>
				uses = replace(uses,"LDM-" & dependency,"")
				'remove FIS-<name>
				uses = replace(uses,"FIS-" & dependency,"")
			next
			for each refName in dictionary.Keys
				'remove LDM-<name>
				uses = replace(uses,"LDM-" & refName,"")
				'remove FIS-<name>
				uses = replace(uses,"FIS-" & refName,"")
			next
	end if
	removeAddionalUses = uses
end function


'returns a dictionary of elements with the name as key and the element as value.
function getDependencies(element)
	dim getDependencySQL
	getDependencySQL =  "select dep.Object_ID from ( t_object dep " & _
						" inner join t_connector con on con.End_Object_ID = dep.Object_ID)   " & _
						" where con.Connector_Type in ('Dependency','Abstraction')  " & _
						" and con.Start_Object_ID = " & element.ElementID   
	set getDependencies = getElementDictionaryFromQuery(getDependencySQL)
end function

function linkDomainClassesWithBusinessRulesInPackage(dictionary,regExp,packageIDString)
	'get a list of all business rules in the selected package
	dim getElementsSQL
	getElementsSQL = "select r.Object_ID from t_object r where r.stereotype = 'Business Rule' and r.Package_ID in (" & packageIDString & ")"
	dim businessRules
	set businessRules = getElementsFromQuery(getElementsSQL)
	linkDomainClassesWithBusinessRules dictionary,regExp, businessRules
end function

function linkDomainClassesWithReportingElementsInPackage(dictionary,regExp,packageIDString)
	'get a list of all reportingElements in the selected package
	dim getElementsSQL
	getElementsSQL = "select r.Object_ID from t_object r where r.stereotype in ('BI-REP', 'BI-Dataset') and r.Package_ID in (" & packageIDString & ")"
	dim reportingElements
	set reportingElements = getElementsFromQuery(getElementsSQL)
	linkDomainClassesWithBusinessRules dictionary,regExp, reportingElements
end function


function linkDomainClassesWithBusinessRules(dictionary,regExp, businessRules)
	Session.Output businessRules.Count &" business rules found"
	dim businessRule as EA.Element
	dim connector as EA.Connector
	dim i
	for each businessRule in BusinessRules
		Repository.WriteOutput outPutName, "Processing Element: " & businessRule.Name, 0
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
	dim classIDToMatch
	'actually link the classes
	for each classIDToMatch in classesToMatch.Items
		linkElementsWithAutomaticTrace businessRule, classIDToMatch
	next
end function

function linkElementsWithAutomaticTrace(sourceElement, targetID)
	dim trace as EA.Connector
	set trace = sourceElement.Connectors.AddNew("","trace")
	trace.Alias = "automatic"
	trace.SupplierID = targetID
	trace.Update
end function

function addToClassDictionary(PackageGUID, dictionary)
	dim package as EA.Package
	set package = Repository.GetPackageByGuid(PackageGUID)
	
	'get the classes in the dictionary (recursively
	addClassesToDictionary package, dictionary
end function

function addClassesToDictionary(package, dictionary)
	dim classRow 
	'search classes using SQL 
	dim packageIDString
	packageIDString = getPackageTreeIDString(package)
	dim sqlGetClasses
	sqlGetClasses = "select o.Name, o.Object_ID from t_object o where o.Object_Type = 'Class' and o.name is not null and o.Package_ID in (" & packageIDString &")"
	dim classNames
	set classNames = getArrayListFromQuery(sqlGetClasses)
	'process owned elements
	for each classRow in classNames
		dim className
		className = classRow(0)
		dim classID
		classID = classRow(1)
		'add to dictionary
		'this works for FISSES as well because they are classes with stereotype Message
		dictionary(className) = classID
		Repository.WriteOutput outPutName, "Adding element: " & className, 0	
	next
end function


'Create a reges pattern like this "\b(name1|name2|name3)s?\b" based on the 
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