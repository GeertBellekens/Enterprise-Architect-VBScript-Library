'[path=\Projects\Project A\Project Browser Group]
'[group=Project Browser Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.Util

'
' Script Name: Validate BR uses
' Author: Geert Bellekens
' Purpose: Lists all Business rules that are linked to a use case but not referenced in the uses column.
' Date: 23/12/2015
'
'name of the output tab
const outPutName = "Check BR links"

sub main

	
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'set timestamp
	Repository.WriteOutput outPutName, "Start check BR links at " & now(), 0
	

	
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
			CheckBRLinksForUseCases usecases
		case otPackage
			' Code for when a package is selected
			dim selectedPackage as EA.Package
			set selectedpackage = Repository.GetTreeSelectedObject()
			'link use domain classes with use cases under the selected package
			CheckBRLinksFromPackage selectedPackage	
		case else
			' Error message
			Repository.WriteOutput outPutName, "Error: wrong type selected. You need to select a package or one or more elements", 0
			
	end select
	Repository.WriteOutput outPutName, "Finished checking BR links " & now(), 0
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

function CheckBRLinksFromPackage(selectedPackage)
	dim packageList 
	set packageList = getPackageTree(selectedPackage)
	dim packageIDString
	packageIDString = makePackageIDString(packageList)
	dim getElementsSQL
	getElementsSQL = "select uc.Object_ID from t_object uc where uc.Object_Type = 'UseCase' and uc.Package_ID in (" & packageIDString & ")"
	dim usecases
	set usecases = getElementsFromQuery(getElementsSQL)
	CheckBRLinksForUseCases usecases
end function

function CheckBRLinksForUseCases(usecases)
	Session.Output usecases.Count & " use cases found"
	dim usecase as EA.Element
	dim businessRule as EA.Element
	'loop de use cases
	for each usecase in usecases
		'Repository.WriteOutput outPutName, "Processing use case: " & usecase.Name, 0
		'first remove all automatic traces
		removeAllAutomaticTraces usecase
		'get the scenarios xml
		dim xmlScenario
		set xmlScenario = getScenariosXML(usecase)
		'Session.output xmlScenario.xml
		'get all dependencies left
		dim dependencies
		set dependencies = getDependencies(usecase)
		for each businessRule in dependencies.Items
			dim businessRuleFound
			businessRuleFound = false
			if businessRule.Stereotype = "Business Rule" then
				'OK we have a business rule
				'first check if <step> exists with the attribute "uses="<usecasename>
				Dim stepNodes, itemnodes, itemnode, attributeNode 
				Set stepNodes = xmlScenario.SelectNodes("//step[contains(@uses," & sanitizeXPathSearch(businessRule.Name) & ")]")
				if stepNodes.length > 0 then
					businessRuleFound = true
				else
					'if not found we look for the node that has the guid of the business rule
					'<item> nodes have an attribute guid= and an attribute oldname=
					set itemnodes = xmlScenario.SelectNodes("//item[@guid='" & businessRule.elementGUID & "']")
					for each itemnode in itemnodes
						for each attributeNode in itemNode.Attributes
							if attributeNode.Name = "oldname" then
								Set stepNodes = xmlScenario.SelectNodes("//step[contains(@uses," & sanitizeXPathSearch(attributeNode.Value) & ")]")
								if stepNodes.length > 0 then
									businessRuleFound = true
									exit for
								end if
							end if
						next
					next
				end if
				'check if businessrule was found
				if not businessRuleFound then
					Repository.WriteOutput outPutName, "Use Case: [" & usecase.Name & "] BR not used: [" & businessRule.Name & "]", usecase.ElementID
				end if
			end if
		next
	next
end function

function sanitizeXPathSearch(searchValue)
	dim searchParts, part, returnValue, first	
	first = true
	returnValue = searchValue
	'first replace double qoutes by &amp;quot;
	returnValue = replace(returnValue,"""","&amp;quot;")
	'then replace any single qotes. "firstpart'secondpart' then becomes "concat('firstpart', "'",'secondpart',...)"
	searchParts = split(returnValue,"'")
	if Ubound(searchParts) > 0 then
		returnValue = "concat("
		for each part in searchParts
			if first then
				first = false
			else
				returnValue = returnValue & ",""'"","
			end if
			returnValue = returnValue & """" & part & """"
		next
		returnValue = returnValue & ")"
		Session.Output "sanitized xpath: " & returnValue
	else
		'enclose in single quotes
		returnValue = "'" & returnValue & "'"
	end if
	sanitizeXPathSearch = returnValue
end function

function getScenariosXML(usecase)
		dim sqlGet, xmlQueryResult
		sqlGet = "select ucs.XMLContent from t_objectscenarios ucs where ucs.Object_ID = " & usecase.ElementID
		xmlQueryResult = Repository.SQLQuery(sqlGet)
		
		xmlQueryResult = replace(xmlQueryResult,"&lt;","<")
		xmlQueryResult = replace(xmlQueryResult,"&gt;",">")
		'Repository.WriteOutput outPutName, "xmlQueryResult: " & xmlQueryResult  , 0
		Dim xDoc 
		Set xDoc = CreateObject( "MSXML2.DOMDocument.4.0" )
		'load the resultset in the xml document
		xDoc.LoadXML xmlQueryResult
		'return value
		set getScenariosXML = xDoc
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
		'Repository.WriteOutput outPutName, "Processing Business Rule: " & businessRule.Name, 0
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
		'this works for FISSES as well because they are classes with stereotype Message
		if classElement.Type = "Class" AND len(classElement.Name) > 0 AND not dictionary.Exists(classElement.Name) then
			Repository.WriteOutput outPutName, "Adding element: " & classElement.Name, 0
			dictionary.Add classElement.Name,  classElement
		end if	
	next
	'process subpackages
	for each subpackage in package.Packages
		addClassesToDictionary subpackage, dictionary
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