'[path=\Projects\Project A\Search Group]
'[group=Search Group]

' Script Name: Validate Message Rules (in selected package)
' Author: Geert Bellekens
' Purpose: shows the Message Rules that are invalid in the search window
' Date: 2017-04-12

option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

const outputTabName = "Validate Message Rules"
dim allRootNodes
set allRootNodes = CreateObject("Scripting.Dictionary")

sub main
	Repository.CreateOutputTab outputTabName
	Repository.ClearOutput outputTabName
	Repository.EnsureOutputVisible outputTabName
	Repository.WriteOutput outputTabName, now() & " : Starting Validate Message Rules" ,0
	'get the current package ID string
	dim currentPackageIDString
	currentPackageIDString = getCurrentPackageTreeIDString()
	'get all Message Test Rules in the selected package
	dim testRules
	set testRules = getAllTestRules(currentPackageIDString)
	Repository.WriteOutput outputTabName, now() & " : Number Message Test Rules Found: " & testRules.Count ,0
	'Validate each test rule
	dim invalidTestRules
	set invalidTestRules = CreateObject("Scripting.Dictionary")
	dim testRule
	for each testRule in testRules
		dim validationError
		validationError = validateTestRule(testRule)
		if validationError <> "valid" then
			invalidTestRules.Add testRule, validationError
		end if
	next
	'get the required output
	dim outputRows
	set outputRows = getOutputRows(invalidTestRules)
	'show the output
	showOutput outputRows
	Repository.WriteOutput outputTabName, now() & " : Finished Validate Message Rules" ,0
end sub 

function getAllTestRules(currentPackageIDString)
	set getAllTestRules = CreateObject("System.Collections.ArrayList")
	dim SQLGetAllTestRules
	SQLGetAllTestRules = "select o.Object_ID from t_object o " & _
						" where o.Stereotype = 'Message Test Rule' " & _
						" and o.Package_ID IN (" & currentPackageIDString & ")" & _
						" order by o.Name"
	dim testRuleElements
	set testRuleElements = getElementsFromQuery(SQLGetAllTestRules)
	dim testRuleElement
	for each testRuleElement in testRuleElements
		dim messageTestRule
		set messageTestRule = new MessageValidationRule
		messageTestRule.initialiseWithTestElement(testRuleElement)
		getAllTestRules.Add messageTestRule
	next
end function

function validateTestRule(testRule)
	'tell the user what we are doing
	Repository.WriteOutput outputTabName, now() & " : Validating Rule: " & testRule.RuleID , testRule.TestElement.ElementID
	dim rootNode
	dim relatedRootNodes
	set relatedRootNodes = getRelatedRootNodes(testRule)
	'if there are no related root nodes then there is a problem as well
	if relatedRootNodes.Count = 0 then
		validateTestRule = "Rule is not linked to any message"
	end if
	for each rootNode in relatedRootNodes
		dim ruleLinked
		ruleLinked = rootNode.linkRuletoNode(testRule, testRule.Path)
		if ruleLinked then 
			validateTestRule = "valid"
		else
			dim messagePackage as EA.Package
			set messagePackage = Repository.GetPackageByID(rootNode.SourceElement.PackageID)
			validateTestRule = "Invalid path '" & Join(testRule.Path.ToArray(),".") & "' not found in message :" & messagePackage.Name
		end if
	next
end function

function getRelatedRootNodes(testRule)
	dim relatedRootNodes
	set relatedRootNodes = CreateObject("System.Collections.ArrayList")
	dim sqlGetRelatedRootNodeElements
	sqlGetRelatedRootNodeElements = "select o.Object_ID from t_object o                         " & _
									" inner join t_connector c on c.End_Object_ID = o.Object_ID " & _
									" where o.Stereotype = 'XSDtopLevelElement'                 " & _
									" and c.Start_Object_ID = " & testRule.TestElement.ElementID         
	dim relatedRootNodeElements 
	set relatedRootNodeElements = getElementsFromQuery(sqlGetRelatedRootNodeElements)
	dim currentRootNodeElement as EA.Element
	for each currentRootNodeElement in relatedRootNodeElements
		'first check if the rootnode is already present in the list of rootnodes
		if allRootNodes.Exists(currentRootNodeElement.ElementID) then
			relatedRootNodes.Add allRootNodes.Item(currentRootNodeElement.ElementID)
		else
			dim currentRootNode
			set currentRootNode = new MessageNode
			currentRootNode.intitializeWithSource currentRootNodeElement,nothing,"1..1",nothing,nothing
			relatedRootNodes.Add currentRootNode
			'also add to the list of all root nodes
			allRootNodes.add currentRootNodeElement.ElementID, currentRootNode
		end if
	next
	'return the rootnodes
	set getRelatedRootNodes = relatedRootNodes
end function

function getOutputRows(invalidTestRules)
	dim currentTestRule
	dim outputRows
	set outputRows = CreateObject("System.Collections.ArrayList")
	for each currentTestRule in invalidTestRules.Keys
		Repository.WriteOutput outputTabName, now() & " Creating output for Test Rule: " & currentTestRule.RuleId , currentTestRule.TestElement.ElementID
		dim currentRow
		dim validation 
		validation = invalidTestRules.Item(currentTestRule)
		set currentRow = getOutputRow(currentTestRule, validation)
		outputRows.add currentRow
	next
	'return outputrows
	set getOutputRows = outputRows
end function



function getOutputRow(testRule,validation)
	'create outputrow
	dim outputRow
	set outputRow = CreateObject("System.Collections.ArrayList")
	'add data
	outputRow.Add testRule.TestElement.ElementGUID
	outputRow.Add "Test"
	outputRow.Add testRule.RuleId
	outputRow.Add validation
	'get packages
	'get package0
	dim package0 as EA.Package
	set package0 = Repository.GetPackageByID(testRule.TestElement.PackageID)
	'get package1
	dim package1 as EA.Package
	if package0.ParentID > 0 then
		set package1 = Repository.GetPackageByID(package0.ParentID)
	end if
	'get package2
	dim package2 as EA.Package
	if not package1 is nothing and package1.ParentID > 0 then
		set package2 = Repository.GetPackageByID(package1.ParentID)
	else
		set package2 = nothing
	end if
	'get package2
	dim package3 as EA.Package
	if not package2 is nothing and package2.ParentID > 0 then
		set package3 = Repository.GetPackageByID(package2.ParentID)
	else
		set package3 = nothing
	end if
	if not package0 is nothing then
		outputRow.Add package0.Name
	else
		outputRow.Add ""
	end if
	if not package1 is nothing then
		outputRow.Add package1.Name
	else
		outputRow.Add ""
	end if
	if not package2 is nothing then
		outputRow.Add package2.Name
	else
		outputRow.Add ""
	end if
	if not package3 is nothing then
		outputRow.Add package3.Name
	else
		outputRow.Add ""
	end if
	'return row
	set getOutputRow = outputRow
end function

function showOutput(outputRows)
	'get the headers for the output
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.Add "CLASSGUID"
	headers.Add "CLASSTYPE"
	headers.Add "Message Test Rule ID"
	headers.Add "Error"
	headers.Add "Package_level1 "
	headers.Add "Package_level2"
	headers.Add "Package_level3"
	headers.Add "Package_level4"
	'create the output object
	dim searchOutput
	set searchOutput = new SearchResults
	searchOutput.Name = "Validate Message Rules"
	searchOutput.Fields = headers
	'put the contents in the output
	dim row
	for each row in outputRows
		'add row the the output
		searchOutput.Results.Add row
	next
	'show the output
	searchOutput.Show
end function






main