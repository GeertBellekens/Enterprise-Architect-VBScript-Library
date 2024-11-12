'[path=\Projects\Project A\Project Browser Package Group]
'[group=Project Browser Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Clean up Demo model
' Author: Geert Bellekens
' Purpose: Does some housekeeping to store elements in their correct package
' Date: 2023-01-06
'
dim stereotypeMapping
set stereotypeMapping = CreateObject("Scripting.Dictionary")

dim actionRulemapping
set actionRulemapping = CreateObject("Scripting.Dictionary")
'-------------CONIGURATION------------------'
stereotypeMapping.Add "Action Rule", "AM;ARS"
stereotypeMapping.Add "Existence Law", "AM;EL"
stereotypeMapping.Add "Actorrole:True", "CM;CTAR"
stereotypeMapping.Add "TransactionKind:True", "CM;MTK"
stereotypeMapping.Add "custom table", "CM;Tabellen"
stereotypeMapping.Add "Transactorrole", "CM;TAR"
stereotypeMapping.Add "Derived Fact Specification", "FM;DFS"
stereotypeMapping.Add "Entity Type", "FM;ET"
stereotypeMapping.Add "Product Kind", "FM;PK"
stereotypeMapping.Add "Actorrole:False", "PM;AR"
stereotypeMapping.Add "TransactionKind:False", "PM;TK"

actionRulemapping.Add "revoke declare", "Initiator"
actionRulemapping.Add "revoke promise", "Initiator"
actionRulemapping.Add "revoke request", "Executor"
actionRulemapping.Add "revoke accept", "Executor"
actionRulemapping.Add "declare", "Initiator"
actionRulemapping.Add "request", "Executor"
actionRulemapping.Add "promise", "Executor"

'-------------CONIGURATION------------------'



const outPutName = "Clean up"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get the selected package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	if package is nothing then
		exit sub
	end if
	'let the user know we started
	Repository.WriteOutput outPutName, now() & " Starting clean up for package '" & package.Name & "'", 0
	Repository.WriteOutput outPutName, now() & " Getting package Dictonary", 0
	'create the package dictionary
	dim packageDictionary
	set packageDictionary = CreateObject("Scripting.Dictionary")
	'do the actual work
	cleanup package,  package, packageDictionary 
	'refresh
	Repository.ReloadPackage package.PackageID
	'let the user know we finished
	Repository.WriteOutput outPutName, now() & " Finished clean up for package '" & package.Name & "'", 0
end sub


function cleanup(package, rootPackage,packageDictionary )
	Repository.WriteOutput outPutName, now() & " Processing package '"& package.Name & "'", 0
	dim continue
	continue = true
	dim element as EA.Element
	for each element in package.Elements
		if continue then
			continue = cleanupElement(element, rootPackage, packageDictionary)
		end if
	next
	'process subpackages
	dim subPackage as EA.Package
	for each subPackage in package.Packages
		if continue then
			continue = cleanup(subPackage,  rootPackage, packageDictionary)
		end if
	next
	'return
	cleanup = continue
end function

function cleanupElement(element, rootPackage, packageDictionary)
	cleanupElement = true
	'determine the key
	dim key 
	key = getStereotypeKey(element)
	'check if stereotype is known
	if stereotypeMapping.Exists(key) then
		'remove composite diagrams for (multiple)TransactionKind
		if key = "TransactionKind:True" then
			deleteEmptySubDiagram element
		end if
		dim packageNames
		packageNames = stereotypeMapping(key)
		'for Action Rules, we create a subpackage based on the alias
		if key = "Action Rule" then
			packageNames = packageNames & determineActionRulePackage(element)
		end if
		'check if we know this package
		dim targetPackage as EA.Package

		set targetPackage = getPackageFromPath(rootPackage, packageNames, packageDictionary)
		if not targetPackage is nothing then
			if not element.PackageID = targetPackage.PackageID _
			  or not element.ParentID = 0 then
				'move the element to the correct package
				Repository.WriteOutput outPutName, now() & " Moving element '" & element.Name & "' to '" & targetPackage.Name &"'", 0
				element.PackageID = targetPackage.PackageID
				element.ParentID = 0
				element.Update
			end if
		else
			'report error
			Repository.WriteOutput outPutName, now() & " ERROR: could not find package with name '" & packageName &"'", 0
			dim response
			response = msgbox("ERROR: could not find package with name '" & packageName &"'" & vbNewLine & "Quit script?", vbYesNo+vbCritical, "Target package not found")
			if response = vbYes then
				cleanupElement = false
				exit function
			end if
		end if
	end if
	'process subelements
	dim subElement as EA.Element
	dim continue
	continue = true
	for each subElement in element.Elements
		if continue then
			continue = cleanupElement(subelement, rootPackage, packageDictionary)
		end if
	next
	'return
	cleanupElement = continue
end function

function determineActionRulePackage(element)
	dim packagePath
	'determine main path
	dim elementAlias
	elementAlias = element.Alias
	'if there is a dot in the alias, we take everything left of the dot
	dim dotPosition
	dotPosition = instr(elementAlias, ".")
	if dotPosition > 0 then
		elementAlias = left(elementAlias, dotposition - 1)
	end if
	if len(elementAlias) > 0 then
		packagePath = ";TK" & elementAlias
	end if
	'determine initiator or executor
	dim lcaseName
	lcaseName = lcase(element.name)
	dim nameKey
	for each nameKey in actionRulemapping
		if instr(lcaseName, namekey) > 0 then
			packagePath = packagePath & ";" & actionRulemapping(nameKey)
			exit for
		end if
	next
	'return
	determineActionRulePackage = packagePath
end function

function deleteEmptySubDiagram(element)
	dim subDiagram as EA.Diagram
	for each subDiagram in element.Diagrams
		if subDiagram.DiagramObjects.Count = 0 then
			element.Diagrams.DeleteAt 0,True
		end if
		'we only delete a single diagram, so exit after first one
		exit for
	next
end function

function getStereotypeKey(element)
	if element.Stereotype = "Actorrole" or _
      element.Stereotype = "TransactionKind" then
		dim addition
		addition = ":False"
		dim taggedValue as EA.TaggedValue
		for each taggedValue in element.TaggedValues
			if taggedValue.Name = "multiple" and lcase(taggedValue.Value) = "true" then
				addition = ":True"
			end if
		next
		getStereotypeKey = element.Stereotype & addition
	else
		getStereotypeKey = element.Stereotype
	end if
	
end function

function getPackageFromPath(startPackage, packagePath, packageDictionary)
    'Repository.WriteOutput outPutName, now() & " Finding package for path '" & packagePath &"'", 0
	'check if package exists in dictionary
	if not packageDictionary.Exists(packagePath) then
		dim package as EA.Package
		set package = getPackageFromPackagePath(startPackage, packagePath)
		'add to dictionary
		packageDictionary.Add packagePath, package
	end if
	'return package
	set getPackageFromPath = packageDictionary(packagePath)
end function


function getPackageFromPackagePath(startPackage, packagePath)
	'parse packagePath
	dim packageNames
	set packageNames = CreateObject("System.Collections.ArrayList")
	dim packageNamesArray
	packageNamesArray = split(packagePath, ";")
	dim packageName
	for each packageName in packageNamesArray
		packageNames.Add packageName
	next
	'start by the top level package
	set getPackageFromPackagePath = getPackageFromPackageNameList(startPackage,packageNames, true)
end function

function getPackageFromPackageNameList(startPackage,packageNames, lookUp)
	dim foundPackage as EA.Package
	set foundPackage = nothing 'initialize
	if packageNames.Count = 0 then
		set getPackageFromPackageNameList = startPackage
		exit function
	end if
	'reset packages collection
	startPackage.Packages.Refresh
	'get the first package
	dim firstPackageName 
	firstPackageName = packageNames(0)
	'first try to find it in the subPackages
	set foundPackage = findPackageDown(startPackage, firstPackageName)
	'if we haven't found it down, we go up
	if foundPackage is nothing and lookUp then
		'then try to find it up
		set foundPackage = findPackageUp(startPackage, firstPackageName)
	end if
	if foundPackage is nothing then
		'if still haven't found it, then we create the whole set of packages
		set foundPackage = createPackages(startPackage, packageNames)
	end if
	if packageNames.Count > 1 _
	  and not foundPackage is nothing then
		'pop first name
		packageNames.RemoveAt 0
		set foundPackage = getPackageFromPackageNameList(foundPackage,packageNames, false)
	end if
	'return
	set getPackageFromPackageNameList = foundPackage
end function

function createPackages(startPackage, packageNames)
	dim newPackage as EA.Package
	if packageNames.Count > 0 then
		'create package
		set newPackage = startPackage.Packages.AddNew(packageNames(0), "")
		newPackage.Update
		'pop name from list and continue
		if packageNames.Count > 1 then
			packageNames.RemoveAt 0
			set newPackage = createPackages(newPackage, packageNames)
		end if
	else
		'return the startPackage
		set newPackage = startPackage
	end if
	'return new Package
	set createPackages = newPackage 
end function

function findPackageUp(startPackage, packageName)
	dim foundPackage as EA.Package
	set foundPackage = nothing 'initialize
	'look at sibling packages
	set foundPackage = findPackageinSiblings(startPackage, packageName)
	if foundPackage is nothing _
	  and startPackage.parentID > 0 then
		dim parentPackage
		set parentPackage = Repository.GetPackageByID(startPackage.ParentID)
		'loot at sibling package of parent package
		set foundPackage = findPackageinSiblings(parentPackage, packageName)
	end if
	'return
	set findPackageUp = foundPackage
end function

function findPackageinSiblings(startPackage, packageName)
	dim foundPackage as EA.Package
	set foundPackage = nothing 'initialize
	dim parentPackage as EA.Package
	if startPackage.ParentID > 0 then
		set parentPackage = Repository.GetPackageByID(startPackage.ParentID)
		dim siblingPackage as EA.Package
		for each siblingPackage in parentPackage.Packages
			if siblingPackage.Name = packageName then
				set foundPackage = siblingPackage
			end if
		next
	end if
	'return
	set findPackageinSiblings = foundPackage
end function

function findPackageDown(startPackage, packageName)
	dim foundPackage as EA.Package
	set foundPackage = nothing 'initialize
	'check if the current package name corresponds
	if startPackage.Name = packageName then
		set foundPackage = startPackage
	else
		'not found, so we look down	
		dim subPackage
		for each subPackage in startPackage.Packages
			set foundPackage = findPackageDown(subPackage, packageName)
			if not foundPackage is nothing then
				'found it, exit the loop
				exit for
			end if
		next
	end if
	'return
	set findPackageDown = foundPackage
end function

main