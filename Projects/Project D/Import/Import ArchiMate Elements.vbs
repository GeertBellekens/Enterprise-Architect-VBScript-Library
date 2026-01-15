'[path=\Projects\Project D\Import]
'[group=Import]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Import ArchiMate Elements from Excel
' Author: Geert Bellekens
' Purpose: import the archimate elemnets from the excel in EA
' Date: 2025-11-12
'

'name of the output tab
const outPutName = "Import ArchiMate Elements"

sub main	
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage
	if not package is nothing then
		'tell the user we are finished
		Repository.WriteOutput outPutName, now() & " Starting " &  outPutName ,0
		'import into the selected pacakge
		importArchiMateFromExcel package 
		'add composite diagrams
		addCompositeDiagrams package 
		'tell the user we are finished
		Repository.WriteOutput outPutName, now() & " Finished " & outPutName , 0
	end if
end sub

function addCompositeDiagrams(package)
	'add diagram for current package
	addCompositeDiagram package
	'add diagrams for subPackages
	dim subPackage
	for each subPackage in package.Packages
		addCompositeDiagrams subPackage
	next
end function

function addCompositeDiagram(package)
	Repository.WriteOutput outPutName, now() & " Adding diagram for package '" & package.Name & "'",0
	'check if diagram exists with same name
	dim diagram
	set diagram = getMainDiagram(package)
	'add elements
	'get main element
	dim mainElement
	set mainElement = getMainElement(package)
	'get connected elements
	if not mainElement is nothing then
		dim connectedElements
		set connectedElements = getConnectedElements(mainElement)
		connectedElements.Add mainElement
		'add all elements to the diagram
		addElementsToDiagram connectedElements, diagram
		'set composite diagram
		setCompositeDiagram mainElement, diagram
	end if
	'layout the diagram
	dim XMLdiagramID
	XMLdiagramID = Repository.GetProjectInterface().GUIDtoXML(diagram.DiagramGUID)
	Repository.GetProjectInterface().LayoutDiagramEx XMLdiagramID,lsDiagramDefault, 4, 20, 20, false
end function



function getConnectedElements(mainElement)
	dim sqlGetData
	sqlGetData = "select c.End_Object_ID as Object_ID from t_connector c " & vbNewLine & _
				" where c.Stereotype = 'ArchiMate_Composition'         " & vbNewLine & _
				" and c.Start_Object_ID = " & mainElement.ElementID &" " & vbNewLine & _
				" union                                                " & vbNewLine & _
				" select c.Start_Object_ID from t_connector c          " & vbNewLine & _
				" where c.Stereotype = 'ArchiMate_Composition'         " & vbNewLine & _
				" and c.End_Object_ID = " & mainElement.ElementID
	dim elements
	set elements = getElementsFromQuery(sqlGetData)
	'return
	set getConnectedElements = elements
end function


function addElementsToDiagram(elements, diagram)
	dim element as EA.Element
	for each element in elements
		dim diagramObject as EA.DiagramObject
		set diagramObject = diagram.DiagramObjects.AddNew("", "")
		diagramObject.ElementID = element.ElementID
		diagramObject.Update
	next
end function

function getMainElement(package)
	dim element as EA.Element
	set element = nothing
	dim currentElement
	for each currentElement in package.Elements
		if currentElement.Name = package.Name then
			set element = currentElement
		end if
	next
	'return
	set getMainElement = element
end function

function getMainDiagram(package)
	dim diagram as EA.Diagram
	set diagram = nothing
	dim currentDiagram
	for each currentDiagram in package.Diagrams
		if currentDiagram.Name = package.Name then
			set diagram = currentDiagram
			exit for
		end if
	next
	if diagram is nothing then
		set diagram = package.Diagrams.AddNew(package.Name, "ArchiMate3::Business")
		diagram.Update
	end if
	'return
	set getMainDiagram = diagram
end function

function importArchiMateFromExcel(package)
	dim sourceExcelFile
	set sourceExcelFile = new ExcelFile
	sourceExcelFile.openUserSelectedFile()
	if len(sourceExcelFile.FileName)> 0 then
		dim elementDictionary
		set elementDictionary = CreateObject("Scripting.Dictionary")
		dim parentDictionary
		set parentDictionary = CreateObject("Scripting.Dictionary")
		'get the contents of the sheet
		dim sheet
		for each sheet in sourceExcelFile.worksheets
			dim contents
			contents = sourceExcelFile.getContents(sheet)
			dim i
			dim j
			for i = 2 to Ubound(contents,1)
				dim stereotype
				dim elementID
				dim name
				dim notes
				dim parent
				dim externalID
				dim source
				stereotype = contents(i, 1)
				elementID = contents(i, 2)
				name = contents(i, 4)
				notes = contents(i, 5)
				parent = contents(i, 6)
				externalID = contents(i, 8)
				source = contents(i, 9)
				'inform user of progress
				Repository.WriteOutput outPutName, now() & " Adding element " & i -1  & " of " & Ubound(contents,1) -1 & "  '" & name & "' " , 0
				dim element
				set element = createArchiMateElement(package, stereotype, elementID, name, notes, externalID, source)
				'add element to dictionary
				if not elementDictionary.Exists(elementID) then
					elementDictionary.Add elementID, element
					parentDictionary.Add elementID, parent
				end if
			next
			'process parent relations
			processParentRelations elementDictionary, parentDictionary
		next
	end if
end function

function processParentRelations(elementDictionary, parentDictionary)
	dim element as EA.Element
	dim key
	dim packageDictionary
	set packageDictionary = CreateObject("Scripting.Dictionary")
	for each key in elementDictionary.Keys
		set element = elementDictionary(key)
		dim parentKey
		parentKey = parentDictionary(key)
		if elementDictionary.Exists(parentKey) then
			dim parentElement as EA.Element
			set parentElement = elementDictionary(parentKey)
			Repository.WriteOutput outPutName, now() & " Adding composition between '" & parentElement.name & "' and '" & element.Name & "'"   , 0
			dim parentPackage as EA.Package
			set parentPackage = getParentPackage(parentElement, packageDictionary)
			dim elementPackage as EA.Package
			if packageDictionary.Exists(element.PackageID) then
				set elementPackage = packageDictionary(element.PackageID)
			else
				set elementPackage = Repository.GetPackageByID(element.PackageID)
				packageDictionary.Add element.PackageID, elementPackage
			end if
			if elementPackage.Name = element.Name then
				'move element package
				elementPackage.ParentID = parentPackage.PackageID
				elementPackage.Update
			else
				'move element to package
				element.PackageID = parentPackage.PackageID
				element.Update
			end if
			'add composite relationship
			dim connector as EA.Connector
			set connector = parentElement.Connectors.AddNew("", "ArchiMate3::ArchiMate_Composition")
			connector.SupplierID = element.ElementID
			connector.Update
			connector.ClientEnd.Aggregation = 2
			connector.ClientEnd.Update
		else
			Repository.WriteOutput outPutName, now() & " WARNING: parent key '" & parentKey & "' not found"   , 0
		end if
	next
end function

function getParentPackage(parentElement, packageDictionary)
	dim package as EA.Package
	'get from dictionary if exists
	if packageDictionary.Exists(parentElement.PackageID) then
		set package = packageDictionary(parentElement.PackageID)
	else
		set package = Repository.GetPackageByID(parentElement.PackageID)
		packageDictionary.Add parentElement.PackageID, package
	end if
	if not package.name = parentElement.Name then
		dim parentPackage as EA.Package
		set parentPackage = package.Packages.AddNew(parentElement.Name, "")
		parentPackage.Update
		parentElement.PackageID = parentPackage.PackageID
		parentElement.Update
	else
		set parentPackage = package
	end if
	'return
	set getParentPackage = parentPackage
end function

function createArchiMateElement(package, stereotype, elementID, name, notes, externalID, source)
	dim element as EA.Element
	set element = package.Elements.AddNew(name, stereotype)
	element.Alias = elementID
	element.Notes = notes
	element.Update
	'add tagged values for externalID and source
	dim tag as EA.TaggedValue
	set tag = element.TaggedValues.AddNew("ExternalID", "")
	tag.Value = externalID
	tag.Update
	set tag = element.TaggedValues.AddNew("Source", "")
	tag.Value = source
	tag.Update
	'return
	set createArchiMateElement = element
end function




function createMessageRule(path, ruleName, ruleID, ruleReason, ownerPackage)
	Repository.WriteOutput outPutName, now() & " Adding Rule'" & ruleID & "'",0
	'create MessageRule
	dim messageRule as EA.Element
	set messageRule = ownerPackage.Elements.AddNew(ruleID,"Test")
	messageRule.StereotypeEx = "Message Test Rule"
	messageRule.Notes = ruleName
	messageRule.Update
	'add the tagged for the rule reason
	dim reasonTaggedValue as EA.TaggedValue
	set reasonTaggedValue = getExistingOrNewTaggedValue(messageRule, "Error Reason")
	reasonTaggedValue.Notes = ruleReason
	reasonTaggedValue.Update
	'add the tagged value with the concatenated path
	dim pathTaggedValue as EA.TaggedValue
	set pathTaggedValue = getExistingOrNewTaggedValue(messageRule, "Constraint Path")
	pathTaggedValue.Value = Join(path.ToArray(),".")
	pathTaggedValue.Update
	'link the message rule to the message it is related to	
	dim relatedMessageObjects
	set relatedMessageObjects = getRelatedMessageObjects(path)
	dim relatedMessage as EA.Element
	for each relatedMessage in relatedMessageObjects
		dim messageLink as EA.Connector
		set messageLink = messageRule.Connectors.AddNew("","Dependency")
		messageLink.SupplierID = relatedMessage.ElementID
		messageLink.Update
	next
end function

function getRelatedMessageObjects(path)
	dim getMessageObjectsSQL 
	getMessageObjectsSQL = "select o.Object_ID from t_object o                     " & _
							" inner join t_package p on o.Package_ID = p.Package_ID" & _
							" where o.Stereotype = 'XSDtopLevelElement'             " & _
							" and p.name = '" & path(0) & " '                      "
	dim messageObjects
	set messageObjects = getElementsFromQuery(getMessageObjectsSQL)
	set getRelatedMessageObjects = messageObjects
end function
main