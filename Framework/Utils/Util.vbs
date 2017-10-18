!INC Local Scripts.EAConstants-VBScript
'[path=\Framework\Utils]
'[group=Utils]
'
' Script Name: Util
' Author: Geert Bellekens
' Purpose: serves as library for other scripts
' Date: 28/09/2015
'
' Synchronises the names of the selected objects or BPMN Activities with their classifier/called activity ref.
' Will also set the composite diagram to that of the classifier/ActivityRef in order to facilitate click-through
function synchronizeElement (element)
	'first check if this is an object or an action
	if not element is Nothing then
		if (element.Type = "Object" OR element.Type = "Action") _
		AND element.ClassifierID > 0 then
			dim classifier
			set classifier = Repository.GetElementByID(element.ClassifierID)
			if not classifier is nothing AND classifier.name <> element.name then
				element.Name = classifier.Name
				element.Stereotype = classifier.Stereotype
				element.Update
				Repository.AdviseElementChange(element.ElementID)
			end if
			'elements of type object should also point to the composite diagram of the classifier
			if element.Type = "Object" then
				dim compositeDiagram
				set compositeDiagram = classifier.CompositeDiagram
				if not compositeDiagram is nothing then
					setCompositeDiagram element, compositeDiagram
				end if
			end if
		elseif element.Type = "Activity" AND element.Stereotype = "Activity" then
			'BPMN activities that call another BPMN activity need to get the same name and same composite diagram
			dim calledActivityTV as EA.TaggedValue
			set calledActivityTV = element.TaggedValues.GetByName("isACalledActivity")
			dim referenceActivityTV as EA.TaggedValue
			set referenceActivityTV = element.TaggedValues.GetByName("calledActivityRef")
			if not calledActivityTV is nothing and not referenceActivityTV is nothing then
				'only do something when the Activity is types a CalledActivity
				'Session.Output "calledActivityTV.Value : " & calledActivityTV.Value 
				'Session.Output "referenceActivityTV.Value :" & referenceActivityTV.Value
				if calledActivityTV.Value = "true" then
					dim calledActivity as EA.Element
					set calledActivity = Repository.GetElementByGuid(referenceActivityTV.Value)
					if not calledActivity is nothing then
						'set name to that of the called activity
						element.Name = calledActivity.Name
						element.Update
						'Set composite diagram to that of the called activity
						setCompositeDiagram element, calledActivity.CompositeDiagram
					end if
				end if
			end if
		end if
	end if
end function



'set the given diagram as composite diagram for this element
function setCompositeDiagram (element, diagram)
	if not diagram is nothing then
		'Tell EA this element is composite
		dim objectQuery
		objectQuery = "update t_object set NType = 8 where Object_ID = " & element.ElementID
		Repository.Execute objectQuery
		if element.Type = "Object" then
			'Tell EA which diagram is the composite diagram
			dim xrefQuery
			xrefquery = "insert into t_xref (XrefID, Name, Type, Visibility, Partition, Client, Supplier) values ('"&CreateGuid&"', 'DefaultDiagram', 'element property', 'Public', '0', '"& element.ElementGUID & "', '"& diagram.DiagramGUID &"')"
			Repository.Execute xrefquery
		elseif element.Type = "Activity" then
			'for activities we need to update PDATA1 with the diagramID
			dim updatequery
			updatequery = "update t_object set PDATA1 = "& diagram.DiagramID & " where Object_ID = " & element.ElementID
			Repository.Execute updatequery
		end if
	end if
end function

' Returns a unique Guid on every call. Removes any cruft.
Function CreateGuid()
    CreateGuid = Left(CreateObject("Scriptlet.TypeLib").Guid,38)
End Function


'make an action into a calling activity
function makeCallingActivity(action, activity)
	action.Type = "Activity"
	action.ClassfierID = 0
	action.Stereotype = "Activity"
	action.Update
	action.SynchTaggedValues "BPMN2.0","Activity"
	action.TaggedValues.Refresh
	'first copy the tagged values values
	copyTaggedValuesValues activity, action
	'set tagged values correctly
	dim calledActivityTV as EA.TaggedValue
	set calledActivityTV = action.TaggedValues.GetByName("isACalledActivity")
	calledActivityTV.Value = "true"
	calledActivityTV.Update
	dim referenceActivityTV as EA.TaggedValue
	set referenceActivityTV = action.TaggedValues.GetByName("calledActivityRef")
	referenceActivityTV.Value = activity.ElementGUID
	referenceActivityTV.Update
	action.TaggedValues.Refresh()
end function

'copies values of the tagged values of the source to the values of the corresponding tagged values at the target
function copyTaggedValuesValues (source, target)
	dim taggedValue as EA.TaggedValue
	for each taggedValue in source.TaggedValues
		dim targetTaggedValue as EA.TaggedValue
		set targetTaggedValue = target.TaggedValues.GetByName(taggedValue.Name)
		if not targetTaggedValue is nothing then
			targetTaggedValue.Value = taggedValue.Value
			targetTaggedValue.Update
		end if
	next
end function

'copies the tagged values from the source to the target
function copyTaggedValues(source, target)
	dim sourceTag as EA.TaggedValue
	dim targetTag as EA.TaggedValue
	for each sourceTag in source.TaggedValues
		set targetTag = nothing
		'first try to find target tag
		dim tag as EA.TaggedValue
		for each tag in target.TaggedValues
			if tag.Name = sourceTag.Name then
				set targetTag = tag
				exit for
			end if
		next
		'if not found then create new
		if targetTag is nothing then
			set targetTag = target.TaggedValues.AddNew(sourceTag.Name,"TaggedValue")
		end if
		'set value
		if not targetTag is nothing then
			targetTag.Value = sourceTag.Value
			targetTag.Update
			target.Update
		end if
	next
end function

function setFontOnDiagramObject(diagramObject, font, size )
	dim styleParts
	styleParts = Split (diagramObject.Style , ";") 
	dim i
	dim stylepart
	dim fontpart 
	fontpart = "font=" & font
	dim fontSet
	fontSet = false
	dim sizePart
	sizePart = "fontsz=" & size * 10
	dim sizeSet
	sizeSet = false
	for i = 0 to Ubound(styleParts) -1
		stylepart = styleParts(i)
		if Instr(stylepart,"font=") > 0 then
			styleParts(i) = fontpart
			fontSet = true
		elseif Instr(stylepart,"fontsz=") > 0 then
			styleParts(i) = sizePart
			sizeSet = true
		end if
	next
	diagramObject.Style = join(styleParts,";")
	if not fontSet then
		diagramObject.Style =  diagramObject.Style & fontpart & ";"
	end if
	if not sizeSet then
		diagramObject.Style =  diagramObject.Style & sizePart & ";"
	end if
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

'get the package id string of the currently selected package tree
function getCurrentPackageTreeIDString()
	'initialize at "0"
	getCurrentPackageTreeIDString = "0"
	dim packageTree
	dim currentPackage as EA.Package
	'get selected package
	set currentPackage = Repository.GetTreeSelectedPackage()
	if not currentPackage is nothing then
		'get the whole tree of the selected package
		set packageTree = getPackageTree(currentPackage)
		' get the id string of the tree
		getCurrentPackageTreeIDString = makePackageIDString(packageTree)
	end if 
end function

'get the package id string of the given package tree
function getPackageTreeIDString(package)
	'initialize at "0"
	getPackageTreeIDString = "0"
	dim packageTree
	dim currentPackage as EA.Package
	if not package is nothing then
		'get the whole tree of the selected package
		set packageTree = getPackageTree(package)
		' get the id string of the tree
		getPackageTreeIDString = makePackageIDString(packageTree)
	end if 
end function

'returns an ArrayList of the given package and all its subpackages recursively
function getPackageTree(package)
	dim packageList
	set packageList = CreateObject("System.Collections.ArrayList")
	addPackagesToList package, packageList
	set getPackageTree = packageList
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
	if idString = "" then
		idString = "0"
	end if
	'return idString
	makePackageIDString = idString
end function

'make an id string out of the ID's of the given elements
function makeIDString(elements)
	dim element as EA.Element
	dim idString
	idString = ""
	dim addComma 
	addComma = false
	for each element in elements
		if addComma then
			idString = idString & ","
		else
			addComma = true
		end if
		idString = idString & element.ElementID
	next 
	'if there are no elements then we return "0"
	if idString = "" then
		idString = "0"
	end if
	'return idString
	makeIDString = idString
end function

'returns the elements in an ArrayList of the given type from the given diagram
function getElementsFromDiagram(diagram, elementType)
	dim selectedElements
	set selectedElements = CreateObject("System.Collections.ArrayList")
	dim diagramObject as EA.DiagramObject
	dim element as EA.Element
	for each diagramObject in diagram.DiagramObjects
		set element = Repository.GetElementByID(diagramObject.ElementID)
		if element.Type = elementType then
			selectedElements.Add element
		end if
	next
	'return selected Elements
	set getElementsFromDiagram = selectedElements
end function

'returns the diagram objects in an ArrayList for elements of the given type from the given diagram
function getDiagramObjects(diagram, elementType)
	dim selectedElements
	set selectedElements = CreateObject("System.Collections.ArrayList")
	dim diagramObject as EA.DiagramObject
	dim element as EA.Element
	for each diagramObject in diagram.DiagramObjects
		set element = Repository.GetElementByID(diagramObject.ElementID)
		if element.Type = elementType then
			selectedElements.Add diagramObject
		end if
	next
	'return selected Elements
	set getDiagramObjects = selectedElements
end function

'returns the elements in an ArrayList of the given type from the given diagram
'the boundary element should be passed as a DiagramObject
function getElementsFromDiagramInBoundary(diagram, elementType,boundary)
	'dim boundary as EA.DiagramObject
	dim selectedElements
	set selectedElements = CreateObject("System.Collections.ArrayList")
	dim diagramObject as EA.DiagramObject
	dim element as EA.Element
	for each diagramObject in diagram.DiagramObjects
		if (diagramObject.left >= boundary.left and _
			diagramObject.left =< boundary.right and _
			diagramObject.top =< boundary.top and _
			diagramObject.top >= boundary.bottom) then
			'get the element and check the type
			set element = Repository.GetElementByID(diagramObject.ElementID)
			if element.Type = elementType then
				selectedElements.Add element
			end if
		end if
	next
	'return selected Elements
	set getElementsFromDiagramInBoundary = selectedElements
end function

function getWC()
	if Repository.RepositoryType = "JET" then
		getWC = "*"
	else
		getWC = "%"
	end if
end function


function addElementToDiagram(element, diagram, y, x)
	dim diagramObject as EA.DiagramObject
	dim positionString
	'determine height and width
	dim width 
	dim height
	dim elementType
	dim setVPartition 
	setVPartition = false
	elementType = element.Type
	select case elementType       
		case "Event"
			width = 30
			height = 30
		case "Object"
			width = 40
			height = 25
		case "Activity"
			width = 110
			height = 60
		case "ActivityPartition"
			width = 190
			height = 60
			setVPartition = true
		case "Package"
			width = 75
			height = 90
		case else
			'default width and height	
			width = 75
			height = 50
	end select
	'to make sure all elements are vertically aligned we subtract half of the width of the x
	x = x - width/2
	'set the position of the diagramObject
	positionString =  "l=" & x & ";r=" & x + width & ";t=" & y & ";b=" & y + height & ";"
	Session.Output "positionString voor element "& element.Name & " : " &  positionString
	set diagramObject = diagram.DiagramObjects.AddNew( positionString, "" )
	diagramObject.ElementID = element.ElementID
	if setVPartition then
		diagramObject.Style = "VPartition=1"
	end if
	diagramObject.Update
	diagram.DiagramObjects.Refresh
	set addElementToDiagram = diagramObject
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

'returns the currently logged in user
'if security is not enabled then the logged in user is defaulted to me
function getUserLogin()
	'get the currently logged in user
	Dim userLogin
	if Repository.IsSecurityEnabled then
		userLogin = Repository.GetCurrentLoginUser(false)
	else
		userLogin = "SYSTEMAT-TCC\BellekensG"
	end if
	getUserLogin = userLogin
end function	

function getArrayFromQuery(sqlQuery)
	dim xmlResult
	xmlResult = Repository.SQLQuery(sqlQuery)
	getArrayFromQuery = convertQueryResultToArray(xmlResult)
end function

'converts the query results from Repository.SQLQuery from xml format to a two dimensional array of strings
Public Function convertQueryResultToArray(xmlQueryResult)
    Dim arrayCreated
    Dim i 
    i = 0
    Dim j 
    j = 0
    Dim result()
    Dim xDoc 
    Set xDoc = CreateObject( "MSXML2.DOMDocument" )
    'load the resultset in the xml document
    If xDoc.LoadXML(xmlQueryResult) Then        
		'select the rows
		Dim rowList
		Set rowList = xDoc.SelectNodes("//Row")

		Dim rowNode 
		Dim fieldNode
		arrayCreated = False
		'loop rows and find fields
		For Each rowNode In rowList
			j = 0
			If (rowNode.HasChildNodes) Then
				'redim array (only once)
				If Not arrayCreated Then
					ReDim result(rowList.Length, rowNode.ChildNodes.Length)
					arrayCreated = True
				End If
				For Each fieldNode In rowNode.ChildNodes
					'write f
					result(i, j) = fieldNode.Text
					j = j + 1
				Next
			End If
			i = i + 1
		Next
		'make sure the array has a dimension even is we don't have any results
		if not arrayCreated then
			ReDim result(0, 0)
		end if
	end if
    convertQueryResultToArray = result
End Function

'let the user select a package
function selectPackage()
	'start from the selected package in the project browser
	dim constructpickerString
	constructpickerString = "IncludedTypes=Package"
	dim treeselectedPackage as EA.Package
	set treeselectedPackage = Repository.GetTreeSelectedPackage()
	if not treeselectedPackage is nothing then
		constructpickerString = constructpickerString &	";Selection=" & treeselectedPackage.PackageGUID
	end if
	dim packageElementID 		
	packageElementID = Repository.InvokeConstructPicker(constructpickerString) 
	if packageElementID > 0 then
		dim packageElement as EA.Element
		set packageElement = Repository.GetElementByID(packageElementID)
		dim package as EA.Package
		set package = Repository.GetPackageByGuid(packageElement.ElementGUID)
	else
		set package = nothing
	end if 
	set selectPackage = package
end function

function getConnectorsFromQuery(sqlQuery)
	dim xmlResult
	xmlResult = Repository.SQLQuery(sqlQuery)
	dim connectorIDs
	connectorIDs = convertQueryResultToArray(xmlResult)
	dim connectors 
	set connectors = CreateObject("System.Collections.ArrayList")
	dim connectorID
	dim connector as EA.Connector
	for each connectorID in connectorIDs
		if connectorID > 0 then
			set connector = Repository.GetConnectorByID(connectorID)
			if not connector is nothing then
				connectors.Add(connector)
			end if
		end if
	next
	set getConnectorsFromQuery = connectors
end function

function getDiagramsFromQuery(sqlQuery)
	dim xmlResult
	xmlResult = Repository.SQLQuery(sqlQuery)
	dim diagramIDs
	diagramIDs = convertQueryResultToArray(xmlResult)
	dim diagrams 
	set diagrams = CreateObject("System.Collections.ArrayList")
	dim diagramID
	dim diagram as EA.Diagram
	for each diagramID in diagramIDs
		if diagramID > 0 then
			set diagram = Repository.GetdiagramByID(diagramID)
			if not diagram is nothing then
				diagrams.Add(diagram)
			end if
		end if
	next
	set getDiagramsFromQuery = diagrams
end function

function getattributesFromQuery(sqlQuery)
	dim xmlResult
	xmlResult = Repository.SQLQuery(sqlQuery)
	dim attributeIDs
	attributeIDs = convertQueryResultToArray(xmlResult)
	dim attributes 
	set attributes = CreateObject("System.Collections.ArrayList")
	dim attributeID
	dim attribute as EA.Attribute
	for each attributeID in attributeIDs
		if attributeID > 0 then
			set attribute = Repository.GetAttributeByID(attributeID)
			if not attribute is nothing then
				attributes.Add(attribute)
			end if
		end if
	next
	set getattributesFromQuery = attributes
end function

'get the description from the given notes 
'that is the text between <NL> and </NL> or <FR> and </FR>
function getTagContent(notes, tag)
	if tag = "" then
		getTagContent = notes
	else
		getTagContent = ""
		dim startTagPosition
		dim endTagPosition
		startTagPosition = InStr(notes,"&lt;" & tag & "&gt;")
		endTagPosition = InStr(notes,"&lt;/" & tag & "&gt;")
		'Session.Output "notes: " & notes & " startTagPosition: " & startTagPosition & " endTagPosition: " &endTagPosition
		if startTagPosition > 0 and endTagPosition > startTagPosition then
			dim startContent
			startContent = startTagPosition + len(tag) + 8
			dim length 
			length = endTagPosition - startContent
			getTagContent = mid(notes, startContent, length)
		end if
	end if 
end function

'Returns the value of the tagged value with the given name (case insensitive)
'If there is no tagged value with the given name, an empty string is returned
'This function can be used with anything that can have tagged values
function getTaggedValueValue(owner, taggedValueName)
	dim taggedValue as EA.TaggedValue
	getTaggedValueValue = ""
	for each taggedValue in owner.TaggedValues
		if lcase(taggedValueName) = lcase(taggedValue.Name) then
			getTaggedValueValue = taggedValue.Value
			exit for
		end if
	next
end function

function getExistingOrNewTaggedValue(owner, tagname)
	dim taggedValue as EA.TaggedValue
	dim returnTag as EA.TaggedValue
	set returnTag = nothing
	'check if a tag with that name alrady exists
	for each taggedValue in owner.TaggedValues
		if taggedValue.Name = tagName then
			set returnTag = taggedValue
			exit for
		end if
	next
	'create new one if not found
	if returnTag is nothing then
		set returnTag = owner.TaggedValues.AddNew(tagname,"")
	end if
	'return
	set getExistingOrNewTaggedValue = returnTag
end function

function isRequireUserLockEnabled()
	dim reqUserLockToEdit
	'default is false
	reqUserLockToEdit = false
	'check if security is enabled
	if Repository.IsSecurityEnabled then
		dim getReqUserLockSQL
		getReqUserLockSQL =	"select sc.Value from t_secpolicies sc " & _
							"where sc.Property = 'RequireLock' "
		dim xmlQueryResult
		xmlQueryResult = Repository.SQLQuery(getReqUserLockSQL)
		dim reqUserLockResults
		reqUserLockResults = convertQueryResultToArray(xmlQueryResult)
		if Ubound(reqUserLockResults) > 0 then
			if reqUserLockResults(0,0) = "1" then
				reqUserLockToEdit = true
			end if
		end if
	end if
	isRequireUserLockEnabled = reqUserLockToEdit
end function

function copyDiagram(diagram, targetOwner)
	dim copiedDiagram as EA.Diagram
	'initialize at nothing
	set copiedDiagram = nothing
	'get the owner package
	dim ownerPackage as EA.Package
	set ownerPackage = Repository.GetPackageByID(diagram.PackageID)
	'check if we need to lock the package to clone it
	if isRequireUserLockEnabled() then
		dim ownerOfOwnerPackage as EA.Package
		if ownerPackage.ParentID > 0 then
			set ownerOfOwnerPackage = Repository.GetPackageByID(ownerPackage.ParentID)
			if not ownerOfOwnerPackage.ApplyUserLock() then
				'tell the user we couldn't do it and then exit the function
				msgbox "Could not lock package " &  ownerPackage.Name & " in order to copy the diagram " & diagram.Name,vbError,"Could not lock Package"
				exit function
			end if
		end if
	end if
	'then actually clone the owner package
	dim clonedPackage as EA.Package
	set clonedPackage = ownerPackage.Clone()
'	if isRequireUserLockEnabled() then
'		clonedPackage.ApplyUserLockRecursive true,true,true
'	end if
	'then get the diagram corresponding to the diagram to copy
	set copiedDiagram = getCorrespondingDiagram(clonedPackage,diagram)
	'set the owner of the copied diagram
	if targetOwner.ObjectType = otElement then
		copiedDiagram.ParentID = targetOwner.ElementID
	else
		copiedDiagram.PackageID = targetOwner.PackageID
	end if
	'save the update to the owner
	copiedDiagram.Update
	'delete the cloned package
	deletePackage(clonedPackage)
	'return the copied diagram
	set copyDiagram = copiedDiagram
end function

function deletePackage(package)
	if package.ParentID > 0 then
		'get parent package
		dim parentPackage as EA.Package
		set parentPackage = Repository.GetPackageByID(package.ParentID )
		dim i
		'delete the pacakge
		for i = parentPackage.Packages.Count -1 to 0 step -1
			dim currentPackage as EA.Package
			set currentPackage = parentPackage.Packages(i)
			if currentPackage.PackageID = package.PackageID then
				parentPackage.Packages.DeleteAt i,false
				exit for
			end if
		next
	end if
end function

function getCorrespondingDiagram(clonedPackage,diagram)
	dim correspondingDiagram as EA.Diagram
	dim candidateDiagrams
	dim getCandidateDiagramsSQL
	dim packageIDs
	packageIDs = getPackageTreeIDString(clonedPackage)
	getCandidateDiagramsSQL = 	"select d.Diagram_ID from t_diagram d " & _
								" where d.name = '" & diagram.Name & "' " & _
								" and d.Package_ID in (" & packageIDs& ") "
	set candidateDiagrams = getDiagramsFromQuery(getCandidateDiagramsSQL)
	'if there is only one candidate then that is the one we take
	if candidateDiagrams.Count = 1 then
		set correspondingDiagram = candidateDiagrams(0)
	end if
	'if there are multiple candidates then we have to filter them
	'first create a dictionary with the diagrams and their owner
	dim candidateDiagramsDictionary
	set candidateDiagramsDictionary = CreateObject("Scripting.Dictionary")
	dim currentDiagram
	for each currentDiagram in candidateDiagrams
		'add the diagram and its owner to the dictionary
		candidateDiagramsDictionary.Add currentDiagram, getOwner(diagram)
	next
	dim currentowner
	set currentOwner = nothing
	'filter the diagrams until we have only one diagram left
	set correspondingDiagram = filterDiagrams(candidateDiagramsDictionary,diagram, clonedPackage, currentOwner)
	'return the diagram
	set getCorrespondingDiagram = correspondingDiagram
end function

function filterDiagrams(candidateDiagramsDictionary,diagram, clonedPackage, currentOwner)
	dim filteredDiagrams
	dim filteredDiagram as EA.Diagram
	'initialize at nothing
	set filteredDiagram = nothing
	set filteredDiagrams = CreateObject("Scripting.Dictionary")
	if currentOwner is nothing then
		set currentOwner = getOwner(diagram)
	end if
	'compare the diagrams and their owner with the current owner
	dim candidateDiagram as EA.Diagram
	dim candidateOwner
	for each candidateDiagram in candidateDiagramsDictionary.Keys
		set candidateOwner = candidateDiagramsDictionary(candidateDiagram)
		if candidateOwner.Name = currentOwner.Name then
			'add the diagram to the new list 
			filteredDiagrams.Add candidateDiagram, getOwner(candidateOwner)
		end if
	next
	'check the number if we have reached he level of the cloned package, or if there is only one diagram left
	if filteredDiagrams.Count = 1 _
	OR currentOwner.ObjectType = otPackage AND currentOwner.ParentID = clonedPackage.PackageID then
		'return the first one
		set filteredDiagram = filteredDiagrams.Keys()(0)
	else
		'go one level deeper to filter the diagrams
		set currentOwner = getOwner(currentOwner)
		set filteredDiagram = filterDiagrams(filteredDiagrams,diagram, clonedPackage, currentOwner)
	end if
	'return filtered diagram
	set filterDiagrams = filteredDiagram
end function

function getOwner(item)
	dim owner
	select case item.ObjectType
		case otElement,otDiagram,otPackage
			'if it has an element as owner then we return the element
			if item.ParentID > 0 then
				set owner = Repository.GetElementByID(item.ParentID)
			else
				if item.ObjectType <> otPackage then
					'else we return the package (not for packages because then we have a root package that doesn't have an owner)
					set owner = Repository.GetPackageByID(item.PackageID)
				end if
			end if
	'TODO: add other cases such as attributes and operations
	end select
	'return owner
	set getOwner = owner
end function

Function lpad(strInput, length, character)
  lpad = Right(String(length, character) & strInput, length)
end function

function makeArrayFromArrayLists(arrayLists)
	dim returnArray()
	'get the dimensions
	dim x
	dim y
	x = arrayLists.Count
	y = arrayLists(0).Count
	'redim the array to the correct dimensions
	redim returnArray(x,y)
	dim i,j
	i = 0
	dim row
	dim field
	for each row in arrayLists
		'reset j
		j = 0
		for each field in row
			if IsObject(field) then
				set returnArray(i,j) = field
			else
				returnArray(i,j) = field
			end if
			j = j + 1
		next
		i = i + 1
	next
	'return the array
	makeArrayFromArrayLists = returnArray
end function

'EA uses a lot of key=value pairs in different types of fields (such as StyleEx etc.)
' each of them separated by a ";"
' this function will search for the value of the key and return the value if it is present in the given search string
function getValueForkey(searchString, key)
	dim returnValue
	returnValue = ""
	'first split int keyvalue pairs using ";"
	dim keyValuePairs
	keyValuePairs = split(searchString,";")
	'then loop the key value pairs
	dim keyValuePairString
	for each keyValuePairString in keyValuePairs
		'and split them usign "=" as delimiter
		dim keyValuePair
		if instr(keyValuePairString,"=") > 0 then
			keyValuePair = split(keyValuePairString,"=")
			if UBound(keyValuePair) = 2 then
				if keyValuePair(1) = key then
					returnValue = keyValuePair(1)
				end if
			end if
		end if
	next
	'return the value
	getValueForkey = returnValue
end function

function copyDiagram(diagram, targetOwner)
	if targetOwner.Objecttype = otPackage then
		'create the new diagram
		dim copiedDiagram as EA.Diagram
		set copiedDiagram = targetOwner.Diagrams.AddNew(diagram.Name, diagram.Type)
		copiedDiagram.Stereotype = diagram.Stereotype
		copiedDiagram.StyleEx = diagram.StyleEx
		copiedDiagram.Notes = diagram.Notes
		copiedDiagram.ExtendedStyle = diagram.ExtendedStyle
		copiedDiagram.ShowDetails = diagram.ShowDetails
		copiedDiagram.ShowPackageContents = diagram.ShowPackageContents
		copiedDiagram.Version = diagram.Version
		copiedDiagram.Update 'hopefully this is enough
		'recreate all diagramObjects
		copyDiagramObjects copiedDiagram, diagram
		'recreate all diagramLinks
		copyDiagramLinks copiedDiagram, diagram
	else
		msgbox "copy diagram currently only supported for copying to packages"
	end if
	'do we need to save the diagram here?
	'diagram.Update
	'return diagram
	set copyDiagram = copiedDiagram
end function 

function copyDiagramObjects(copiedDiagram, diagram)
	dim currentElement as EA.Element
	dim currentDiagramObject as EA.DiagramObject
	dim targetPackage as EA.Element
	set targetPackage = Repository.GetPackageByID(copiedDiagram.PackageID)
	for each currentDiagramObject in diagram.DiagramObjects
		set currentElement = Repository.GetElementByID(currentDiagramObject.ElementID)
		'in case of diagram owned objects we need to copy them as well
		select case currentElement.Type
			case "Note","Boundary","Text"
			set currentElement = copyOwnedElement(currentElement,targetPackage)
		end select
		'copy the diagram object
		dim newDiagramObject as EA.DiagramObject
		set newDiagramObject = copiedDiagram.DiagramObjects.AddNew("","")
		newDiagramObject.ElementID = currentDiagramObject.ElementID
		newDiagramObject.top = currentDiagramObject.top
		newDiagramObject.bottom = currentDiagramObject.bottom
		newDiagramObject.left = currentDiagramObject.left
		newDiagramObject.right = currentDiagramObject.right
		newDiagramObject.fontSize = currentDiagramObject.fontSize
		newDiagramObject.fontName = currentDiagramObject.fontName
		newDiagramObject.FontBold = currentDiagramObject.FontBold
		newDiagramObject.FontColor = currentDiagramObject.FontColor
		newDiagramObject.FontItalic = currentDiagramObject.FontItalic
		newDiagramObject.FontUnderline = currentDiagramObject.FontUnderline
		newDiagramObject.Update
	next
end function

function copyDiagramLinks(copiedDiagram, diagram)
	dim currentDiagramLink as EA.DiagramLink
	for each currentDiagramLink in diagram.DiagramLinks
		'copy each diagram link
		dim newDiagramLink as EA.DiagramLink
		set newDiagramLink = copiedDiagram.DiagramLinks.AddNew("","")
		newDiagramLink.ConnectorID = currentDiagramLink.ConnectorID
		newDiagramLink.Geometry = currentDiagramLink.Geometry
		newDiagramLink.IsHidden = currentDiagramLink.IsHidden
		newDiagramLink.LineStyle = currentDiagramLink.LineStyle
		newDiagramLink.LineColor = currentDiagramLink.LineColor
		newDiagramLink.LineWidth = currentDiagramLink.LineWidth
		newDiagramLink.Path = currentDiagramLink.Path
		newDiagramLink.HiddenLabels = currentDiagramLink.HiddenLabels
		newDiagramLink.Update
	next
end function

function copyOwnedElement(currentElement, targetPackage)
	dim newOwnedElement as EA.Element
	set newOwnedElement = targetPackage.Elements.AddNew(currentElement.Name,currentElement.Type)
	newOwnedElement.Notes = currentElement.Notes
	newOwnedElement.Subtype = currentElement.Subtype
	newOwnedElement.StyleEx = currentElement.StyleEx
	newOwnedElement.Alias = currentElement.Alias
	newOwnedElement.Update 'hopefully this is enough
	'return the object
	set copyOwnedElement = newOwnedElement
end function

function deletePackage(package)
	if package.ParentID > 0 then
		'get parent package
		dim parentPackage as EA.Package
		set parentPackage = Repository.GetPackageByID(package.ParentID )
		dim i
		'delete the pacakge
		for i = parentPackage.Packages.Count -1 to 0 step -1
			dim currentPackage as EA.Package
			set currentPackage = parentPackage.Packages(i)
			if currentPackage.PackageID = package.PackageID then
				parentPackage.Packages.DeleteAt i,false
				exit for
			end if
		next
	end if
end function

function getOwner(item)
	dim owner
	select case item.ObjectType
		case otPackage
			if item.ParentID > 0 then
				set owner = Repository.GetPackageByID(item.ParentID)
			end if
		case otElement,otDiagram
			'if it has an element as owner then we return the element
			if item.ParentID > 0 then
				set owner = Repository.GetElementByID(item.ParentID)
			else
				if item.ObjectType <> otPackage then
					'else we return the package (not for packages because then we have a root package that doesn't have an owner)
					set owner = Repository.GetPackageByID(item.PackageID)
				end if
			end if
	'TODO: add other cases such as attributes and operations
	end select
	'return owner
	set getOwner = owner
end function


'put the given value onto the clipboard
function putOnClipBoard(stringValue)
	dim WshShell
	Set WshShell = CreateObject("WScript.Shell")
	WshShell.Run "cmd.exe /c echo " & stringValue & " | clip", 0, TRUE
end function

'merge two array together. First a1, then a2
Function mergeArrays(a1, a2)
  ReDim aTmp(Ubound(a1, 1) + Ubound(a2,1) + 1, UBound(a1, 2) )
  Dim i, j, k
  For i = 0 To UBound(a1, 1)
      For j = 0 To UBound(aTmp, 2)
          aTmp(i, j) = a1(i, j)
      Next
  Next
  For k = 0 To UBound(a2, 1)
      For j = 0 To UBound(aTmp, 2)
          aTmp(i + k, j) = a2(k, j)
      Next
  Next
  mergeArrays = aTmp
End Function