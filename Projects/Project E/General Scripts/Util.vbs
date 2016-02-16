'[path=\Projects\Project E\General Scripts]
'[group=General Scripts]
!INC Local Scripts.EAConstants-VBScript

'Repository types
'dim rpt_MSQL, rpt_SQLSVR, rpt_ADOJET, rpt_ORACLE, rpt_POSTGRES, rpt_ASA, rpt_OPENEDGE, rpt_ACCESS2007, rpt_FireBird
'rpt_MSQL = 0
'rpt_SQLSVR = 2
'rpt_ADOJET = 3
'rpt_ORACLE = 4 
'rpt_POSTGRES = 5 
'rpt_ASA = 6
'rpt_OPENEDGE = 7 
'rpt_ACCESS2007 = 8
'rpt_FireBird = 9


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

'returns all elements in the package tree (so elements in the package and all subpackages recursively)
function getAllElementsInPackageTree(package)
	dim packageList 
	set packageList = getPackageTree(package)
	dim packageIDString
	packageIDString = makePackageIDString(packageList)
	dim getElementsSQL
	getElementsSQL = "select o.Object_ID from t_object o where o.Package_ID in (" & packageIDString & ")"
	dim elements
	set elements = getElementsFromQuery(getElementsSQL)
	set getAllElementsInPackageTree = elements
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
	if packages.Count = 0 then
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
	if elements.Count = 0 then
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
		dim network
 		Set network = CreateObject("Wscript.Network")
		userLogin = network.UserName
 	end if
	getUserLogin = userLogin
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
	end if
    convertQueryResultToArray = result
End Function

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