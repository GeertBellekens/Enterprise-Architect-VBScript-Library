'[path=\Projects\Project E\General Scripts]
'[group=General Scripts]

!INC Local Scripts.EAConstants-VBScript
!INC General Scripts.DocGenHelpers
!INC General Scripts.Util
'

' Script Name: UseCaseDocuemnt
' Author: Geert Bellekens
' Purpose: Create the virtual document for a Use Case Document based on the given diagram
' Date: 17/02/2016
'


function createServiceDocument( service, documentsPackage)
	'dim service as EA.Element
	'first create a master document
	dim masterDocument as EA.Package
	set masterDocument = makeServiceMasterDocument(service,documentsPackage)
	if not masterDocument is nothing then
		dim servicePackage as EA.Package
		set servicePackage = Repository.GetPackageByID(service.PackageID) 'get the package owning the service
		dim i
		i = 0
		'introduction
		'tell the user what he is expected
		msgbox "Please select the introduction artifact"
		dim introductionElement as EA.Element
		set introductionElement = nothing
		dim introductionID

		introductionID = Repository.InvokeConstructPicker("IncludedTypes=Artifact;Selection=" & servicePackage.PackageGUID & ";")
		if introductionID > 0 then
			set introductionElement = Repository.GetElementByID(introductionID)
		end if
		if not introductionElement is nothing then
			addModelDocument masterDocument, "Enexis Linked Document Template", service.Name & " Inleiding", introductionElement.ElementGUID, i
			i = i + 1
		end if
		'the service part
		addModelDocument masterDocument, "ESD_Service", service.Name & " Service", service.ElementGUID, i
			i = i + 1
		'the service diagram

		addModelDocumentForPackage masterDocument,servicePackage,"Diagram " & service.Name , i, "ESD_Diagram"
		i = i + 1
		
		'add interfaces
		i = addInterfaces(masterDocument, service, i)
		
		'add data model
		dim serviceDeclaration as EA.Package
		set serviceDeclaration = Repository.GetPackageByID(servicePackage.ParentID)
		addModelDocumentForPackage masterDocument,serviceDeclaration,"DataModel " & service.Name , i, "ESD_DataModel"
		i = i + 1
		
		Repository.RefreshModelView(masterDocument.PackageID)
		'select the created master document in the project browser
		Repository.ShowInProjectView(masterDocument)
	end if
end function


function makeServiceMasterDocument(service,documentsPackage)
	'we should ask the user for a version
	dim documentTitle
	dim documentVersion
	dim documentName
	dim diagramName
	set makeServiceMasterDocument = nothing
	'get version of the doucment
	documentVersion = InputBox("Please enter document version", "Document version", service.Version )
	if documentVersion <> "" then
		'OK, we have a version, continue
		documentName = "SVD - " & service.Name & " v. " & documentVersion
		dim masterDocument as EA.Package
		set masterDocument = addMasterDocumentWithDetails(documentsPackage.PackageGUID, documentName,documentVersion,service.Name)
		set makeServiceMasterDocument = masterDocument
	end if
end function

'add the interfaces to the document
function addInterfaces(masterDocument, service, i)
	'get the interfaces from the service
	dim interfaces
	set interfaces = getInterfacesForService(service)
	dim interface
	for each interface in interfaces
		'interface
		addModelDocument masterDocument, "ESD_Interface", "Interface " & interface.Name, interface.ElementGUID, i
		i = i + 1
		'interface operations
		i = addOperations(masterDocument, interface, i)
		
	next
	'return the new i
	addInterfaces = i
end function

'add the operations to the document
function addOperations(masterDocument, interface, i)
	'get the interfaces from the service
	dim operations
	set operations = getOperationsForInterface(interface)
	dim operation
	for each operation in operations
		'operation
		addModelDocument masterDocument, "ESD_Operation", "Operation " & operation.Name, operation.ElementGUID, i
		i = i + 1
		'messages
		i = addMessages(masterDocument, operation, i)
	next
	'return the new i
	addOperations = i
end function

'add the operations to the document
function addMessages(masterDocument, operation, i)
	'get the interfaces from the service
	dim messages
	set messages = getMessagesForOperation(operation)
	dim message
	for each message in messages
		'message
		addModelDocument masterDocument, "ESD_Message", "Message " & message.Name, message.ElementGUID, i
		i = i + 1
		'message diagram
		dim messagePackage
		set messagePackage = getMessagePackage(message)
		if not messagePackage is nothing then
			addModelDocumentForPackage masterDocument,messagePackage,"Diagram " & Message.Name , i, "ESD_Diagram"
			i = i + 1
		end if
	next
	'return the new i
	addMessages = i
end function

function getMessagePackage(message)
	'get the Message assemby linked to the message
	dim elementTypes
	dim stereotypes
	dim connectorTypes
	dim messageAssemblies
	dim messageAssembly as EA.Element
	set getMessagePackage = nothing
	elementTypes = Array("Class")
	stereotypes = Array("MessageAssembly")
	connectorTypes = Array("Association","Aggregation")
	set messageAssemblies = getRelatedElements(message,elementTypes,stereotypes, connectorTypes)
	if messageAssemblies.Count > 0 then
		set messageAssembly = messageAssemblies(0)
		set getMessagePackage = Repository.GetPackageByID(messageAssembly.PackageID)
	end if
end function

function getMessagesForOperation(operation)
	dim elementTypes
	dim stereotypes
	dim connectorTypes
	elementTypes = Array("Class")
	stereotypes = Array("BusinessMessage")
	connectorTypes = Array("Association","Aggregation")
	set getMessagesForOperation = getRelatedElements(operation,elementTypes,stereotypes, connectorTypes)
end function

function getOperationsForInterface(interface)
	dim elementTypes
	dim stereotypes
	dim connectorTypes
	elementTypes = Array("Class")
	stereotypes = Array("Operation")
	connectorTypes = Array("Association","Aggregation")
	set getOperationsForInterface = getRelatedElements(interface,elementTypes,stereotypes, connectorTypes)
end function

function getInterfacesForService(service)
	dim elementTypes
	dim stereotypes
	dim connectorTypes
	elementTypes = Array("Class")
	stereotypes = Array("InterfaceContract")
	connectorTypes = Array("Association","Aggregation")
	set getInterfacesForService = getRelatedElements(service,elementTypes,stereotypes, connectorTypes)
end function

function getRelatedElements(element,elementTypes,stereotypes, connectorTypes)
	dim sqlGet
	sqlGet = "select oo.Object_ID from ((t_object o " &_
			" inner join t_connector c on (o.Object_ID = c.Start_Object_ID " &_
			" 							or o.Object_ID = c.End_Object_ID)) " &_
			" inner join t_object oo on ((oo.Object_ID = c.Start_Object_ID " &_
			" 							or oo.Object_ID = c.End_Object_ID) " &_
			"							and o.Object_ID <> oo.Object_ID)) " &_
			" where " &_ 
			" o.Object_ID = " & element.ElementID &_
			" and c.Connector_Type in ('" & Join(connectorTypes,"','") & "') " &_
			" and oo.Object_Type in ('" & Join(elementTypes,"','") & "') " &_
			" and oo.Stereotype in ('" & Join(stereotypes,"','") & "') "
	'get the elements
	set getRelatedElements = getElementsFromQuery(sqlGet)
end function


function getNestedDiagramOnwerForElement(element, elementType)
	dim diagramOnwer as EA.Element
	set diagramOnwer = nothing
	dim nestedElement as EA.Element
	for each nestedElement in element.Elements
		if nestedElement.Type = elementType and nestedElement.Diagrams.Count > 0 then
			set diagramOnwer = nestedElement
			exit for
		end if
	next
	set getNestedDiagramOnwerForElement = diagramOnwer
end function


'sort the elements in the given ArrayList of EA.Elements by their name 
function sortElementsByName (elements)
	dim i
	dim goAgain
	goAgain = false
	dim thisElement as EA.Element
	dim nextElement as EA.Element
	for i = 0 to elements.Count -2 step 1
		set thisElement = elements(i)
		set nextElement = elements(i +1)
		if  elementIsAfter(thisElement, nextElement) then
			elements.RemoveAt(i +1)
			elements.Insert i, nextElement
			goAgain = true
		end if
	next
	'if we had to swap an element then we go over the list again
	if goAgain then
		set elements = sortElementsByName (elements)
	end if
	'return the sorted list
	set sortElementsByName = elements
end function

'check if the name of the next element is bigger then the name of the first element
function elementIsAfter (thisElement, nextElement)
	dim compareResult
	compareResult = StrComp(thisElement.Name, nextElement.Name,1)
	if compareResult > 0 then
		elementIsAfter = True
	else
		elementIsAfter = False
	end if
end function