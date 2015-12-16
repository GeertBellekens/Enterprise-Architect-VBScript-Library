'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
!INC Local Scripts.EAConstants-VBScript

' EA-Matic
' This script, when used with EA-Matic will maintain auto-updating diagrams for elements.
' A nested diagram with prefix AUTO_ will be considered an auto-updating diagram.
' The diagram will keep track of all elements related to the owner of the auto diagram.
' 
' Author: 	Geert Bellekens
' EA-Matic: http://bellekens.com/ea-matic/
'
'maintain a reference to the connector in context
dim contextConnectorID
dim oldClientID
dim oldSupplierID

'a new connector has been created. Add the related elements to the auto-diagram
function EA_OnPostNewConnector(Info)
	 'get the connector id from the Info
	 dim connectorID
	 connectorID = Info.Get("ConnectorID")
     dim model 
	 'get the model
     set model = getEAAddingFrameworkModel()
	 dim connector
	 set connector = model.getRelationByID(connectorID)
	 'get the related elements
	 dim relatedElements
	 set relatedElements = model.toArrayList(connector.relatedElements)
	 'for i = 0 to attributes.Count - 1
	 if relatedElements.Count = 2 then
		'once with the first
		addRelatedElementoAutoDiagram relatedElements(0), relatedElements(1), model
		'then with the second
		addRelatedElementoAutoDiagram relatedElements(1), relatedElements(0), model 
	 end if 
end function

'adds the related element to the auto_updatediagrams if any
function addRelatedElementoAutoDiagram(element,relatedElement, model)
	'get the diagram owned by this element
	dim ownedDiagrams
	set ownedDiagrams = model.toArrayList(element.ownedDiagrams)
	for each diagram In ownedDiagrams
		'check the name of the diagram
		if Left(diagram.name,LEN("AUTO_")) = "AUTO_" then
			'add the related element to the diagram
			diagram.addToDiagram(relatedElement)
		end if
	next
end function

'layout the auto diagram
function layoutAutoDiagram(diagramID, model)
	dim diagram
	set diagram = model.getDiagramByID(DiagramID)
	'if the diagram is an auto diagram then we do an automatic layout
	if Left(diagram.name,LEN("AUTO_")) = "AUTO_" then
		'auto layout diagram
		dim diagramGUIDXml
		'The project interface needs GUID's in XML format, so we need to convert first.
		diagramGUIDXml = Repository.GetProjectInterface().GUIDtoXML(diagram.wrappedDiagram.DiagramGUID)
		'Then call the layout operation
		Repository.GetProjectInterface().LayoutDiagramEx diagramGUIDXml, lsDiagramDefault, 4, 20 , 20, false
	end if
end function

' A connector will be deleted. Remove the elements from the auto-diagram
function EA_OnPreDeleteConnector(Info)
	 'get the connector id from the Info
	 dim connectorID
	 connectorID = Info.Get("ConnectorID")
     dim model 
	 'get the model
     set model = getEAAddingFrameworkModel()
	 dim connector
	 set connector = model.getRelationByID(connectorID)
	 'get the related elements
	 dim relatedElements
	 set relatedElements = model.toArrayList(connector.relatedElements)
	 'for i = 0 to attributes.Count - 1
	 if relatedElements.Count = 2 then
		'we only need to remove the related element if they are not connected anymore after deleting the connector
		'so only if there is only one relationship between the two elements
		if sharedRelationsCount(relatedElements(0), relatedElements(1), model) <= 1 then
			'once with the first
			removeRelatedElemenFromAutoDiagram relatedElements(0), relatedElements(1), model
			'then with the second
			removeRelatedElemenFromAutoDiagram relatedElements(1), relatedElements(0), model 
		end if
	 end if 
end function

'returns the number of relations that connecto both elements
function sharedRelationsCount(elementA, elementB, model)
	'start counting at zero
	sharedRelationsCount = 0
	'get the relationships for both objects
	dim relationsA
	set relationsA = model.toArrayList(elementA.relationships)
	dim relationsB
	set relationsB = model.toArrayList(elementB.relationships)
	for each relationA in relationsA
		for each relationB in relationsB
			'if both relations have the same ID then we have a shared relation
			if relationA.id = relationB.id then
				sharedRelationsCount = sharedRelationsCount +1
			end if 
		next 
	next 
end function

' Removes the related element from the auto update diagram if any.
function removeRelatedElemenFromAutoDiagram(element,relatedElement, model)
	'get the diagram owned by this element
	dim ownedDiagrams
	set ownedDiagrams = model.toArrayList(element.ownedDiagrams)
	for each diagram In ownedDiagrams
		dim diagram
		set diagram = ownedDiagrams(i)
		'check the name of the diagram
		if Left(diagram.name,LEN("AUTO_")) = "AUTO_" then
			'Removing elements from a diagram in unfortunately not implemented in the EAAddinFramework so we'll have to do it in the script
			dim eaDiagram 
			set eaDiagram = diagram.wrappedDiagram
			for i = 0 to eaDiagram.DiagramObjects.Count -1
				dim diagramObject
				set diagramObject = eaDiagram.DiagramObjects.GetAt(i)
				if diagramObject.ElementID = relatedElement.id then
					'remove the diagramObject
					eaDiagram.DiagramObjects.Delete(i)
					'refresh the diagram after we changed it
					diagram.reFresh()
					'exit the loop we have delete the diagramobject					
					exit for
				end if
			next
		end if
	next
end function

'gets a new instance of the EAAddinFramework and initializes it with the EA.Repository
function getEAAddingFrameworkModel()
	'Initialize the EAAddinFramework model
    dim model 
    set model = CreateObject("TSF.UmlToolingFramework.Wrappers.EA.Model")
    model.initialize(Repository)
	set getEAAddingFrameworkModel = model
end function

'autodiagrams are automatically layouted when opened
function EA_OnPostOpenDiagram(DiagramID)
	dim model 
	'get the model
	set model = getEAAddingFrameworkModel()
	layoutAutoDiagram DiagramID, model	
end function

'autodiagrams are automatically layouted when we tab is switched to them
function EA_OnTabChanged(TabName, DiagramID)
	if  DiagramID > 0 then
		dim model 
		'get the model
		set model = getEAAddingFrameworkModel()
		layoutAutoDiagram DiagramID, model		
	end if	 
end function

'keep a reference to the selected connector
function EA_OnContextItemChanged(GUID, ot)
	'we only do something when the context item is a connector
	if ot = otConnector then
		dim model
		'get the model
		set model = getEAAddingFrameworkModel()
		'get the connector
		dim contextConnector
		set contextConnector = model.getRelationByGUID(GUID)
		'MsgBox(TypeName(contextConnector))
		contextConnectorID = contextConnector.id
		oldClientID = contextConnector.WrappedConnector.ClientID 
		oldSupplierID = contextConnector.WrappedConnector.SupplierID 
	end if
end function

'a connector has changed, we need to update the auto-diagrams
function EA_OnNotifyContextItemModified(GUID, ot)
	'we only do something when the context item is a connector
	if ot = otConnector then
		dim model
		'get the model
		set model = getEAAddingFrameworkModel()
		'get the connector
		dim changedConnector
		set changedConnector = model.getRelationByGUID(GUID)
		'check if we are talking about the same connector
		if changedConnector.WrappedConnector.ConnectorID = contextConnectorID then
		    dim supplier
			dim client
			'check the client side
			if changedConnector.WrappedConnector.ClientID <>  oldClientID then
				'get supplier
				set supplier = model.getElementWrapperByID(changedConnector.WrappedConnector.SupplierID)
				'remove old client from supplier and vice versa
				set client = model.getElementWrapperByID(oldClientID)
				if not client is nothing then
					removeRelatedElemenFromAutoDiagram supplier,client, model
					removeRelatedElemenFromAutoDiagram client, supplier, model
				end if 
				'add new client
				set client = model.getElementWrapperByID(changedConnector.WrappedConnector.ClientID)
				addRelatedElementoAutoDiagram supplier,client, model
			end if
			'check the supplier side
			if changedConnector.WrappedConnector.SupplierID <> oldSupplierID then
				'get client
				set client = model.getElementWrapperByID(changedConnector.WrappedConnector.ClientID)
				'remove old supplier from client and vice versa
				set supplier = model.getElementWrapperByID(oldSupplierID)
				if not supplier is nothing then
					removeRelatedElemenFromAutoDiagram client,supplier, model
					removeRelatedElemenFromAutoDiagram supplier,client, model
				end if 
				'add new supplier
				set supplier = model.getElementWrapperByID(changedConnector.WrappedConnector.SupplierID)
				addRelatedElementoAutoDiagram client,supplier, model
			end if
		end if
	end if
end function