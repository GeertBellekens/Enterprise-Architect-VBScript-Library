'[path=\Projects\Project A\Old Scripts]
'[group=Old Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.Util


'
' Script Name: FixDiagram
' Author: Geert Bellekens
' Purpose: 
'    - Converts all Invocations to local BPMN Activities calling the library Activity
'    - Converts all links to Activities from the librabry to local BPMN Activities calling the library Activity, 
'		including embedded intermediat events
'	 - Converts all linked pools and lanes to local pools and lanes, respecting the nesting structure.
' Date: 18/02/2015
'
sub test
	' Get a reference to the current diagram
	dim currentDiagram as EA.Diagram
	set currentDiagram = Repository.GetTreeSelectedObject()
	
	if not currentDiagram is nothing then
		convertInvocationsToCallingActivities currentDiagram
	end if
end sub

sub main()
	' Get a reference to the current diagram
	dim currentDiagram as EA.Diagram
	set currentDiagram = Repository.GetCurrentDiagram()
	
	if not currentDiagram is nothing then
		convertInvocationsToCallingActivities currentDiagram
	end if
end sub

main
'test

function convertInvocationsToCallingActivities(diagram)
	'save the diagram before anything else
	Repository.SaveDiagram diagram.DiagramID
	dim diagramObject as EA.DiagramObject
	dim ownerElement as EA.Element
	dim ownerPackage as EA.Package
	dim updateDiagramObjectSQL
	'first copy the diagramobject into an array
	dim count 
	count = diagram.DiagramObjects.Count
	dim diagramObjects()
	redim diagramObjects(count)
	dim i
	for i = 1 to count
		set diagramObjects(i) = diagram.DiagramObjects.GetAt(i-1)
		'Session.Output "diagramObjects(i).ElementID: i=" & i & " elementID = " &diagramObjects(i).ElementID
	next
	'Session.Output "diagram.DiagramObjects.Count: " & diagram.DiagramObjects.Count
	'Session.Output "diagramObjects.Count: " & UBound(diagramObjects)
	'then loop the array
	for i = 1 to count
		set diagramObject = diagramObjects(i)
		dim element as EA.Element
		set element = nothing
		set element = Repository.GetElementByID(diagramObject.ElementID)
		if element.Type = "Action" then
			dim action
			set action = element
			'Check if it calls a BPMN Activity
			if action.ClassfierID > 0 then
				dim activity as EA.Element
				set activity = nothing
				set activity = Repository.GetElementByID(action.ClassifierID)
				if not activity is nothing then
					if activity.Stereotype = "Activity" then
						'OK we got one. Make it into a calling activity
						makeCallingActivity action, activity
						synchronizeElement action
						Repository.AdviseElementChange(action.ElementID)
					end if
				end if
			end if
		'Pools and lanes need to be copied locally and stripped of their name in case they have a classifier set (which they should)
		elseif element.Type = "ActivityPartition" AND (element.stereotype = "Pool" OR element.Stereotype = "Lane") then
			dim skipElement
			skipElement = false
			'check if it is local element
			if element.PackageID <> diagram.packageID  OR element.ParentID <> diagram.ParentID then
				'get the correct stereotype
				dim BPMNstereotype
				if element.Stereotype = "Pool" then
					BPMNstereotype = "BPMN2.0::Pool"
				else
					'Lane
					'check if the parent element of the lane is also on this diagram.
					'In that case we don't treat it now but it gets treated when we deal with the parent pool
					if isElementPresentOnDiagram (element.ParentID, diagram, diagramObjects, count) then
						'skip this one
						skipElement = true
					end if 
					BPMNstereotype = "BPMN2.0::Lane"
				end if
				if not skipElement then
					'make a new lane of pool under the owner element or owner package
					dim localActivityPartition as EA.Element
					if diagram.ParentID > 0 then
						set ownerElement = Repository.GetElementByID(diagram.ParentID)
						set localActivityPartition = ownerElement.Elements.AddNew("",BPMNstereotype)
					else
						set ownerPackage = Repository.GetPackageByID(diagram.PackageID)
						set localActivityPartition = ownerPackage.Elements.AddNew("",BPMNstereotype)
					end if
					'check if callingActivity was created
					if not localActivityPartition is Nothing then
						'copy the classifierID
						if element.ClassfierID > 0 then
							localActivityPartition.ClassfierID = element.ClassfierID
						else
							'copy the name if no classifierID present
							localActivityPartition.Name = element.Name
						end if
						'save local element
						localActivityPartition.Update
						'replace element in diagram
						updateDiagramObjectSQL = "update t_diagramobjects set object_id = "& localActivityPartition.ElementID &" where Diagram_ID = " & diagramObject.DiagramID & " and Object_ID = " & element.ElementID
						Repository.Execute updateDiagramObjectSQL
						'set the relations
						set diagram = Repository.GetDiagramByID(diagram.DiagramID)
						setActionRelations localActivityPartition, element, diagram, diagramObjects, count
						'copy owned lanes
						dim ownedLane as EA.Element
						for each ownedLane in element.Elements
							'check if the owned lane is shown on this diagram
							
							' get the diagramObject for the owned lane
							dim laneDiagramObject
							set laneDiagramObject = getDiagramObjectFromArray(ownedLane.elementID, diagramObjects, count)
							if not laneDiagramObject is nothing AND ownedLane.Stereotype = "Lane" then
								'if yes then make a new embedded elementin the callingActivity
								dim newOnwedLane as EA.Element
								set newOnwedLane = localActivityPartition.Elements.AddNew("","BPMN2.0::Lane")
								'copy the classifierID
								if ownedLane.ClassfierID > 0 then
									newOnwedLane.ClassfierID = ownedLane.ClassfierID
								else
									'copy the name if no classifierID present
									newOnwedLane.Name = ownedLane.Name
								end if
								'save new owned lane
								newOnwedLane.Update
								'replace element in diagram
								updateDiagramObjectSQL = "update t_diagramobjects set object_id = "& newOnwedLane.ElementID &" where Diagram_ID = " & diagramObject.DiagramID & " and Object_ID = " & ownedLane.ElementID
								Repository.Execute updateDiagramObjectSQL
								'set the relations
								set diagram = Repository.GetDiagramByID(diagram.DiagramID)
								setActionRelations newOnwedLane, ownedLane, diagram, diagramObjects, count
							end if
						next
					end if
				end if
			end if
		elseif element.Type = "Activity" then 
			'check if it is local activity
			if element.PackageID <> diagram.packageID  OR element.ParentID <> diagram.ParentID then
				'Make a new Activity for this activity
				dim callingActivity as EA.Element
				if diagram.ParentID > 0 then
					set ownerElement = Repository.GetElementByID(diagram.ParentID)
					set callingActivity = ownerElement.Elements.AddNew("","BPMN2.0::Activity")
				else
					set ownerPackage = Repository.GetPackageByID(diagram.PackageID)
					set callingActivity = ownerPackage.Elements.AddNew("","BPMN2.0::Activity")
				end if
				'check if callingActivity was created
				if not callingActivity is Nothing then
					makeCallingActivity callingActivity, element
					'set the element of the diagramObject to the new action
					updateDiagramObjectSQL = "update t_diagramobjects set object_id = "& callingActivity.ElementID &" where Diagram_ID = " & diagramObject.DiagramID & " and Object_ID = " & element.ElementID
					Repository.Execute updateDiagramObjectSQL
					'synchronize						
					synchronizeElement callingActivity
					Repository.AdviseElementChange(callingActivity.ElementID)
					'set the relations
					set diagram = Repository.GetDiagramByID(diagram.DiagramID)
					setActionRelations callingActivity, element, diagram, diagramObjects, count
					'copy embedded elements
					dim embeddedElement as EA.Element
					for each embeddedElement in element.EmbeddedElements
						'check if the embedded element is shown on this diagram
						
						' get the diagramObject for the embedded element
						dim embeddedDiagramObject
						set embeddedDiagramObject = getDiagramObjectFromArray(embeddedElement.elementID, diagramObjects, count)
						if not embeddedDiagramObject is nothing then
							'if yes then make a new embedded elementin the callingActivity
							dim newEmbeddedElement as EA.Element
							set newEmbeddedElement = callingActivity.EmbeddedElements.AddNew("","ObjectNode")
							newEmbeddedElement.Name = embeddedElement.Name
							newEmbeddedElement.Stereotype = "IntermediateEvent"
							newEmbeddedElement.Update()
							newEmbeddedElement.SynchTaggedValues "BPMN2.0","IntermediateEvent"
							newEmbeddedElement.TaggedValues.Refresh
							'Copy tagged values
							copyTaggedValuesValues embeddedElement, newEmbeddedElement
							' set the element id of the diagramobject to the new embedded element
							'embeddedDiagramObject.ElementID = newEmbeddedElement.ElementID
							'embeddedDiagramObject.Update
							'for some reason the update doesn't want to work. so we do it the hard way
							dim updateEmbeddedDiagramObjectSQL
							updateEmbeddedDiagramObjectSQL = "update t_diagramobjects set object_id = "& newEmbeddedElement.ElementID &" where Diagram_ID = " & diagramObject.DiagramID & " and Object_ID = " & embeddedElement.ElementID
							'Session.Output updateEmbeddedDiagramObjectSQL
							Repository.Execute updateEmbeddedDiagramObjectSQL
							'update relations of the embedded element
							setActionRelations newEmbeddedElement, embeddedElement, diagram, diagramObjects, count
						end if
					next
				end if
			end if
		end if
	next
	
	'reload the diagram to be able to click through
	Repository.ReloadDiagram(diagram.DiagramID)
	'tell the user we have finished
	MsgBox "Finished!"
end function

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


function setActionRelations(action, activity, diagram, diagramobjects, count)
	dim activityConnector as EA.Connector
	dim diagramLink as EA.DiagramLink
	dim hidden
	dim visible
	for each activityConnector in activity.Connectors
		hidden = false
		visible = false
		'check if connector visible in this diagram
		for each diagramLink in diagram.DiagramLinks
			if diagramLink.ConnectorID = activityConnector.ConnectorID then
				hidden = diagramLink.IsHidden
			end if
		next
		if not hidden then
			'check if the other element is visible on the diagram
			if activityConnector.ClientID = activity.ElementID then
				visible = isElementPresentOnDiagram(activityConnector.SupplierID, diagram, diagramobjects, count)
			else
				visible = isElementPresentOnDiagram(activityConnector.ClientID, diagram, diagramobjects, count) 
			end if
			if visible then
				'check if the connector is shown on other diagrams
				' if not we can re-use it, else we need to copy it
				if isLinkVisibleOnOtherDiagrams (activityConnector,diagram)  then
					'copy connector
					set activityConnector = copyConnector(action, activityConnector)	
				end if
				'set relations to action iso activity
				if activityConnector.ClientID = activity.ElementID then
					activityConnector.ClientID = action.ElementID
					activityConnector.Update
				end if
				if activityConnector.SupplierID = activity.ElementID then
					activityConnector.SupplierID = action.ElementID
					activityConnector.Update
				end if
			end if
		end if
	next
end function

function isLinkVisibleOnOtherDiagrams (connector, diagram)
	isLinkVisibleOnOtherDiagrams = false
	dim getOtherDiagramsSQL
	getOtherDiagramsSQL = "select dl.DiagramID from t_diagramlinks dl " _
						& " where dl.ConnectorID = " & connector.ConnectorID _
						& " and dl.DiagramID <> " & diagram.DiagramID _
						& " and dl.Hidden = 0"
	dim result 
	result = Repository.SQLQuery(getOtherDiagramsSQL)
	if InStr(result, "DiagramID") > 0 then
		isLinkVisibleOnOtherDiagrams = true
	end if
end function

function isElementPresentOnDiagram (elementID, diagram, diagramobjects, count)
	if not getDiagramObjectFromArray(elementID, diagramobjects, count) is nothing then
		isElementPresentOnDiagram = true
	else 
		isElementPresentOnDiagram = false
	end if
end function

'copy the given connector on this element
function copyConnector(element, connector)
	'@STEREO;Name=SequenceFlow;GUID={D48F475E-6647-4e93-9439-753FFCB06902};FQName=BPMN2.0::SequenceFlow;@ENDSTEREO;
	'first figure out the type of the connector.
	dim connectorTypeSQL
	dim connectorTypeResult
	connectorTypeSQL = "select x.description from (t_connector c "_
						& " left join t_xref x on x.Client = c.ea_guid) "_
						& " where c.Connector_ID = " & connector.ConnectorID _
						&" and x.Name = 'Stereotypes' " 
	Session.Output connectorTypeSQL
	connectorTypeResult = Repository.SQLQuery(connectorTypeSQL)
	Session.Output connectorTypeResult
	dim beginFQName
	dim connectorType
	connectorType = connector.Type
	beginFQName = InStr(connectorTypeResult, "FQName=")
	if beginFQName > 0 then
		'get the type of the connector from the string
		dim endFQName
		beginFQName = beginFQName + len("FQName=")
		endFQName = InStr(beginFQName,connectorTypeResult, ";")
		if endFQName > beginFQName then
			connectorType = Mid(connectorTypeResult, beginFQName , endFQName - beginFQName)
			'debug
			Session.Output "connectorType: " & connectorType
		end if
	end if
	set copyConnector = element.Connectors.AddNew("",connectorType)
	'copy connectors attributes
	copyConnector.Name = connector.Name
	copyConnector.ClientID = connector.ClientID
	copyConnector.SupplierID = connector.SupplierID
	copyConnector.Direction = connector.Direction
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

function getDiagramObjectForElementID(elementID, diagram)
	set getDiagramObjectForElementID = nothing
	dim diagramObject as EA.DiagramObject
	for each diagramObject in diagram.DiagramObjects
		if diagramObject.ElementID = elementID then
			'MsgBox "diagramObject.ElementID = " & diagramObject.ElementID  & " elementID = "& elementID 
			set getDiagramObjectForElementID = diagramObject
			exit for
		end if
	next
end function

function getDiagramObjectFromArray(elementID, diagramObjects, count)
	set getDiagramObjectFromArray = nothing
	dim i
	for i = 1 to count
		if diagramObjects(i).ElementID = elementID then
			set getDiagramObjectFromArray = diagramObjects(i)
			exit for
		end if
	next
end function