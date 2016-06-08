'[path=\Projects\Project A\Diagram Group]
'[group=Diagram Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.Util

' Script Name: Synchronize
' Author: Geert Bellekens
' Purpose: Synchronises the names of the selected objects or BPMN Activities with their classifier/called activity ref.
' Will also set the composite diagram to that of the classifier/ActivityRef in order to facilitate click-through
' Date: 27/03/2015
'

'
' Diagram Script main function
'


sub OnDiagramScript()
	' Get a reference to the current diagram
	dim currentDiagram as EA.Diagram
	set currentDiagram = Repository.GetCurrentDiagram()
	
	if not currentDiagram is nothing then
		' Get a reference to any selected connector/objects
		'save the diagram before anything else
		Repository.SaveDiagram currentDiagram.DiagramID
		dim selectedObjects as EA.Collection
		set selectedObjects = currentDiagram.SelectedObjects
		'if nothing is selected then we do synchronize on all objects
		if selectedObjects.Count < 1 then
			set selectedObjects = currentDiagram.DiagramObjects
		end if
		dim selectedObject as EA.DiagramObject
		for each selectedObject in selectedObjects
			synchronizeObjectNames selectedObject, currentDiagram
		next
		'reload the diagram to be able to click through
		Repository.ReloadDiagram(currentDiagram.DiagramID)
	else
		Session.Prompt "This script requires a diagram to be visible", promptOK
	end if
end sub

'Sets the object name to that of the classifier, and set the composite diagram for objects
function synchronizeObjectNames(diagramObject, diagram)
	dim element as EA.Element
	set element = Repository.GetElementByID(diagramObject.ElementID)
	synchronizeElement element
	'set default size for message objects
	if element.Type = "Object" AND (element.Stereotype = "Message" or element.Stereotype = "FIS") then
		diagramObject.bottom = diagramObject.top - 25
		diagramObject.right = diagramObject.left + 40
		setFont diagramObject
		diagramObject.Update
'		copyMessageDirection(element)
	end if
	'check if it is local activity
	if element.Type = "Activity" and element.Stereotype = "Activity" and (element.PackageID <> diagram.packageID ) then
		'Make a new Activity for this activity
		dim callingActivity as EA.Element
		dim ownerElement as EA.Element
		dim ownerPackage as EA.Package
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
			dim updateDiagramObjectSQL
			'set the element of the diagramObject to the new action
			updateDiagramObjectSQL = "update t_diagramobjects set object_id = "& callingActivity.ElementID &" where Diagram_ID = " & diagramObject.DiagramID & " and Object_ID = " & element.ElementID
			Repository.Execute updateDiagramObjectSQL
			'synchronize						
			synchronizeElement callingActivity
			Repository.AdviseElementChange(callingActivity.ElementID)
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
				end if
			next
		end if
	end if
end function

function getOrCreateTaggedValue(element, taggedValueName)
		'add tagged value if not exists yet
		dim taggedValue as EA.TaggedValue
		dim taggedValueExists
		taggedValueExists = false
		for each taggedValue in element.TaggedValues
			if taggedValue.Name = taggedValueName then
				taggedValueExists = true
				exit for
			end if
		next
		'create tagged value is not existing yet
		if taggedValueExists = false then
			set taggedValue = element.TaggedValues.AddNew(taggedValueName,"")
			taggedValue.Update
		end if
		set getOrCreateTaggedValue = taggedValue
end function

function copyMessageDirection(message)
	dim tv as EA.TaggedValue
	dim messageClassifier as EA.Element
	'get the classifier message
	if message.ClassifierID > 0 then
		set messageClassifier = Repository.GetElementByID(message.ClassifierID)
		if not messageClassifier is nothing then
			dim parentDirection
			parentDirection = getDirection(messageClassifier)
			if parentDirection = "In" or parentDirection = "Out" then
				setDirection message, parentDirection
			end if
		end if
	end if
end function

function getDirection(message)
	dim tv as EA.TaggedValue
	getDirection = ""
	for each tv in message.TaggedValues
		if tv.Name = "Atrias::Direction" then
			getDirection = tv.Value 
			exit for
		end if
	next
end function

function setDirection(message, value)
	dim tv as EA.TaggedValue
	set tv = getOrCreateTaggedValue(message, "Atrias::Direction")
	tv.Value = value
	tv.Update
end function

function setFont(diagramObject)
	dim styleParts
	styleParts = Split (diagramObject.Style , ";") 
	dim i
	dim stylepart
	dim fontpart 
	fontpart = "font=Arial Narrow"
	dim fontSet
	fontSet = false
	dim sizePart
	sizePart = "fontsz=120"
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

OnDiagramScript
'test
sub test
	' Get a reference to the current diagram
	dim currentDiagram as EA.Diagram
	set currentDiagram = Repository.GetDiagramByGuid("{4FBDF6B6-B684-4b40-839F-D2A65A9F5418}")
	
	if not currentDiagram is nothing then
		' Get a reference to any selected connector/objects
		'save the diagram before anything else
		Repository.SaveDiagram currentDiagram.DiagramID
		dim selectedObjects as EA.Collection
		set selectedObjects = currentDiagram.SelectedObjects
		'if nothing is selected then we do synchronize on all objects
		if selectedObjects.Count < 1 then
			set selectedObjects = currentDiagram.DiagramObjects
		end if
		dim selectedObject as EA.DiagramObject
		for each selectedObject in selectedObjects
			synchronizeObjectNames selectedObject, currentDiagram
		next
		'reload the diagram to be able to click through
		Repository.ReloadDiagram(currentDiagram.DiagramID)
	else
		Session.Prompt "This script requires a diagram to be visible", promptOK
	end if
end sub