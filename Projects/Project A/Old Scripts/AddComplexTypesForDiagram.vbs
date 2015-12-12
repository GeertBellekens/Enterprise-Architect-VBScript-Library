'[path=\Projects\Project A\Old Scripts]
'[group=Old Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' This code has been included from the default Project Browser template.
' If you wish to modify this template, it is located in the Config\Script Templates
' directory of your EA install path.   
'
' Script Name: AddComplexTypesForDiagram 
' Purpose: on diagram, add Complex Types that are being referred to 
' Date: 2014-10-13
'
' Project Browser Script main function
'

' list of ID's (of ComplexTypes) already added to the diagram
dim newComplexTypesIds

' position for new object to be added
dim x
dim y

sub OnProjectBrowserScript()

	' Get the type of element selected in the Project Browser
	dim treeSelectedType
	treeSelectedType = Repository.GetTreeSelectedItemType()

	' Handling Code: Uncomment any types you wish this script to support
	' NOTE: You can toggle comments on multiple lines that are currently
	' selected with [CTRL]+[SHIFT]+[C].
	select case treeSelectedType
	
'		case otElement
'			' Code for when an element is selected
'			dim theElement as EA.Element
'			set theElement = Repository.GetTreeSelectedObject()
'					
'		case otPackage
'			' Code for when a package is selected
'			dim thePackage as EA.Package
'			set thePackage = Repository.GetTreeSelectedObject()
'			
		case otDiagram
			' Code for when a diagram is selected
			dim theDiagram as EA.Diagram
			set theDiagram = Repository.GetTreeSelectedObject()
			'Debug: Session.Output "Script started for " + theDiagram.Name
			AddComplexTypesForDiagram theDiagram
			Repository.ReloadDiagram theDiagram.DiagramID
			Session.Prompt "Script finished. Look at output tab for log details.", promptOK
'			
'		case otAttribute
'			' Code for when an attribute is selected
'			dim theAttribute as EA.Attribute
'			set theAttribute = Repository.GetTreeSelectedObject()
'			
'		case otMethod
'			' Code for when a method is selected
'			dim theMethod as EA.Method
'			set theMethod = Repository.GetTreeSelectedObject()
		
		case else
			' Error message
			Session.Prompt "This script does not support items of this type.", promptOK
			
	end select
	
end sub

' AddComplexTypesForDiagram
'   for each of the complex on the diagram, add ComplexTypes it has as attribute
sub AddComplexTypesForDiagram(theDiagram)
 	dim EADiagramElement as EA.DiagramObject
	dim EAElement as EA.Element
 
	'Debug: Session.Output "Treating diagram: " + theDiagram.Name
	
	' Reset array
	Set newComplexTypesIds = CreateObject("Scripting.Dictionary")
	' Initialize location for new objects
	x = 0
	y = 0
	
    ' Browse all elements of the diagram
    for each EADiagramElement In theDiagram.DiagramObjects
		set EAElement=Repository.GetElementByID(EADiagramElement.ElementID)
		AddDependentForObject theDiagram, EAElement
	next
end sub

' AddDependentForObject
'   for each ComplexType, add to diagram if not present yet
'   (recursive)
sub AddDependentForObject(theDiagram, theElement)
	dim EAAttribute as EA.Attribute
	dim classifierElement as EA.Element
 
	'Debug: Session.Output "Treating element: " + theElement.Name
	'Debug: Session.Output " on diagram " + theDiagram.Name
	if (theElement.Stereotype = "XSDcomplexType") then
		AddComplexTypeToDiagram theDiagram,theElement
		for each EAAttribute in theElement.Attributes
			if (EAAttribute.ClassifierID <> 0) then
				set classifierElement=Repository.GetElementByID(EAAttribute.ClassifierID)
				'Debug: Session.Output " Element: " & classifierElement.Name
				'Debug:	Session.Output "   type: " & classifierElement.StereoType
				if (classifierElement.Stereotype = "XSDcomplexType") then
					' Add complex types as dependent
					AddDependentForObject theDiagram, classifierElement
					'AddOrUpdateDependencyConnector theElement,EAAttribute,classifierElement
				end if
			end if
		next
		
	end if
end sub 

' AddComplexTypeToDiagram
'   add the ComplexType to the diagram
'   add also to list of already added ComplexTypes
sub AddComplexTypeToDiagram(theDiagram, theComplexType)
	Dim newDiagramObject as EA.DiagramObject
	Dim position
	if not onDiagram(theDiagram,theComplexType) then
		'Debug: Session.Output "Add element: " + theComplexType.Name + " to the diagram " + theDiagram.Name
		x = x + 25
		y = y + 25
		position = "l=" + CStr(x) + "t=" + CStr(y)
		set newDiagramObject = theDiagram.DiagramObjects.AddNew(position,"")
		If Not newDiagramObject Is Nothing Then
			newDiagramObject.ElementID = theComplexType.ElementID
			newDiagramObject.Style = "BCol=4443520" '4443520 = #43cd80 = SeaGreen3 (http://www.color-hex.com/color-names.html)
			newDiagramObject.Update
			' Add the connector to know attribute connector ID list
			newComplexTypesIds.Add theComplexType.ElementID,theComplexType.Name
		End If
	end if
end sub

' onDiagram
'   Check if ComplexType already on diagram
function onDiagram(theDiagram, theElement)
	dim EADiagramElement as EA.DiagramObject
	dim EAElement as EA.Element
	
	if newComplexTypesIds.Exists(theElement.ElementID) then
		'Debug: Session.Output "Already added to diagram: " + theElement.Name
		onDiagram = true
		Exit Function
	end if
	
	'Debug: Session.Output "Already on diagram? " + theElement.Name
	for each EADiagramElement In theDiagram.DiagramObjects
        set EAElement=Repository.GetElementByID(EADiagramElement.ElementID)
        if EAElement.ElementID = theElement.ElementID then	
			'Debug: Session.Output "Already on diagram: " + theElement.Name
			onDiagram = true
			Exit Function
		end if
    next
		
	'Debug: Session.Output "Not yet on diagram: " + theElement.Name
	onDiagram = false
end function

OnProjectBrowserScript