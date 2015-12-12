'[path=\Projects\Project A\Old Scripts]
'[group=Old Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
' Project Browser Script main function
'
sub OnProjectBrowserScript()
	
	' Get the type of element selected in the Project Browser
	dim treeSelectedType
	treeSelectedType = Repository.GetTreeSelectedItemType()
	dim selectedElement as EA.Element
	set selectedElement = nothing
	' Handling Code: Uncomment any types you wish this script to support
	' NOTE: You can toggle comments on multiple lines that are currently
	' selected with [CTRL]+[SHIFT]+[C].
	select case treeSelectedType
	
		case otElement
			' Code for when an element is selected
			set selectedElement = Repository.GetTreeSelectedObject()
					
		case otPackage
			' Code for when a package is selected
			dim thePackage as EA.Package
			set thePackage = Repository.GetTreeSelectedObject()
			set selectedElement = thePackage.Element
			
'		case otDiagram
'			' Code for when a diagram is selected
'			dim theDiagram as EA.Diagram
'			set theDiagram = Repository.GetTreeSelectedObject()
			
		case otAttribute
			' Code for when an attribute is selected
			dim theAttribute as EA.Attribute
			set theAttribute = Repository.GetTreeSelectedObject()
			set selectedElement = Repository.GetElementByID(theAttribute.ParentID)
			
		case otMethod
			' Code for when a method is selected
			dim theMethod as EA.Method
			set theMethod = Repository.GetTreeSelectedObject()
			set selectedElement = Repository.GetElementByID(theMethod.ParentID)
		
		case else
			' Error message
			Session.Prompt "This script does not support items of this type.", promptOK
			
	end select
	if not selectedElement is nothing then
		Repository.RunModelSearch "Diagrams By ElementName",selectedElement.ElementGUID,"",""
	end if
	
end sub

OnProjectBrowserScript