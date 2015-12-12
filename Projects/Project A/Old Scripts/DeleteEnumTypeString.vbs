'[path=\Projects\Project A\Old Scripts]
'[group=Old Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
'
' Script Name: Delete Enum Type String
' Author: Tom Geerts
' Purpose: Deletes the "string" type of enumeration attributes, which should not have a type, because they are literal values.
' Use: Continued use if needed.
' Date: 01/07/2015
'
' Project Browser Script main function
'
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
'			dim enumElement as EA.Element
'			set enumElement = Repository.GetTreeSelectedObject()
'					
		case otPackage
			' Code for when a package is selected
			dim thePackage as EA.Package
			set thePackage = Repository.GetTreeSelectedObject()
'			
'		case otDiagram
'			' Code for when a diagram is selected
'			dim theDiagram as EA.Diagram
'			set theDiagram = Repository.GetTreeSelectedObject()
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
	'	dim enumElement as EA.Element

'		if enumElement.Type = "Enumeration" then
'			dim enumAttribute as EA.Attribute
'			for each enumAttribute in enumElement.Attributes
'				if enumAttribute.Type = "string" then
'					enumAttribute.Type = ""
'					enumAttribute.Update
'				end if
'			next
'		end if
'	next
	dim enumElement as EA.Element
	for each enumElement in thePackage.Elements
		if enumElement.Type = "Enumeration" then
			dim enumAttribute as EA.Attribute
			for each enumAttribute in enumElement.Attributes
				if enumAttribute.Type = "string" then
					enumAttribute.Type = ""
					enumAttribute.Update
				end if
			next
		end if
	next
msgbox "Script has finished"
end sub

OnProjectBrowserScript