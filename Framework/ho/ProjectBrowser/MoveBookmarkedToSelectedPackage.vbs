option explicit
'[path=\Framework\ho\ProjectBrowser]
'[group=hoProjectBrowser]
!INC Local Scripts.EAConstants-VBScript
!INC ho.Move

'
'-------------------------------------------------------------------
' Move bookmarked elements to selected package
'-------------------------------------------------------------------
' Items to move:
' - Elements
' Show all bookmarked item by standard search/find:
' - Diagrams, Find bookmarked Elements (Version 12.1)
'
'----------------- 
'Procedure:
'1. Bookmark the elements you want to handle
'2. Check if correct (Standard Search):  
'   Diagrams, Find bookmarked Elements (Version 12.1)
'3. Project Browser:
'   Select the the target package to move to
'4. Project Browser:
'   Right Click, Scripts, MoveBookmarkedToSelectedPackage
'5. Output in System Output, Tab Script
'6. Sometimes you have to reload Project Browser
'   Select root package
'   Right Click, File, Reload Project  (12.1, other EA versions may differ)
'------------------
'Prerequisites:
'- Script has to be in ProjectBrowser Group:
'
'
sub OnProjectBrowserScript()
	
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
		case otPackage
			' Code for when a package is selected
			dim count
			dim thePackage as EA.Package
			set thePackage = Repository.GetTreeSelectedObject()
			count = moveBookmarkedItems(thePackage)
			Session.Output CStr(count) + " elements moved to package '" + thePackage.Name + "'."
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
			Session.Prompt "This script only supports Packages.", promptOK
			
	end select
	
end sub

OnProjectBrowserScript