option explicit
'[path=\Framework\ho\ProjectBrowser]
'[group=hoProjectBrowser]
!INC Local Scripts.EAConstants-VBScript
!INC ho.Move

'
'-------------------------------------------------------------------
' Move clipboarded items to selected package (project browser group)
'-------------------------------------------------------------------
' Items to move:
' - Elements
' - Diagrams
' - Packages
'----------------- 
'Procedure:
'1. Run SQL Model Search to find your items to move 
'    (Example: hoClipboard)
'2. Search:            
'   Copy the required rows to Clipboard
'   (e.g. CTR+A,CTRL+C for all)
'3. Project Browser:
'   Select the the target package to move to
'4. Project Browser:
'   Right Click, Scripts, MoveClipboardedToSelectedPackage
'5. Output in System Output, Tab Script
'6. Sometimes you have to reload Project Browser
'   '- Script has to be in ProjectBrowser Group:
'   Select root package
'   Right Click, File, Reload Project  (12.1, other EA versions may differ)
'------------------
'Prerequisites:
'- Projectbrowser Group
'- Each raw in Search result contains a GUID of the item to move
'-- Elements to move
'-- Diagrams to move
'-- Packages to move  
'------------------- 
' Example SQL Model Search
' select ea_guid, name, object_Type, "Element" from t_object where name like "XX*" UNION
 'select ea_guid, name, diagram_Type, "Diagram" from t_diagram where name like "XX*" UNION
' select ea_guid, name, "Package", "Package" from t_package where name like "XX*"
'
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
'			dim theElement as EA.Element
'			set theElement = Repository.GetTreeSelectedObject()
'					
		case otPackage
			' Code for when a package is selected
			dim count
			dim thePackage as EA.Package
			set thePackage = Repository.GetTreeSelectedObject()
			count = moveClipboardItems(thePackage)
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