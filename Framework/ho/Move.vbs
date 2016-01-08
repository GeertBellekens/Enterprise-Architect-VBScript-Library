'[path=\Framework\ho]
'[group=ho]
'Function:  Move copied elements to a new position
'File:      Move.vbs
'Author:    Helmut Ortmann
'Date: 2015-12-29
!INC Local Scripts.EAConstants-VBScript
!INC ho.Clipboard
!INC ho.Bookmark


sub testMove
Session.output "Test function 'move' started"
Session.output "Test function 'move' finished"
end sub

'------------------------------------------------------
' Move all items from clipboard to the passed package
'------------------------------------------------------
' Precondition:
' - Each raw of Clipboard shall contain a GUID to identify the item to move
' - Supported element types
' -- Elements
' -- Packages
' -- Diagrams
'
'Possible select statement:
' select ea_guid, name, object_Type, "Element" from t_object where name like "XX*" UNION
' select ea_guid, name, diagram_Type, "Diagram" from t_diagram where name like "XX*" UNION
' select ea_guid, name, "Package", "Package" from t_package where name like "XX*"
Function moveClipboardItems(pkg)
	Dim lGuids ' List of GUIDs in "System.Collections.ArrayList"
   
	Set lGuids = getGuidsFromClipboard()
    moveClipboardItems = moveByGuids(lGuids, pkg)
End function

'------------------------------------------------------
' Move all bookmarked items to the passed package
'------------------------------------------------------
' Precondition:
' - Supported element types
' -- Elements
' -- Packages
' -- Diagrams
'
Function moveBookmarkedItems(pkg)
	Dim lGuids ' List of GUIDs in "System.Collections.ArrayList"
   
	Set lGuids = getBookmarks()
    moveBookmarkedItems = moveByGuids(lGuids, pkg)
End function





Function moveByGuids(lGuids, pkg)
    Dim count
	Dim pkgId
	Dim guid
	Dim trgtPkgId
	
	Dim trgtPkg As EA.Package
	Dim eaPkg As EA.Package
	Dim eaDia As EA.Diagram
	Dim eaEl As EA.Element
	
	Dim dictSrcPkg
	Set dictSrcPkg = CreateObject("Scripting.Dictionary")
	
	count = 0
	pkgId = 0
	Set trgtPkg = pkg 
	trgtPkgId =  trgtPkg.PackageID
	Session.Output "Started with package '" + trgtPkg.Name + "' "
	for each guid in lGuids
	    Set eaPkg = nothing
		Set eaDia = nothing
		Set eaEl = nothing
	
	    ' Element, an element might be a package and therefore isn't counted or logged to screen
		Set eaEl = Repository.GetElementByGuid(guid)
		if not eaEl is nothing then
		  eaEl.PackageID = trgtPkg.PackageID
		  eaEl.Update()
		  trgtPkg.Update()
		  ' Element possibly is a package, only count during copying the package
		  if eaEl.Type <> "Package" Then
			Session.Output "Moved Element '" + eaEl.Name  + ":" + eaEl.Type + "' to Package " + trgtPkg.Name 
			count = count + 1
		  End if
		End if
		
		' Package
		Set eaPkg = Repository.GetPackageByGuid(guid)
		if not eaPkg is nothing then
		  eaPkg.ParentID = trgtPkg.PackageID
		  eaPkg.Update()
		  Session.Output "Moved Package '" + eaPkg.Name + "' to Package " + trgtPkg.Name 
		  count = count + 1
		End if
		
		' Diagram
		on error resume next
		Set eaDia = Repository.GetDiagramByGuid(guid)
		if Err.Number = 0 then
			if not eaDia is nothing then
			  pkgId = eaDia.PackageID
			  eaDia.PackageID = trgtPkg.PackageID
			  eaDia.Update()
			  
			  ' Remeber package to later refresh them, posible EA error
			  if not dictSrcPkg.Exists(pkgId) then
					dictSrcPkg.Add pkgId, 0 
			  end if

			  Session.Output "Moved Diagram '" + eaDia.Name + "' to Package " + trgtPkg.Name 
			  count = count + 1
			End if
		End if
		On Error GoTo 0
		
	Next
    trgtPkg.Update()
	Repository.RefreshModelView(trgtPkgId)
	
	' Refresh all source packages
	For each pkgId in dictSrcPkg
		Repository.RefreshModelView(pkgId)	
	Next
	
	
	' select the target package
	Repository.ShowInProjectView(trgtPkg)
	
	moveByGuids = count
End Function	


'Test
'testMove