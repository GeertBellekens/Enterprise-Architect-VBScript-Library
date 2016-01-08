'[path=\Framework\ho]
'[group=ho]
'Function:  Delete copied/bookmarked elements 
'File:      Delete.vbs
'Author:    Helmut Ortmann
'Date: 2015-12-29
!INC Local Scripts.EAConstants-VBScript
!INC ho.Clipboard
!INC ho.Bookmark


sub testDelete
Session.output "Test function 'delete' started"
Session.output "Test function 'delete' finished"
end sub

'------------------------------------------------------
' Delete all items from clipboard 
'------------------------------------------------------
' Precondition:
' - Each raw of Clipboard shall contain a GUID to identify the item to delete
' - Supported element types
' -- Elements
' -- Packages
' -- Diagrams
'
'Possible select statement:
' select ea_guid, name, object_Type, "Element" from t_object where name like "XX*" UNION
' select ea_guid, name, diagram_Type, "Diagram" from t_diagram where name like "XX*" UNION
' select ea_guid, name, "Package", "Package" from t_package where name like "XX*"
Function deleteClipboardItems()
	Dim lGuids ' List of GUIDs in "System.Collections.ArrayList"
   
	Set lGuids = getGuidsFromClipboard()
    deleteClipboardItems = deleteByGuids(lGuids)
End Function

'------------------------------------------------------
' Delete all bookmarked items to the passed package
'------------------------------------------------------
' Precondition:
' - Supported element types
' -- Elements
' -- Packages
' -- Diagrams
'
Function deleteBookmarkedItems()
	Dim lGuids ' List of GUIDs in "System.Collections.ArrayList"
   
	Set lGuids = getBookmarks()
    deleteBookmarkedItems = deleteByGuids(lGuids)
End function





Function deleteByGuids(lGuids)
    Dim count
	Dim i
	Dim itemName
	Dim itemPkgId
	Dim itemType
	Dim guid
	Dim itemId
	
	Dim eaPkg As EA.Package
	Dim eaDia As EA.Diagram
	Dim eaEl As EA.Element
	Dim col As EA.Collection
	
	' List of packages to update
	Dim dictSrcPkg
	Set dictSrcPkg = CreateObject("Scripting.Dictionary")
	
	count = 0
	for each guid in lGuids
	    Set eaPkg = nothing
		Set eaDia = nothing
		Set eaEl = nothing

	
	    ' Element, an element might be a package and therefore isn't counted or logged to screen
		Set eaEl = Repository.GetElementByGuid(guid)
		if not eaEl is nothing then
		  itemPkgId = eaEl.PackageID
		  itemName = eaEl.Name
		  itemType = eaEl.Type
		  itemId = eaEl.ElementID

		  ' delete element
		  Set eaPkg = Repository.GetPackageByID(itemPkgId)
		  Set col = eaPkg.Elements
		  i = 0
		  For each eaEl in col
			if eaEl.ElementID = itemId then
				col.Delete(i)
			End if
			i = i + 1
		  Next
		  
		  
		  ' Remeber package to later refresh/update package, posible EA error
		  if not dictSrcPkg.Exists(itemPkgId) then
				dictSrcPkg.Add itemPkgId, 0 
		  end if

		  ' Element possibly is a package, only count during copying the package
		  if itemType <> "Package" Then
			Session.Output "Delete Element '" + itemName
			count = count + 1
		  End if
		End if
		
		' Package
		Set eaPkg = Repository.GetPackageByGuid(guid)
		if not eaPkg is nothing then
		  itemPkgId = eaPkg.ParentID
		  itemName = eaPkg.Name
		  itemId = eaPkg.PackageID
		  
		  ' delete Package
		  Set eaPkg = Repository.GetPackageByID(itemPkgId)
		  Set col = eaPkg.Packages
		  i = 0
		  For each eaPkg in col
			if eaPkg.PackageID = itemId then
				col.Delete(i)
			End if
			i = i + 1
		  Next

		  ' Remeber package to later refresh/update package, posible EA error
		  'if not dictSrcPkg.Exists(itemPkgId) then
		  '		dictSrcPkg.Add itemPkgId, 0 
		  'end if
		  Session.Output "Delete Package '" + itemName
		  count = count + 1
		End if
		
		' Diagram
		on error resume next
		Set eaDia = Repository.GetDiagramByGuid(guid)
		if Err.Number = 0 then
			if not eaDia is nothing then
			  itemPkgId = eaDia.PackageID
		      itemName = eaDia.Name
			  itemId = eaDia.DiagramID
			  
			  ' delete diagram
			  Set eaPkg = Repository.GetPackageByID(itemPkgId)
			  Set col = eaPkg.Diagrams
			  i = 0
			  For each eaDia in col
				if eaDia.DiagramID = itemId then
					col.Delete(i)
				End if
				i = i + 1
			  Next
			  
			  ' Remeber package to later refresh them, posible EA error
			  if not dictSrcPkg.Exists(pkgId) then
			   	   dictSrcPkg.Add itemPkgId, 0 
			  end if

			  Session.Output "Delete Diagram '" + itemName  
			  count = count + 1
			End if
		End if
		On Error GoTo 0
		
	Next
	
	' Refresh all source packages
    For each itemPkgId in dictSrcPkg
	    Set eaPkg = Repository.GetPackageByID(itemPkgId)
		if not eaPkg is nothing then
			eaPkg.Update()
			Repository.RefreshModelView(itemPkgId)	
		End if
	Next
	
	
	deleteByGuids = count
End Function	


'Test
'testDelete