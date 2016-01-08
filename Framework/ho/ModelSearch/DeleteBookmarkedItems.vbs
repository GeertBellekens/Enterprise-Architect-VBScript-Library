option explicit
'[path=\Framework\ho\ModelSearch]
'[group=hoModelSearch]
!INC Local Scripts.EAConstants-VBScript
!INC ho.Bookmark
!INC ho.Delete

'------------------------------------------------------------------------------------
' Delete bookmarked elements 
'------------------------------------------------------------------------------------
' Items to delete:
' - Elements
' Show all bookmarked item by standard search/find:
' - Diagrams, Find bookmarked Elements (Version 12.1)
'----------------- 
'Procedure:
'1. Bookmark the elements you want to handle
'2. Check if correct (Standard Search):  
'   Diagrams, Find bookmarked Elements (Version 12.1)
'3. Search:
'   Right Click, Scripts, DeleteBookmarkedItems
'4. Output in System Output, Tab Script
'5. Sometimes you have to reload Project Browser
'   Project Browser:
'   Select root package
'   Right Click, File, Reload Project  (12.1, other EA versions may differ)
'------------------
'Prerequisites:
'- Script is in ModelSearch Group
'
sub OnProjectBrowserScript()
	
	Dim promptResult
	Dim count
	Dim lGuids  ' List of guids <System.Collections.ArrayList>
	Set lGuids = getBookmarks()
    If lGuids.Count > 0 then	
		promptResult = Session.Prompt("Do you really want to delete " + CStr(lGuids.Count) + " bookmarked items?", promptYESNO)
		if promptResult = resultYes Then
			count = deleteBookmarkedItems()
			Session.Output CStr(count) + " bookmarked items deleted!!"
		End If
	else
		Session.Output "No elements bookmarked, break!"
	end if
	
end sub

OnProjectBrowserScript