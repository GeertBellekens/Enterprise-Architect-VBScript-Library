option explicit
'[path=\Framework\ho\ModelSearch]
'[group=hoModelSearch]
!INC Local Scripts.EAConstants-VBScript
!INC ho.Clipboard
!INC ho.Delete

'
'------------------------------------------------------------------------------------
' Delete clipboarded items (Model Search Group)
'------------------------------------------------------------------------------------
' Items to delete:
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
'3. Search:
'   Right Click, Scripts, DeleteClipboardedItems
'5. Output in System Output, Tab Script
'6. Sometimes you have to reload Project Browser
'   Project Browser:
'   Select root package
'   Right Click, File, Reload Project  (12.1, other EA versions may differ)
'------------------
'Prerequisites:
'- Script has to be in ModelSearch Group
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
sub OnProjectBrowserScript()
    Dim promptResult
	Dim count
	Dim lGuids  ' List of guids <System.Collections.ArrayList>
	Set lGuids = getGuidsFromClipboard()
    If lGuids.Count > 0 then	
		promptResult = Session.Prompt("Do you really want to delete " + CStr(lGuids.Count) + " clipboarded items?", promptYESNO)
		if promptResult = resultYes Then
			count = deleteClipboardItems()
			Session.Output CStr(count) + " clipboarded items deleted!!"
		End If
	else
		Session.Output "No elements on Clipboard, break!"
	end if
end sub

OnProjectBrowserScript