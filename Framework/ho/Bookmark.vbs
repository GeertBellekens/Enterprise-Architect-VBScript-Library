'[path=\Framework\ho]
'[group=ho]
'Function:  Handle Bookmarks (package & elements, diagrams planned)
'File:      Bookmarks.vbs
'Author:    Helmut Ortmann
'Date: 2015-12-29
!INC Utils.Include

'----------------------------------------------------------------------------------------------------------
'Supported function:
'-------------------
'delAllBookmarks()                Delete all bookmarks
'delBookmarks(lGuid)              Delete bookmarks according to the passed list of GUIDs (System.Collections.ArrayList)
'lGuid=getBookmarks()             Get list of bookmarks (System.Collections.ArrayList)
'setBookmarks(lGuid)              Set bookmarks according to the passed list of GUIDs (System.Collections.ArrayList)
'updateBookmarks(lGuid, value)    Update bookmarks according to the passed list of GUIDs (System.Collections.ArrayList)
'                                 Delete: value=0; Set: value=1
'------------------
' Remarks:
' Bookmarks are EA Elements with the field Tagged=1
' Only Elements and packages can be bookmarked, no Digarams!

'----------------------------------------------------------------
' Testscript: Modify/Delete/Report selected elements from Search Window
'----------------------------------------------------------------
' Example SQL Search: 
' Make sure to use GUIDs
' SELECT t_object.ea_guid AS CLASSGUID, 
'        t_object.Object_Type AS CLASSTYPE, t_object.Name, t_object.ea_guid, t_object.Object_Type
'    from t_object
' Test:
' 1. Run query
' 2. Select rows, copy them to clipboard
' 3. Run script

sub testBookmark
  Dim guid, el
  Dim lGuids
  set lGuids = getBookmarks()
  Session.Output "--------- Start test bookmark --------------"
  for each guid in lGuids
	' here you may put in your action according to the meta type
	Set el = Repository.GetElementByGuid(guid)
	if not el is nothing then
		Select Case el.MetaType
			Case "Package"
				Session.Output "  Package:" + Left(el.Name,30) + " " + el.MetaType+ " " + guid
			Case "Class"
				Session.Output "  Class  :" + Left(el.Name,30) + " " + guid
			Case Else
				Session.Output  "  " + el.MetaType + ":" + Left(el.Name,30) + " " + guid
		End Select
	end if
  Next
  Session.Output "  Found Bookmarks: " + CStr(lGuids.Count)
  
  ' delete all bookmarks
  delAllBookmarks()
  Session.Output "All bookmarks deleted" 
  
  ' reread all bookmarksd
  set lGuids = getBookmarks()
  Session.Output "  Found Bookmarks: " + CStr(lGuids.Count)
  
  ' create bookmarks
  setBookmarks(lGuids)
  Session.Output "All bookmarks restored" 
  
  ' reread all bookmarksd
  set lGuids = getBookmarks()
  Session.Output "  Found Bookmarks: " + CStr(lGuids.Count)
  
  Session.Output "--------- End test bookmark --------------"
  
end sub

'--------------------------------------------------------------
' delAllBookmarks
'--------------------------------------------------------------
Function delAllBookmarks()
    Dim sqlDelAll
	Dim queryResult
    sqlDelAll = "update t_object " & _
		            " set tagged = 0; " 
	sqlDelAll = "UPDATE t_object SET tagged=0"
    Repository.Execute sqlDelAll
End Function

'--------------------------------------------------------------
' delBookmarks(ArrayList<GUID>)
'--------------------------------------------------------------
Function delBookmarks(lGuid)
	updateBookmarks lGuid,0
End Function
'--------------------------------------------------------------
' setBookmarks ArrayList<GUID>
'--------------------------------------------------------------
Function setBookmarks(lGuid)
	updateBookmarks lGuid,1 
End Function

Function updateBookmarks(lGuid, value)
	Dim guids
	Dim guid
	Dim sqlSet
	
	' construct sql update 
	guids = ""
	for each guid in lGuid
		Session.Output "'"+guid+"'"
	    if len(guid) > 37 then
			guids = guids + "'" + guid + "',"
		End if
	Next
	if guid <> "" Then 
		' remove last comma
		if (len(guids) > 0) then
			guids = mid(guids, 1, len(guids) -1)
		End if
		sqlSet = "update t_object " & _
				 " set tagged=" & CStr(value) & _
				 " where ea_guid in (" + guids + ");"
		' Session.Output "Set bookmarks '" + sqlSet + "' "
		Repository.Execute sqlSet
	End if

End Function
'--------------------------------------------------------------
' ArrayList<GUID> getBookmarks
'--------------------------------------------------------------
' Prerequisition:
'
'
Function getBookmarks( )
		Dim sqlGet
		Dim resultArray
		Dim queryResult
		Dim lGuid
		Set lGuid = CreateObject("System.Collections.ArrayList")
		sqlGet = "select o.ea_guid " & _
		            " from t_object o " & _
					" where o.tagged = 1 "
        queryResult = Repository.SQLQuery(sqlGet)
		resultArray = convertQueryResultToArray(queryResult)
		dim i
		For i = LBound(resultArray) To UBound(resultArray) 
			if len(resultArray(i,0)) > 37 then 
				lGuid.Add resultArray(i,0)
			End if
		Next

		Set getBookmarks = lGuid
 End Function

' Call main function
'testBookmark