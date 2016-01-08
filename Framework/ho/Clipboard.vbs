'[path=\Framework\ho]
'[group=ho]
'Function:  Handle Cipboard, read GUIDs from Clipboard
'File:      Clipboard.vbs
'Author:    Helmut Ortmann
'Date: 2015-12-29
!INC Local Scripts.EAConstants-VBScript
!INC ho.StringFunctions    


'----------------------------------------------------------------------------------------------------------
'Supported function:
'-------------------
' lGuid = getGuidsFromClipboard()     Get list of GUIDs from vlipboard in the array of GUIDs of type (System.Collections.ArrayList)
' setTextInClipboard(text)            Copy text to Clipboard (On some Windows systems there is a IExplorer security warning)


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

sub testGetGuidsFromClipboard
  Dim GUID
  Dim lGuids
  Dim oType
  Dim el As EA.Element
  Dim dia As EA.Diagram
  ' get list of GUIDs, usually one per row
  set lGuids = getGuidsFromClipboard()
  
  for each guid in lGuids
	' Here you may put in your action according to the meta type
	' Element or package
	Set el = Repository.GetElementByGuid(guid)
	if not el is nothing then
		Select Case el.MetaType
			Case "Package"
				Session.Output "Package:" + Rpad(el.Name,"_",30) + " " + guid
			Case "Class"
				Session.Output "Class  :" + Rpad(el.Name,"_",30) + " " + guid
			Case Else
				Session.Output  el.MetaType + ":" + RPad(el.Name,"_",30) + " " + guid
		End Select
	else 
		Set dia = Repository.GetDiagramByGuid(guid)
		if not dia is nothing then
			Session.Output "Diagram:" + Rpad(dia.Name,"_",30) + " " + guid			
		end if
	
	end if
  Next
  Session.Output "Found GUIDs in clipboard:" + CStr(lGuids.Count)
  
  clipboard = getTextFromClipboard

  'Put something to clipboard
  ret = setTextInClipboard("Beatiful weather")
  

end sub


'--------------------------------------------------------------
' ArrayList<GUID> getGuidsFromClipboard
'--------------------------------------------------------------
' Prerequisition:
' - One GUID for each row is used
'
'
Function getGuidsFromClipboard( )
Dim objHtml
Dim re
Dim colMatch
Dim textClipboard
Dim lGuid


  ' Get Clipboard content of type text
  textClipboard = getTextFromClipboard()
  
  ' Extract GUIDs with regExp
  Set re = CreateObject("VBScript.RegExp")
  With re
	.Pattern = "{[0-9abcdefABCDEF-]+}"
	.Global = True
	.IgnoreCase = False
  End With
  Set colMatch = re.Execute(textClipboard)
  ' Collect Guids in ArrayList
  Set lGuid = CreateObject("System.Collections.ArrayList")
  for each colMatch in colMatch
	lGuid.Add colMatch.Value
  Next
  ' return ArrayList of GUIDs
  Set getGuidsFromClipboard = lGuid
End Function

Function getTextFromClipboard()
  Dim objHTML 
  Set objHTML = CreateObject("htmlfile")
  getTextFromClipboard = objHTML.ParentWindow.ClipboardData.GetData("Text")
End Function

' Note: IE gives a security warning. I haven' found another way to copy
' It's not that fast
Function setTextInClipboard(text)
	Set objIE = CreateObject("InternetExplorer.Application")
	objIE.Navigate("about:blank")
	objIE.document.parentwindow.clipboardData.SetData "text", text
	objIE.Quit
	
	
End Function

' test function, if useful remove tick mark
'testGetGuidsFromClipboard