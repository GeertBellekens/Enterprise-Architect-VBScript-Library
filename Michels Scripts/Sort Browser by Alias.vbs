'[group=Michels Scripts]
option explicit
!INC Local Scripts.EAConstants-VBScript
!INC EAScriptLib.VBScript-Logging
'
' Script Name: SortbyAlias 
' Author: Guillaume FINANCE, guillaume[at]umlchannel.com
' Purpose: Sort elements contained in the selected package from the Project Browser by the Alias name
' Date: 03/04/2014
'
sub main
 OnProjectBrowserScript
end sub

sub SortDictionary (objDict)
   ' constants
   Const dictKey  = 1
   Const dictItem = 2
   ' variables
   Dim strDict()
   Dim objKey
   Dim strKey,strItem
   Dim X,Y,Z
   ' get the dictionary count
   Z = objDict.Count 
   ' sorting needs more than one item
   If Z > 1 Then
     ' create an array to store dictionary information
     ReDim strDict(Z,2)
     X = 0
     ' populate the string array
     For Each objKey In objDict
         strDict(X,dictKey)  = CStr(objKey)
         strDict(X,dictItem) = CStr(objDict(objKey))
         X = X + 1
     Next 
     ' perform a a shell sort of the string array
     For X = 0 To (Z - 2)
       For Y = X To (Z - 1)
         If StrComp(strDict(X,1),strDict(Y,1),vbTextCompare) > 0 Then
             strKey  = strDict(X,dictKey)
             strItem = strDict(X,dictItem)
             strDict(X,dictKey)  = strDict(Y,dictKey)
             strDict(X,dictItem) = strDict(Y,dictItem)
             strDict(Y,dictKey)  = strKey
             strDict(Y,dictItem) = strItem
         End If
       Next
     Next
     ' erase the contents of the dictionary object
     objDict.RemoveAll
     ' repopulate the dictionary with the sorted information
     For X = 0 To (Z - 1)
       objDict.Add strDict(X,dictKey), strDict(X,dictItem)
     Next
 ' sort the package elements based on the new sorting order
 dim newOrder
 newOrder = 0
 dim theItem
 dim eaelement
  for each objKey in objDict
  theItem = objDict.Item(objKey)
  Set eaelement = Repository.GetElementByGuid(theItem)
  'change the position of the element in the package to the new sorting order value
  eaelement.TreePos = CLng(newOrder)
  eaelement.Update()
  newOrder = newOrder + 1
 next
   end if
end sub

sub sortElementsbyAlias (selectedPackage)
 LOGInfo("Processing selected package " & selectedPackage.Name)
 dim elements as EA.Collection
 dim i
 dim processedElements
 set processedElements = CreateObject( "Scripting.Dictionary" )
 set elements = selectedPackage.Elements  
 for i = 0 to elements.Count - 1
   dim currentElement as EA.Element   
   set currentElement = elements.GetAt( i )
    LOGInfo("Processing " & currentElement.Type & " no " & i & " with alias " & currentElement.Alias & "(" &  currentElement.ElementGUID & ")")
    processedElements.Add currentElement.Alias, currentElement.ElementGUID
 next
 LOGInfo("Sorting package elements")
 SortDictionary processedElements
end sub
'
' Project Browser Script main function
'
sub OnProjectBrowserScript()
 Repository.ClearOutput "Script"
 LOGInfo( "Starting SortbyAlias script" )
 LOGInfo( "==============================" )
 ' Get the type of element selected in the Project Browser
 dim treeSelectedType
 treeSelectedType = Repository.GetTreeSelectedItemType()
 select case treeSelectedType
   case otPackage
   ' Code for when a package is selected
   dim thePackage as EA.Package
   set thePackage = Repository.GetTreeSelectedObject()
   sortElementsbyAlias thePackage
   Repository.RefreshModelView (thePackage.PackageID)
  case else
   ' Error message
   Session.Prompt "This script does not support items of this type.", promptOK
 end select
end sub
OnProjectBrowserScript
