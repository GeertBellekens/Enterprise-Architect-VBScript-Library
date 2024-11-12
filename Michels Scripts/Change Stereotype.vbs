'[group=Michels Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC EAScriptLib.VBScript-Logging



sub ChangeStereoType()
 dim elements as EA.Collection
 Dim i
 Dim stereotype
 Dim currentElement as EA.Element
 set elements = Repository.GetTreeSelectedElements()
 'set elements = GetTreeSelectedElements()
 stereotype = InputBox("Naar welk stereotype wenst u dit element te transformeren?", "Kies stereotype", "Zin")
 Session.Output elements.Count
 For i = 0 To elements.Count - 1
  set currentElement = elements.GetAt(i)
  Session.Output currentElement.Stereotype
  If currentElement.Stereotype <> stereotype then
   LOGTrace("Changing stereotype of element " & currentElement.ElementGUID)
   currentElement.Stereotype = stereotype
   currentElement.Update
  End if
  
 Next
 
end sub

ChangeStereoType
