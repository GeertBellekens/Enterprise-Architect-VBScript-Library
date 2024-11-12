'[path=\Projects\Project A\Template fragments]
'[group=Template fragments]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: LinkedDocument
' Author: Geert Bellekens
' Purpose: Gets the linked document for the given objectID and returns it as an RTF string. 
'   To be used in document generation as a Document Script.
'       Call as getLinkedDocument(#OBJECTID#)
' Date: 2018-04-25
'
function getLinkedDocument(objectID)
 dim element as EA.Element
 set element = Repository.GetElementByID(objectID)
 getLinkedDocument = element.GetLinkedDocument
end function

'test
'Session.Output getLinkedDocument(435910)