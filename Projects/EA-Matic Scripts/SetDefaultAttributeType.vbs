'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: SetDefaultAttributeType
' Author: Geert Bellekens
' Purpose: Set the default attribute type to None
' Date: 2020-10-12
'
'EA-Matic


'function EA_OnPostNewAttribute(Info)
'	msgbox "EA_OnPostNewAttribute"
'	dim attributeID 
'	attributeID = Info.Get("AttributeID")
'	dim attribute
'	set attribute = Repository.GetAttributeByID(attributeID)
'	msgbox "Attribute Name " & attribute.Name
'	'set type to none
'	attribute.Type = ""
'	attribute.Update
'end function