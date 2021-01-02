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


function EA_OnPostNewAttribute(Info)
	dim attributeID 
	attributeID = Info.Get("AttributeID")
	dim attribute
	set attribute = Repository.GetAttributeByID(attributeID)
	'set type to none
	attribute.Type = ""
	attribute.Update
end function