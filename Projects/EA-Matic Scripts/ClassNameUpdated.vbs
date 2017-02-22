'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
option explicit
'EA-Matic
!INC Local Scripts.EAConstants-VBScript

'
' Script Name: ClassNameUpdated
' Author: Geert Bellekens
' Purpose: Will show a message when the name of a class with a certain tagged value has been changed
' Date: 2016-12-12
'

Dim contextName
Dim contextGUID
function EA_OnContextItemChanged(GUID, ot)
	 if ot = otElement then
		Dim contextElement 
		set contextElement = Repository.GetElementByGuid(GUID)
		if not contextElement is nothing AND isUsedByBPM(contextElement) then
			contextName = contextElement.Name
			contextGUID = contextElement.ElementGUID
		end if
	end if
end function

function EA_OnNotifyContextItemModified(GUID, ot)
	 'check if the name has been changed
	 if GUID = contextGUID then
		Dim contextElement 
		set contextElement = Repository.GetElementByGuid(GUID)
		if contextName <> contextElement.Name then
			msgbox "Element with name '" & contextName & "' has been changed to '" & contextElement.Name & "'"
		end if
	 end if
end function

function isUsedByBPM(contextElement)
	dim taggedValue
	isUsedByBPM = false
	for each taggedValue in contextElement.TaggedValues
		if taggedValue.Name = "BPM_ID" and len(taggedValue.Value) > 0 then
			isUsedByBPM = true
			exit for
		end if
	next
end function