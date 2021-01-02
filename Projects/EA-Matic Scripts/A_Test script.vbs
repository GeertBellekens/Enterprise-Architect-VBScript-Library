'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

dim counter
counter = 0
'function EA_OnAttributeTagEdit(AttributeID, TagName, TagValue, TagNotes)
'	 Msgbox "AttributeTagEdit"
'end function
function EA_OnPostNewPackage(Info)
	counter = counter + 1
	 Msgbox "new package try: "  & counter
end function