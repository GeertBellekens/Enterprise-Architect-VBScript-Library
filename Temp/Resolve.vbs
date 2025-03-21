'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
	resolve "{E4E483CE-DF27-4ae6-B4F1-7A17CDA11F1A}"
end sub

function resolve (itemGuid)
	resolve = false 'initial value
	dim item
	set item = Repository.GetAttributeByGuid(itemGuid)
	if item is nothing then
		set item = Repository.GetElementByGuid(itemGuid)
	end if
	dim itemName
	itemName = trim(item.Name)
	dim itemNotes
	itemNotes = item.Notes
	dim invalidChars
	invalidChars = Array("[","–","«","»","“","”","‘","’","]")
	dim invalidChar
	for each invalidChar in invalidChars
		itemName = replace(itemName, invalidChar, "_")
		itemNotes =  replace(itemNotes, invalidChar, "_")
	next
	'check if we need to update
	if itemName <> item.Name _
	  or itemNotes <> item.Notes then
		item.Name = itemName
		item.Notes = itemNotes
		item.Update
	end if
	resolve = true 
end function

main