'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Export Message Detail
' Author: Geert Bellekens
' Purpose: Export the details of a message into a the search window and possibly into a CSV file
' Date: 2017-03-14
'

const outputTabName = "Export Message Detail"
	
sub main
	'setup output
	Repository.CreateOutputTab outputTabName
	Repository.ClearOutput outputTabName
	Repository.EnsureOutputVisible outputTabName
	
	'select the root object
	dim rootObject as EA.Element
	set rootObject = Repository.GetContextObject()
	if rootObject.ObjectType = otElement then
		exportDetails rootObject
	end if
end sub

function exportDetails(rootObject)
	
end function

main