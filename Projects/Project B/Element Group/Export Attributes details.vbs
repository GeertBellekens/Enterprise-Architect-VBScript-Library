'[path=\Projects\Project B\Element Group]
'[group=Element Group]

option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Export Attributes details
' Author: Geert Bellekens
' Purpose: export the details from the attributes to a .csv file
' Date: 2020-12-04
'

sub main
	dim element as EA.Element
	set element = Repository.GetContextObject()
	if not element is nothing then
		if element.ObjectType = otElement then
			exportAttributeDetails(element)
		end if
	end if
end sub

function exportAttributeDetails(element)
	dim sqlGetData
	sqlGetData = "select a.Name, a.Notes, a.Type,isnull(a.LowerBound, 0) as Mandatory   " & vbNewLine & _
				" from t_attribute a                                                    " & vbNewLine & _
				" where a.Object_ID = " & element.ElementID & "                         "
	dim data
	set data = getArrayListFromQuery(sqlGetData)
	'format notes
	formatNotes data, 1
	'add headers
	dim headers
	set headers = getHeaders()	
	data.Insert 0, headers
	'write data to file
	dim csvFile
	set csvFile = new CSVFile
	csvFile.FullPath =  Repository.InvokeFileDialog("Character Separated Values|*.csv",1,0)
	csvFile.Contents = data
	csvFile.Save
end function

function formatNotes(data, notesIndex)
	dim row
	for each row in data
		dim notesField
		notesField = row(notesIndex)
		'convert to plain text
		notesField = Repository.GetFormatFromField("TXT",notesField)
		'put it back
		row(notesIndex) = notesField
	next
end function

function getHeaders()
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.add("Name") '0
	headers.add("Notes") '1
	headers.add("Type") '2
	headers.add("Mandatory") '3
	set getHeaders = headers
end function

main