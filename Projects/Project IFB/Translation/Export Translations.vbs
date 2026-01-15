'[path=\Projects\Project IFB\Translation]
'[group=Translation]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Export Translations
' Author: Geert Bellekens
' Purpose: Export the translations for al elements in the selected pacakge
' Date: 2023-12-07
'
'name of the output tab
const outPutName = "Export Translations"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'set timestamp
	Repository.WriteOutput outPutName, now() &  " Starting " & outPutName , 0
	'Actually do the thing
	exportTranslations
	'set timestamp
	Repository.WriteOutput outPutName, now() &  " Finished " & outPutName , 0
end sub


function exportTranslations()
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	'check if package found
	if package is nothing then
		exit function
	end if
	'get elements with possible translations
	dim elements
	set elements = getTranslationElements(package)
	'get translations
	dim translationData
	set translationData = getTranslationData(elements)
	'export to excel file
	exportTranslationsToExcel translationData
end function

function getTranslationElements(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o                         " & vbNewLine & _
				" where o.Object_Type not in ('Text', 'Note', 'Boundary')   " & vbNewLine & _
				" and o.Package_ID in ( " & packageTreeIDString & ")        "
	dim data
	set data = getElementsFromQuery(sqlGetData)
	'return
	set getTranslationElements = data
end function

function getTranslationData(elements)
	'get the secundary languages
	dim secundaryLanguages
	set secundaryLanguages = getSecondaryLanguages()
	'create data
	dim data
	set data = CreateObject("System.Collections.ArrayList")
	'get the headers first
	dim headers
	set headers = getHeaders(secundaryLanguages)
	data.add headers
	'add acutal element data
	dim element as EA.Element
	for each element in elements
		dim row
		set row = CreateObject("System.Collections.ArrayList")
		row.Add element.Name '0
		row.add element.Type '1
		row.add element.Stereotype '2
		row.add element.ElementGUID '3
		row.add Repository.GetFormatFromField("TXT",element.Notes) '4
		'add translations for each secundary language
		dim language
		for each language in secundaryLanguages
			row.add element.GetTXName(language, 0)
			row.add Repository.GetFormatFromField("TXT",element.GetTXNote(language, O))
		next
		'add the row to the data
		data.add row
	next
	'return
	set getTranslationData = data
end function

function exportTranslationsToExcel(translationData)
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	'create a two dimensional array
	dim excelContents
	excelContents = makeArrayFromArrayLists(translationData)
	'add the output to a sheet in excel
	excelOutput.createTab "Translations", excelContents, true, "TableStyleMedium4"
	'save the excel file
	excelOutput.save
end function

function getHeaders(secundaryLanguages)
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.add("Name")
	headers.add("Type")
	headers.Add("Stereotype")
	headers.Add("UniqueID")
	headers.Add("Notes")
	dim language
	for each language in secundaryLanguages
		headers.Add "Name - " & ucase(language)
		headers.Add "Notes - " & ucase(language)
	next
	set getHeaders = headers
end function

function getApplicationComponents(package)
	dim applicationComponents
	set applicationComponents = nothing
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	dim sqlGetData
	sqlGetData =  "select o.Object_ID from t_object o                       " & vbNewLine & _
				  " where o.Stereotype = 'ArchiMate_ApplicationComponent'   " & vbNewLine & _
				  " and o.Package_ID in ( " & packageTreeIDString & ")"
	set applicationComponents = getElementsFromQuery(sqlGetData)
	dim element as EA.Element
	for each element in applicationComponents
		if element.FQStereotype = "ArchiMate3::ArchiMate_ApplicationComponent" then
			Repository.WriteOutput outPutName, now() &  " Found " & element.Name & " " & element.ElementGUID, 0
		end if
	next
	'return
	set getApplicationComponents = applicationComponents
end function



main