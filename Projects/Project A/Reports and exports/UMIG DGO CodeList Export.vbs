'[path=\Projects\Project A\Reports and exports]
'[group=Reports and exports]

option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: UMIG DGO CodeList Export
' Author: Geert Bellekens
' Purpose: Create an export of all enumeration values used in UMIG DGO
' Date: 2019-10-04
'

const excelTemplate = "G:\Projects\80 Enterprise Architect\Output\Excel export templates\UMIG DGO - SD - XD - 05 - Code Lists.xltx"

sub main
	'get the selected package
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage
	if selectedPackage is nothing then
		msgbox "Please select a package in the project browser before running this script",vbOKOnly+vbExclamation,"No package selected!"
		exit sub
	end if
	'get the package tree id's
	dim packageTreeIDs
	packageTreeIDs = getPackageTreeIDString(selectedPackage)
	'get the results
	dim sqlGetCodeList
	sqlGetCodeList = getSQLGetCodeList(packageTreeIDs)
	dim codeListResults
	set codeListResults = getArrayListFromQuery(sqlGetCodeList)
	dim headers
	set headers = getHeaders()
	'add the headers to the results
	codeListResults.Insert 0, headers
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	'load the template
	excelOutput.NewFile excelTemplate
	'create a two dimensional array
	dim excelContents
	excelContents = makeArrayFromArrayLists(codeListResults)
	'add the output to a sheet in excel
	dim sheet
	set sheet = excelOutput.createTab("UMIG DGO IM Code Lists", excelContents, true, "TableStyleMedium4")
	'set headers to atrias red
	dim headerRange
	set headerRange = sheet.Range(sheet.Cells(1,1), sheet.Cells(1, headers.Count))
	excelOutput.formatRange headerRange, atriasRed, "default", "default", "default", "default", "default"
	'save the excel file
	excelOutput.save
end sub


function getSQLGetCodeList(packageTreeIDs)
	getSQLGetCodeList = "select distinct o.name AS 'Enumeration', '''' + a.name AS 'LiteralValue', a.style AS 'Alias'   " & vbNewLine & _
						" from t_attribute a                                                                     " & vbNewLine & _
						" inner join t_object o ON a.object_id = o.object_id                                     " & vbNewLine & _
						" where 1=1                                                                              " & vbNewLine & _
						" and o.Object_Type = 'Enumeration'                                                      " & vbNewLine & _
						" and o.Package_ID in (" & packageTreeIDs & ")                                           " & vbNewLine & _
						" order by 1,2                                                                           "
end function

function getHeaders()
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.add("Code List")
	headers.add("Code")
	headers.Add("Description")
	set getHeaders = headers
end function

main