'[path=\Projects\Project E\Documents]
'[group=Documents]


!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Data Model Export
' Author: Geert Bellekens
' Purpose: Export Datamodel to excel under the selected pacakge
' Date: 2023-03-24
'

'const outPutName = "Export Data Model"

dim DataModelDocument
set DataModelDocument = new Document

DataModelDocument.Name = "DataModel"
DataModelDocument.Description = "A Excel document with all the details of all classes and attributes in this package branch"
DataModelDocument.GenerateFunction = "DataModelExport"
DataModelDocument.ValidQuery = "select top 1 o.Object_ID from t_object o where o.Object_Type = 'Class' and isnull(o.Stereotype, 'Class') like '%Class' and o.Package_ID in (#Branch#)"



sub DataModelExport
	'do the actual work
	exportDataModel()
end sub

function exportDataModel()
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	'get data
	dim data
	set data = getDataModelData(package)
	'export to Excel
	exportDataModelToExcel data
	
end function

function getDataModelData(package)
	dim data
	set data = CreateObject("System.Collections.ArrayList")
	Repository.WriteOutput outPutName, now() & " Getting data from package '" & package.Name & "'" , 0
	'get data
	dim sqlGetData
	sqlGetData = getDataModelSQLGetData(package)
	dim xmlResult
	xmlResult = Repository.SQLQuery(sqlGetData)
	set data = convertQueryResultToArrayList(xmlResult)
	'return
	set getDataModelData = data
end function

function exportDataModelToExcel(data)
	Repository.WriteOutput outPutName, now() & " Exporting to Excel" , 0
	dim headers
	set headers = getDataModelHeaders()
	'add the headers to the results
	data.Insert 0, headers
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	'format description field
	dim row
	for each row in data
		'format notes
		row(8) = Repository.GetFormatFromField("TXT", row(8))
		'remove ownerfield and pos
		row.RemoveAt(3)
		row.RemoveAt(2)
	next
	'create a two dimensional array
	dim excelContents
	excelContents = makeArrayFromArrayLists(data)
	'add the output to a sheet in excel
	dim sheet
	set sheet = excelOutput.createTab("Datamodel Export", excelContents, true, "TableStyleMedium4")
	'set headers to atrias red
	dim headerRange
	set headerRange = sheet.Range(sheet.Cells(1,1), sheet.Cells(1, headers.Count))
	excelOutput.formatRange headerRange, eliaOrange, white, "default", "default", "default", "default"
	'delete first default sheet
	excelOutput.deleteTabAtIndex 1
	'save the excel file
	excelOutput.save
end function

function sortExcelSheet(sheet)
	With sheet.ListObjects("Table1").Sort
		.SortFields.Clear
		.SortFields.Add sheet.Range("Table1[Subset]")
		.SortFields.Add sheet.Range("Table1[Domain]")
		.SortFields.Add sheet.Range("Table1[Entity]")
		.SortFields.Add sheet.Range("Table1[ItemType]"), xlSortOnValues ,xlAscending, "Class,Attribute,Association"
		.SortFields.Add sheet.Range("Table1[Property]")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlSortColumns
        .Apply
    End With
end function


function getDataModelHeaders()
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.add("Type") '0
	headers.add("Guid") '1
	headers.Add("OwnerField") '2
	headers.Add("Pos") '3
	headers.add("Class Name") '4
	headers.Add("Attribute Name") '5
	headers.Add("Alias") '6
	headers.Add("Stereotype") '7
	headers.Add("Notes") '8
	headers.Add("Datatype") '9
	headers.Add("Data360Url") '10
	set getDataModelHeaders = headers
end function


function getDataModelSQLGetData(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	getDataModelSQLGetData = "select o.Object_Type as CLASSTYPE, o.ea_guid AS CLASSGUID,  o.ea_guid as ownerField, o.TPos as Pos,                 " & vbNewLine & _
					" o.Name as ClassName, null as AttributeName, o.Alias, o.Stereotype, o.Note AS Notes_formatted, null as Datatype,    " & vbNewLine & _
					" tv1.Value as Data360Url                                                                                            " & vbNewLine & _
					" from t_object o                                                                                                    " & vbNewLine & _
					" left join t_objectproperties tv1 on tv1.Object_ID = o.Object_ID and tv1.Property = 'Data360 Url'                   " & vbNewLine & _
					"                                                                                                                    " & vbNewLine & _
					" where o.Package_ID in (" & packageTreeIDString & ")                                                                " & vbNewLine & _
					" and ( o.Object_Type ='Class' and isnull(o.Stereotype, 'Class') like '%Class'                                       " & vbNewLine & _
					"     or  o.Object_Type = 'Enumeration')                                                                             " & vbNewLine & _
					" union all                                                                                                          " & vbNewLine & _
					" select case when o.Object_Type = 'Enumeration' then 'Enum Value' else 'Attribute' end as CLASSTYPE,                " & vbNewLine & _
					" a.ea_guid as CLASSGUID, o.ea_guid as ownerField, a.Pos as Pos,                                                     " & vbNewLine & _
					" o.Name as ClassName, a.Name as AttributeName, a.Style as Alias, a.Stereotype, a.Notes as Notes_Formatted,          " & vbNewLine & _
					" a.Type as Datatype, null as Data360Url                                                                             " & vbNewLine & _
					" from t_attribute a                                                                                                 " & vbNewLine & _
					" inner join t_object o on o.Object_ID = a.Object_ID                                                                 " & vbNewLine & _
					" where o.Package_ID in (" & packageTreeIDString & ")                                                                " & vbNewLine & _
					" order by ownerField, CLASSTYPE desc, Pos                                                                           "
end function
