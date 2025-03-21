'[path=\Projects\Project E\Documents]
'[group=Documents]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Ambition to Capability export
' Author: Geert Bellekens
' Purpose: Export the Project Dependencies under this package, and the links to gaps, projects and capabilities
' Date: 2023-06-16
'

'create the document object and add to global list of all documents

dim ProjectDependenciesDocument
set ProjectDependenciesDocument = new Document

ProjectDependenciesDocument.Name = "ProjectDependencies"
ProjectDependenciesDocument.Description = "An Excel export of all Project Dependencies in this package, and the links to gaps, projects and capabilities"
ProjectDependenciesDocument.ValidQuery = "select top 1 o.Object_ID from t_object o where o.Stereotype = 'Elia_Project' and o.Package_ID in (#Branch#)"

sub ProjectDependenciesExport
	'do the actual work
	exportProjectDependencies()
end sub

function exportProjectDependencies()
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	'get data
	dim data
	set data = getProjectDependenciesExportData(package)
	'export to Excel
	exportProjectDependenciesToExcel data
	
end function

function getProjectDependenciesExportData(package)
	dim data
	set data = CreateObject("System.Collections.ArrayList")
	Repository.WriteOutput outPutName, now() & " Getting data from package '" & package.Name & "'" , 0
	'get data
	dim sqlGetData
	sqlGetData = getProjectDependenciesSQLGetData(package)
	dim xmlResult
	xmlResult = Repository.SQLQuery(sqlGetData)
	set data = convertQueryResultToArrayList(xmlResult)
	'return
	set getProjectDependenciesExportData = data
end function

function exportProjectDependenciesToExcel(data)
	Repository.WriteOutput outPutName, now() & " Exporting to Excel" , 0
	dim headers
	set headers = getProjectDependenciesHeaders()
	'add the headers to the results
	data.Insert 0, headers
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	'format description field
'	dim row
'	for each row in data
'		'format notes
'		row(1) = Repository.GetFormatFromField("TXT", row(1))
'		row(3) = Repository.GetFormatFromField("TXT", row(3))
'	next
	'create a two dimensional array
	dim excelContents
	excelContents = makeArrayFromArrayLists(data)
	'add the output to a sheet in excel
	dim sheet
	set sheet = excelOutput.createTab("Project Dependencies", excelContents, true, "TableStyleMedium4")
	'set headers to atrias red
	dim headerRange
	set headerRange = sheet.Range(sheet.Cells(1,1), sheet.Cells(1, headers.Count))
	excelOutput.formatRange headerRange, eliaOrange, white, "default", "default", "default", "default"
	'delete first default sheet
	excelOutput.deleteTabAtIndex 1
	'save the excel file
	excelOutput.save
end function

function getProjectDependenciesHeaders()
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.add("Project Source") '0
	headers.add("Gap Source") '1
	headers.Add("Ambition Source") '2
	headers.Add("L0 Source") '3
	headers.Add("L1 Source") '4
	headers.add("L2 Source") '5
	headers.Add("Conveyed") '6
	headers.Add("Project Target") '7
	headers.Add("Gap Target") '8
	headers.add("Ambition Target") '9
	headers.Add("L0 Target") '10
	headers.Add("L1 Target") '11
	headers.Add("L2 Target") '12
	set getProjectDependenciesHeaders = headers
end function


function getProjectDependenciesSQLGetData(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	getProjectDependenciesSQLGetData = "select pr.Name as Project_Source, gap.Name as Gap_Source, amb.Name as Ambition_Source    " & vbNewLine & _
					" ,l0.name as L0_Source, l1.name as L1_Source, l2.name as L2_Source,                      " & vbNewLine & _
					" isnull (conv.Name , '<missing conveyed object> ' + isnull(c.Name, '')) as Conveyed      " & vbNewLine & _
					" ,prt.Name as Project_Target, gapt.Name as Gap_Target, amb.Name as Ambition_Source       " & vbNewLine & _
					" ,l0t.Name as L0Target, l1t.name as L1Target, l2t.Name as L2Target                       " & vbNewLine & _
					" from t_object l2                                                                        " & vbNewLine & _
					" inner join t_connector c on c.Start_Object_ID = l2.Object_ID                            " & vbNewLine & _
					" 							and c.Connector_Type = 'InformationFlow'                      " & vbNewLine & _
					" left join t_xref x on x.Client = c.ea_guid                                              " & vbNewLine & _
					" 					and x.Name = 'MOFProps'                                               " & vbNewLine & _
					" 					and x.Type = 'connector property'                                     " & vbNewLine & _
					" 					and x.Behavior = 'conveyed'                                           " & vbNewLine & _
					" left join t_object conv on x.Description like '%' + conv.ea_guid + '%'                  " & vbNewLine & _
					" inner join t_objectproperties tvl on tvl.Object_ID = l2.Object_ID                       " & vbNewLine & _
					" 									and tvl.Property = 'Level'                            " & vbNewLine & _
					" 									and tvl.Value = 'L2'                                  " & vbNewLine & _
					" left join (select l2l1.End_Object_ID, l1.Object_ID, l1.name                             " & vbNewLine & _
					" 			from t_connector l2l1                                                         " & vbNewLine & _
					" 			inner join t_object l1 on l1.Object_ID = l2l1.Start_Object_ID                 " & vbNewLine & _
					" 			where l2l1.Stereotype = 'ArchiMate_Composition'                               " & vbNewLine & _
					" 			) l1 on l1.End_Object_ID = l2.Object_ID                                       " & vbNewLine & _
					" left join (select l1l0.End_Object_ID, l0.Object_ID , l0.Name                            " & vbNewLine & _
					" 			from t_connector l1l0                                                         " & vbNewLine & _
					" 			inner join t_object l0 on l0.Object_ID = l1l0.Start_Object_ID                 " & vbNewLine & _
					" 			where l1l0.Stereotype = 'ArchiMate_Composition'                               " & vbNewLine & _
					" 			) l0 on l0.End_Object_ID = l1.Object_ID                                       " & vbNewLine & _
					" inner join t_object l2t on l2t.Object_ID = c.End_Object_ID                              " & vbNewLine & _
					" 							and l2t.Stereotype =  'Archimate_Capability'                  " & vbNewLine & _
					" left join (select l2l1.End_Object_ID, l1.Object_ID, l1.name                             " & vbNewLine & _
					" 			from t_connector l2l1                                                         " & vbNewLine & _
					" 			inner join t_object l1 on l1.Object_ID = l2l1.Start_Object_ID                 " & vbNewLine & _
					" 			where l2l1.Stereotype = 'ArchiMate_Composition'                               " & vbNewLine & _
					" 			) l1t on l1t.End_Object_ID = l2t.Object_ID                                    " & vbNewLine & _
					" left join (select l1l0.End_Object_ID, l0.Object_ID , l0.Name                            " & vbNewLine & _
					" 			from t_connector l1l0                                                         " & vbNewLine & _
					" 			inner join t_object l0 on l0.Object_ID = l1l0.Start_Object_ID                 " & vbNewLine & _
					" 			where l1l0.Stereotype = 'ArchiMate_Composition'                               " & vbNewLine & _
					" 			) l0t on l0t.End_Object_ID = l1t.Object_ID                                    " & vbNewLine & _
					" inner join                                                                              " & vbNewLine & _
					" 	(select c.End_Object_ID, gap.Name, gap.Object_ID, gap.Note                            " & vbNewLine & _
					" 	from t_connector c                                                                    " & vbNewLine & _
					" 	inner join t_object gap on gap.Object_ID = c.Start_Object_ID                          " & vbNewLine & _
					" 						and gap.Stereotype = 'Elia_Gap'                                   " & vbNewLine & _
					" 	where c.Stereotype = 'ArchiMate_Association')                                         " & vbNewLine & _
					" 	gap on gap.End_Object_ID = l2.Object_ID                                               " & vbNewLine & _
					" inner join                                                                              " & vbNewLine & _
					" 	(select c.End_Object_ID, pr.Name, pr.Object_ID, pr.Package_ID                         " & vbNewLine & _
					" 	from t_connector c                                                                    " & vbNewLine & _
					" 	inner join t_object pr on pr.Object_ID = c.Start_Object_ID                            " & vbNewLine & _
					" 							and pr.Stereotype = 'Elia_Project'                            " & vbNewLine & _
					" 	where c.Stereotype = 'ArchiMate_Realization')                                         " & vbNewLine & _
					" 	pr on pr.End_Object_ID = gap.Object_ID                                                " & vbNewLine & _
					" left join                                                                               " & vbNewLine & _
					" 	(select c.Start_Object_ID, pr.Name, pr.Object_ID                                      " & vbNewLine & _
					" 	from t_connector c                                                                    " & vbNewLine & _
					" 	inner join t_object pr on pr.Object_ID = c.End_Object_ID                              " & vbNewLine & _
					" 							and pr.Stereotype = 'Elia_Ambition'                           " & vbNewLine & _
					" 	where c.Stereotype = 'ArchiMate_Realization')                                         " & vbNewLine & _
					" 	amb on amb.Start_Object_ID = gap.Object_ID                                            " & vbNewLine & _
					" inner join                                                                              " & vbNewLine & _
					" 	(select c.End_Object_ID, c.Start_Object_ID, gap.Name, gap.Object_ID, gap.Note         " & vbNewLine & _
					" 	from t_connector c                                                                    " & vbNewLine & _
					" 	inner join t_object gap on gap.Object_ID in (c.Start_Object_ID, c.End_Object_ID)      " & vbNewLine & _
					" 						and gap.Stereotype = 'Elia_Gap'                                   " & vbNewLine & _
					" 	where c.Stereotype = 'ArchiMate_Association')                                         " & vbNewLine & _
					" 	gapt on l2t.Object_ID in (gapt.End_Object_ID, gapt.Start_Object_ID)                   " & vbNewLine & _
					" inner join                                                                              " & vbNewLine & _
					" 	(select c.End_Object_ID, pr.Name, pr.Object_ID, pr.Package_ID                         " & vbNewLine & _
					" 	from t_connector c                                                                    " & vbNewLine & _
					" 	inner join t_object pr on pr.Object_ID = c.Start_Object_ID                            " & vbNewLine & _
					" 							and pr.Stereotype = 'Elia_Project'                            " & vbNewLine & _
					" 	where c.Stereotype = 'ArchiMate_Realization')                                         " & vbNewLine & _
					" 	prt on prt.End_Object_ID = gapt.Object_ID                                             " & vbNewLine & _
					" left join                                                                               " & vbNewLine & _
					" 	(select c.Start_Object_ID, pr.Name, pr.Object_ID                                      " & vbNewLine & _
					" 	from t_connector c                                                                    " & vbNewLine & _
					" 	inner join t_object pr on pr.Object_ID = c.End_Object_ID                              " & vbNewLine & _
					" 							and pr.Stereotype = 'Elia_Ambition'                           " & vbNewLine & _
					" 	where c.Stereotype = 'ArchiMate_Realization')                                         " & vbNewLine & _
					" 	ambt on ambt.Start_Object_ID = gapt.Object_ID                                         " & vbNewLine & _
					" where l2.Stereotype = 'Archimate_Capability'                                            " & vbNewLine & _
					" and (pr.Package_ID in (" & packageTreeIDString & ")                                     " & vbNewLine & _
					" or  prt.Package_ID in (" & packageTreeIDString & "))                                    "
end function

