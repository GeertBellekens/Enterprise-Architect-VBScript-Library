'[path=\Projects\Project E\Documents]
'[group=Documents]


!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Ambition to Capability export
' Author: Geert Bellekens
' Purpose: Export the Ambitions under this package, and the links to gaps, projects and capabilities
' Date: 2023-04-18
'

'const outPutName = "Export Ambition to Capability"

dim AmbitionToCapabilityDocument
set AmbitionToCapabilityDocument = new Document
AmbitionToCapabilityDocument.Name = "AmbitionToCapability"
AmbitionToCapabilityDocument.Description = "An Excel export mapping all Ambitions to Capabilities"
AmbitionToCapabilityDocument.ValidQuery = "select top 1 o.Object_ID from t_object o where o.Stereotype = 'Elia_Ambition' and o.Package_ID in (#Branch#)"

sub AmbitionToCapabilityExport
	'do the actual work
	exportDataModel()
end sub

function exportAmbitionToCapabilityDataModel()
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	'get data
	dim data
	set data = getAmbitionToCapabilityData(package)
	'export to Excel
	exportAmbitionToCapabilityToExcel data
	
end function

function getAmbitionToCapabilityData(package)
	dim data
	set data = CreateObject("System.Collections.ArrayList")
	Repository.WriteOutput outPutName, now() & " Getting data from package '" & package.Name & "'" , 0
	'get data
	dim sqlGetData
	sqlGetData = getAmbitionToCapabilitySQLGetData(package)
	dim xmlResult
	xmlResult = Repository.SQLQuery(sqlGetData)
	set data = convertQueryResultToArrayList(xmlResult)
	'return
	set getAmbitionToCapabilityData = data
end function

function exportAmbitionToCapabilityToExcel(data)
	Repository.WriteOutput outPutName, now() & " Exporting to Excel" , 0
	dim headers
	set headers = getAmbitionToCapabilityHeaders()
	'add the headers to the results
	data.Insert 0, headers
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	'format description field
	dim row
	for each row in data
		'format notes
		row(1) = Repository.GetFormatFromField("TXT", row(1))
		row(3) = Repository.GetFormatFromField("TXT", row(3))
	next
	'create a two dimensional array
	dim excelContents
	excelContents = makeArrayFromArrayLists(data)
	'add the output to a sheet in excel
	dim sheet
	set sheet = excelOutput.createTab("AmbitionToCapability", excelContents, true, "TableStyleMedium4")
	'set headers to atrias red
	dim headerRange
	set headerRange = sheet.Range(sheet.Cells(1,1), sheet.Cells(1, headers.Count))
	excelOutput.formatRange headerRange, eliaOrange, white, "default", "default", "default", "default"
	'delete first default sheet
	excelOutput.deleteTabAtIndex 1
	'save the excel file
	excelOutput.save
end function

function getAmbitionToCapabilityHeaders()
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.add("Ambition") '0
	headers.add("Ambition Notes") '1
	headers.Add("Gap") '2
	headers.Add("Gap Description") '3
	headers.Add("Elia Onshore") '4
	headers.add("Elia Offshore") '5
	headers.Add("50Hz Onshore") '6
	headers.add("50Hz Offshore") '7
	headers.Add("Gap Maturity") '8
	headers.Add("Project") '9
	headers.Add("Project Elia/50Hz") '10
	headers.add("Project Onshore/Offshore") '11
	headers.Add("Project Maturity") '12
	headers.Add("Capability") '13
	headers.Add("Capability Elia/50Hz") '14
	headers.add("Capability Onshore/Offshore") '15
	headers.Add("Capability Maturity") '16
	headers.add("Capability L1") '17
	headers.Add("Capability L0") '18
	set getAmbitionToCapabilityHeaders = headers
end function


function getAmbitionToCapabilitySQLGetData(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	getAmbitionToCapabilitySQLGetData = "select o.name as Ambition, o.note as ambitionNotes                                          " & vbNewLine & _
					" , gap.Name as Gap, gap.Note as GapNotes                                                    " & vbNewLine & _
					" , case when gap.EliaOnshore = 'TBD' then null else gap.EliaOnshore end as EliaOnshore      " & vbNewLine & _
					" , case when gap.EliaOffshore = 'TBD' then null else gap.EliaOffshore end as EliaOffshore   " & vbNewLine & _
					" , case when gap.FHzOnshore = 'TBD' then null else gap.FHzOnshore end as FHzOnshore         " & vbNewLine & _
					" , case when gap.FHzOffshore = 'TBD' then null else gap.FHzOffshore end as FHzOffshore      " & vbNewLine & _
					" , gap.Maturity as GapMaturity                                                              " & vbNewLine & _
					" , pr.Name as Project                                                                       " & vbNewLine & _
					" , pr.Scope as PrScope, pr.Shore as PrShore, pr.Maturity as PrMaturity                      " & vbNewLine & _
					" , cap.Name as Capability                                                                   " & vbNewLine & _
					" , cap.Scope as CapScope, cap.Shore as CapShore, cap.Maturity as CapMaturity                " & vbNewLine & _
					" , cap1.Name as Cap1, cap2.name as cap2                                                     " & vbNewLine & _
					" from t_object o                                                                            " & vbNewLine & _
					" inner join t_package p on p.Package_ID = o.Package_ID                                      " & vbNewLine & _
					" left join                                                                                  " & vbNewLine & _
					" 	(select c.End_Object_ID, gap.Name, gap.Object_ID, gap.Note                               " & vbNewLine & _
					" 	,tveo.Value as EliaOnshore, tvef.Value as EliaOffshore                                   " & vbNewLine & _
					" 	, tvfo.Value as FHzOnshore, tvff.Value as FHzOffshore                                    " & vbNewLine & _
					" 	, tvm.Value as Maturity                                                                  " & vbNewLine & _
					" 	from t_connector c                                                                       " & vbNewLine & _
					" 	inner join t_object gap on gap.Object_ID = c.Start_Object_ID                             " & vbNewLine & _
					" 						and gap.Stereotype = 'Elia_Gap'                                      " & vbNewLine & _
					" 	left join t_objectproperties tveo on tveo.Object_ID = gap.Object_ID                      " & vbNewLine & _
					" 									and tveo.Property = 'Elia Onshore'                       " & vbNewLine & _
					" 	left join t_objectproperties tvef on tvef.Object_ID = gap.Object_ID                      " & vbNewLine & _
					" 									and tvef.Property = 'Elia Offshore'                      " & vbNewLine & _
					" 	left join t_objectproperties tvfo on tvfo.Object_ID = gap.Object_ID                      " & vbNewLine & _
					" 									and tvfo.Property = '50Hz Onshore'                       " & vbNewLine & _
					" 	left join t_objectproperties tvff on tvff.Object_ID = gap.Object_ID                      " & vbNewLine & _
					" 									and tvff.Property = '50Hz Offshore'                      " & vbNewLine & _
					" 	left join t_objectproperties tvm on tvm.Object_ID = gap.Object_ID                        " & vbNewLine & _
					" 									and tvm.Property = 'Maturity'                            " & vbNewLine & _
					" 	where c.Stereotype = 'ArchiMate_Realization')                                            " & vbNewLine & _
					" 	gap on gap.End_Object_ID = o.Object_ID                                                   " & vbNewLine & _
					" left join                                                                                  " & vbNewLine & _
					" 	(select c.End_Object_ID, pr.Name, pr.Object_ID                                           " & vbNewLine & _
					" 	,tvs.Value as Scope, tvo.Value as Shore, tvm.Value as Maturity                           " & vbNewLine & _
					" 	from t_connector c                                                                       " & vbNewLine & _
					" 	inner join t_object pr on pr.Object_ID = c.Start_Object_ID                               " & vbNewLine & _
					" 							and pr.Stereotype = 'Elia_Project'                               " & vbNewLine & _
					" 	left join t_objectproperties tvs on tvs.Object_ID = pr.Object_ID                         " & vbNewLine & _
					" 									and tvs.Property = 'Elia/50Hz'                           " & vbNewLine & _
					" 	left join t_objectproperties tvo on tvo.Object_ID = pr.Object_ID                         " & vbNewLine & _
					" 									and tvo.Property = 'Onshore/Offshore'                    " & vbNewLine & _
					" 	left join t_objectproperties tvm on tvm.Object_ID = pr.Object_ID                         " & vbNewLine & _
					" 									and tvm.Property = 'Maturity'                            " & vbNewLine & _
					" 	where c.Stereotype = 'ArchiMate_Realization')                                            " & vbNewLine & _
					" 	pr on pr.End_Object_ID = gap.Object_ID                                                   " & vbNewLine & _
					" left join                                                                                  " & vbNewLine & _
					" 	(select c.Start_Object_ID, cap.Name, cap.Object_ID                                       " & vbNewLine & _
					" 	,tvs.Value as Scope, tvo.Value as Shore, tvm.Value as Maturity                           " & vbNewLine & _
					" 	from t_connector c                                                                       " & vbNewLine & _
					" 	inner join t_object cap on cap.Object_ID = c.End_Object_ID                               " & vbNewLine & _
					" 						and cap.Stereotype = 'ArchiMate_Capability'                          " & vbNewLine & _
					" 	left join t_objectproperties tvs on tvs.Object_ID = cap.Object_ID                        " & vbNewLine & _
					" 									and tvs.Property = 'Elia/50Hz'                           " & vbNewLine & _
					" 	left join t_objectproperties tvo on tvo.Object_ID = cap.Object_ID                        " & vbNewLine & _
					" 									and tvo.Property = 'Onshore/Offshore'                    " & vbNewLine & _
					" 	left join t_objectproperties tvm on tvm.Object_ID = cap.Object_ID                        " & vbNewLine & _
					" 									and tvm.Property = 'Maturity'                            " & vbNewLine & _
					" 	where c.Stereotype = 'ArchiMate_Association')                                            " & vbNewLine & _
					" 	cap on cap.Start_Object_ID = gap.Object_ID                                               " & vbNewLine & _
					" left join                                                                                  " & vbNewLine & _
					" 	(select c.End_Object_ID, cap.Name, cap.Object_ID                                         " & vbNewLine & _
					" 	from t_connector c                                                                       " & vbNewLine & _
					" 	inner join t_object cap on cap.Object_ID = c.Start_Object_ID                             " & vbNewLine & _
					" 						and cap.Stereotype = 'ArchiMate_Capability'                          " & vbNewLine & _
					" 	where c.Stereotype = 'ArchiMate_Composition')                                            " & vbNewLine & _
					" 	cap1 on cap1.End_Object_ID = cap.Object_ID                                               " & vbNewLine & _
					" left join                                                                                  " & vbNewLine & _
					" 	(select c.End_Object_ID, cap.Name, cap.Object_ID                                         " & vbNewLine & _
					" 	from t_connector c                                                                       " & vbNewLine & _
					" 	inner join t_object cap on cap.Object_ID = c.Start_Object_ID                             " & vbNewLine & _
					" 						and cap.Stereotype = 'ArchiMate_Capability'                          " & vbNewLine & _
					" 	where c.Stereotype = 'ArchiMate_Composition')                                            " & vbNewLine & _
					" 	cap2 on cap2.End_Object_ID = cap1.Object_ID                                              " & vbNewLine & _
					" where 1=1                                                                                  " & vbNewLine & _
					" and o.Stereotype = 'Elia_Ambition'                                                         " & vbNewLine & _
					" and p.Package_ID in (" & packageTreeIDString & ")                                          "
end function
