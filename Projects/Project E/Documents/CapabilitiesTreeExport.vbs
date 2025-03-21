'[path=\Projects\Project E\Documents]
'[group=Documents]


!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Capabilities tre export
' Author: Geert Bellekens
' Purpose: Export all capabilities under the selected package
' Date: 2023-01-13
'
dim CapabilityTreeDocument
set CapabilityTreeDocument = new Document
CapabilityTreeDocument.Name = "CapabilityTree"
CapabilityTreeDocument.Description = "An Excel export of the hierarchy of all capabilities under this package branch"
CapabilityTreeDocument.ValidQuery = "select top 1 o.Object_ID from t_object o where o.Stereotype = 'ArchiMate_Capability' and o.Package_ID in (#Branch#)"


sub CapabilityTreeExport
	'do the actual work
	exportCapabilitiesTree()
end sub

function exportCapabilitiesTree()
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	'get data
	dim data
	set data = getCapabilityTreeData(package)
	'export to Excel
	exportCapabilityTreeToExcel data
	
end function

function getCapabilityTreeData(package)
	dim data
	set data = CreateObject("System.Collections.ArrayList")
	Repository.WriteOutput outPutName, now() & " Getting data from package '" & package.Name & "'" , 0
	'get data
	dim sqlGetData
	sqlGetData = getCapabilityTreeSQLGetData(package)
	dim xmlResult
	xmlResult = Repository.SQLQuery(sqlGetData)
	set data = convertQueryResultToArrayList(xmlResult)
	'return
	set getCapabilityTreeData = data
end function

function exportCapabilityTreeToExcel(data)
	Repository.WriteOutput outPutName, now() & " Exporting to Excel" , 0
	dim headers
	set headers = getCapabilityTreeHeaders()
	'add the headers to the results
	data.Insert 0, headers
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	'format description field
	dim row
	for each row in data
		row(15) = Repository.GetFormatFromField("TXT", row(15))
	next
	'create a two dimensional array
	dim excelContents
	excelContents = makeArrayFromArrayLists(data)
	'add the output to a sheet in excel
	dim sheet
	set sheet = excelOutput.createTab("Capabilitiy Tree", excelContents, true, "TableStyleMedium4")
	'set headers to atrias red
	dim headerRange
	set headerRange = sheet.Range(sheet.Cells(1,1), sheet.Cells(1, headers.Count))
	excelOutput.formatRange headerRange, eliaOrange, white, "default", "default", "default", "default"
	'delete first default sheet
	excelOutput.deleteTabAtIndex 1
	'save the excel file
	excelOutput.save
end function

function getCapabilityTreeHeaders()
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.add("GUID")
	headers.Add("TYPE")
	headers.Add("Stereotype")
	headers.add("Name")
	headers.Add("Level")
	headers.Add("LevelTag")
	headers.Add("Maturity")
	headers.Add("Elia/50Hz")
	headers.Add("Onshore/Offshore")
	headers.Add("BusinessFunction")
	headers.Add("Capability_Category")
	headers.Add("Capability_L0")
	headers.Add("Capability_L1")
	headers.Add("Capability_L2")
	headers.Add("Capability_L3")
	headers.Add("Description")
	set getCapabilityTreeHeaders = headers
end function


function getCapabilityTreeSQLGetData(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	getCapabilityTreeSQLGetData = "select                                                                                                                        " & vbNewLine & _
					" o.ea_guid AS CLASSGUID,  o.Object_Type AS CLASSTYPE, o.Stereotype                                                            " & vbNewLine & _
					" ,o.name as Name, c.Level                                                                                                     " & vbNewLine & _
					" ,tvl.Value as LevelTag, tvm.Value as Maturity                                                                                " & vbNewLine & _
					" , tvef.Value as Elia_50Hz, tvo.Value as Onshore_Offshore                                                                     " & vbNewLine & _
					" , c.BusinessFunction BusinessFunction                                                                                        " & vbNewLine & _
					" , c.Capability_Category, c.Capability_L0, c.Capability_L1, c.Capability_L2, c.Capability_L3                                  " & vbNewLine & _
					" , o.Note as Description                                                                                                      " & vbNewLine & _
					" from                                                                                                                         " & vbNewLine & _
					" 	(                                                                                                                          " & vbNewLine & _
					" 	select p.ea_guid as PackageGUID,c0.Object_ID as CapabilityID, cc.Name as Capability_Category                               " & vbNewLine & _
					" 	, c0.name as Capability_L0, null as Capability_L1, null as Capability_L2, null as Capability_L3                            " & vbNewLine & _
					" 	, 0 as Level, null as BusinessFunction                                                                                     " & vbNewLine & _
					" 	from t_object cc                                                                                                           " & vbNewLine & _
					" 	inner join (select c.Start_Object_ID, o.*                                                                                  " & vbNewLine & _
					" 				from t_object o                                                                                                " & vbNewLine & _
					" 				inner join t_connector c on c.End_Object_ID = o.Object_ID                                                      " & vbNewLine & _
					" 										and c.Stereotype = 'Archimate_Composition'                                             " & vbNewLine & _
					" 				where o.Stereotype = 'ArchiMate_Capability'                                                                    " & vbNewLine & _
					" 				) c0 on c0.Start_Object_ID = cc.Object_ID                                                                      " & vbNewLine & _
					" 	inner join t_package p on p.Package_ID = c0.Package_ID                                                                     " & vbNewLine & _
					" 	where cc.Stereotype = 'ArchiMate_Capability'                                                                               " & vbNewLine & _
					" 	and not exists (select c.Connector_ID from t_connector c                                                                   " & vbNewLine & _
					" 				inner join t_object o on o.Object_id = c.Start_Object_ID                                                       " & vbNewLine & _
					" 										and o.Stereotype = 'ArchiMate_Capability'                                              " & vbNewLine & _
					" 				where c.Stereotype = 'Archimate_Composition'                                                                   " & vbNewLine & _
					" 				and c.End_Object_ID = cc.Object_ID)                                                                            " & vbNewLine & _
					" 	union                                                                                                                      " & vbNewLine & _
					" 	select p.ea_guid as PackageGUID, c1.Object_ID as CapabilityID, cc.Name as Capability_Category                              " & vbNewLine & _
					" 	, c0.name as Capability_L0, c1.name as Capability_L1, null as Capability_L2, null as Capability_L3                         " & vbNewLine & _
					" 	, 1 as Level, null as BusinessFunction                                                                                     " & vbNewLine & _
					" 	from t_object cc                                                                                                           " & vbNewLine & _
					" 	inner join (select c.Start_Object_ID, o.*                                                                                  " & vbNewLine & _
					" 				from t_object o                                                                                                " & vbNewLine & _
					" 				inner join t_connector c on c.End_Object_ID = o.Object_ID                                                      " & vbNewLine & _
					" 										and c.Stereotype = 'Archimate_Composition'                                             " & vbNewLine & _
					" 				where o.Stereotype = 'ArchiMate_Capability'                                                                    " & vbNewLine & _
					" 				) c0 on c0.Start_Object_ID = cc.Object_ID                                                                      " & vbNewLine & _
					" 	inner join (select c.Start_Object_ID, o.*                                                                                  " & vbNewLine & _
					" 				from t_object o                                                                                                " & vbNewLine & _
					" 				inner join t_connector c on c.End_Object_ID = o.Object_ID                                                      " & vbNewLine & _
					" 										and c.Stereotype = 'Archimate_Composition'                                             " & vbNewLine & _
					" 				where o.Stereotype = 'ArchiMate_Capability'                                                                    " & vbNewLine & _
					" 				) c1 on c1.Start_Object_ID = c0.Object_ID                                                                      " & vbNewLine & _
					" 	inner join t_package p on p.Package_ID = c0.Package_ID                                                                     " & vbNewLine & _
					" 	where cc.Stereotype = 'ArchiMate_Capability'                                                                               " & vbNewLine & _
					" 	and not exists (select c.Connector_ID from t_connector c                                                                   " & vbNewLine & _
					" 			inner join t_object o on o.Object_id = c.Start_Object_ID                                                           " & vbNewLine & _
					" 									and o.Stereotype = 'ArchiMate_Capability'                                                  " & vbNewLine & _
					" 			where c.Stereotype = 'Archimate_Composition'                                                                       " & vbNewLine & _
					" 			and c.End_Object_ID = cc.Object_ID)                                                                                " & vbNewLine & _
					" 	union                                                                                                                      " & vbNewLine & _
					" 	select p.ea_guid as PackageGUID, c2.Object_ID as CapabilityID, cc.Name as Capability_Category                              " & vbNewLine & _
					" 	, c0.name as Capability_L0, c1.name as Capability_L1, c2.name as Capability_L2, null as Capability_L3                      " & vbNewLine & _
					" 	, 2 as Level, null as BusinessFunction                                                                                     " & vbNewLine & _
					" 	from t_object cc                                                                                                           " & vbNewLine & _
					" 	inner join (select c.Start_Object_ID, o.*                                                                                  " & vbNewLine & _
					" 				from t_object o                                                                                                " & vbNewLine & _
					" 				inner join t_connector c on c.End_Object_ID = o.Object_ID                                                      " & vbNewLine & _
					" 										and c.Stereotype = 'Archimate_Composition'                                             " & vbNewLine & _
					" 				where o.Stereotype = 'ArchiMate_Capability'                                                                    " & vbNewLine & _
					" 				) c0 on c0.Start_Object_ID = cc.Object_ID                                                                      " & vbNewLine & _
					" 	inner join (select c.Start_Object_ID, o.*                                                                                  " & vbNewLine & _
					" 				from t_object o                                                                                                " & vbNewLine & _
					" 				inner join t_connector c on c.End_Object_ID = o.Object_ID                                                      " & vbNewLine & _
					" 										and c.Stereotype = 'Archimate_Composition'                                             " & vbNewLine & _
					" 				where o.Stereotype = 'ArchiMate_Capability'                                                                    " & vbNewLine & _
					" 				) c1 on c1.Start_Object_ID = c0.Object_ID                                                                      " & vbNewLine & _
					" 	inner join (select c.Start_Object_ID, o.*                                                                                  " & vbNewLine & _
					" 				from t_object o                                                                                                " & vbNewLine & _
					" 				inner join t_connector c on c.End_Object_ID = o.Object_ID                                                      " & vbNewLine & _
					" 										and c.Stereotype = 'Archimate_Composition'                                             " & vbNewLine & _
					" 				where o.Stereotype = 'ArchiMate_Capability'                                                                    " & vbNewLine & _
					" 				) c2 on c2.Start_Object_ID = c1.Object_ID                                                                      " & vbNewLine & _
					" 	inner join t_package p on p.Package_ID = c0.Package_ID                                                                     " & vbNewLine & _
					" 	where cc.Stereotype = 'ArchiMate_Capability'                                                                               " & vbNewLine & _
					" 	and not exists (select c.Connector_ID from t_connector c                                                                   " & vbNewLine & _
					" 				inner join t_object o on o.Object_id = c.Start_Object_ID                                                       " & vbNewLine & _
					" 										and o.Stereotype = 'ArchiMate_Capability'                                              " & vbNewLine & _
					" 				where c.Stereotype = 'Archimate_Composition'                                                                   " & vbNewLine & _
					" 				and c.End_Object_ID = cc.Object_ID)                                                                            " & vbNewLine & _
					" 	union                                                                                                                      " & vbNewLine & _
					" 	select p.ea_guid as PackageGUID, c3.Object_ID as CapabilityID, cc.Name as Capability_Category                              " & vbNewLine & _
					" 	, c0.name as Capability_L0, c1.name as Capability_L1, c2.name as Capability_L2, c3.name as Capability_L3                   " & vbNewLine & _
					" 	, 3 as Level, null as BusinessFunction                                                                                     " & vbNewLine & _
					" 	from t_object cc                                                                                                           " & vbNewLine & _
					" 	inner join (select c.Start_Object_ID, o.*                                                                                  " & vbNewLine & _
					" 				from t_object o                                                                                                " & vbNewLine & _
					" 				inner join t_connector c on c.End_Object_ID = o.Object_ID                                                      " & vbNewLine & _
					" 										and c.Stereotype = 'Archimate_Composition'                                             " & vbNewLine & _
					" 				where o.Stereotype = 'ArchiMate_Capability'                                                                    " & vbNewLine & _
					" 				) c0 on c0.Start_Object_ID = cc.Object_ID                                                                      " & vbNewLine & _
					" 	inner join (select c.Start_Object_ID, o.*                                                                                  " & vbNewLine & _
					" 				from t_object o                                                                                                " & vbNewLine & _
					" 				inner join t_connector c on c.End_Object_ID = o.Object_ID                                                      " & vbNewLine & _
					" 										and c.Stereotype = 'Archimate_Composition'                                             " & vbNewLine & _
					" 				where o.Stereotype = 'ArchiMate_Capability'                                                                    " & vbNewLine & _
					" 				) c1 on c1.Start_Object_ID = c0.Object_ID                                                                      " & vbNewLine & _
					" 	inner join (select c.Start_Object_ID, o.*                                                                                  " & vbNewLine & _
					" 				from t_object o                                                                                                " & vbNewLine & _
					" 				inner join t_connector c on c.End_Object_ID = o.Object_ID                                                      " & vbNewLine & _
					" 										and c.Stereotype = 'Archimate_Composition'                                             " & vbNewLine & _
					" 				where o.Stereotype = 'ArchiMate_Capability'                                                                    " & vbNewLine & _
					" 				) c2 on c2.Start_Object_ID = c1.Object_ID                                                                      " & vbNewLine & _
					" 	inner join (select c.Start_Object_ID, o.*                                                                                  " & vbNewLine & _
					" 				from t_object o                                                                                                " & vbNewLine & _
					" 				inner join t_connector c on c.End_Object_ID = o.Object_ID                                                      " & vbNewLine & _
					" 										and c.Stereotype = 'Archimate_Composition'                                             " & vbNewLine & _
					" 				where o.Stereotype = 'ArchiMate_Capability'                                                                    " & vbNewLine & _
					" 				) c3 on c3.Start_Object_ID = c2.Object_ID                                                                      " & vbNewLine & _
					" 	inner join t_package p on p.Package_ID = c0.Package_ID                                                                     " & vbNewLine & _
					" 	where cc.Stereotype = 'ArchiMate_Capability'                                                                               " & vbNewLine & _
					" 	and not exists (select c.Connector_ID from t_connector c                                                                   " & vbNewLine & _
					" 				inner join t_object o on o.Object_id = c.Start_Object_ID                                                       " & vbNewLine & _
					" 										and o.Stereotype = 'ArchiMate_Capability'                                              " & vbNewLine & _
					" 				where c.Stereotype = 'Archimate_Composition'                                                                   " & vbNewLine & _
					" 				and c.End_Object_ID = cc.Object_ID)                                                                            " & vbNewLine & _
					" 	union                                                                                                                      " & vbNewLine & _
					" 	select p.ea_guid as PackageGUID, bf.Object_ID as CapabilityID, cc.Name as Capability_Category                              " & vbNewLine & _
					" 	, c0.name as Capability_L0, c1.name as Capability_L1, c2.name as Capability_L2, c3.name as Capability_L3                   " & vbNewLine & _
					" 	, case when bf.ParentID = c1.Object_ID then 1                                                                              " & vbNewLine & _
					" 		  when  bf.ParentID = c2.Object_ID then 2                                                                              " & vbNewLine & _
					" 		  when  bf.ParentID = c3.Object_ID then 3                                                                              " & vbNewLine & _
					" 	  end as Level                                                                                                             " & vbNewLine & _
					" 	, bf.Name as BusinessFunction                                                                                              " & vbNewLine & _
					" 	from t_object cc                                                                                                           " & vbNewLine & _
					" 	inner join (select c.Start_Object_ID, o.*                                                                                  " & vbNewLine & _
					" 				from t_object o                                                                                                " & vbNewLine & _
					" 				inner join t_connector c on c.End_Object_ID = o.Object_ID                                                      " & vbNewLine & _
					" 										and c.Stereotype = 'Archimate_Composition'                                             " & vbNewLine & _
					" 				where o.Stereotype = 'ArchiMate_Capability'                                                                    " & vbNewLine & _
					" 				) c0 on c0.Start_Object_ID = cc.Object_ID                                                                      " & vbNewLine & _
					" 	left join (select c.Start_Object_ID, o.*                                                                                   " & vbNewLine & _
					" 				from t_object o                                                                                                " & vbNewLine & _
					" 				inner join t_connector c on c.End_Object_ID = o.Object_ID                                                      " & vbNewLine & _
					" 										and c.Stereotype = 'Archimate_Composition'                                             " & vbNewLine & _
					" 				where o.Stereotype = 'ArchiMate_Capability'                                                                    " & vbNewLine & _
					" 				) c1 on c1.Start_Object_ID = c0.Object_ID                                                                      " & vbNewLine & _
					" 	left join (select c.Start_Object_ID, o.*                                                                                   " & vbNewLine & _
					" 				from t_object o                                                                                                " & vbNewLine & _
					" 				inner join t_connector c on c.End_Object_ID = o.Object_ID                                                      " & vbNewLine & _
					" 										and c.Stereotype = 'Archimate_Composition'                                             " & vbNewLine & _
					" 				where o.Stereotype = 'ArchiMate_Capability'                                                                    " & vbNewLine & _
					" 				) c2 on c2.Start_Object_ID = c1.Object_ID                                                                      " & vbNewLine & _
					" 	left join (select c.Start_Object_ID, o.*                                                                                   " & vbNewLine & _
					" 				from t_object o                                                                                                " & vbNewLine & _
					" 				inner join t_connector c on c.End_Object_ID = o.Object_ID                                                      " & vbNewLine & _
					" 										and c.Stereotype = 'Archimate_Composition'                                             " & vbNewLine & _
					" 				where o.Stereotype = 'ArchiMate_Capability'                                                                    " & vbNewLine & _
					" 				) c3 on c3.Start_Object_ID = c2.Object_ID                                                                      " & vbNewLine & _
					" 	inner join t_object bf on bf.parentID in (c1.Object_ID, c2.Object_ID, c3.Object_ID)                                        " & vbNewLine & _
					" 					and bf.Stereotype = 'ArchiMate_BusinessFunction'                                                           " & vbNewLine & _
					" 	inner join t_package p on p.Package_ID = c0.Package_ID                                                                     " & vbNewLine & _
					" 	where cc.Stereotype = 'ArchiMate_Capability'                                                                               " & vbNewLine & _
					" 	and not exists (select c.Connector_ID from t_connector c                                                                   " & vbNewLine & _
					" 				inner join t_object o on o.Object_id = c.Start_Object_ID                                                       " & vbNewLine & _
					" 										and o.Stereotype = 'ArchiMate_Capability'                                              " & vbNewLine & _
					" 				where c.Stereotype = 'Archimate_Composition'                                                                   " & vbNewLine & _
					" 				and c.End_Object_ID = cc.Object_ID)                                                                            " & vbNewLine & _
					" 	) c                                                                                                                        " & vbNewLine & _
					" inner join t_object o on o.Object_ID = c.CapabilityID                                                                        " & vbNewLine & _
					" left join t_objectproperties tvl on tvl.Object_ID = o.Object_ID                                                              " & vbNewLine & _
					" 									and tvl.Property = 'Level'                                                                 " & vbNewLine & _
					" left join t_objectproperties tvm on tvm.Object_ID = o.Object_ID                                                              " & vbNewLine & _
					" 									and tvm.Property = 'Maturity'                                                              " & vbNewLine & _
					" left join t_objectproperties tvo on tvo.Object_ID = o.Object_ID                                                              " & vbNewLine & _
					" 									and tvo.Property = 'Onshore/Offshore'                                                      " & vbNewLine & _
					" left join t_objectproperties tvef on tvef.Object_ID = o.Object_ID                                                            " & vbNewLine & _
					" 									and tvef.Property = 'Elia/50Hz'                                                            " & vbNewLine & _
					" where o.Package_ID in (" & packageTreeIDString & ")                                                                          " & vbNewLine & _
					" order by Capability_Category, Capability_L0, Capability_L1, Capability_L2, Capability_L3, BusinessFunction                   "
end function
