'[path=\Projects\Project E\Documents]
'[group=Documents]


!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Capability Informationflows export
' Author: Geert Bellekens
' Purpose: Export all Capability Informationflows under the selected pacakge
' Date: 2023-03-24
'

dim InformationFlowsDocument
set InformationFlowsDocument = new Document
InformationFlowsDocument.Name = "InformationFlows"
InformationFlowsDocument.Description = "An Excel export with the details of all information flows under this package"
InformationFlowsDocument.ValidQuery = "select top 1 o.Object_ID from t_object o where o.Stereotype = 'ArchiMate_Capability' and o.Package_ID in (#Branch#)"


sub InformationFlowsExport
	'do the actual work
	exportInformationFlows
end sub

function exportInformationFlows()
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	'get data
	dim data
	set data = getInformationFlowsData(package)
	'export to Excel
	exportInformationFlowsToExcel data
	
end function

function getInformationFlowsData(package)
	dim data
	set data = CreateObject("System.Collections.ArrayList")
	Repository.WriteOutput outPutName, now() & " Getting data from package '" & package.Name & "'" , 0
	'get data
	dim sqlGetData
	sqlGetData = getInformationFlowsSQLGetData(package)
	dim xmlResult
	xmlResult = Repository.SQLQuery(sqlGetData)
	set data = convertQueryResultToArrayList(xmlResult)
	'return
	set getInformationFlowsData = data
end function

function exportInformationFlowsToExcel(data)
	Repository.WriteOutput outPutName, now() & " Exporting to Excel" , 0
	dim headers
	set headers = getInformationFlowsHeaders()
	'add the headers to the results
	data.Insert 0, headers
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	'format description field
'	dim row
'	for each row in data
'		'format notes
'		row(8) = Repository.GetFormatFromField("TXT", row(8))
'		'remove ownerfield and pos
'		row.RemoveAt(3)
'		row.RemoveAt(2)
'	next
	'create a two dimensional array
	dim excelContents
	excelContents = makeArrayFromArrayLists(data)
	'add the output to a sheet in excel
	dim sheet
	set sheet = excelOutput.createTab("InformationFlows", excelContents, true, "TableStyleMedium4")
	'set headers to atrias red
	dim headerRange
	set headerRange = sheet.Range(sheet.Cells(1,1), sheet.Cells(1, headers.Count))
	excelOutput.formatRange headerRange, eliaOrange, white, "default", "default", "default", "default"
	'delete first default sheet
	excelOutput.deleteTabAtIndex 1
	'save the excel file
	excelOutput.save
end function

function getInformationFlowsHeaders()
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.add("L0 source") '0
	headers.add("L1 source") '1
	headers.Add("L2 source") '2
	headers.Add("Conveyed") '3
	headers.Add("Conveyed Owner") '4
	headers.add("L2 target") '5
	headers.Add("L1 target") '6
	headers.Add("L0 target") '7
	set getInformationFlowsHeaders = headers
end function


function getInformationFlowsSQLGetData(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	getInformationFlowsSQLGetData = "select l0.name as L0, l1.name as L1, l2.name as L2,                                   " & vbNewLine & _
					" isnull (conv.Name , '<missing conveyed object> ' + isnull(c.Name, '')) as Conveyed   " & vbNewLine & _
					" ,null as ConveyedOwner                                                               " & vbNewLine & _
					" , l2t.Name as L2Target, l1t.name as L1Target, l0t.Name as L0Target                   " & vbNewLine & _
					" from t_object l2                                                                     " & vbNewLine & _
					" inner join t_connector c on c.Start_Object_ID = l2.Object_ID                         " & vbNewLine & _
					" 							and c.Connector_Type = 'InformationFlow'                   " & vbNewLine & _
					" left join t_xref x on x.Client = c.ea_guid                                           " & vbNewLine & _
					" 					and x.Name = 'MOFProps'                                            " & vbNewLine & _
					" 					and x.Type = 'connector property'                                  " & vbNewLine & _
					" 					and x.Behavior = 'conveyed'                                        " & vbNewLine & _
					" left join t_object conv on x.Description like '%' + conv.ea_guid + '%'               " & vbNewLine & _
					" inner join t_objectproperties tvl on tvl.Object_ID = l2.Object_ID                    " & vbNewLine & _
					" 									and tvl.Property = 'Level'                         " & vbNewLine & _
					" 									and tvl.Value = 'L2'                               " & vbNewLine & _
					" left join (select l2l1.End_Object_ID, l1.Object_ID, l1.name                          " & vbNewLine & _
					" 			from t_connector l2l1                                                      " & vbNewLine & _
					" 			inner join t_object l1 on l1.Object_ID = l2l1.Start_Object_ID              " & vbNewLine & _
					" 			where l2l1.Stereotype = 'ArchiMate_Composition'                            " & vbNewLine & _
					" 			) l1 on l1.End_Object_ID = l2.Object_ID                                    " & vbNewLine & _
					" left join (select l1l0.End_Object_ID, l0.Object_ID , l0.Name                         " & vbNewLine & _
					" 			from t_connector l1l0                                                      " & vbNewLine & _
					" 			inner join t_object l0 on l0.Object_ID = l1l0.Start_Object_ID              " & vbNewLine & _
					" 			where l1l0.Stereotype = 'ArchiMate_Composition'                            " & vbNewLine & _
					" 			) l0 on l0.End_Object_ID = l1.Object_ID                                    " & vbNewLine & _
					" inner join t_object l2t on l2t.Object_ID = c.End_Object_ID                           " & vbNewLine & _
					" 							and l2t.Stereotype =  'Archimate_Capability'               " & vbNewLine & _
					" left join (select l2l1.End_Object_ID, l1.Object_ID, l1.name                          " & vbNewLine & _
					" 			from t_connector l2l1                                                      " & vbNewLine & _
					" 			inner join t_object l1 on l1.Object_ID = l2l1.Start_Object_ID              " & vbNewLine & _
					" 			where l2l1.Stereotype = 'ArchiMate_Composition'                            " & vbNewLine & _
					" 			) l1t on l1t.End_Object_ID = l2t.Object_ID                                 " & vbNewLine & _
					" left join (select l1l0.End_Object_ID, l0.Object_ID , l0.Name                         " & vbNewLine & _
					" 			from t_connector l1l0                                                      " & vbNewLine & _
					" 			inner join t_object l0 on l0.Object_ID = l1l0.Start_Object_ID              " & vbNewLine & _
					" 			where l1l0.Stereotype = 'ArchiMate_Composition'                            " & vbNewLine & _
					" 			) l0t on l0t.End_Object_ID = l1t.Object_ID                                 " & vbNewLine & _
					" where l2.Stereotype = 'Archimate_Capability'                                         " & vbNewLine & _
					" and (l2.Package_ID in (" & packageTreeIDString & ")                                  " & vbNewLine & _
					" or  l2t.Package_ID in (" & packageTreeIDString & ")) 								   " & vbNewLine & _
					" union                                                                                " & vbNewLine & _
					" select l0.name as L0, l1.name as L1, l2.name as L2,                                  " & vbNewLine & _
					" conv2.Name as Conveyed, conv.Name as ConveyedOwner                                   " & vbNewLine & _
					" , l2t.Name as L2Target, l1t.name as L1Target, l0t.Name as L0Target                   " & vbNewLine & _
					" from t_object l2                                                                     " & vbNewLine & _
					" inner join t_connector c on c.Start_Object_ID = l2.Object_ID                         " & vbNewLine & _
					" 							and c.Connector_Type = 'InformationFlow'                   " & vbNewLine & _
					" left join t_xref x on x.Client = c.ea_guid                                           " & vbNewLine & _
					" 					and x.Name = 'MOFProps'                                            " & vbNewLine & _
					" 					and x.Type = 'connector property'                                  " & vbNewLine & _
					" 					and x.Behavior = 'conveyed'                                        " & vbNewLine & _
					" left join t_object conv on x.Description like '%' + conv.ea_guid + '%'               " & vbNewLine & _
					" inner join t_objectproperties tvl on tvl.Object_ID = l2.Object_ID                    " & vbNewLine & _
					" 									and tvl.Property = 'Level'                         " & vbNewLine & _
					" 									and tvl.Value = 'L2'                               " & vbNewLine & _
					" left join (select l2l1.End_Object_ID, l1.Object_ID, l1.name                          " & vbNewLine & _
					" 			from t_connector l2l1                                                      " & vbNewLine & _
					" 			inner join t_object l1 on l1.Object_ID = l2l1.Start_Object_ID              " & vbNewLine & _
					" 			where l2l1.Stereotype = 'ArchiMate_Composition'                            " & vbNewLine & _
					" 			) l1 on l1.End_Object_ID = l2.Object_ID                                    " & vbNewLine & _
					" left join (select l1l0.End_Object_ID, l0.Object_ID , l0.Name                         " & vbNewLine & _
					" 			from t_connector l1l0                                                      " & vbNewLine & _
					" 			inner join t_object l0 on l0.Object_ID = l1l0.Start_Object_ID              " & vbNewLine & _
					" 			where l1l0.Stereotype = 'ArchiMate_Composition'                            " & vbNewLine & _
					" 			) l0 on l0.End_Object_ID = l1.Object_ID                                    " & vbNewLine & _
					" inner join t_object l2t on l2t.Object_ID = c.End_Object_ID                           " & vbNewLine & _
					" 							and l2t.Stereotype =  'Archimate_Capability'               " & vbNewLine & _
					" left join (select l2l1.End_Object_ID, l1.Object_ID, l1.name                          " & vbNewLine & _
					" 			from t_connector l2l1                                                      " & vbNewLine & _
					" 			inner join t_object l1 on l1.Object_ID = l2l1.Start_Object_ID              " & vbNewLine & _
					" 			where l2l1.Stereotype = 'ArchiMate_Composition'                            " & vbNewLine & _
					" 			) l1t on l1t.End_Object_ID = l2t.Object_ID                                 " & vbNewLine & _
					" left join (select l1l0.End_Object_ID, l0.Object_ID , l0.Name                         " & vbNewLine & _
					" 			from t_connector l1l0                                                      " & vbNewLine & _
					" 			inner join t_object l0 on l0.Object_ID = l1l0.Start_Object_ID              " & vbNewLine & _
					" 			where l1l0.Stereotype = 'ArchiMate_Composition'                            " & vbNewLine & _
					" 			) l0t on l0t.End_Object_ID = l1t.Object_ID                                 " & vbNewLine & _
					" inner join t_object conv2 on conv2.ParentID = conv.Object_ID                         " & vbNewLine & _
					"                             and conv2.Object_Type = 'Class'                          " & vbNewLine & _
					" where l2.Stereotype = 'Archimate_Capability'                                         " & vbNewLine & _
					" and (l2.Package_ID in (" & packageTreeIDString & ")                                  " & vbNewLine & _
					" or  l2t.Package_ID in (" & packageTreeIDString & ")) 								   " & vbNewLine & _
					" order by L0, l1, l2, Conveyed, L0Target, L1Target, L2Target                          "
end function
