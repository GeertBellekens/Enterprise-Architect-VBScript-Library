'[path=\Projects\Project A\Reports and exports]
'[group=Reports and exports]

option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Create UMIG CodeList Export
' Author: Geert Bellekens
' Purpose: Create an export of all enumeration values used in UMIG
' Date: 2018-06-07
'

const excelTemplate = "G:\Projects\80 Enterprise Architect\Output\Excel export templates\UMIG - SD - XD - 05 - Code Lists.xltx"

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
	set sheet = excelOutput.createTab("Code Lists", excelContents, true, "TableStyleMedium4")
	'set headers to atrias red
	dim headerRange
	set headerRange = sheet.Range(sheet.Cells(1,1), sheet.Cells(1, headers.Count))
	excelOutput.formatRange headerRange, atriasRed, "default", "default", "default", "default", "default"
	'save the excel file
	excelOutput.save
end sub


function getSQLGetCodeList(packageTreeIDs)
	getSQLGetCodeList = "select distinct o.Name as Enumeration,'''' + a.Name as Value, isnull(acd.VALUE, inh.CodeName) as Description, " & vbNewLine & _
						" '''' + inh.CodeListAgency as ListAgencyIdentifier, '''' + inh.CodeListID as ListIdentifier                   " & vbNewLine & _
						" from t_object o                                                                                              " & vbNewLine & _
						" inner join (select c.Connector_ID,c.End_Object_ID, sso.Name as subName, sso.Object_Type as subObjectType     " & vbNewLine & _
						" 		, sso.Object_ID, sso.ea_guid                                                                           " & vbNewLine & _
						" 		from t_connector c                                                                                     " & vbNewLine & _
						" 		inner join t_object sso on sso.Object_ID = c.Start_Object_ID                                           " & vbNewLine & _           
						" 		inner join t_package ssp on ssp.Package_ID = sso.Package_ID                                            " & vbNewLine & _
						" 		inner join t_object sspo on sspo.ea_guid = ssp.ea_guid                                                 " & vbNewLine & _
						" 							  and sspo.Stereotype = 'DOCLibrary'                                               " & vbNewLine & _
						" 		where c.Connector_Type = 'Abstraction'                                                                 " & vbNewLine & _
						" 		and c.Stereotype = 'trace'                                                                             " & vbNewLine & _
						" 		and ssp.Package_ID in (" & packageTreeIDs & ")                                                         " & vbNewLine & _
						" 		) sub on sub.End_Object_ID = o.Object_ID                                                               " & vbNewLine & _
						" 				and sub.subObjectType = o.Object_Type                                                          " & vbNewLine & _                        
						" inner join t_objectproperties oca on oca.Object_ID = sub.Object_ID                                           " & vbNewLine & _
						" 								and oca.Property = 'CodeListAgencyID'                                          " & vbNewLine & _
						" inner join t_objectproperties oid on oid.Object_ID = sub.Object_ID                                           " & vbNewLine & _
						" 								and oid.Property = 'uniqueID'                                                  " & vbNewLine & _
						" inner join t_objectproperties ovr on ovr.Object_ID = sub.Object_ID                                           " & vbNewLine & _
						" 								and ovr.Property = 'versionID'                                                 " & vbNewLine & _
						" left join t_attribute a on a.Object_ID = sub.Object_ID                                                       " & vbNewLine & _
						" left join t_attributetag acd on acd.ElementID = a.ID                                                         " & vbNewLine & _
						" 							and acd.Property = 'CodeName'                                                      " & vbNewLine & _
						" left join (select gen.Start_Object_ID, og.Object_ID, og.Name as Enumeration, ag.Name as EnumValue,           " & vbNewLine & _
						" 		ogca.Value as CodeListAgency, ogid.Value as CodeListID, agcd.Value as CodeName                         " & vbNewLine & _
						" 		from t_connector gen 				                                                                   " & vbNewLine & _
						" 		inner join t_object og on og.Object_ID = gen.End_Object_ID                                             " & vbNewLine & _
						" 		inner join t_attribute ag on ag.Object_ID = og.Object_ID                                               " & vbNewLine & _
						" 		inner join t_objectproperties ogca on ogca.Object_ID = og.Object_ID                                    " & vbNewLine & _
						" 											and ogca.Property = 'CodeListAgencyID'                             " & vbNewLine & _
						" 		left  join t_objectproperties ogid on ogid.Object_ID = og.Object_ID                                    " & vbNewLine & _
						" 											and ogid.Property = 'CodeListID'                                   " & vbNewLine & _
						" 		left join t_attributetag agcd on agcd.ElementID = ag.ID                                                " & vbNewLine & _
						" 							and agcd.Property = 'CodeName'                                                     " & vbNewLine & _
						" 		where  gen.Connector_Type = 'Generalization'                                                           " & vbNewLine & _
						" 		) inh on inh.Start_Object_ID = o.Object_ID                                                             " & vbNewLine & _
						" 			     and inh.EnumValue = a.Name                                                                    " & vbNewLine & _
						" where o.Object_Type = 'Enumeration'                                                                          " & vbNewLine & _                                                      
						" order by Enumeration, Value, ListAgencyIdentifier, ListIdentifier                                            "
end function

function getHeaders()
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.add("Code List")
	headers.add("Value")
	headers.Add("Description")
	headers.Add("ListAgencyIdentifier")
	headers.Add("ListIdentifier")
	set getHeaders = headers
end function

main