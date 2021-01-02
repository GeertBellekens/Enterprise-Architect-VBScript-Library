'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Create UMIG CodeList Export
' Author: Geert Bellekens
' Purpose: Create an export of all enumeration values used in UMIG
' Date: 2018-06-07
'
sub main
	'get the results
	dim sqlGetCodeList
	sqlGetCodeList = getSQLGetCodeList()
	dim codeListResults
	set codeListResults = getArrayListFromQuery(sqlGetCodeList)
	dim headers
	set headers = getHeaders()
	'add the headers to the results
	codeListResults.Insert 0, headers
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	'create a two dimensional array
	dim excelContents
	excelContents = makeArrayFromArrayLists(codeListResults)
	'add the output to a sheet in excel
	excelOutput.createTab "Code Lists", excelContents, true, "TableStyleMedium4"
	'save the excel file
	excelOutput.save
end sub


function getSQLGetCodeList()
	getSQLGetCodeList = "select distinct o.Name as Enumeration,'''' + a.Name as Value, isnull(acd.VALUE, inh.CodeName) as Description,    " & _
						" inh.CodeListAgency as ListAgencyIdentifier, inh.CodeListID as ListIdentifier                                     " & _
						"  from t_object o                                                                                                 " & _  
						"  inner join t_objectproperties oca on oca.Object_ID = o.Object_ID                                                " & _  
						"  									and oca.Property = 'CodeListAgencyID'                                          " & _ 
						"  inner join t_objectproperties oid on oid.Object_ID = o.Object_ID                                                " & _  
						"  									and oid.Property = 'uniqueID'                                                  " & _ 
						"  inner join t_objectproperties ovr on ovr.Object_ID = o.Object_ID                                                " & _  
						"  									and ovr.Property = 'versionID'                                                 " & _ 
						"  left join t_attribute a on a.Object_ID = o.Object_ID                                                            " & _  
						"  left join t_attributetag acd on acd.ElementID = a.ID                                                            " & _  
						"  								and acd.Property = 'CodeName'                                                      " & _ 
						"  left join (select gen.Start_Object_ID, og.Object_ID, og.Name as Enumeration, ag.Name as EnumValue,              " & _  
						"  			ogca.Value as CodeListAgency, ogid.Value as CodeListID, agcd.Value as CodeName                         " & _   
						"  			from t_connector gen 				                                                                   " & _ 
						"  			inner join t_object og on og.Object_ID = gen.End_Object_ID                                             " & _ 
						"  			inner join t_attribute ag on ag.Object_ID = og.Object_ID                                               " & _ 
						"  			inner join t_objectproperties ogca on ogca.Object_ID = og.Object_ID                                    " & _ 
						"  												and ogca.Property = 'CodeListAgencyID'                             " & _ 
						"  			left  join t_objectproperties ogid on ogid.Object_ID = og.Object_ID                                    " & _ 
						"  												and ogid.Property = 'CodeListID'                                   " & _   
						"  			left join t_attributetag agcd on agcd.ElementID = ag.ID                                                " & _ 
						"  								and agcd.Property = 'CodeName'                                                     " & _ 
						"  			where  gen.Connector_Type = 'Generalization' ) inh                                                     " & _ 
						"  						on inh.Start_Object_ID = o.Object_ID                                                       " & _ 
						"  						and inh.EnumValue = a.Name                                                                 " & _ 
						"  where o.Object_Type = 'Enumeration'                                                                             " & _  
						"  and exists (select c.Connector_ID from t_connector c                                                            " & _  
						"  			inner join t_object sso on sso.Object_ID = c.Start_Object_ID                                           " & _ 
						"  							       and sso.Name = o.Name                                                           " & _ 
						"  								   and sso.Object_Type = o.Object_Type                                             " & _ 
						"  			inner join t_package ssp on ssp.Package_ID = sso.Package_ID                                            " & _ 
						"  			inner join t_object sspo on sspo.ea_guid = ssp.ea_guid                                                 " & _ 
						"  								  and sspo.Stereotype = 'DOCLibrary'                                               " & _ 
						"  			where c.End_Object_ID = o.Object_ID                                                                    " & _ 
						"  			and c.Connector_Type = 'Abstraction'                                                                   " & _ 
						"  			and c.Stereotype = 'trace')                                                                            " & _ 
						"  order by Enumeration, Value, ListAgencyIdentifier, ListIdentifier                                               "
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
