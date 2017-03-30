'[path=\Projects\Project EC\EAC scripts]
'[group=EAC scripts]

option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Export Programme Tree to Excel
' Author: Geert Bellekens
' Purpose: Exports the Programme tree to an excel file
' Date: 2017-03-30
'
'name of the output tab
const outPutName = "Export Programme Tree"

sub main

	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get the selected element
	msgbox "Please select the package that contains the program tree"
	dim selectedPackage as EA.Package
	set selectedPackage = selectPackage()
	if selectedPackage.ObjectType = otPackage then
		'tell the user we are starting
		Repository.WriteOutput outPutName, now() & " Starting Export Proramme Tree '" & selectedPackage.Name & "'", selectedPackage.Element.ElementID
		'do the actual export
		exportProgrammeTree(selectedPackage)
		'tell the user we are finished
		Repository.WriteOutput outPutName, now() & " Finished Export Proramme Tree '" & selectedPackage.Name & "'", selectedPackage.Element.ElementID
	end if
	
end sub

function exportProgrammeTree(selectedPackage)
	dim packageTreeIDs
	packageTreeIDs = getPackageTreeIDString(selectedPackage)
	dim getPogrammeTreeContents
	getPogrammeTreeContents = 	"select 'Programme' as Programme, 'Area' as Area ,'Key Action' as KeyAction, 'Action' as Action,'Action Type' as ActionType, 'Activity Type' as   " & _
								"  ActivityType,'Package Name' as PackageName, 'Package_1' as PackageLevel1, 'Package_2'as PackageLevel2 , 'Package_3' as PackageLevel3 		  " & _
								"union                                                                                                                                   		  " & _
								"select o.Name as Programme, ar.[Name] as Area, ka.Name as KeyAction, ac.Name as Action, at.Name as ActionType, avt.Name as ActivityType		  " & _
								" ,package.name as PackageName ,package_p1.name as PackageLevel1,package_p2.name as PackageLevel2 ,package_p3.name as PackageLevel3               " & _
								" from (((((((((((((( t_object o                                                                                                                  " & _
								" left join t_connector o_ar on (o_ar.[End_Object_ID] = o.[Object_ID]                                                                             " & _
								"                               and o_ar.[Connector_Type] in ('Association', 'Aggregation')))                                                     " & _
								" left join t_object ar on (o_ar.[Start_Object_ID] = ar.[Object_ID]                                                                               " & _
								"                          and ar.[Stereotype] = 'Area'))                                                                                         " & _
								" left join t_connector ar_ka on (ar_ka.[End_Object_ID] = ar.[Object_ID]                                                                          " & _
								"                               and ar_ka.[Connector_Type] in ('Association', 'Aggregation')))                                                    " & _
								" left join t_object ka on (ar_ka.[Start_Object_ID] = ka.[Object_ID]                                                                              " & _
								"                          and ka.[Stereotype] = 'Key Action'))                                                                                   " & _
								" left join t_connector ka_ac on (ka_ac.[End_Object_ID] = ka.[Object_ID]                                                                          " & _
								"                               and ka_ac.[Connector_Type] in ('Association', 'Aggregation')))                                                    " & _
								" left join t_object ac on (ka_ac.[Start_Object_ID] = ac.[Object_ID]                                                                              " & _
								"                          and ac.[Stereotype] = 'Action'))                                                                                       " & _
								" left join t_connector ac_at on (ac_at.[End_Object_ID] = ac.[Object_ID]                                                                          " & _
								"                               and ac_at.[Connector_Type] in ('Association', 'Aggregation')))                                                    " & _
								" left join t_object at on (ac_at.[Start_Object_ID] = at.[Object_ID]                                                                              " & _
								"                          and at.[Stereotype] = 'Action Type'))                                                                                  " & _
								" left join t_connector at_avt on (at_avt.[End_Object_ID] = at.[Object_ID]                                                                        " & _
								"                               and at_avt.[Connector_Type] in ('Association', 'Aggregation')))                                                   " & _
								" left join t_object avt on (at_avt.[Start_Object_ID] = avt.[Object_ID]                                                                           " & _
								"                          and avt.[Stereotype] = 'Activity Type'))     						                                                  " & _
								" inner join t_package package on o.package_id = package.package_id)                                                                              " & _
								" left join t_package package_p1 on package_p1.package_id = package.parent_id)                                                                    " & _
								" left join t_package package_p2 on package_p2.package_id = package_p1.parent_id)                                                                 " & _
								" left join t_package package_p3 on package_p3.package_id = package_p2.parent_id)                                                                 " & _
								" where o.Package_ID in (" & packageTreeIDs & ")                                                                                                  " & _
								" and o.[Stereotype] = 'Programme'                                                                                                                " & _
								" order by 1 desc                                                                                                              					  " 
	dim arrayResult 
	arrayResult = getArrayFromQuery(getPogrammeTreeContents)
	if Ubound(arrayResult) > 0 then
		'create the excel file
		dim excelOutput
		set excelOutput = new ExcelFile
		excelOutput.createTab "Programme Tree", arrayResult, true, "TableStyleMedium13"
		excelOutput.save
	end if
end function

main