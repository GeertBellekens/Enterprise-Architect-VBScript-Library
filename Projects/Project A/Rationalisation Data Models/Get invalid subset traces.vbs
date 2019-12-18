'[path=\Projects\Project A\Rationalisation Data Models]
'[group=Rationalisation Data Models]

option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Get invalid subset traces
' Author: Geert Bellekens
' Purpose: Get all traces that do not lead tot he master LDM
' Date: 2019-02-21
'
const LDMGUID = "{9CC085FC-8701-4aa4-8E6A-AC4530EB55C0}"

sub main
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage
	dim LDMPackage as EA.Package
	set LDMPackage = Repository.GetPackageByGuid(LDMGUID)
	'get the results
	dim sqlInvalidTraces
	sqlInvalidTraces = getSQLInvalidTraces(selectedPackage,LDMPackage )
	dim invalidTracesResults
	set invalidTracesResults = getArrayListFromQuery(sqlInvalidTraces)
	
	dim headers
	set headers = getHeaders()
	'add the headers to the results
	'create the output object
	dim searchOutput
	set searchOutput = new SearchResults
	searchOutput.Name = "Invalid Traces"
	searchOutput.Fields = headers
	'put the contents in the output
	dim row
	for each row in invalidTracesResults
		'add row the the output
		searchOutput.Results.Add row
	next
	'show the output
	searchOutput.Show
end sub


function getSQLInvalidTraces(selectedPackage, LDMPackage)
	dim selectedPackageIDString
	selectedPackageIDString = getPackageTreeIDString(selectedPackage)
	dim LDMPackageIDString
	LDMPackageIDString = getPackageTreeIDString(LDMPackage)
	getSQLInvalidTraces = "select o.ea_guid as CLASSGUID,o.Object_Type as CLASSTYPE,o.Name as Name,o.Stereotype, o.Name as Owner, '' as Otherend,               " & vbNewLine & _
							" package.name as PackageName ,package_p1.name as PackageLevel1,package_p2.name as PackageLevel2 ,package_p3.name as PackageLevel3  " & vbNewLine & _
							" from t_object o                                                                                                                   " & vbNewLine & _
							" left join (select c.Start_Object_ID, po.Object_ID from t_connector c                                                              " & vbNewLine & _
							" 		inner join t_object po on po.Object_ID = c.End_Object_ID                                                                    " & vbNewLine & _
							" 		where c.Connector_Type = 'Abstraction'                                                                                      " & vbNewLine & _
							" 		and c.Stereotype = 'trace'                                                                                                  " & vbNewLine & _
							" 		and po.Package_ID not in ("& LDMPackageIDString &")) parent on parent.Start_Object_ID = o.Object_ID                         " & vbNewLine & _
							" inner join t_package package on o.package_id = package.package_id                                                                 " & vbNewLine & _
							" left join t_package package_p1 on package_p1.package_id = package.parent_id                                                       " & vbNewLine & _
							" left join t_package package_p2 on package_p2.package_id = package_p1.parent_id                                                    " & vbNewLine & _
							" left join t_package package_p3 on package_p3.package_id = package_p2.parent_id                                                    " & vbNewLine & _
							" where o.Object_Type in ('Class', 'DataType', 'Enumeration', 'Datatype')                                                           " & vbNewLine & _
							" and parent.Object_ID is not null                                                                                                  " & vbNewLine & _
							" and o.Package_ID in ("& selectedPackageIDString &")                                                                               " & vbNewLine & _
							" union all                                                                                                                         " & vbNewLine & _
							" select a.ea_guid as CLASSGUID,'Attribute' as CLASSTYPE,a.Name as Name,a.Stereotype,o.Name as Owner, '' as Otherend,               " & vbNewLine & _
							" package.name as PackageName ,package_p1.name as PackageLevel1,package_p2.name as PackageLevel2 ,package_p3.name as PackageLevel3  " & vbNewLine & _
							" from t_attribute a                                                                                                                " & vbNewLine & _
							" left join (select tv.ElementID, pa.Object_ID, tv.VALUE from t_attributetag tv                                                     " & vbNewLine & _
							" 			inner join t_attribute pa on pa.ea_guid = tv.VALUE                                                                      " & vbNewLine & _
							" 			inner join t_object po on po.Object_ID = pa.Object_ID                                                                   " & vbNewLine & _
							" 			where tv.Property = 'sourceAttribute'                                                                                   " & vbNewLine & _
							" 			and po.Package_ID not in ("& LDMPackageIDString &"))  parent on  parent.ElementID = a.ID                                " & vbNewLine & _
							" inner join t_object o on o.Object_ID = a.Object_ID                                                                                " & vbNewLine & _
							" inner join t_package package on o.package_id = package.package_id                                                                 " & vbNewLine & _
							" left join t_package package_p1 on package_p1.package_id = package.parent_id                                                       " & vbNewLine & _
							" left join t_package package_p2 on package_p2.package_id = package_p1.parent_id                                                    " & vbNewLine & _
							" left join t_package package_p3 on package_p3.package_id = package_p2.parent_id                                                    " & vbNewLine & _
							" where parent.Object_ID is not null                                                                                                " & vbNewLine & _
							" and o.Package_ID in ("& selectedPackageIDString &")                                                                               " & vbNewLine & _
							" union all                                                                                                                         " & vbNewLine & _
							" select o.ea_guid as CLASSGUID,c.Connector_Type as CLASSTYPE,c.Name as Name,c.Stereotype,o.Name as Owner, oe.Name as Otherend,     " & vbNewLine & _
							" package.name as PackageName ,package_p1.name as PackageLevel1,package_p2.name as PackageLevel2 ,package_p3.name as PackageLevel3  " & vbNewLine & _
							" from t_connector c                                                                                                                " & vbNewLine & _
							" left join (select tv.ElementID, co.Start_Object_ID as Object_ID from t_connectorTag tv                                            " & vbNewLine & _
							" 			inner join t_connector co on co.ea_guid = tv.Value                                                                      " & vbNewLine & _
							" 			inner join t_object po on po.Object_ID = co.Start_Object_ID                                                             " & vbNewLine & _
							" 			where tv.Property = 'sourceAssociation'                                                                                 " & vbNewLine & _
							" 			and po.Package_ID not in ("& LDMPackageIDString &")) parent on  parent.ElementID = c.Connector_ID                       " & vbNewLine & _
							" inner join t_object o on o.Object_ID = c.Start_Object_ID                                                                          " & vbNewLine & _
							" inner join t_object oe on oe.Object_ID = c.End_Object_ID                                                                          " & vbNewLine & _
							" inner join t_package package on o.package_id = package.package_id                                                                 " & vbNewLine & _
							" left join t_package package_p1 on package_p1.package_id = package.parent_id                                                       " & vbNewLine & _
							" left join t_package package_p2 on package_p2.package_id = package_p1.parent_id                                                    " & vbNewLine & _
							" left join t_package package_p3 on package_p3.package_id = package_p2.parent_id                                                    " & vbNewLine & _
							" where parent.Object_ID is not  null                                                                                               " & vbNewLine & _
							" and c.Connector_Type in ('Association', 'Aggregation')                                                                            " & vbNewLine & _
							" and o.Package_ID in ("& selectedPackageIDString &")                                                                               "
end function

function getHeaders()
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.add("CLASSGUID")
	headers.add("CLASSTYPE")
	headers.add("Name")
	headers.add("Stereotype")
	headers.add("Owner")
	headers.add("Otherend")
	headers.add("PackageName")
	headers.add("PackageLevel1")
	headers.add("PackageLevel2")
	headers.add("PackageLevel3")

	set getHeaders = headers
end function

main