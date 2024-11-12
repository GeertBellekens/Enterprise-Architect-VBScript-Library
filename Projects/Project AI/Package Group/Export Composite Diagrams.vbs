'[path=\Projects\Project AI\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Export Composite diagrams
' Author: Geert Bellekens
' Purpose: Generate a CSV file with for each composite diagram
' 		- element name
' 		- element stereotype
' 		- element guid
' 		- linked diagram name
' 		- linked diagram type
' 		- linked diagram guid
' Date: 2024-09-20
'

'name of the output tab
const outPutName = "Export Composite Diagram"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'set timestamp
	Repository.WriteOutput outPutName, now() &  " Starting " & outPutName , 0
	'Actually do the thing
	exportCompositeDiagrams
	'set timestamp
	Repository.WriteOutput outPutName, now() &  " Finished " & outPutName , 0
	
end sub

function exportCompositeDiagrams()
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	'check if package found
	if package is nothing then
		exit function
	end if
	'get application components
	dim data
	set data = getCompositeDiagramData(package)
	'add headers
	dim headers
	set headers = getHeaders()
	data.Insert 0, headers
	'export to excel CSVFile
	dim csvFile
	set csvFile = new CSVFile
	csvFile.Contents = data
	csvFile.Save
end function

function getCompositeDiagramData(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	dim sqlGetData
	sqlGetData =  "select cd.o_name, cd.o_stereotype, cd.o_ea_guid, cd.d_name, coalesce( d_MDGType, cd.d_type) as d_type, cd.d_ea_guid                                                                                  " & vbNewLine & _
				" from t_object o                                                                                                                                                                                     " & vbNewLine & _
				" inner join	                                                                                                                                                                                      " & vbNewLine & _
				" 	(select o.name as o_name, o.Stereotype as o_stereotype, o.ea_guid as o_ea_guid                                                                                                                    " & vbNewLine & _
				" 	, d.Name as d_name, d.Diagram_Type as d_type, d.ea_guid as d_ea_guid                                                                                                                              " & vbNewLine & _
				" 	,case when CHARINDEX('MDGDgm=', d.styleEx)> 0 and CHARINDEX(';', d.styleEx, CHARINDEX('MDGDgm=', d.styleEx)) > 0                                                                                  " & vbNewLine & _
				" 		then SUBSTRING(d.styleEx, CHARINDEX('MDGDgm=', d.styleEx) + LEN('MDGDgm='), CHARINDEX(';', d.styleEx, CHARINDEX('MDGDgm=', d.styleEx)) - CHARINDEX('MDGDgm=', d.styleEx) - LEN('MDGDgm='))    " & vbNewLine & _
				" 		else NULL end AS d_MDGType                                                                                                                                                                    " & vbNewLine & _
				" 	from t_object o                                                                                                                                                                                   " & vbNewLine & _
				" 	inner join t_diagram d on CONVERT(nvarchar(255), d.Diagram_ID) = o.PDATA1                                                                                                                         " & vbNewLine & _
				" 						and o.NType = 8                                                                                                                                                               " & vbNewLine & _
				" 	union                                                                                                                                                                                             " & vbNewLine & _
				" 	select o.name as o_name, o.Stereotype as o_stereotype, o.ea_guid as o_ea_guid                                                                                                                     " & vbNewLine & _
				" 	, d.Name as d_name, d.Diagram_Type as d_type, d.ea_guid as d_ea_guid                                                                                                                              " & vbNewLine & _
				" 	,case when CHARINDEX('MDGDgm=', d.styleEx)> 0 and CHARINDEX(';', d.styleEx, CHARINDEX('MDGDgm=', d.styleEx)) > 0                                                                                  " & vbNewLine & _
				" 		then SUBSTRING(d.styleEx, CHARINDEX('MDGDgm=', d.styleEx) + LEN('MDGDgm='), CHARINDEX(';', d.styleEx, CHARINDEX('MDGDgm=', d.styleEx)) - CHARINDEX('MDGDgm=', d.styleEx) - LEN('MDGDgm='))    " & vbNewLine & _
				" 		else NULL end AS d_MDGType                                                                                                                                                                    " & vbNewLine & _
				" 	from t_object o                                                                                                                                                                                   " & vbNewLine & _
				" 	inner join t_xref x on x.Client = o.ea_guid                                                                                                                                                       " & vbNewLine & _
				" 						and x.Name = 'DefaultDiagram'                                                                                                                                                 " & vbNewLine & _
				" 	inner join t_diagram d on d.ea_guid = x.Supplier                                                                                                                                                  " & vbNewLine & _
				" 	) cd on cd.o_ea_guid = o.ea_guid                                                                                                                                                                  " & vbNewLine & _
				" where o.Package_ID in (" & packageTreeIDString & ")                                                                                                                                                 " & vbNewLine & _
				" order by cd.o_name                                                                                                                                                                                  "
	dim data
	set data = getArrayListFromQuery(sqlGetData)
	'return
	set getCompositeDiagramData = data
end function

Private Function getHeaders()
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.add("Element Name")
	headers.add("Element Stereotype")
	headers.Add("Element GUID")
	headers.Add("Diagram Name")
	headers.Add("Diagram Type")
	headers.Add("Diagram GUID")
	set getHeaders = headers
end Function

main