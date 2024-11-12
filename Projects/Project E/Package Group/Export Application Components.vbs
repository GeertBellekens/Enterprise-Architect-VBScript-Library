'[path=\Projects\Project E\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Export Application Components
' Author: Geert Bellekens
' Purpose: Export Application Components
' Date: 2023-12-07
'
'name of the output tab
const outPutName = "Export Application Components"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'set timestamp
	Repository.WriteOutput outPutName, now() &  " Starting " & outPutName , 0
	'Actually do the thing
	exportApplicationComponents
	'set timestamp
	Repository.WriteOutput outPutName, now() &  " Finished " & outPutName , 0
	
end sub

function exportApplicationComponents()
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	'check if package found
	if package is nothing then
		exit function
	end if
	'get application components
	dim applicationComponentData
	set applicationComponentData = getApplicationComponentData(package)
	'export to excel file
	exportComponentsToExcel applicationComponentData
end function

function getApplicationComponentData(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	dim sqlGetData
	sqlGetData =  "select o.Name, o.Alias, o.Note, tve.Value as Eigenaar, tvt.Value as Type    " & vbNewLine & _
				" from t_object o                                                            " & vbNewLine & _
				" inner join t_objectproperties tve on tve.Object_ID = o.Object_ID           " & vbNewLine & _
				" 									and tve.Property = 'Eigenaar'            " & vbNewLine & _
				" inner join t_objectproperties tvt on tvt.Object_ID = o.Object_ID           " & vbNewLine & _
				" 									and tvt.Property = 'Type'                " & vbNewLine & _
				" where o.Stereotype = 'EDSN_Applicatie'                                     " & vbNewLine & _ 
				" and o.Package_ID in ( " & packageTreeIDString & ")"
	dim data
	set data = getArrayListFromQuery(sqlGetData)
	'return
	set getApplicationComponentData = data
end function

function exportComponentsToExcel(applicationComponentData)
	'convert description to plain TXT
	dim row
	for each row in applicationComponentData
		row(2) = Repository.GetFormatFromField("TXT", row(2))
	next
	dim headers
	set headers = getHeaders()
	'add the headers to the results
	applicationComponentData.Insert 0, headers
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	'create a two dimensional array
	dim excelContents
	excelContents = makeArrayFromArrayLists(applicationComponentData)
	'add the output to a sheet in excel
	excelOutput.createTab "Applications", excelContents, true, "TableStyleMedium4"
	'save the excel file
	excelOutput.save
end function

function getHeaders()
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.add("Application Name")
	headers.add("Alias")
	headers.Add("Description")
	headers.Add("Eigenaar")
	headers.Add("Type")
	set getHeaders = headers
end function

function getApplicationComponents(package)
	dim applicationComponents
	set applicationComponents = nothing
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	dim sqlGetData
	sqlGetData =  "select o.Object_ID from t_object o                       " & vbNewLine & _
				  " where o.Stereotype = 'ArchiMate_ApplicationComponent'   " & vbNewLine & _
				  " and o.Package_ID in ( " & packageTreeIDString & ")"
	set applicationComponents = getElementsFromQuery(sqlGetData)
	dim element as EA.Element
	for each element in applicationComponents
		if element.FQStereotype = "ArchiMate3::ArchiMate_ApplicationComponent" then
			Repository.WriteOutput outPutName, now() &  " Found " & element.Name & " " & element.ElementGUID, 0
		end if
	next
	'return
	set getApplicationComponents = applicationComponents
end function



main