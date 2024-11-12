'[path=\Projects\Project BF]
'[group=Belfius]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Export Application Portfolio
' Author: Geert Bellekens
' Purpose: Export applications
' Date: 2023-11-07

const outPutName = "Export Application PortFolio"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName, 0
	'Actual work
	exportApplications
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName, 0
end sub

function exportApplications()
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage
	dim applicationData
	set applicationData = getApplicationData(package)
	dim headers
	set headers = getHeaders()
	'add the headers to the results
	applicationData.Insert 0, headers
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	'create a two dimensional array
	dim excelContents
	excelContents = makeArrayFromArrayLists(applicationData)
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
	headers.Add("ApplicationID")
	headers.Add("Package Name")
	headers.Add("Version")
	headers.Add("Team")
	set getHeaders = headers
end function

function getApplicationData(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	dim sqlGetData
	sqlGetData = "select o.Name, o.Alias, o.Note, tvID.Value as ApplicationID, p.Name, o.Version, tvTID.Value as TeamID    " & vbNewLine & _
				" from t_object o                                                                                 " & vbNewLine & _
				" inner join t_objectproperties tvID on tvID.Object_ID = o.Object_ID                              " & vbNewLine & _
				" 							and tvID.Property = 'applicationID'                                   " & vbNewLine & _
				" inner join t_package p on p.Package_ID = o.Package_ID                                           " & vbNewLine & _
				" left join t_objectproperties tvTID on tvTID.Object_ID = o.Object_ID                             " & vbNewLine & _
				" 							and tvTID.Property = 'teamID'                                         " & vbNewLine & _
				" where o.Stereotype = 'ArchiMate_ApplicationComponent'                                           " & vbNewLine & _
				" and o.Package_ID in (" & packageTreeIDString & ")                                               "
	dim result
	set result = getArrayListFromQuery(sqlGetData)
	'return
	set getApplicationData = result
end function

main