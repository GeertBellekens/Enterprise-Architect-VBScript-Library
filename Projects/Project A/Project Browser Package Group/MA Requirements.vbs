'[path=\Projects\Project A\Project Browser Package Group]
'[group=Project Browser Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: M&A Requirements
' Author: Matthias Van der Elst
' Purpose: Lists al the M&A Requirements in the selected package, create an Excel report.
' Date: 2017-06-02

const outputTabName = "M&A Requirements"

sub main
	Repository.CreateOutputTab outputTabName
	Repository.ClearOutput outputTabName
	Repository.EnsureOutputVisible outputTabName
	
	'get the package id's of the currently selected package tree
	dim currentPackageTreeIDString
	currentPackageTreeIDString = getCurrentPackageTreeIDString()
	
	dim arr1, arr2, arr3
	dim row
	
	redim row(7)
	row(0) = "ID Application Requirement"
	row(1) = "Context"
	row(2) = "Title"
	row(3) = "Description"
	row(4) = "Category"
	row(5) = "Domain"
	row(6) = "Importance"
	Repository.WriteOutput outputTabName, now() & " Starting List M&A Requirements" ,0
	arr1 = getRequirements(currentPackageTreeIDString)
	Repository.WriteOutput outputTabName, now() & " Starting Fixing Array M&A Requirements" ,0
	arr1 = addRowToArray(arr1,row)
	Repository.WriteOutput outputTabName, now() & " Finished List M&A Requirements" ,0
	
	redim row(2)
	row(0) = "ID Application Requirement"
	row(1) = "ID Use Case"
	Repository.WriteOutput outputTabName, now() & " Starting List M&A Use Cases" ,0
	arr2 = getUseCases(currentPackageTreeIDString)
	Repository.WriteOutput outputTabName, now() & " Starting Fixing Array M&A Use Cases" ,0
	arr2 = addRowToArray(arr2, row) 
	Repository.WriteOutput outputTabName, now() & " Finished List M&A Use Cases" ,0
	
	redim row(3)
	row(0) = "ID Application Requirement"
	row(1) = "ID Application Interface"
	row(2) = "Note Realize Relationship"
	Repository.WriteOutput outputTabName, now() & " Starting List M&A Application Interfaces" ,0
	arr3 = getInterfaces(currentPackageTreeIDString)
	Repository.WriteOutput outputTabName, now() & " Starting Fixing Array M&A Application Interfaces" ,0
	arr3 = addRowToArray(arr3, row) 
	Repository.WriteOutput outputTabName, now() & " Finished List M&A Application Interfaces" ,0
	
	Repository.WriteOutput outputTabName, now() & " Starting Saving To Excel" ,0
	saveToExcel arr1, arr2, arr3
	Repository.WriteOutput outputTabName, now() & " Finished Saving To Excel" ,0
	
end sub

function addRowToArray(array(), row())
	dim headers, columns, rows
	headers = ubound(row) 'Headers to be added
	'columns = ubound(array, 2) 'Columns in results
	On Error Resume Next
	rows = ubound(array,1) 'Rows in results
	If Err.Number <> 0 Then
		rows = 0
	End If
		
	dim arr()
	redim arr(rows + 1, headers)
	
	'Fill headers
	dim i, j
	for i = 0 to headers
		arr(0,i) = row(i)
	next
	'Fill the rest of the array
	for i = 1 to rows	
		for j = 0 to headers
			arr(i,j) = array(i-1,j)
		next	
	next
	addRowToArray = arr

end function

function getRequirements(currentPackageTreeIDString)
	dim getRequirementsSQL
	getRequirementsSQL = 	"select o.name as IDApplicationRequirement, opContext.Value as Context, opTitle.Value as Title, o.Note as Description, opCat.Value as Category,  " & _ 
							"opDomain.Value as Domain, opImp.Value as Importance  " & _
							"from (((((t_object o " & _
							"inner join t_objectproperties opContext  " & _
							"on o.Object_ID = opContext.Object_ID)  " & _
							"inner join t_objectproperties opTitle  " & _
							"on o.Object_ID = opTitle.Object_ID)  " & _
							"inner join t_objectproperties opCat  " & _
							"on o.Object_ID = opCat.Object_ID)  " & _
							"inner join t_objectproperties opDomain  " & _
							"on o.Object_ID = opDomain.Object_ID)  " & _
							"inner join t_objectproperties opImp  " & _
							"on o.Object_ID = opImp.Object_ID)  " & _
							"where o.Package_ID in (" & currentPackageTreeIDString & ") " & _
							"and o.Object_Type = 'Requirement'  " & _
							"and opContext.Property = 'Context'  " & _
							"and opTitle.Property = 'Title'  " & _
							"and opCat.Property = 'Category'  " & _
							"and opDomain.Property = 'Domain'  " & _
							"and opImp.Property = 'Importance' "				
								
	getRequirements = getArrayFromQuery(getRequirementsSQL)
end function

function getUseCases(currentPackageTreeIDString)
	dim getUseCasesSQL
	getUseCasesSQL = 	"select o.Name as req, uc.Name as uc " & _
						"from ((t_object o " & _
						"inner join t_connector con " & _
						"on o.Object_ID = con.End_Object_ID) " & _
						"inner join t_object uc " & _
						"on con.Start_Object_ID = uc.Object_ID) " & _
						"where o.Package_ID in (" & currentPackageTreeIDString & ") " & _
						"and o.Object_Type = 'Requirement' " & _
						"and uc.Object_Type = 'UseCase'"
								
	getUseCases = getArrayFromQuery(getUseCasesSQL)
end function

function getInterfaces(currentPackageTreeIDString)
	dim getInterfacesSQL
	getInterfacesSQL = 	"select o.name as IDApplicationRequirement, ai.Name as IDApplicationInterface, ai.Note as NoteRealizeRelationship " & _
						"from ((t_object o " & _
						"left join t_connector con " & _
						"on o.Object_ID = con.End_Object_ID) " & _
						"left join t_object ai " & _
						"on (con.Start_Object_ID = ai.Object_ID and ai.Stereotype = 'archimate_applicationinterface')) " & _
						"where ai.Package_ID in (" & currentPackageTreeIDString & ") " & _
						"and o.Object_Type = 'Requirement'"
	'Dim TextFile
	'set TextFile = new TextFile
	'TextFile.Contents = getInterfacesSQL
	'TextFile.FullPath = "\\intra.atrias.be\dfs\Data\Home\vanderelstm\M&A\output.txt"
	'TextFile.Save
								
	getInterfaces = getArrayFromQuery(getInterfacesSQL)
end function

function saveToExcel(arr1, arr2, arr3)
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	
	'create tab for the mapping application interface
	excelOutput.createTab "Mapping Application Interface", arr3, true, "TableStyleMedium4"
	
	'create tab for the mapping use cases
	excelOutput.createTab "Mapping Use Cases", arr2, true, "TableStyleMedium4"
	
	'create tab for the application requirements
	excelOutput.createTab "Application Requirements", arr1, true, "TableStyleMedium4"
	
	
	
	'save the excel file
	excelOutput.save
end function


main