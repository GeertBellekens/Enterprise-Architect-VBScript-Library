'[path=\Framework\Wrappers\]
'[group=Wrappers]


!INC Utils.Include
!INC Local Scripts.EAConstants-VBScript


'
' Script Name: ExcelReport
' Author: Geert Bellekens
' Purpose: Class to facilitate creating excel reports
' Date: 2022-02-23
'
Class ExcelReport
	'private variables
	dim m_sqlQuery
	dim m_ExcelFile
	dim m_Headers
	
	Private Sub Class_Initialize
		set m_ExcelFile = new ExcelFile
		m_sqlQuery = ""
		m_headers = ""
	End Sub
	
	' ExcelFile property.
	Public Property Get WrappedExcelFile
	  Set WrappedExcelFile = m_ExcelFile
	End Property
	
	Public Function Save()
		'save the wrapped excel file
		m_ExcelFile.Save
	end Function
	
	Public function CreateReportFromQuery(sqlGetData, headers, template, sheetName, formatAsTable, tableStyle)
		dim data
		set data = getArrayListFromQuery(sqlGetData)
		'return
		set CreateReportFromQuery = CreateReportFromArrayLists(data, headers, template, sheetname, formatAsTable, tableStyle)
	end function
	
	Public function CreateReportFromArrayLists(data, headers, template, sheetName, formatAsTable, tableStyle)
		'add headers if needed
		if not headers is nothing then
			'add the headers to the results
			data.Insert 0, headers
		end if
		if len(template) > 0 then
			'load the template
			m_ExcelFile.NewFile template
		end if
		'create a two dimensional array
		dim excelContents
		excelContents = makeArrayFromArrayLists(data)
		'create the sheet
		dim sheet
		set sheet = m_ExcelFile.createTab(sheetName, excelContents, formatAsTable, tableStyle)
		'return sheet 
		set CreateReportFromArrayLists = sheet
	end function
	
End Class