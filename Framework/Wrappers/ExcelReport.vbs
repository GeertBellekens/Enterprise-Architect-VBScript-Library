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
	dim m_ExcelFile
	dim m_data
	dim m_sheet
	dim m_defaultSheetDeleted
	dim m_formattedTextColumns
	
	Private Sub Class_Initialize
		set m_ExcelFile = new ExcelFile
		set m_sheet = nothing
		set m_data = nothing
		m_defaultSheetDeleted = false
		m_formattedTextColumns = Array()
	End Sub
	
	public property Get FormattedTextColumns
		FormattedTextColumns = m_formattedTextColumns
	end Property
	public property Let FormattedTextColumns(value)
		m_formattedTextColumns = value
	end Property
	
	' ExcelFile property.
	Public Property Get WrappedExcelFile
	  Set WrappedExcelFile = m_ExcelFile
	End Property
	
	Public Function Save()
		'save the wrapped excel file
		m_ExcelFile.Save
	end Function
	
	Public function CreateReportFromQuery(sqlGetData, headers, template, sheetName, formatAsTable, tableStyle)
		'return
		set CreateReportFromQuery = CreateReportFromArrayLists(getArrayListFromQuery(sqlGetData), headers, template, sheetname, formatAsTable, tableStyle)
	end function
	
	Public function CreateReportFromArrayLists(data, headers, template, sheetName, formatAsTable, tableStyle)
		set m_data = data
		'add headers if needed
		if not headers is nothing then
			'add the headers to the results
			m_data.Insert 0, headers
		end if
		if len(template) > 0 then
			'load the template
			m_ExcelFile.NewFile template
		end if
		'format text columns
		formatTextColumns
		'create a two dimensional array
		dim excelContents
		excelContents = makeArrayFromArrayLists(m_data)
		'create the sheet
		set m_sheet = m_ExcelFile.createTab(sheetName, excelContents, formatAsTable, tableStyle)
		'delete the default first sheet
		if not m_defaultSheetDeleted then
			m_ExcelFile.deleteTabAtIndex 1
			m_defaultSheetDeleted = true
		end if
		'return sheet 
		set CreateReportFromArrayLists = m_sheet
	end function
	
	private function formatTextColumns()
		'only process if there are
		if not UBound(me.FormattedTextColumns) >= 0 then
			exit function
		end if
		dim row
		for each row in m_data
			dim columnIndex
			for each columnIndex in me.FormattedTextColumns
				dim notesText
				'get notes
				notesText = row(columnIndex)
				'format as Text
				notesText = Repository.GetFormatFromField("TXT",notesText)
				'put it back
				row(columnIndex) = notesText
			next
		next
	end function
	
	public function FormatHeaders(backColor, fontColor)
		dim headerRange
		set headerRange = m_sheet.Range(m_sheet.Cells(1,1), m_sheet.Cells(1, m_data(0).Count))
		m_ExcelFile.formatRange headerRange, backColor, fontColor, "default", "default", "default", "default"
	end function
	
	public function newFile(templateFilePath)
		m_ExcelFile.NewFile templateFilePath
		m_defaultSheetDeleted = true
	end function
	
End Class