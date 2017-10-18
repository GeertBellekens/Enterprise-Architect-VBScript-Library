'[path=\Framework\Utils]
'[group=Utils]

'Name: ExcelFile
'Author: Geert Bellekens
'Purpose: Wrapper script class for Excel files
'Date: 2017-03-20

!INC Utils.Include

Class ExcelFile
	'private variables
	Private m_ExcelApp
	Private m_FileName
	Private m_WorkBook

	Private Sub Class_Initialize
		m_FileName = ""
		set m_ExcelApp = CreateObject("Excel.Application")
		set m_WorkBook = nothing
	End Sub
	
	
	' FileName property.
	Public Property Get FileName
	  FileName = m_FileName
	End Property
	Public Property Let FileName(value)
	  m_FileName = value
	End Property
	public Property Get worksheets
		set worksheets = m_WorkBook.Sheets
	end property
	
	'public operations
	'create a tab with the given name. The contents should parameter should be a two dimensional array
	public Function createTab(tabName, contents,formatAsTable, tableStyle)
		'check if the workbook has been created already
		if m_WorkBook is nothing then
			set m_WorkBook = m_ExcelApp.Workbooks.Add()
		end if
		'create the tab at the end
		Dim ws
		Set ws = m_WorkBook.Sheets.Add()
		ws.Name = tabName
		'fill the contents
		dim targetRange
		set targetRange = ws.Range(ws.Cells(1,1), ws.Cells(Ubound(contents,1) +1, Ubound(Contents,2) +1))
		targetRange.Value2 = contents
		'format as table if needed
		if formatAsTable then
			formatSheetAsTable ws, targetRange, tableStyle
		end if
	end function
	
	public function formatSheetAsTable(worksheet, targetRange, tableStyle)
		dim table
		Set table = worksheet.ListObjects.Add(1, targetRange, 1, 1)
		table.TableStyle = tableStyle
	end function
	
	public Function getUserSelectedFileName()
		dim selectedFileName
		dim project
		set project = Repository.GetProjectInterface()
		me.FileName = project.GetFileNameDialog ("", "Excel Files|*.xls;*.xlsx;*.xlsm", 1, 2 ,"", 1) 'save as with overwrite prompt: OFN_OVERWRITEPROMPT
	end function
	
	public Function openUserSelectedFile()
		dim selectedFileName
		dim project
		set project = Repository.GetProjectInterface()
		me.FileName = project.GetFileNameDialog ("", "Excel Files|*.xls;*.xlsx;*.xlsm", 1, 0 ,"", 0) 'save as with overwrite prompt: OFN_OVERWRITEPROMPT
		me.Open me.FileName
	end function
	
	public function Open(filePath)
		me.FileName = filePath
		set m_WorkBook = m_ExcelApp.Workbooks.Open(me.FileName)
	end function
	
	public function getContents(sheet)
		getContents = sheet.UsedRange.Value2
	end function
	
	public Function save()
		'make sure we have a filename
		if len(me.FileName) = 0 then
			getUserSelectedFileName
		end if
		'if the file name is still empty then exit
		if len(me.FileName) = 0 then
			exit function
		end if
		'Delete the existing file if it exists
		dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
		if fso.FileExists(me.FileName) then
			fso.DeleteFile me.FileName
		end if
		'save the workbook at the given filename
		m_WorkBook.Saveas me.FileName
		'make excel visible
		m_ExcelApp.visible = True
		m_ExcelApp.WindowState = -4137 'xlMaximized
	end function
	
	
end Class