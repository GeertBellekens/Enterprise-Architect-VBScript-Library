'[path=\Framework\Utils]
'[group=Utils]

'Name: ExcelFile
'Author: Geert Bellekens
'Purpose: Wrapper script class for Excel files
'Date: 2017-03-20

!INC Utils.Include

const xlCalculationAutomatic	= -4105	'Excel controls recalculation.
const xlCalculationManual	= -4135	'Calculation is done when the user requests it.
const xlCenter = -4108
const xlLeft = -4131
const xlBelow = 1
const xlAbove = 0

Class ExcelFile
	'private variables
	Private m_ExcelApp
	Private m_FileName
	Private m_WorkBook
	private m_isExisting

	Private Sub Class_Initialize
		m_FileName = ""
		set m_ExcelApp = CreateObject("Excel.Application")
		set m_WorkBook = nothing
		m_isExisting = false
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
	
	public function freezePanes(ws, row, column)
		'select the worksheet
		ws.Activate
		m_ExcelApp.ActiveWindow.SplitRow = row
		m_ExcelApp.ActiveWindow.SplitColumn = column
		m_ExcelApp.ActiveWindow.FreezePanes = true
	end function
	
	
	private function createNewTab(tabName,beforeSheetIndex)
		'check if the workbook has been created already
		if m_WorkBook is nothing then
			set m_WorkBook = m_ExcelApp.Workbooks.Add()
		end if
		Dim ws
		set ws = nothing
		dim currentWs
		'check if it exists
		for each currentWs in m_WorkBook.Sheets
			if currentWs.Name = tabName then
				set ws = currentWs
				exit for
			end if
		next
		'if not exist yet then create
		if ws is nothing then
			'check the beforeIndex. In -1 then add in the back
			if beforeSheetIndex > 0 and beforeSheetIndex <= m_Workbook.Sheets.Count then
				Set ws = m_WorkBook.Sheets.Add(m_Workbook.Sheets(beforeSheetIndex)) 'add before the given sheetIndex
			else
				Set ws = m_WorkBook.Sheets.Add(,m_Workbook.Sheets(m_Workbook.Sheets.Count)) 'add after the last one
			end if
			ws.Name = tabName
		end if
		'return
		set createNewTab = ws
	end function 
		
	'public operations
	'create a tab with the given name. The contents should parameter should be a two dimensional array
	'anything int he contents that starts with "=" will be interpreted as a formula
	public Function createTabWithFormulas(tabName, contents,formatAsTable, tableStyle, beforeSheetIndex)
		'turn off automatic calculation
		m_ExcelApp.Calculation = xlCalculationManual
		'create the tab
		Dim ws
		set ws = createNewTab(tabName,beforeSheetIndex)
		'fill the contents
		'loop content
		dim i
		dim j
		for i = 0 to Ubound(contents,1)
			for j = 0 to Ubound(Contents,2)
				dim cellValue
				cellValue = contents(i,j)
				if left(cellValue,1) = "=" then
					ws.Cells(i + 1,j + 1).Formula = cellValue
				else
					ws.Cells(i + 1,j + 1).Value = cellValue
				end if
			next
		next
		dim targetRange
		set targetRange = ws.Range(ws.Cells(1,1), ws.Cells(Ubound(contents,1) +1, Ubound(Contents,2) +1))
		'format as table if needed
		if formatAsTable then
			formatSheetAsTable ws, targetRange, tableStyle
		end if
		'turn on automatic calculation
		m_ExcelApp.Calculation = xlCalculationAutomatic
	end function
	'public operations
	'create a tab with the given name. The contents should parameter should be a two dimensional array
	public Function createTab(tabName, contents,formatAsTable, tableStyle)
		'return 
		set createTab = createTabAtIndex(tabName, contents,formatAsTable, tableStyle, -1)
	end function
	
	public function createTabAtIndex(tabName, contents,formatAsTable, tableStyle, beforeSheetIndex)
		'create the tab
		Dim ws
		set ws = createNewTab(tabName,beforeSheetIndex)
		'fill the contents
		dim targetRange
		set targetRange = ws.Range(ws.Cells(1,1), ws.Cells(Ubound(contents,1) +1, Ubound(Contents,2) +1))
		targetRange.Value2 = contents
		'format as table if needed
		if formatAsTable then
			formatSheetAsTable ws, targetRange, tableStyle
		end if
		'set autofit
		targetRange.Columns.Autofit
		'return 
		set createTabAtIndex = ws
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
		me.FileName = project.GetFileNameDialog ("", "Excel Files|*.xls;*.xlsx;*.xlsm|Excel Templates|*.xlt;*.xltx;*.xltm", 1, 0 ,"", 0) 'save as with overwrite prompt: OFN_OVERWRITEPROMPT
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
		if fso.FileExists(me.FileName) then
			'check the extension
			dim extension
			extension = lcase(fso.GetExtensionName(me.FileName))
			select case extension
				case "xlt","xltx","xltm"
					me.NewFile me.FileName
				case else
					me.Open me.FileName
					m_isExisting = true
			end select
		end if
	end function
	
	public function Open(filePath)
		me.FileName = filePath
		set m_WorkBook = m_ExcelApp.Workbooks.Open(me.FileName)
	end function
	
	public function NewFile(filePath)
		dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
		if fso.FileExists(filePath) then
			set m_WorkBook = m_ExcelApp.Workbooks.Add(filePath)
		end if
		'reset filename
		me.FileName = ""
	end function
	
	public function formatRange (range, backColor, fontColor, fontName, fontSize, bold, horizontalAlignment)
		if backColor <> "default" then 
			range.Interior.Color = backColor 
		end if
		if fontColor <> "default" then
			range.Font.Color = fontColor
		end if
		if fontName <> "default" then
			range.Font.Name = fontName
		end if
		if fontSize <> "default" then
			range.Font.Size = fontSize
		end if
		if horizontalAlignment <> "default" then
			range.HorizontalAlignment = horizontalAlignment
		end if
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
		if m_isExisting then
			m_WorkBook.Save
		else
			'Delete the existing file if it exists
			dim fso
			Set fso = CreateObject("Scripting.FileSystemObject")
			if fso.FileExists(me.FileName) then
				fso.DeleteFile me.FileName
			end if
			'save the workbook at the given filename
			m_WorkBook.Saveas me.FileName
		end if
		'make excel visible
		m_ExcelApp.visible = True
		m_ExcelApp.WindowState = -4137 'xlMaximized
	end function
	
	
end Class