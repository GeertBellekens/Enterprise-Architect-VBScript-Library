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
const xlUnderlineStyleSingle = 2
const xlUnderlineStyleNone = -4142	
const xlContinuous = 1
const xlThin = 2
const xlNone = -4142
const xlEdgeBottom = 9
const xlOpenXMLWorkbookMacroEnabled	= 52


Class ExcelFile
	'private variables
	Private m_ExcelApp
	Private m_FileName
	Private m_WorkBook
	private m_isExisting

	Private Sub Class_Initialize
		m_FileName = ""
		set m_ExcelApp = CreateObject("Excel.Application")
		m_ExcelApp.DisplayAlerts = false 'turn off alerts
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
	
	public function copyWorksheet(fileName, sheetName, targetSheet)
		dim copyWorkbook
		set copyWorkbook = m_ExcelApp.Workbooks.Open(fileName)
		dim copySheet
		set copySheet = nothing
		dim sheet
		for each sheet in copyWorkbook.Sheets
			if lcase(sheet.Name) = lcase(sheetName) then
				set copySheet = sheet
				exit for
			end if
		next
		'exit if not found
		if copySheet is nothing then
			Repository.WriteOutput outPutName, now() & " ERROR: could not find sheet " & sheetName & " in workbook " & fileName, 0
			copyWorkbook.Close
			exit function
		end if
		'find existing BusinessRulesSheet in this workbook
		dim businessRulesSheet
		dim currentSheet
		dim previousSheet
		set previousSheet = nothing 'initialize
		for each currentSheet in me.worksheets
			if lcase(trim(currentSheet.Name)) = lcase(targetSheet) then
				currentSheet.Delete
				exit for
			else
				set previousSheet = currentSheet
			end if
		next
		if previousSheet is nothing then
			set previousSheet = me.worksheets(1)
		end if
		'copy the business rules sheet after the previous sheet
		copySheet.Copy , previousSheet
		'rename sheet
		dim newSheet
		set newSheet = me.getTab(copySheet.Name)
		newSheet.Name = targetSheet
		'close the copy workbook again
		copyWorkbook.Close
	end function
	
	private function createNewTab(tabName)
		'check if the workbook has been created already
		if m_WorkBook is nothing then
			set m_WorkBook = m_ExcelApp.Workbooks.Add()
		end if
		Dim ws
		set ws = nothing
		dim currentWs
		'check if it exists
		for each currentWs in m_WorkBook.Sheets
			if lcase(currentWs.Name) = lcase(tabName) then
				set ws = currentWs
				exit for
			end if
		next
		'no exact match, check if a sheet exists with template name
		if ws is nothing then
			dim tabNameSuffix
			tabNameSuffix = lcase(getSuffix(tabName))
			for each currentWs in m_WorkBook.Sheets
				if left(currentWs.Name, 1) = "%" then
					'compare suffixes
					dim wsSuffix
					wsSuffix = lcase(getSuffix(currentWs.Name))
					if wsSuffix = tabNameSuffix then
						set ws = currentWs
						'rename the tab
						ws.Name = tabName
						exit for
					end if
				end if
			next
		end if
		'if not exist yet then create
		if ws is nothing then
			Set ws = m_WorkBook.Sheets.Add()
			ws.Name = right(tabName, 31)
		end if
		'return
		set createNewTab = ws
	end function 
	'return the sheet with the given name (case insensitive
	public function getTab(tabName)
		set getTab = nothing
		dim sheet
		for each sheet in me.worksheets
			if lcase(sheet.Name) = lcase(tabName) then
				set getTab = sheet
				exit function
			end if
		next
	end function
	
	private function getSuffix(stringValue)
		dim suffix
		suffix = ""
		dim rev
		rev = StrReverse(stringValue)
		dim i
		i = instr(rev,"_")
		if i > 1 then
			suffix = StrReverse(left(rev, i -1))
		end if
		getSuffix = suffix
	end function
	
	'public operations
	'create a tab with the given name. The contents should parameter should be a two dimensional array
	'anything int he contents that starts with "=" will be interpreted as a formula
	public Function createTabWithFormulas(tabName, contents,formatAsTable, tableStyle)
		'turn off automatic calculation
		m_ExcelApp.Calculation = xlCalculationManual
		'create the tab
		Dim ws
		set ws = createNewTab(tabName)
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
		createTabWithOffset tabName, contents,formatAsTable, tableStyle, 1, 1
	end function
	
	public Function createTabWithOffset(tabName, contents,formatAsTable, tableStyle, startRow, startColumn)
		'create the tab
		Dim ws
		set ws = createNewTab(tabName)
		'Clear the existing contents
		dim lastRow
		lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count
		dim clearRange
		set clearRange = ws.Range(ws.Cells(startRow,startColumn), ws.Cells(lastRow, Ubound(Contents,2) + startColumn))
		clearRange.ClearContents
		'fill the contents
		dim targetRange
		set targetRange = ws.Range(ws.Cells(startRow,startColumn), ws.Cells(Ubound(contents,1) + startRow, Ubound(Contents,2) + startColumn))
		targetRange.Value2 = contents
		'format as table if needed
		if formatAsTable then
			formatSheetAsTable ws, targetRange, tableStyle
		end if
		'set autofit
		set targetRange = ws.Range(ws.Cells(1,startColumn), ws.Cells(Ubound(contents,1) + startRow, Ubound(Contents,2) + startColumn))
		targetRange.Columns.Autofit
		'return sheet
		set createTabWithOffset = ws
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
	
	public function getContents(sheet)
		getContents = sheet.UsedRange.Value2
	end function
	
	public function ungroupAll(sheet)
		sheet.UsedRange.ClearOutline
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
			'check if the extension is xlsm
			if lcase(right(me.FileName,5)) = ".xlsm" then
				m_WorkBook.Saveas me.FileName, xlOpenXMLWorkbookMacroEnabled
			else
				'save the workbook at the given filename
				m_WorkBook.Saveas me.FileName
			end if
		end if
		'make excel visible
		m_ExcelApp.visible = True
		m_ExcelApp.WindowState = -4137 'xlMaximized
	end function
	
	
end Class