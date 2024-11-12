'[path=\Framework\Utils]
'[group=Utils]

'Name: ExcelFile
'Author: Geert Bellekens
'Purpose: Wrapper script class for Excel files
'Date: 2017-03-20

!INC Utils.Include

const xlCalculationAutomatic = -4105 'Excel controls recalculation.
const xlCalculationManual = -4135 'Calculation is done when the user requests it.

Class ExcelFile
 'private variables
 Private m_ExcelApp
 Private m_FileName
 Private m_WorkBook

 Private Sub Class_Initialize
  m_FileName = ""
  set m_ExcelApp = CreateObject("Excel.Application")
  m_ExcelApp.Visible = true
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
 
 private function createNewTab(tabName)
  'check if the workbook has been created already
  if m_WorkBook is nothing then
   set m_WorkBook = m_ExcelApp.Workbooks.Add()
  end if
  'create the tab at the end
  Dim ws
  Set ws = m_WorkBook.Sheets.Add()
  ws.Name = tabName
  set createNewTab = ws
 end function 
 
 'Returns the data in from the sheet with the given name as a 2 dimensional array
 public function getData(sheetname)
  dim sheet
  for each sheet in me.worksheets
   'return data for first sheet if sheetName is empty
   if len(sheetName) = 0 then
    getData = sheet.UsedRange.Value
   end if
   'find sheet with the given name
   if sheet.Name = sheetname then
    getData = sheet.UsedRange.Value
    exit for
   end if
  next
 end function
 
 'public operations
 'create a tab with the given name. The contents should parameter should be a two dimensional array
 'anything int he contents that starts with "=" will be interpreted as a formula
 public Function createTabWithFormulas(tabName, contents,formatAsTable, tableStyle)
  'turn off automatic calculation
  'm_ExcelApp.Calculation = xlCalculationManual
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
  'return
  set createTabWithFormulas = ws
  'turn on automatic calculation
  'm_ExcelApp.Calculation = xlCalculationAutomatic
 end function
 'public operations
 'create a tab with the given name. The contents should parameter should be a two dimensional array
 public Function createTab(tabName, contents,formatAsTable, tableStyle)
  'create the tab
  Dim ws
  set ws = createNewTab(tabName)
  'fill the contents
  dim targetRange
  set targetRange = ws.Range(ws.Cells(1,1), ws.Cells(Ubound(contents,1) +1, Ubound(Contents,2) +1))
  targetRange.Value2 = contents
  'format as table if needed
  if formatAsTable then
   formatSheetAsTable ws, targetRange, tableStyle
  end if
  'return sheet
  set createTab = ws
 end function
 
 public function formatSheetAsTable(worksheet, targetRange, tableStyle)
  dim table
  Set table = worksheet.ListObjects.Add(1, targetRange, 1, 1)
  table.TableStyle = tableStyle
  'set autofit
  targetRange.Columns.Autofit
 end function
 
 public Function getUserSelectedFileName()
  Session.Output "getUserSelectedFileName"
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
 
end Class