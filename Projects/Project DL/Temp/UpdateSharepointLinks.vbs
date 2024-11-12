'[path=\Projects\Project DL\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: UpdateSharepointLinks
' Author: Geert Bellekens
' Purpose: Update the sharepoint links based on the SiteURsMapping
' Date: 2023-02-18
'


'name of the output tab
const outPutName = "UpdateSharepointLinks"

sub main
 'create output tab
 Repository.CreateOutputTab outPutName
 Repository.ClearOutput outPutName
 Repository.EnsureOutputVisible outPutName
 'set timestamp for start
 Repository.WriteOutput outPutName,now() & " Start " & outPutName  , 0
 'open the excel file
 dim excel
 set excel = new ExcelFile
 excel.openUserSelectedFile
 Repository.WriteOutput outPutName,now() & " Reading excel file"  , 0 
 'get the data from the "import" sheet
 dim siteMappingData
 siteMappingData = excel.getData("")
 'actually import the data
 updateUrls siteMappingData
 'set timestamp for finish
 Repository.WriteOutput outPutName,now() & " Finished " & outPutName  , 0
end sub

function updateUrls(siteMappingData)
 'loop the data
 dim i
 dim j
 dim sortedSiteMappings
 set sortedSiteMappings = sortSiteMappingData(siteMappingData)
 Repository.WriteOutput outPutName,now() & " Updating Urls"  , 0 
 'skip first row
 dim row
 for each row in sortedSiteMappings
  dim oldUrl
  dim newUrl
  oldUrl = replace(row(0), " ", "%20")
  newUrl = replace(row(1), " ", "%20")
  if len(oldUrl) > 0 and len(newUrl) > 0 then
   'check how many results we have
   dim results
   dim sqlGetData
   dim updateSQL
   'check results in objectfiles
   sqlGetData = "select fileName from t_objectfiles where fileName like '%" & replace(oldUrl, "'", "''") & "%'"
   set results = getArrayListFromQuery(sqlGetData)
   if results.Count > 0 then
    Repository.WriteOutput outPutName,now() & " Updating:  " & results.Count &  " hits in links for mapping '" & oldUrl & "' to '" & newUrl & "'", 0  
    updateSQL = "update t_objectfiles set FileName = replace(filename, '" & replace(oldUrl, "'", "''") & "', '" & replace(newUrl, "'", "''") & "') where fileName like '%" & replace(replace(oldUrl, "'", "''"), "%", "[%]") & "%'"
    repository.Execute updateSQL
   end if
   'check results in notes
   sqlGetData = "select o.ea_guid from t_object o where o.note like '%" & replace(oldUrl, "'", "''") & "%'"
   set results = getArrayListFromQuery(sqlGetData)
   if results.Count > 0 then
    Repository.WriteOutput outPutName,now() & " Updating:  " & results.Count &  " hits in notes for mapping '" & oldUrl & "' to '" & newUrl & "'", 0  
    updateSQL = "update t_object set note = replace(note, '" & replace(oldUrl, "'", "''") & "', '" & replace(newUrl, "'", "''") & "') where note like '%" & replace(replace(oldUrl, "'", "''"), "%", "[%]") & "%'"
    repository.Execute updateSQL
   end if
  end if
 next
end function

function sortSiteMappingData(siteMappingData)
 Repository.WriteOutput outPutName,now() & " Sorting Urls"  , 0 
 dim siteMappings
 set siteMappings = CreateObject("System.Collections.ArrayList")
 'loop the data
 dim i
 'skip first row
 for i = 2 to Ubound(siteMappingData)  'rows 
  'get data
  dim oldUrl
  dim newUrl
  dim urlLength
  oldUrl = siteMappingData(i,2)
  newUrl = siteMappingData(i,3)
  urlLength = len(oldUrl)
  'create row 
  dim row
  set row = CreateObject("System.Collections.ArrayList")
  row.Add oldUrl
  row.Add newUrl
  row.Add urlLength 
  'add row to the siteMappings list
  dim siteMapping
  dim added
  dim j
  j = 0
  added = false
  for each siteMapping in siteMappings
   if urlLength >= siteMapping(2) then
    'insert before
    siteMappings.Insert j, row
    added = true
    exit for
   end if
   j = j + 1
  next
  'if not added yet, then add to the back
  if not added then
   siteMappings.add row
  end if
 next
 'return
 set sortSiteMappingData = siteMappings
end function

main