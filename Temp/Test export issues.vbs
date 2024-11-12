'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
'
' Script Name: Export Risico's
' Author: Michel De Coninck
' Purpose:
' 1. De Risisco's 'per package) exportenen uit EA
' 2. Publiceren op sharepoint
' Date: 25/03/2020
'
!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
'
'** guid aanpassen indien de top level package van de PUR wijzigt.
'const PURpackageGUID = "{7E289E8A-07AD-4ebd-9C07-2E06DA1283C1}"

Private Sub main
 Session.Output "Working on Risk export to Excel. Please hold on..."
 Repository.EnsureOutputVisible "Script"
 'get the results
 dim sqlGetExport
 Session.Output "getSQLPurExport"
 sqlGetExport = getSQLExport()
 dim purResults
 Session.Output "getArrayListFromQuery"
 set purResults = getArrayListFromQuery(sqlGetExport)
 'udpate results
 Session.Output "updateResults"
 updateResults(purResults)
 'get the headers
 dim headers
 Session.Output "getHeaders"
 set headers = getHeaders()
 'add the headers to the results
 purResults.Insert 0, headers
 'create the excel file
 dim excelOutput
 set excelOutput = new ExcelFile
 'create a two dimensional array
 dim excelContents
 Session.Output "makeArrayFromArrayLists"
 excelContents = makeArrayFromArrayLists(purResults)
 'add the output to a sheet in excel
 excelOutput.createTab "Risk", excelContents, true, "TableStyleMedium4"
 'save the excel file
 excelOutput.save
 Session.Output "Finished successfully. Please review result in Excel."
 Repository.EnsureOutputVisible "Script"
end Sub

Private Function updateResults(purResults)
 dim row
 for each row in purResults
  'add a single quote to the ref field to get Excel to accept it as text
  'row(0) = "'" & row(0)
  row(2) = Repository.GetFormatFromField("TXT", row(2)) 'get text format for notes
  'remove the orderfield located at index 5
  'row.RemoveAt 5
 next
end Function

Private Function getSQLExport()
 getSQLExport =  "SELECT DISTINCT p.name as Package, "    & vbNewLine & _
      "o.Name AS Object, "        & vbNewLine & _
      "o.Note AS Note, "         & vbNewLine & _
      "o.Author, "         & vbNewLine & _
      "o.PDATA3 as x, "         & vbNewLine & _
      "o.PDATA2 as y, "         & vbNewLine & _
      "(SELECT R.VALUE "            & vbNewLine & _
      "FROM t_objectproperties R "     & vbNewLine & _
      "WHERE "          & vbNewLine & _
      "O.Object_ID = R.Object_ID "     & vbNewLine & _
      "AND R.Property ='Rating' "      & vbNewLine & _
      ") as Rating, "        & vbNewLine & _
      "Object_Type AS [Type], "       & vbNewLine & _
      "o.Stereotype, "        & vbNewLine & _
      "Scope, "           & vbNewLine & _
      "Status, "           & vbNewLine & _
      "Phase, "           & vbNewLine & _
      "FORMAT (O.CreatedDate, 'dd-MMM-yyyy')  as d ," & vbNewLine & _
      "FORMAT (O.ModifiedDate, 'dd-MMM-yyyy')  as e " & vbNewLine & _
    "FROM t_object o, t_package p "       & vbNewLine & _
     "WHERE "            & vbNewLine & _
     "o.Object_Type = 'Issue' "        & vbNewLine & _
     "and o.package_ID = p.Package_ID "
end Function

Private Function getHeaders()
 dim headers
 set headers = CreateObject("System.Collections.ArrayList")
 headers.add("PACKAGE NAAM")
 headers.add("RISICO NAAM")
 headers.Add("RISICO BESCHRIJVING")
 headers.Add("AUTEUR")
 headers.Add("DIFFICULTY")
 headers.Add("PRIORITY")
 headers.Add("RATING")
 headers.Add("TYPE")
 headers.Add("STEREOTYPE")
 headers.Add("SCOPE")
 headers.Add("STATUS")
 headers.Add("FASE")
 headers.Add("CREATIEDATUM")
 headers.Add("LAATST GEWIJZIGD")
 set getHeaders = headers
end Function

main