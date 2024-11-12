'[path=\Projects\Project DL\DL Scripts]
'[group=De Lijn Scripts]

Option Explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
'
' Script Name: Export Glossary
' Author: Geert Bellekens
' Purpose: Export the contents of the glossary to excel
' Date: 2023-09-14
'


const outPutName = "Export Glossary"

sub main
 'create output tab
 Repository.CreateOutputTab outPutName
 Repository.ClearOutput outPutName
 Repository.EnsureOutputVisible outPutName
 'inform user
 Repository.WriteOutput outPutName, now() & " Starting " & outPutName, 0
 'do the actual work
 exportGlossary
 'inform user
 Repository.WriteOutput outPutName, now() & " Finished " & outPutName, 0
end sub

function exportGlossary
 dim sqlGetData
 sqlGetData = "select o.Name as Term,  d.Name as Type, o.Note as [Meaning], 'True' as [ModelItem]                                   " & vbNewLine & _
    " from t_object o                                                                                                      " & vbNewLine & _
    " inner join t_diagramobjects do on do.Object_ID = o.Object_ID                                                         " & vbNewLine & _
    " inner join t_diagram d on d.Diagram_ID = do.Diagram_ID                                                               " & vbNewLine & _
    "       and d.StyleEx like '%MDGDgm=Glossary Item Lists::GlossaryItemList;%'         " & vbNewLine & _
    " union                                                                                                                " & vbNewLine & _
    " select o.Name as Term,  p.Name as Type, o.Note as [Meaning], 'True' as [ModelItem]                                  " & vbNewLine & _
    " from t_object o                                                                                                      " & vbNewLine & _
    " inner join t_package p on p.Package_ID = o.Package_ID                                                                " & vbNewLine & _
    " where o.Stereotype = 'GlossaryEntry'                                                                                 " & vbNewLine & _
    " union                                                                                                                " & vbNewLine & _
    " select g.Term, g.Type, g.Meaning, 'False' as [ModelItem]                                                            " & vbNewLine & _
    " from t_glossary g                                                                                                    " & vbNewLine & _
    " order by 1, 2, 3                                                                                                     "
 Dim results
 'inform user
 Repository.WriteOutput outPutName, now() & " Getting data", 0
 Set results = getArrayListFromQuery(sqlGetData)
 'inform user
 Repository.WriteOutput outPutName, now() & " Formatting notes", 0
 ' format notes
 formatNotesField results, 2
 ' Get the headers
 Dim headers
 Set headers = getHeaders()
 ' Add the headers to the results
 results.Insert 0, headers
 Dim excelOutput
 Set excelOutput = new ExcelFile
 ' Create a two dimensional array
 Dim excelContents
 excelContents = makeArrayFromArrayLists(results)
 Repository.WriteOutput outPutName, now() & " Exporting to excel", 0
 ' Add the output to a sheet in excel
 dim ws
 set ws = excelOutput.createTabWithFormulas("Glossary", excelContents, true, "TableStyleMedium4")
 'save
 ' Ask for location, then Save the excel file
 excelOutput.save
end function

Function formatNotesField(results, notesIndex)
 Dim row
 For Each row in results
  row(notesIndex) = Repository.GetFormatFromField("TXT", row(notesIndex))
 Next
End Function

Function getHeaders()
 dim headers
 set headers = CreateObject("System.Collections.ArrayList")
 headers.add("Term")
 headers.add("Type")
 headers.Add("Meaning")
 headers.Add("Model Item")
 set getHeaders = headers
End Function

main