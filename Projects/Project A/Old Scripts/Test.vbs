'[path=\Projects\Project A\Old Scripts]
'[group=Old Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Document Generation for UI Navigation
' Author: Kristof Smeyers
' Purpose: Creates an Excel file containing UI navigation
' Date: 2018-01-02
'
sub main

 ' Get Output
 dim body
 body = getOutPut()
 ' create headers 
 dim headers(0,5)
 headers(0,0) = "Source ID"
 headers(0,1) = "Source UI"
 headers(0,2) = "Connector Notes"
 headers(0,3) = "Target ID"
 headers(0,4) = "Object_Type"
 headers(0,5) = "target UI"
 'combine headers and content
 dim content
 content = mergeArrays(headers, body)
 'write the output to excel
 writeToExcel(content)

end sub

function writeToExcel(content)
 'create the excel file
 dim excelOutput
 set excelOutput = new ExcelFile
 'create the tab
 excelOutput.createTab "Documentgeneration for UI navigation", content, true, "TableStyleMedium4"
 'save the excel file
 excelOutput.save
end function

function getOutPut()
 dim sqlGetContent
 sqlGetContent = " SELECT uc1.ea_guid AS CLASSGUID, uc1.Object_Type AS CLASSTYPE  " & _ 
    " ,uc1.Object_ID AS 'Source ID',uc1.name AS 'source UI'   " & _ 
    " , con.Notes AS 'Connector Notes'      " & _ 
     ",uc2.Object_ID AS 'Target ID',uc2.Object_Type,uc2.name AS 'target UI' " & _ 
    "FROM t_object uc1         " & _ 
    "  INNER JOIN t_connector con ON (uc1.Object_ID = con.Start_Object_ID)" & _ 
    " INNER JOIN t_object uc2 ON (con.End_Object_ID = uc2.Object_ID) " & _ 
    " WHERE uc1.stereotype = 'ArchiMate_ApplicationInterface'   " & _  
    " AND con.stereotype LIKE 'ArchiMate_Triggering'     " & _ 
    " AND uc2.stereotype = 'ArchiMate_ApplicationInterface'   " & _ 
    " UNION ALL          " & _ 
    " SELECT uc1.ea_guid AS CLASSGUID, uc1.Object_Type AS CLASSTYPE  " & _ 
    " ,uc1.Object_ID AS 'Source ID',uc1.name AS 'source UI'   " & _ 
    " , con.Notes AS 'Connector Notes'      " & _ 
    " ,uc2.Object_ID AS 'Target ID',uc2.Object_Type,uc2.name AS 'target UI' " & _ 
    " FROM t_object uc1         " & _ 
    "  INNER JOIN t_connector con ON (uc1.Object_ID = con.End_Object_ID) " & _ 
    "  INNER JOIN t_object uc2 ON (con.Start_Object_ID = uc2.Object_ID) " & _ 
    " WHERE uc1.stereotype = 'ArchiMate_ApplicationInterface'    " & _ 
    " AND con.stereotype LIKE 'ArchiMate_Triggering'     " & _ 
    " AND uc2.stereotype = 'ArchiMate_ApplicationInterface'   " 
 dim outputArray
 outputArray = getArrayFromQuery(sqlGetContent)
 getOutPut = formatOutput(outputArray)
end function

function formatOutput(outputArray)
 'The third column is in formatted text. We convert it to plain text
 dim i
 for i = 0 to Ubound(outputArray)
  outputArray(i,2) = Repository.GetFormatFromField("TXT",outputArray(i,2))
 next
 formatOutput = outputArray
end function


main