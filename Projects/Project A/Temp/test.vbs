'[path=\Projects\Project A\Temp]
'[group=Temp]

!INC Wrappers.Include

function resolve (itemGuid)
	resolve = false 'initial value
	'make sure item is writable
	if not isItemWritable(itemGuid) then
	exit function
	end if
	dim diagram as EA.Diagram
	set diagram = Repository.GetDiagramByGuid(itemGuid)
	dim element as EA.Element
	set element = Repository.GetElementByID(diagram.ParentID)
	if not element is nothing AND not diagram is nothing then
		if not diagram.Name = element.Name then
			diagram.Name = element.Name
			diagram.Update
			resolve = true
		end if
	end if
end function

function isItemWritable(itemGuid)
	isDiagramWritable = false 'default = false
	dim userGUID
	userGUID = Repository.GetCurrentLoginUser(true)
	dim sqlGetData
	sqlGetData = "select l.UserID from t_seclocks l                            " & vbNewLine & _
" where l.UserID = '" & userGUID & "'                          " & vbNewLine & _
" and l.EntityID = '" & itemGuid & "'                          "
	dim xmlResult
	xmlResult = Repository.SQLQuery(sqlGetData)				
	dim userIDs
	userIDs = convertQueryResultToArray(xmlResult)
	dim userID
	for each userID in userIDs
	isItemWritable = true
	exit function
	next
end function

'converts the query results from Repository.SQLQuery from xml format to a two dimensional array of strings
Public Function convertQueryResultToArray(xmlQueryResult)
	Dim arrayCreated
	Dim i 
	i = 0
	Dim j 
	j = 0
	Dim result()
	Dim xDoc 
	Set xDoc = CreateObject( "MSXML2.DOMDocument" )
	'load the resultset in the xml document
	If xDoc.LoadXML(xmlQueryResult) Then        
	'select the rows
	Dim rowList
	Set rowList = xDoc.SelectNodes("//Row")

	Dim rowNode 
	Dim fieldNode
	arrayCreated = False
	'loop rows and find fields
	For Each rowNode In rowList
		j = 0
		If (rowNode.HasChildNodes) Then
'redim array (only once)
If Not arrayCreated Then
	ReDim result(rowList.Length, rowNode.ChildNodes.Length)
	arrayCreated = True
End If
For Each fieldNode In rowNode.ChildNodes
	'write f
	result(i, j) = fieldNode.Text
	j = j + 1
Next
		End If
		i = i + 1
	Next
	end if
	convertQueryResultToArray = result
End Function

'MsgBox resolve("{EADBAB99-3875-4263-B303-2B93B54DC781}")
const outPutName = "Test"
sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'start measuring
	dim startTimeStamp
	dim endTimeStamp

	dim i
	dim ms
	ms = 0
	dim iterations
	iterations = 1
	for i = 0 to iterations 
		startTimeStamp = Timer()
		dim xmlQueryResult
		xmlQueryResult = Repository.SQLQuery("select top 1000 * from t_object o order by o.Object_ID desc")
		dim result
		set result = convertQueryResultToArrayList(xmlQueryResult)
		endTimeStamp = Timer()
		dim dif
		dif = (endTimeStamp - startTimeStamp)*1000 
		ms = ms + dif
		Repository.WriteOutput outPutName, now &  " Executed in " &  dif & " ms", 0
	next
	Repository.WriteOutput outPutName, now &  " Average execution API " &  ms/(iterations +1) & " ms", 0
	
	ms = 0
	Dim connection
	dim resultSet
	Set connection = CreateObject("ADODB.Connection")
	'connect.ConnectionString = "Provider=SQLOLEDB;Server=server\instance;Database=pm;Trusted_Connection=True;"
	connection.ConnectionString ="Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TMF;Data Source=DESKTOP-BGN5EL4"
	connection.Open
	dim sqlGetData
	sqlGetData="select top 1000 * from t_object o order by o.Object_ID desc"

	for i = 0 to iterations 
		startTimeStamp = Timer()
		Set resultSet = connection.Execute(sqlGetData)
		'createResultsetXML(resultSet)
		set result = createArrayListsFromResultSet(resultSet)
		endTimeStamp = Timer()
		dif = (endTimeStamp - startTimeStamp)*1000 
		ms = ms + dif
		Repository.WriteOutput outPutName, now &  " Executed in " &  dif & " ms", 0
	next
	connection.Close

	Repository.WriteOutput outPutName, now &  " Average execution ADODB Connection " &  ms/(iterations + 1) & " ms", 0
	
end sub

function createArrayListsFromResultSet(resultSet)
	dim result
	set result = CreateObject("System.Collections.ArrayList")
	resultSet.MoveFirst
	'add headers
	dim row
	set row = CreateObject("System.Collections.ArrayList")
	dim field
	for each field in resultSet.Fields
		row.add field.Name
	next
	result.add row
	'add the values
	Do While Not resultSet.eof
		set row = CreateObject("System.Collections.ArrayList")
		for each field in resultSet.Fields
			row.Add field.Value & ""
		next
		result.Add row
		resultSet.MoveNext
	Loop
	resultSet.Close
	'return
	set createArrayListsFromResultSet = result
end function

function createResultsetXML(resultSet)
	dim xmlDOM 
'	set  xmlDOM = CreateObject( "Microsoft.XMLDOM" )
	set  xmlDOM = CreateObject( "MSXML2.DOMDocument" )

	
	xmlDOM.validateOnParse = false
	xmlDOM.async = false
	 
	dim node 
	set node = xmlDOM.createProcessingInstruction( "xml", "version='1.0'")
    xmlDOM.appendChild node
'
	dim xmlRoot 
	set xmlRoot = xmlDOM.createElement( "EADATA" )
	xmlDOM.appendChild xmlRoot

	dim xmlDataSet
	set xmlDataSet = xmlDOM.createElement( "Dataset_0" )
	xmlRoot.appendChild xmlDataSet
	 
	dim xmlData 
	set xmlData = xmlDOM.createElement( "Data" )
	xmlDataSet.appendChild xmlData
	
	resultSet.MoveFirst
	Do While Not resultSet.eof
		dim xmlRow
		set xmlRow = xmlDOM.createElement( "Row" )
		xmlData.appendChild xmlRow
		dim field
		for each field in resultSet.Fields
			dim xmlField
			set xmlField = xmlDOM.createElement(field.Name)	
			xmlField.text = field.Value & ""
			xmlRow.appendChild xmlField
		next
		resultSet.MoveNext
	Loop
	resultSet.Close
	'return xmldom
	set createResultsetXML = xmlDOM
end function

main