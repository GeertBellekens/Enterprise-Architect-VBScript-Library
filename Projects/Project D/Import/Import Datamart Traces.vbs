'[path=\Projects\Project D\Import]
'[group=Import]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Import Datamart Traces
' Author: Geert Bellekens
' Purpose: Create traces from the datamart tables to the source views
' Date: 2025-11-15

const outPutName = "Import Datamart traces"

sub main	
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'tell the user we are finished
	Repository.WriteOutput outPutName, now() & " Starting " &  outPutName ,0
	'Import the traces based on the csv file
	importTraces
	'tell the user we are finished
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName , 0
end sub

function importTraces()
	'read mapping file
	dim mappingFile
	set mappingFile = New CSVFile
	if not mappingFile.UserSelect("") then
		Repository.WriteOutput outPutName, now() & " Cancelled by user" , 0
		exit function
	end if
	'create dictionary for source systems
	dim sourceSystemDictionary
	set sourceSystemDictionary = getSourceSystemDictionary()
	dim mappingFileContents
	set mappingFileContents = mappingFile.Contents
	dim row
	for each row in mappingFileContents
		dim sourceSystem
		dim sourceView
		dim targetTable
		if row.Count >= 3 then
			sourceSystem = row(0)
			sourceView = row(1)
			targetTable = row(2)
			'skip if we don't know the source system
			if sourceSystemDictionary.Exists(lcase(sourceSystem)) then
				importTrace sourceSystemDictionary(lcase(sourceSystem)), sourceView, targetTable
			end if
		end if
	next
end function

function importTrace(sourceSystemPackageIDs, sourceViewName, targetTableName)
	'get the target table
	dim targetTable as EA.Element
	set targetTable = getTargetTable(targetTableName)
	if targetTable is nothing then
		Repository.WriteOutput outPutName, now() & " ERROR: target table not found '" & targetTableName & "'" , 0
		exit function
	end if
	'get the source view
	dim sourceView as EA.Element
	set sourceView = getSourceView(sourceSystemPackageIDs, sourceViewName)
	if sourceView is nothing then
			Repository.WriteOutput outPutName, now() & " ERROR: source view not found '" & sourceViewName & "'" , 0
		exit function
	end if
	'check if trace exists before creating
	if not traceExists( targetTable, sourceView) then
		Repository.WriteOutput outPutName, now() & " Adding trace from '" & targetTableName & "' to '" & sourceViewName & "'"  , 0
		dim trace as EA.Connector
		set trace = targetTable.Connectors.AddNew("", "Trace")
		trace.SupplierID = sourceView.ElementID
		trace.Update
	else
		Repository.WriteOutput outPutName, now() & " Found existing trace from '" & targetTableName & "' to '" & sourceViewName & "'"  , 0
	end if
	
end function

function getSourceView(sourceSystemPackageIDs, sourceViewName)
	dim sourceView
	set sourceView = nothing
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o                     " & vbNewLine & _
				" where o.Stereotype = 'view'                           " & vbNewLine & _
				" and o.Name = '" & sourceViewName & "'                 " & vbNewLine & _
				" and o.Package_ID in (" & sourceSystemPackageIDs & ")   "
	dim result
	set result = getElementsFromQuery(sqlGetData)
	if result.Count > 0 then
		set sourceView = result(0)
	end if
	'return
	set getSourceView = sourceView
end function


function getTargetTable(targetTableName)
	dim targetTable as EA.Element
	set targetTable = nothing
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o         " & vbNewLine & _
	             " where o.stereotype = 'table'              " & vbNewLine & _
				 " and o.Name = '" & targetTableName & "'    " 
	dim result
	set result = getElementsFromQuery(sqlGetData)
	if result.Count > 0 then
		set targetTable = result(0)
	end if
	'return
	set getTargetTable = targetTable
end function

function getSourceSystemDictionary()
	dim sqlGetData
	sqlGetData = "select p.Package_ID from t_package p              " & vbNewLine & _
				" inner join t_object o on o.ea_guid = p.ea_guid   " & vbNewLine & _
				" where o.Stereotype = 'database'                  "
	dim result
	set result = getPackagesFromQuery(sqlGetData)
	dim sourceSystemDictionary
	set sourceSystemDictionary = CreateObject("Scripting.Dictionary")
	dim package as EA.Package
	for each package in result
		dim key
		key = lcase(package.Name)
		if not sourceSystemDictionary.Exists(key) then
			dim packageTreeIDString
			packageTreeIDString = getPackageTreeIDString(package)
			'add to dictionary
			sourceSystemDictionary.Add key, packageTreeIDString
		else
			Repository.WriteOutput outPutName, now() & " ERROR: duplicate database found: '" & package.Name & "'" , 0
		end if
	next
	'return
	set getSourceSystemDictionary = sourceSystemDictionary
end function



function traceExists(sourceElement, targetElement)
	dim sqlGetTrace
	sqlGetTrace = "select c.Connector_ID from t_connector c             " & vbNewLine & _
					" where c.Stereotype = 'trace'                      " & vbNewLine & _
					" and c.Start_Object_ID = " & sourceElement.ElementID &  vbNewLine & _
					" and c.End_Object_ID = " & targetElement.ElementID
	dim connectors
	set connectors = getConnectorsFromQuery(sqlGetTrace)
	'check if exists
	if connectors.Count > 0 then
		traceExists = true
	else
		traceExists = false
	end if
end function






main