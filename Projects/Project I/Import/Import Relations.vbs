'[path=\Projects\Project I\Import]
'[group=Import]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Import Connectors
' Author: Geert Bellekens
' Purpose: Create connectors for the data in the csv file
' Date: 2018-01-26
const outPutName = "Import Relations"


sub main
	dim mappingFile
	set mappingFile = New TextFile
	'first select the mapping file
	if mappingFile.UserSelect("","CSV Files (*.csv)|*.csv") then
		'create output tab
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		'set timestamp
		Repository.WriteOutput outPutName, now() & " Starting import connectors ", 0
		'counter
		dim i
		'split into lines
		dim lines
		lines = Split(mappingFile.Contents, vbCrLf)
		dim line
		for each line in lines
			i = i + 1
			'split into logical and physical part
			dim parts
			parts = Split(line,";")
			if Ubound(parts) = 1 then
				dim sourceName
				dim targetName
				sourceName = parts(0)
				targetName = parts(1)
				'get the source object
				dim sourceElement as EA.Element
				set sourceElement = getRequirementByName(sourceName)
				'get the target object
				dim targetElement as EA.Element
				set targetElement = getRequirementByName(targetName)
				if not sourceElement is nothing and not targetElement is nothing then
					'check if trace doesn't exist already
					if traceExists(sourceElement, targetElement) then
						' trace already exists
						Repository.WriteOutput outPutName, now() & " line : " & i & " Trace already exists between " & sourceName & " and " & targetName , 0
					else
						'create the connector
						dim connector as EA.Connector
						set connector = sourceElement.Connectors.AddNew(connectorName, "Abstraction")
						connector.Stereotype = "trace"
						connector.SupplierID = targetElement.ElementID
						on error resume next
						connector.Update
						if Err.Number <> 0 then
							'set timestamp
							Repository.WriteOutput outPutName, now() & " line : " & i & " ERROR: Count not create connector between " & sourceName & " and " & targetName , 0
							Err.Clear
						else
							'set timestamp
							Repository.WriteOutput outPutName, now() & " line : " & i & " Added connector between " & sourceName & " and " & targetName , 0
						end if
					end if
				end if
			end if	
		next
		'set timestamp
		Repository.WriteOutput outPutName, now() & " Finished import connectors " , 0
		Repository.EnableUIUpdates = true
	end if
end sub

function getRequirementByName(elementName)
	dim foundElement as EA.Element
	set foundElement = nothing
	dim sqlGetRequirement
	sqlGetRequirement = "select o.Object_ID from t_object o   " & vbNewLine & _
						" where o.Object_Type = 'Requirement'  " & vbNewLine & _
						" and o.Name = '" & elementName & "'   "
	dim result
	set result = getElementsFromQuery(sqlGetRequirement)
	if result.Count = 1 then
		'only return if we have exactly one result
		set foundElement = result(0) 
	elseif result.Count > 1 then
		Repository.WriteOutput outPutName, now() & " ERROR: More than one requirement found with name '" & elementName & "'" , 0
	else
		Repository.WriteOutput outPutName, now() & " ERROR: No requirement found with name '" & elementName & "'" , 0
	end if	
	'return foundElement
	set getRequirementByName = foundElement
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