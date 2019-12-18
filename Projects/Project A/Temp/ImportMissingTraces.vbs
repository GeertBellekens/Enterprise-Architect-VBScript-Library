'[path=\Projects\Project A\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include


'
' Script Name: ImportMissingTraces
' Author: geert Bellekens
' Purpose: import the mising traces based on the csv file
' Date: 2019-04-02
'

'Input file create with the following query

'select c.ea_guid, o.ea_guid, ot.ea_guid, c.[StyleEx]
'from ((t_connector c
'inner join t_object o on o.Object_ID = c.Start_Object_ID)
'inner join t_object ot on ot.[object_id] = c.[End_Object_ID])
'
'where 
'o.[Stereotype] = 'Business Rule'
'and c.Connector_Type = 'Abstraction'
'and c.Stereotype = 'trace'
'and ot.Object_Type = 'class'
'and (ot.Stereotype is null or ot.Stereotype = 'LDM')


const outPutName = "Fix use case traces"

sub main
	dim mappingFile
	set mappingFile = New TextFile
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	
	'first select the mapping file
	if mappingFile.UserSelect("","Txt Files (*.txt)|*.txt") then
		'split into lines
		dim lines
		lines = Split(mappingFile.Contents, vbCrLf)
		dim line
		dim i
		i = 0
		dim useCasesDictonary 
		set useCasesDictonary = CreateObject("Scripting.Dictionary")
		dim ldmDictionary
		set ldmDictionary = CreateObject("Scripting.Dictionary")
		for each line in lines
			i = i + 1
			dim linecontent
			lineContent = Split(line, ";")
			dim connectorGUID
			connectorGUID = lineContent(0)
			dim sourceGUID
			sourceGUID = lineContent(1)
			dim targetGUID
			targetGUID = lineContent(2)
			dim automatic
			'Session.Output i & ": Ubound = " & ubound(lineContent)
			if ubound(lineContent) > 3 then
				automatic = true
			else
				automatic = false
			end if
			'get the connector
			dim connector as EA.Connector
'			set connector = Repository.GetConnectorByGuid(connectorGUID)
'			if connector is nothing then
			'check if a trace already exists between the two guid's
			dim traceExisting
			traceExisting = traceExists(sourceGUID, targetGUID)
			if not traceExisting then
				dim usecase as EA.Element
				'check if use case is known
				if useCasesDictonary.Exists(sourceGUID) then
					set usecase = useCasesDictonary(sourceGUID)
				else
					set usecase = Repository.GetElementByGuid(sourceGUID)
					useCasesDictonary.Add sourceGUID, usecase
				end if
				dim targetClass as EA.Element
				if ldmDictionary.Exists(targetGUID) then
					set targetClass = ldmDictionary(targetGUID)
				else
					set targetClass = Repository.GetElementByGuid(targetGUID)
					ldmDictionary.Add targetGUID, targetClass
				end if
				if not usecase is nothing _
				 and not targetClass is nothing then
					Repository.WriteOutput outPutName, now() & " " & i &  ": Creating link between " & usecase.Name & " and " & targetClass.Name, 0
					on error resume next
					linkElementsWithAutomaticTrace usecase, targetClass, automatic
					if Err.number <> 0 then
						Err.clear
						Repository.WriteOutput outPutName, now() & " " & i & ": ERROR (locking?) creating link for " & connectorGUID, 0
					end if
					on error goto 0
				else
					Repository.WriteOutput outPutName, now() & " " & i &  ": ERROR creating link for "  & connectorGUID, 0
				end if
			else
				Repository.WriteOutput outPutName, now() & " " & i &  ": Connector already exists "  & connectorGUID, 0
			end if
		next
	end if
end sub

function traceExists(sourceGUID, targetGUID)
	dim sqlGetTraces
	sqlGetTraces =  "select c.Connector_ID from t_connector c                    " & _
					" inner join t_object so on so.Object_ID = c.Start_Object_ID " & _
					" inner join t_object eo on eo.Object_ID = c.End_Object_ID   " & _
					" where                                                      " & _
					" c.Stereotype = 'trace'                                     " & _
					" and so.ea_guid = '" & sourceGUID & "'                      " & _
					" and eo.ea_guid = '" & targetGUID & "'                      "
	dim results
	results = getArrayFromQuery(sqlGetTraces)
	on error resume next
	if ubound(results) > 0 then
		traceExists = true
	else
		traceExists = false
	end if
	If Err.Number <> 0 Then
		traceExists = false
		Err.Clear
	end if
	on error goto 0
end function

function linkElementsWithAutomaticTrace(sourceElement, TargetElement, automatic)
	dim trace as EA.Connector
	set trace = sourceElement.Connectors.AddNew("","trace")
	if automatic then
		trace.Alias = "automatic"
	end if
	trace.SupplierID = TargetElement.ElementID
	trace.Update
end function

main