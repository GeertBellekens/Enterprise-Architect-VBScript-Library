'[path=\Projects\Project A\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include


'
' Script Name: FixImportStubConnectors
' Author: geert Bellekens
' Purpose: create an sql update query to fix the import stub connectors
' Date: 2019-03-07
'
sub main
	dim mappingFile
	set mappingFile = New TextFile
	
	'first select the mapping file
	if mappingFile.UserSelect("","Txt Files (*.txt)|*.txt") then
		'split into lines
		dim lines
		lines = Split(mappingFile.Contents, vbCrLf)
		dim guid
		dim query
		dim i
		i = 0
		dim fileCounter
		fileCounter = 1
		for each guid in lines
			i = i + 1
			'get the connector
			dim guidWildCards
			guidWildCard = Replace(guid, "{", "<")
			guidWildCard = Replace(guidWildCard, "}", ">")
			dim connector as EA.Connector
			set connector = Repository.GetConnectorByGuid(guid)
			if not connector is nothing then
				dim targetClass as EA.Element
				set targetClass = Repository.GetElementByID(connector.SupplierID)
				query = query & "update c set c.End_Object_ID = (select o.object_ID from t_object o where o.ea_guid = '" & targetClass.ElementGUID & "'), c.ea_guid = '" & connector.ConnectorGUID & "' from t_connector c where c.ea_guid = '" & guidWildCard &"'" & vbNewLine
				
			else
				'Report error missing guid
				Session.Output "ERROR: Line " & i & " could not find connector with guid: " & guid
			end if
			'save file for each 2000 lines
			if i >= filecounter * 2000 then
				Session.Output "Saving file: " & filecounter
				saveFile fileCounter, query
				'reset query
				query = ""
				'up fileCounter
				filecounter = filecounter + 1
			end if
		next
		'output to file
		saveFile fileCounter + 1, query
	end if
end sub

function saveFile(fileCounter, query)
		'output to file
		dim outputFile
		set outputFile = New TextFile
		outputfile.Contents = query
		outputfile.FullPath = "H:\temp\Fix Import stubs\fixQuery" & fileCounter & ".sql"
		outputFile.Save
end function

main