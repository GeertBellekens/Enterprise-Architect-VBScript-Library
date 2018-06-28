'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Import Connectors
' Author: Geert Bellekens
' Purpose: Create connectors for the data in the csv file
' Date: 2018-01-26
const outPutName = "Import database mappings"


sub main
	dim mappingFile
	set mappingFile = New TextFile
	'first select the mapping file
	if mappingFile.UserSelect("H:\Temp\","Txt Files (*.txt)|*.txt") then
		'create output tab
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		'set timestamp
		Repository.WriteOutput outPutName, now() & "Starting import connectors ", 0
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
			if Ubound(parts) = 5 then
				dim sourceGUID
				dim connectorName
				dim connectorType
				dim connectorStereotype
				dim targetGUID
				sourceGUID = parts(0)
				connectorName = parts(1)
				connectorType = parts(2)
				connectorStereotype = parts(3)
				targetGUID = parts(4)
				'get the source object
				dim sourceElement as EA.Element
				set sourceElement = Repository.GetElementByGuid(sourceGUID)
				'get the target object
				dim targetElement as EA.Element
				set targetElement = Repository.GetElementByGuid(targetGUID)
				if not sourceElement is nothing and not targetElement is nothing then
					'create the connector
					dim connector as EA.Connector
					set connector = sourceElement.Connectors.AddNew(connectorName, connectorType)
					connector.Stereotype = connectorStereotype
					connector.SupplierID = targetElement.ElementID
					on error resume next
					connector.Update
					if Err.Number <> 0 then
						'set timestamp
						Repository.WriteOutput outPutName, now() & " line : " & i & " ERROR! Count not create connector between existing " & sourceGUID & " and " & targetGUID , 0
						Err.Clear
					else
						'set timestamp
						Repository.WriteOutput outPutName, now() & " line : " & i & " Added connector between " & sourceGUID & " and " & targetGUID , 0
					end if
				else
					'set timestamp
					Repository.WriteOutput outPutName, now() & " line : " & i & " ERROR! Count not create connector between non existing " & sourceGUID & " and " & targetGUID , 0
				end if
			end if	
		next
		'set timestamp
		Repository.WriteOutput outPutName, now() & " Finished import connectors " , 0
		Repository.EnableUIUpdates = true
	end if
end sub




main