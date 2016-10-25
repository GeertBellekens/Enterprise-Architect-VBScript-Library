'[path=\Projects\Project AC]
'[group=Acerta Scripts]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Import Database mappings
' Author: Geert Bellekens
' Purpose: Import the database mappings from a csv file exported from MEGA
' Date: 2016-07-07
const outPutName = "Import database mappings"


sub main
	dim mappingFile
	set mappingFile = New TextFile
	'select source logical
	dim logicalPackage as EA.Package
	msgbox "select the logical package root (S-OAA-...)"
	set logicalPackage = selectPackage()
	'select source database
	dim physicalPackage as EA.Package
	msgbox "select the database package (example: «database» GBDOAA01)"
	set physicalPackage = selectPackage()
	'first select the mapping file
	if mappingFile.UserSelect("C:\Temp\","CSV Files (*.csv)|*.csv") _
		AND not logicalPackage is nothing _
		AND not physicalPackage is nothing then
		'create output tab
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		'set timestamp
		Repository.WriteOutput outPutName, "Starting import database mappings " & now(), 0
		'split into lines
		dim lines
		lines = Split(mappingFile.Contents, vbCrLf)
		dim line
		dim associationDictionary
		set associationDictionary = CreateObject("Scripting.Dictionary")
		dim packagesTraced
		packagesTraced = false
		for each line in lines
			'replace any "." with "::" 
			line = Replace(line,".","::")
			'split into logical and physical part
			dim parts
			parts = Split(line,";")
			if Ubound(parts) = 1 then
				dim logicalPath
				dim physicalPath
				logicalPath = parts(0)
				physicalPath = parts(1)
				dim logicalObject
				dim physicalObject
				
				set logicalObject = selectObjectFromQualifiedName(logicalPackage,nothing, logicalPath, "::")
				set physicalObject = selectObjectFromQualifiedName(physicalPackage,nothing, physicalPath, "::")
				if not physicalObject is nothing then
					if not logicalObject is nothing then
						Repository.WriteOutput outPutName, "logical: type = " & TypeName(logicalObject) & " name= " & logicalObject.Name & " => Physical: " & physicalObject.Name , 0
						if logicalObject.ObjectType = otElement AND physicalObject.ObjectType = otElement then
							'if the packages are not traced yet then we do that now
							if not packagesTraced then
								tracePackages logicalObject,physicalPackage
								packagesTraced = true
							end if
							'make a trace between the elements
							traceElements logicalObject, physicalObject
						elseif physicalObject.ObjectType = otAttribute and logicalObject.ObjectType = otAttribute then
							traceAttributes logicalObject, physicalObject
						elseif physicalObject.ObjectType = otAttribute and logicalObject.ObjectType = otConnector then
							traceFeatureToAssociation logicalObject, physicalObject
						end if
					elseif physicalObject.ObjectType = otMethod then
						'logical object is not found and physicalObject an operation so we might be dealing with an association to FK mapping
						'Associations have their mega ID in the name like [F19bgUYLMv70]
						if not associationDictionary.Exists(logicalPath) then
							associationDictionary.Add logicalPath, physicalObject
						end if
					end if
				end if
			end if
		next
		Repository.WriteOutput outPutName, "Mapping associations" , 0
		mapAssociations logicalPackage, associationDictionary
		'set timestamp
		Repository.WriteOutput outPutName, "Finished import database mappings " & now(), 0
	end if
end sub

function tracePackages (logicalObject,physicalPackage)
	dim pkconPackage as EA.Package
	set pkconPackage = Repository.GetPackageByID(logicalObject.PackageID)
	if not pkconPackage is nothing then
		Repository.WriteOutput outPutName, "Making trace between package " & pkconPackage.Name & " and package " & physicalPackage.Name,0
			'check if the connector exists already
		for each connector in physicalPackage.Connectors
			if connector.SupplierID = pkconPackage.Element.ElementID _
			AND connector.Type = "Abstraction" _
			AND connector.Stereotype = "trace" then
				traceExists = true
				exit for
			end if
		next
		'if it doesn't exist yet we create a new one
		if traceExists = false then
			set trace = physicalPackage.Connectors.AddNew("","Abstraction")
			trace.SupplierID = pkconPackage.Element.ElementID
			trace.Stereotype = "trace"
			trace.Update
		end if
	end if
end function


function mapAssociations (logicalPackage, associationDictionary)
	dim logicalPath,bracketPosition, associationName
	dim FKOperation as EA.Method
	for each logicalPath in associationDictionary.Keys
		set FKOperation = associationDictionary.Item(logicalPath)
		'get the actual associationName
		bracketPosition = InstrRev(logicalPath, "[", -1, 1)
		if bracketPosition > 0 then
			associationName = Mid(logicalPath,1,bracketPosition -2)
			dim sqlGetAssociations, associations, association
			sqlGetAssociations = "select c.Connector_ID  from t_connector c " & _
								" inner join t_object sob on c.Start_Object_ID = sob.Object_ID " & _
								" inner join t_object tob on c.End_Object_ID = tob.Object_ID " & _
								" inner join t_connector dob_sob on dob_sob.End_Object_ID in (sob.Object_ID, tob.Object_ID) " & _
																  " and dob_sob.Connector_Type = 'Abstraction' " & _
																  " and dob_sob.Stereotype = 'trace' " & _
								" inner join t_object dob on dob_sob.Start_Object_ID = dob.Object_ID " & _
														   " and dob.Stereotype = 'table' " & _
								" inner join t_operation op on op.Object_ID = dob.Object_ID " & _
								" inner join t_connector fkcon on fkcon.Start_Object_ID = dob.Object_ID " & _
																" and fkcon.SourceRole = op.Name " & _
								" inner join t_object dtob on dtob.Object_ID = fkcon.End_Object_ID " & _
								" inner join t_connector dtobtr on dtobtr.Start_Object_ID = dtob.Object_ID " & _
																" and dtobtr.Connector_Type = 'Abstraction' " & _
																" and dtobtr.Stereotype = 'trace' " & _
																" and dtobtr.End_Object_ID in (c.Start_Object_ID, c.End_Object_ID) " & _							
								" where '" & associationName & "' in (c.DestRole + ' ' + c.SourceRole, c.SourceRole + ' ' + c.DestRole) " & _
								" and op.OperationID = " & FKOperation.MethodID
			set associations = getConnectorsFromQuery(sqlGetAssociations)
			for each association in associations
				'map FKOperation to association
				traceFeatureToAssociation association, FKOperation
			next
		end if
	next
end function



function traceElements (originalElement,copyElement)
'	dim originalElement as EA.Element
'	dim copyElement as EA.Element
	'add trace relation
	dim trace as EA.Connector
	dim connector as EA.Connector
	dim traceExists
	traceExists = false
	'check if the connector exists already
	for each connector in copyElement.Connectors
		if connector.SupplierID = originalElement.ElementID _
		AND connector.Type = "Abstraction" _
		AND connector.Stereotype = "trace" then
			traceExists = true
			exit for
		end if
	next
	'if it doesn't exist yet we create a new one
	if traceExists = false then
		set trace = copyElement.Connectors.AddNew("","Abstraction")
		trace.SupplierID = originalElement.ElementID
		trace.Stereotype = "trace"
		trace.Update
	end if
	'set the alias on the original element with the name of the copyElement
	originalElement.Alias = copyElement.Name
	originalElement.Update
end function

function traceAttributes(originalAttribute,copyAttribute)
	'add trace tag
	dim traceTag as EA.AttributeTag
	set traceTag = getExistingOrNewTaggedValue(copyAttribute,"sourceAttribute")
	traceTag.Value = originalAttribute.AttributeGUID
	traceTag.Update
	'set the alias on the original element with the name of the copyElement
	originalAttribute.Alias = copyAttribute.Name
	if not copyAttribute.AllowDuplicates then
		originalAttribute.LowerBound = "0"
	end if
	originalAttribute.Update
end function

function traceFeatureToAssociation(originalConnector,copyFeature)
	'add trace tag
	dim traceTag as EA.ConnectorTag
	set traceTag = getExistingOrNewTaggedValue(copyFeature,"sourceAssociation")
	traceTag.Value = originalConnector.ConnectorGUID
	traceTag.Update
	'set the alias on the original element with the name of the copyElement
	originalConnector.Alias = copyFeature.Name
	originalConnector.Update
end function



main