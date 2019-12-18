'[path=\Projects\Project A\Schema Scripts]
'[group=Schema Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Set traceability
' Author: Geert Bellekens
' Purpose: Adds traceability from the copy package selected in the project browser to the original package selected by the user.
' Date: 2016-02-08
'

dim outputTabName
outputTabName = "Set Traceability"

sub main()

	Repository.CreateOutputTab outputTabName
	Repository.ClearOutput outputTabName
	Repository.EnsureOutputVisible outputTabName
	Repository.EnableUIUpdates = false
	
	dim copyPackage as EA.Package
	dim originalPackage as EA.Package
	'the copy packag is the one selected in the project browser
	set copyPackage = Repository.GetTreeSelectedPackage
	msgbox "Please select the original package"
	'the original package is the one selected by the user
	dim originalPackageID
	dim originalPackageElementID
	originalPackageElementID = Repository.InvokeConstructPicker("IncludedTypes=Package;Selection=" & copyPackage.PackageGUID)
	if originalPackageElementID > 0 then
		dim originalPackageElement as EA.Element
		set originalPackageElement = Repository.GetElementByID(originalPackageElementID)
		set originalPackage = Repository.GetPackageByGuid(originalPackageElement.ElementGUID)
		'tell the user we are starting
		Repository.WriteOutput outputTabName, now() & " Starting adding traces",0
		'start the trace		
		tracePackageElements originalPackage, copyPackage
		Repository.WriteOutput outputTabName, now() & " Finished adding traces",0
	end if
	Repository.EnableUIUpdates = true
	Repository.RefreshModelView copyPackage.PackageID
end sub

function tracePackageElements(originalPackage, copyPackage)
	dim originalElement as EA.Element
	dim copyElement as EA.Element
	'find corresponding element
	for each originalElement in originalPackage.Elements
		
		'only process elements that have a name and are Class, Enumeration, Datatype, PrimitiveType
		if len(originalElement.Name) > 0 and _
		  (originalElement.Type = "Class" or originalElement.Type = "Enumeration" _
     	  or originalElement.Type = "DataType" or originalElement.Type = "PrimitiveType" ) then
			'Repository.WriteOutput outputTabName, now() & " Processing " & originalElement.Type & ": " & originalElement.Name ,0
			dim matchFound
			matchFound = false
			for each copyElement in copyPackage.Elements
				if copyElement.Name = originalElement.Name _
				  and copyElement.Type = originalElement.Type then
					'found a match
					traceElements originalElement,copyElement
					matchFound = true
					exit for
				end if
			next
			if matchFound then
				Repository.WriteOutput outputTabName, now() & " Match found for " & originalElement.Type & ": " & originalElement.Name ,0
			else
				Repository.WriteOutput outputTabName, now() & " Match NOT found for " & originalElement.Type & ": " & originalElement.Name ,0
			end if
		end if
	next
	'process subpackages
	dim originalSubPackage
	dim copySubpackage
	for each originalSubPackage in originalPackage.Packages
		for each copySubpackage in copyPackage.Packages
			if originalSubPackage.Name = copySubpackage.Name then
				'found a match
				tracePackageElements originalSubPackage, copySubpackage
				exit for
			end if
		next
	next
end function

function traceElements (originalElement,copyElement)
	'add trace relation
	dim trace as EA.Connector
	dim connector as EA.Connector
	'delete all existing traces
	dim i
	for i = copyElement.Connectors.Count -1 to 0 step -1
		set connector = copyElement.Connectors.GetAt(i)
		if connector.Type = "Abstraction" _
		  AND connector.ClientID = copyElement.ElementID _
		  AND connector.Stereotype = "trace" then
			copyElement.Connectors.DeleteAt i,false
		end if
	next
	'refresh connectors
	copyElement.Connectors.Refresh
	'create new trace
	set trace = copyElement.Connectors.AddNew("","Abstraction")
	trace.SupplierID = originalElement.ElementID
	trace.Stereotype = "trace"
	trace.Update
	'trace attributes
	traceAttributes originalElement,copyElement
	'trace associations
	traceAssociations originalElement,copyElement
end function


function traceAttributes(originalElement,copyElement)
	dim originalAttribute as EA.Attribute
	dim copyAttribute as EA.Attribute
	for each originalAttribute in originalElement.Attributes
		for each copyAttribute in copyElement.Attributes
			if copyAttribute.Name = originalAttribute.Name then
				'found match, add trace tag
				dim traceTag as EA.AttributeTag
				set traceTag = getExistingOrNewTaggedValue(copyAttribute,"sourceAttribute")
				traceTag.Value = originalAttribute.AttributeGUID
				traceTag.Update
				exit for
			end if
		next
	next
end function

function traceAssociations (originalElement,copyElement)
	'make sure the connectors are refreshed
	copyElement.Connectors.Refresh
	originalElement.Connectors.Refresh
	
	dim originalConnector as EA.Connector
	dim copyConnector as EA.Connector
	for each originalConnector in originalElement.Connectors
		'we process only associations that start from the original element
		if (originalConnector.Type = "Association" or originalConnector.Type = "Aggregation") _
			AND originalConnector.ClientID =  originalElement.ElementID then
			for each copyConnector in copyElement.Connectors
				if copyConnector.Type = originalConnector.Type _
					AND copyConnector.Name = originalConnector.Name _
					AND copyConnector.ClientEnd.Role = originalConnector.ClientEnd.Role _
					AND copyConnector.ClientEnd.Aggregation = originalConnector.ClientEnd.Aggregation _
					AND copyConnector.SupplierEnd.Role = originalConnector.SupplierEnd.Role _
					AND copyConnector.SupplierEnd.Aggregation = originalConnector.SupplierEnd.Aggregation then
					'AND copyConnector.ClientEnd.Cardinality = originalConnector.ClientEnd.Cardinality _
					'AND copyConnector.SupplierEnd.Cardinality = originalConnector.SupplierEnd.Cardinality _
					'connector properties match, now check the other ends
					dim originalOtherEnd as EA.Element
					dim copyOtherEnd as EA.Element
					set originalOtherEnd = Repository.GetElementByID(originalConnector.SupplierID)
					set copyOtherEnd = Repository.GetElementByID(copyConnector.SupplierID)
					if copyOtherEnd.Name = originalOtherEnd.Name then
						'found a match, add trace tag
						dim traceTag as EA.ConnectorTag
						set traceTag = getExistingOrNewTaggedValue(copyConnector,"sourceAssociation")
						traceTag.Value = originalConnector.ConnectorGUID
						traceTag.Update
						exit for
					end if
				end if
			next
		end if
	next
end function

main