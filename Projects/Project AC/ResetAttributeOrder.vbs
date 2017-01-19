'[path=\Projects\Project AC]
'[group=Acerta Scripts]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Reset Attribute Order
' Author: Geert Bellekens
' Purpose: Reset the logical attributes according to the order of the corresponding columns in the database
' Date: 2017-01-19
const outPutName = "Reset Attribute order"


sub main
	dim mappingFile
	set mappingFile = New TextFile
	'select source logical
	dim logicalPackage as EA.Package
	msgbox "select the Logical Data Model Package (PK CON ...)"
	set logicalPackage = selectPackage()
	'select source database
	dim physicalPackage as EA.Package
	msgbox "select the database package (example: «database» GBDOAA01)"
	set physicalPackage = selectPackage()
	if not physicalPackage is nothing and not logicalPackage is nothing then
		'first select the mapping file
		Repository.EnableUIUpdates = false
		'create output tab
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		'set timestamp
		Repository.WriteOutput outPutName, now() & ": Starting reset attribute order " , 0
		resetAttributeOrder logicalPackage, physicalPackage
		Repository.WriteOutput outPutName, now() & ": Finished reset attribute order " , 0
		Repository.EnableUIUpdates = true
		Repository.RefreshModelView logicalPackage.PackageID
	end if
end sub

function resetAttributeOrder(logicalPackage, physicalPackage)
	dim table as EA.Element
	for each table in physicalPackage.elements
		if lcase(table.Stereotype) = "table" then
			'get the corresponding logical classes
			dim logicalClasses 
			set logicalClasses = getLocialClasses(table)
			'set the attributes order
			setAttributeOrder table, logicalClasses 
		end if
	next
	'process subPackages
	dim subPackage as EA.Package
	for each subPackage in physicalPackage.Packages
		resetAttributeOrder logicalPackage, subPackage
	next
end function

function setAttributeOrder(table, logicalClasses)
	'tell the user what we are doing
	Repository.WriteOutput outPutName,  "Processign table: " & table.Name , 0
	dim column as EA.Attribute
	for each column in table.Attributes
		dim logicalAttributes
		set logicalAttributes = getLogicalAttributes(column, logicalClasses)
		dim logicalAttribute as EA.Attribute
		for each logicalAttribute in logicalAttributes
			logicalAttribute.Pos = column.Pos
			logicalAttribute.Update
		next
	next
end function

function getLogicalAttributes(column, logicalClasses)
	dim logicalAttributes
	set logicalAttributes = CreateObject("System.Collections.ArrayList")
	'get the GUID's of the tagged values tracing to the attributes
	dim traceTag as EA.AttributeTag
	for each traceTag in column.TaggedValues
		if traceTag.Name = "sourceAttribute" and traceTag.Value <> "" then
			dim logicalAttribute as EA.Attribute
			set logicalAttribute = Repository.GetAttributeByGuid(traceTag.Value)
			if not logicalAttribute is nothing then
				'check if the parent is in the lis of logical classes
				if attributeIsOwnedBy(logicalAttribute,logicalClasses) then
					logicalAttributes.Add logicalAttribute
				end if
			end if
		end if
	next
	set getLogicalAttributes = logicalAttributes
end function

function attributeIsOwnedBy(logicalAttribute,logicalClasses)
	attributeIsOwnedBy = false
	dim logicalClass as EA.Element
	for each logicalClass in logicalClasses
		if logicalClass.ElementID = logicalAttribute.ParentID then
			attributeIsOwnedBy = true
			exit for
		end if
	next
end function

function getLocialClasses(table)
	dim logicalClasses
	set logicalClasses = CreateObject("System.Collections.ArrayList")
	dim trace as EA.Connector
	for each trace in table.Connectors
		if trace.ClientID = table.ElementID _
		AND trace.Type = "Abstraction" _
		AND trace.Stereotype = "trace" then
			dim logicalClass as EA.Element
			set logicalClass = Repository.GetElementByID(trace.SupplierID)
			if logicalClass.Type = "Class" then
				logicalClasses.Add logicalClass
			end if
		end if
	next
	set getLocialClasses = logicalClasses
end function


main