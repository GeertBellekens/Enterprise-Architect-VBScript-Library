'[path=\Framework\EAWrappers]
'[group=EAWrappers]

' Author: Geert Bellekens
' Purpose: A wrapper class for a EA.Package
' Date: 2023-05-09


'"static" property propertyNames
dim EAPackagePropertyNames
set EAPackagePropertyNames = nothing
'static global cache
dim EAPackageCache
set EAPackageCache = nothing

'const EAPackageColumns = "Object_ID, Object_Type, Diagram_ID, Name, Alias, Author, Version, Note, Package_ID, Stereotype, NType, CreatedDate, ModifiedDate, Status, Abstract, PDATA1, PDATA2, PDATA3, PDATA4, PDATA5,GenType, GenFile, GenOption, GenLinks, Classifier, ea_guid, ParentID, RunState, Classifier_guid, TPos, IsRoot, Multiplicity, StyleEx"

'initializes the metadata for EA packages (containing all columnNames of t_object
function initializeEAPackagePropertyNames()
	dim result
	set result = getArrayListFromQueryWithHeaders("select top 1 * from t_package")
	dim headersRow 
	set headersRow = result(0)
	set EAPackagePropertyNames = CreateObject("Scripting.Dictionary")
	dim i
	for i = 0 to headersRow.Count -1
		EAPackagePropertyNames.Add lcase(headersRow(i)), i
	next
end function

Class EAPackage
	Private m_properties
	Private m_elements
	Private m_element
	
	'constructor
	Private Sub Class_Initialize
		set m_properties = Nothing
		set m_elements = Nothing 
		set m_element = nothing
		if EAPackagePropertyNames is nothing then
			initializeEAPackagePropertyNames
		end if
		if EAPackageCache is nothing then
			set EAPackageCache = CreateObject("Scripting.Dictionary")
		end if
	end sub
	
	public default function Item (propertyName)
		dim index
		index = EAPackagePropertyNames(lcase(propertyName))
		Item = m_properties.Item(index)
	end function
	
	Public Property Get Properties
		set Properties = m_properties
	End Property
		
	Public Property Get ObjectType
		ObjectType = "EAPackage"
	End Property
	
	Public Property Get Name
		Name = me("Name")
	End Property
	
	Public Property Get ParentID
		ParentID = me("Parent_ID")
	End Property
	
	Public Property Get PackageID
		PackageID = me("Package_ID")
	End Property
	
	Public Property Get Version
		Version = me("Version")
	End Property
	
	Public Property Get Element
		if m_element is nothing then
			set m_element = getEAElementForGUID(me("ea_guid"))
		end if
		set Element = m_element
	End Property
	
	Public Property Get Elements
		if m_elements is nothing then
			set m_elements = getAllEAElementsForPackageIDs(me("Package_ID"))
			Session.Output "Getting my own elements is expensive"
		end if
		set Elements = m_elements
	End Property
	
	Public function AddElement(element)
		if m_elements is nothing then
			set m_elements = CreateObject("Scripting.Dictionary")
		end if
		if not m_elements.Exists(element("Object_ID")) then
			m_elements.Add element("Object_ID"), element
		end if
	end function
	
	

	Public function initializeProperties(propertyList)
		set m_properties = propertyList
	end function
	
	
end class

function getEAPackageForGUID(guid)
	dim sqlGetData
	sqlGetData = "select p.Package_ID from t_package p where p.ea_guid = '" & guid & "'"
	dim packageIDs
	set packageIDs = getFirstColumnArrayListFromQuery(sqlGetData)
	dim eaPackage
	set eaPackage = nothing
	if packageIDs.Count > 0 then
		set eaPackage = getEAPackageForPackageID(packageIDs(0))
	end if
	'return
	set getEAPackageForGUID = eaPackage
end function

function getEAPackageForPackageID(packageID)
	dim package
	set package = nothing
	dim packages
	set packages = getAllEAPackagesForPackageIDs(packageID)
	if packages.Count > 0 then
		if packages.Exists(packageID) then
			set package = packages(packageID)
		else
			Session.Output "ERROR: getEAPackageForPackageID(packageID) for packageID: " & packageID
		end if
	end if
	'return
	set getEAPackageForPackageID = package
end function

function getAllEAPackagesForPackageIDs(packageIDstring)
	'initialize if needed
	if EAPackageCache is nothing then
		set EAPackageCache = CreateObject("Scripting.Dictionary")
	end if
	'remove the packageID's already in the cache
	dim newPackageIDs
	set newPackageIDs = CreateObject("System.Collections.ArrayList") 
	dim packageIDs
	packageIDs = Split(packageIDstring, ",")
	dim packageID
	for each packageID in packageIDs
		if not EAPackageCache.Exists(packageID) then
			newPackageIDs.Add packageID
		end if
	next
	'get the remaining 
	if newPackageIDs.Count > 0 then
		dim sqlGetdata
		sqlGetdata = "select * from t_package p where p.Package_ID in (" & Join(newPackageIDs.ToArray(), ",") & ")"
		dim queryResults
		set queryResults = getArrayListFromQuery(sqlGetdata)
		dim row
		for each row in queryResults
			dim newPackage
			set newPackage = New EAPackage
			newPackage.initializeProperties row
			'add to cache based on packageID
			EAPackageCache.Add newPackage("Package_ID"), newPackage
		next
	end if
	dim packagesDictionary
	set packagesDictionary = CreateObject("Scripting.Dictionary")
	'get the total list from the cache
	for each packageID in packageIDs
		packagesDictionary.Add packageID, EAPackageCache(packageID)
	next
	'return
	set getAllEAPackagesForPackageIDs = packagesDictionary
end function

function loadPackageContentsForPackageIDs(packageIDstring)
	'Loading packages
	Repository.WriteOutput outPutName, now() & " Loading Packages",0
	getAllEAPackagesForPackageIDs packageIDString
	'Loading elements
	Repository.WriteOutput outPutName, now() & " Loading Elements",0
	getAllEAElementsForPackageIDs packageIDstring
	'Loading attributes
	Repository.WriteOutput outPutName, now() & " Loading Attributes",0
	getEAAttributesForPackageIDs packageIDstring
	'Loading connectors
	Repository.WriteOutput outPutName, now() & " Loading Associations",0
	getEAConnectorsForPackageIDs packageIDString, "Association','Aggregation"
end function

function loadAllPackageContentsForPackageGUID(packageGUID)
	'get the package
	dim package
	set package = getEAPackageForGUID(packageGUID)
	if package is nothing then
		exit function
	end if
	'get the packageIDstring
	dim packageIDString
	packageIDString = getPackageTreeIDString(package)
	'load the package tree
	loadPackageContentsForPackageIDs packageIDstring 
end function



function test
	dim startTime
	startTime = Timer()
	dim package as EA.Package 
	'set package = Repository.GetPackageByGuid("{EEF07879-A6AB-4c06-9D32-0E6F9D04C817}") 'UMIG DGO Measure
	'set package = Repository.GetPackageByGuid("{E325EEC9-C25B-47e3-9D8B-6203A0E54055}") 'Test package
	'set package = Repository.GetPackageByGuid("{4D91445D-438D-458f-8757-A386A0718250}") 'UMIG Measure | Exchange Meter Read
    set package = Repository.GetTreeSelectedPackage
	
	dim packageIDString
	packageIDString = GetPackageTreeIDString(package)
	dim packages
	Session.Output now & " Caching Packages"
	'caching packages
	set packages = getAllEAPackagesForPackageIDs(packageIDString)
	Session.Output now & " Number of Packages cached: " & packages.Count
	Session.Output now & " Caching Elements"
	'with pre caching of elements
	getAllEAElementsForPackageIDs packageIDstring
	Session.Output now & " Number of Elements cached: " & EAElementCache.Count
	Session.Output now & " Caching attributes"
	'with pre-caching of all the attributes
	getEAAttributesForPackageIDs packageIDstring
	Session.Output now & " Number of Attributes cached: " & EAAttributeCache.Count
	Session.Output now & " Caching Associations"
	'caching of connectors
	getEAConnectorsForPackageIDs packageIDString, "Association','Aggregation"
	Session.Output now & " Number of Associations cached: " & EAConnectorCache.Count
	dim packageCount
	packageCount = 0
	dim elementCount
	elementCount = 0
	dim attributeCount
	attributeCount = 0
	dim associationCount
	associationCount = 0 
	'loop packages
	dim EApackage
	for each EApackage in packages.Items
		packageCount = packageCount + 1
		'Session.Output now() & " PackageName: " & EApackage("Name")
		dim element
		for each element in EApackage.Elements.Items
		elementCount = elementCount + 1
			'Session.Output now() & " ElementName: " & element("Name")
			dim attribute
			for each attribute in element.Attributes.Items
				attributeCount = attributeCount + 1
				'Session.Output now() & " AttributeName: " & attribute("Name")
			next
			dim connector
			for each connector in element.Associations.Items
				'Session.Output now() & " Connector: " & connector("Connector_Type") & " " & connector("Stereotype")
				associationCount = associationCount + 1
			next
		next
	next
	Session.Output "Number of packages: " & packageCount
	Session.Output "Number of Elements: " & elementCount
	Session.Output "Number of Attributes: " & attributeCount
	Session.Output "Number of Associations: " & associationCount
	EndTime = Timer()
	Session.Output now() & " Runtime: " & FormatNumber(EndTime - StartTime, 2)

end function

function testcompare
	dim startTime
	startTime = Timer()
	dim package as EA.Package 
	set package = Repository.GetPackageByGuid("{EEF07879-A6AB-4c06-9D32-0E6F9D04C817}")
	compareLoopPackage(package)
	EndTime = Timer()
	Session.Output now() & " Runtime: " & FormatNumber(EndTime - StartTime, 2)
end function

function compareLoopPackage(package)
	dim element as EA.Element
	for each element in package.Elements
		Session.Output now() & " elementName: " &  element.Name
		dim attribute as EA.Attribute
		for each attribute in element.Attributes
			Session.Output now() & " AttributeName: " & attribute.Name
		next
		dim connector as EA.Connector
		for each connector in element.Connectors
			Session.Output now() & " Connector: " & connector.Type & " " & connector.Stereotype
		next
	next
	dim subPackage
	for each subPackage in package.Packages
		compareLoopPackage subPackage 
	next
end function

function testLoadMapping
	Session.Output now() & " Start load mapping"
	Repository.WriteOutput outPutName, now() & " Start load mapping", 0
	dim startTime
	startTime = Timer()
	loadAllPackageContentsForPackageGUID("{9CC085FC-8701-4aa4-8E6A-AC4530EB55C0}")
	EndTime = Timer()
	Repository.WriteOutput outPutName, now() & " Runtime: " & FormatNumber(EndTime - StartTime, 2), 0
end function 

'testLoadMapping

'test
'testcompare