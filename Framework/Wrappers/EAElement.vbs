'[path=\Framework\Wrappers]
'[group=Wrappers]

!INC Utils.Include
!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.EAAttribute
!INC Wrappers.EAConnector
!INC Wrappers.EATaggedValue
' Author: Geert Bellekens
' Purpose: A wrapper class for a EA.Element
' Date: 2023-05-09

'"static" property propertyNames
dim EAElementPropertyNames
set EAElementPropertyNames = nothing
'static global cache
dim EAElementCache
set EAElementCache = nothing

const EAElementColumns = "Object_ID, Object_Type, Diagram_ID, Name, Alias, Author, Version, Note, Package_ID, Stereotype, NType, CreatedDate, ModifiedDate, Status, Abstract, PDATA1, PDATA2, PDATA3, PDATA4, PDATA5,GenType, GenFile, GenOption, GenLinks, Classifier, ea_guid, ParentID, RunState, Classifier_guid, TPos, IsRoot, Multiplicity, StyleEx"

'initializes the metadata for EA elements (containing all columnNames of t_object
function initializeEAElementPropertyNames()
	dim result
	set result = getArrayListFromQueryWithHeaders("select top 1 " & EAElementColumns & " from t_object")
	set EAElementPropertyNames = result(0) 'get the headers
end function

Class EAElement
	Private m_properties
	Private m_attributes
	Private m_connectors
	Private m_associations
	
	'constructor
	Private Sub Class_Initialize
		set m_properties = Nothing
		set m_attributes = Nothing 
		set m_connectors = Nothing
		set m_associations = Nothing
		if EAElementPropertyNames is nothing then
			initializeEAElementPropertyNames
		end if
		if EAElementCache is nothing then
			set EAElementCache = CreateObject("Scripting.Dictionary")
		end if
	end sub
	
	public default function Item (propertyName)
		Item = me.Properties.Item(propertyName)
	end function
	
	Public Property Get Properties
		set Properties = m_properties
	End Property
	
	Public Property Get Attributes
		if m_attributes is nothing then
			set m_attributes = getEAAttributesForElementIDs(me("Object_ID"))
		end if
		set Attributes = m_attributes
	End Property
	
	Public Property Get Connectors
		if m_connectors is nothing then
			set m_connectors = getEAConnectorsForElementID(me("Object_ID"), "")
		end if
		set Connectors = m_connectors
	End Property
	
	Public Property Get Associations
		if m_associations is nothing then
			set m_associations = getEAConnectorsForElementID(me("Object_ID"), "Association','Aggregation")
		end if
		set Associations = m_associations
	End Property
	

	Public function initializeProperties(propertyList)
		'initialize with new Dictionary
		set m_properties = CreateObject("Scripting.Dictionary")
		dim i
		i = 0
		dim propertyName
		for each propertyName in EAElementPropertyNames
			'fill the dictionary
			m_properties.Add propertyName, propertyList(i)
			'add the counter
			i =  i + 1
		next
	end function
	
	
end class

function getAllEAElementsForPackageIDs(packageIDs)
	'initialize if needed
	if EAElementCache is nothing then
		set EAElementCache = CreateObject("Scripting.Dictionary")
	end if
	dim elementIDs
	'get all elementID's for the elements in the given package
	set elementIDs = getEAElementIDsforPackageIDs(packageIDs)
	'remove the elementID's already in the cache
	dim newElementIDs
	set newElementIDs = CreateObject("System.Collections.ArrayList") 
	dim elementID
	for each elementID in elementIDs
		if not EAElementCache.Exists(elementID) then
			newElementIDs.Add elementID
		end if
	next
	'get the remaining 
	if newElementIDs.Count > 0 then
		dim sqlGetdata
		sqlGetdata = "select " & EAElementColumns & "  from t_object o where o.Object_ID in (" & Join(newElementIDs.ToArray(), ",") & ")"
		dim queryResults
		set queryResults = getArrayListFromQuery(sqlGetdata)
		dim row
		for each row in queryResults
			dim newElement
			set newElement = New EAElement
			newElement.initializeProperties row
			'add to cache based on elementID
			EAElementCache.Add newElement("Object_ID"), newElement
		next
	end if
	dim elementsDictionary
	set elementsDictionary = CreateObject("Scripting.Dictionary")
	'get the total list from the cache
	for each elementID in elementIDs
		elementsDictionary.Add elementID, EAElementCache(elementID)
	next
	'return
	set getAllEAElementsForPackageIDs = elementsDictionary
end function



function getEAElementIDsforPackageIds(packageIDstring)
	dim elementIDs
	set elementIDs = CreateObject("System.Collections.ArrayList") 'initialize
	dim sqlGetData 
	sqlGetData = "select o.Object_ID from t_object o where o.Package_ID in ( " & packageIDstring & ")"
	dim result
	set result = getVerticalArrayListFromQuery(sqlGetData)
	if result.Count > 0 then
		set elementIDs = result(0)
	end if
	'return
	set getEAElementIDsforPackageIds = elementIDs
end function


function test
	dim startTime
	startTime = Timer()
	dim package as EA.Package 
	set package = Repository.GetPackageByGuid("{A7890C8D-258C-4e2d-96BE-29BED3FFA7FA}")
	dim packageIDString
	packageIDString = GetPackageTreeIDString(package)
	dim elements
	set elements = getAllEAElementsForPackageIDs(packageIDString)
	dim element
	'with pre-caching of all the attributes
'	dim elementIDs
'	elementIDs = Join(elements.Keys, ",")
'	loadEAAttributesForElementIDs elementIDs
	'loop elements
	for each element in elements.Items
		Session.Output now() & " elementName: " & element("Name")
		dim attribute
		for each attribute in element.Attributes.Items
			Session.Output now() & " AttributeName: " & attribute("Name")
		next
		dim connector
		for each connector in element.Connectors.Items
			Session.Output now() & " Connector: " & connector("Connector_Type") & " " & connector("Stereotype")
		next
	next
	EndTime = Timer()
	Session.Output now() & " Runtime: " & FormatNumber(EndTime - StartTime, 2)

end function

function testcompare
	dim startTime
	startTime = Timer()
	dim package as EA.Package 
	set package = Repository.GetPackageByGuid("{A7890C8D-258C-4e2d-96BE-29BED3FFA7FA}")
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
test
'testcompare