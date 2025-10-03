'[path=\Framework\EAWrappers]
'[group=EAWrappers]

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
	dim headersRow 
	set headersRow = result(0)
	set EAElementPropertyNames = CreateObject("Scripting.Dictionary")
	dim i
	for i = 0 to headersRow.Count -1
		EAElementPropertyNames.Add lcase(headersRow(i)), i
	next
end function

Class EAElement
	Private m_properties
	Private m_attributes
	Private m_connectors
	Private m_associations
	Private m_baseClasses
	Private m_taggedValues
	
	'constructor
	Private Sub Class_Initialize
		set m_properties = Nothing
		set m_attributes = Nothing 
		set m_connectors = Nothing
		set m_associations = Nothing
		set m_baseClasses = nothing
		set m_taggedValues = Nothing
		if EAElementPropertyNames is nothing then
			initializeEAElementPropertyNames
		end if
		if EAElementCache is nothing then
			set EAElementCache = CreateObject("Scripting.Dictionary")
		end if
	end sub
	
	public default function Item (propertyName)
		dim index
		index = EAElementPropertyNames(lcase(propertyName))
		Item = m_properties.Item(index)
	end function
	
	Public Property Get Properties
		set Properties = m_properties
	End Property
	
	Public Property Get ObjectType
		ObjectType = "EAElement"
	End Property
	
	Public Property Get Name
		Name = me("Name")
	End Property
		
	Public Property Get Stereotype
		Stereotype = me("Stereotype")
	End Property
	
	Public Property Get PackageID
		PackageID = me("Package_ID")
	end Property
	
	Public Property Get ElementType
		ElementType = me("Object_Type")
	end Property
	
	Public Property Get ElementID
		ElementID = me("Object_ID")
	end Property

	Public Property Get ElementGUID
		ElementGUID = me("ea_guid")
	end Property
	
	Public Property Get Genlinks
		Genlinks = me("Genlinks")
	end Property
	
	Public Property Get TaggedValues
		if m_taggedValues is nothing then
			set m_taggedValues = getEATaggedValuesForElementID(me("Object_ID"), otElement)
		end if
		'return only the tagged values themselves
		TaggedValues = m_taggedValues.Items()
	end Property
	
	Public Property Get BaseClasses
		if m_baseClasses is nothing then
			set m_baseClasses = CreateObject("Scripting.Dictionary") 'initialize emtpy
			'first get the ID's iwht a query, and then get the actual elements
			dim sqlGetData 
			sqlGetData = "select c.End_Object_ID from t_connector c    " & vbNewLine & _
						"  where c.Connector_Type = 'Generalization'   " & vbNewLine & _
						"  and c.Start_Object_ID = " & me("Object_ID")
			dim result
			set result = getVerticalArrayListFromQuery(sqlGetData)
			dim row
			for each row in result
				if row.Count > 0 then
					dim baseClassID
					baseClassID = row(0)
					set m_baseClasses = getEAElementsForElementIDs(Array(baseClassID))
				end if
				exit for
			next
		end if
		set BaseClasses = m_baseClasses
	End Property
	
	Public Property Get Attributes
		if m_attributes is nothing then
			'check if this objectype can even have attributes
			dim elemenType
			elemenType = lcase(me("Object_Type"))
			if elemenType = "class" _
			  or elemenType = "enumeration"_
			  or elemenType = "datatype"_
			  or elemenType = "component" _
			  or elemenType = "interface" then
				set m_attributes = getEAAttributesForElementIDs(me("Object_ID"))
				'msgbox "Getting my own attributes is expensive"
			else
				'set empty dictionary
				set m_attributes = CreateObject("Scripting.Dictionary")
			end if
		end if
		set Attributes = m_attributes
	End Property
	
	Public function AddAttribute(attribute)
		if m_attributes is nothing then
			set m_attributes = CreateObject("Scripting.Dictionary")
		end if
		if not m_attributes.Exists(attribute("ID")) then
			m_attributes.Add attribute("ID"), attribute
		end if
	end function
	
	Public Property Get Connectors
		if m_connectors is nothing then
			set m_connectors = getEAConnectorsForElementIDs(me("Object_ID"), "")
		end if
		set Connectors = m_connectors
	End Property
	
	Public Property Get Associations
		if m_associations is nothing then
			set m_associations = getEAConnectorsForElementIDs(me("Object_ID"), "Association','Aggregation")
		end if
		set Associations = m_associations
	End Property
	
	Public function AddConnector(connector)
		if m_connectors is nothing then
			set m_connectors = CreateObject("Scripting.Dictionary")
		end if
		if not m_connectors.Exists(connector("Connector_ID")) then
			m_connectors.Add connector("Connector_ID"), connector
		end if
		'check if the connector is an association (or aggregation)
		if connector("Connector_Type") = "Association" _
		  or connector("Connector_Type") = "Aggregation" then
			if m_associations is nothing then
				set m_associations = CreateObject("Scripting.Dictionary")
			end if
			if not m_associations.Exists(connector("Connector_ID")) then
				m_associations.Add connector("Connector_ID"), connector
			end if
		end if
	end function
	

	Public function initializeProperties(propertyList)
		set m_properties = propertyList
		'add myself to the parent package
		if not EAPackageCache.Exists(me("Package_ID")) then
			getAllEAPackagesForPackageIDs(me("Package_ID"))
		end if
		if EAPackageCache.Exists(me("Package_ID")) then
			dim ownerPackage
			set ownerPackage = EAPackageCache(me("Package_ID"))
			ownerPackage.AddElement(me)
		end if
		'add myself to the parent element
		'TODO
	end function
	

	
end class


function getEAElementForGUID(guid)
	dim eaElement 
	set eaElement  = nothing
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o where o.ea_guid = '" & guid & "'"
	dim elementIDs
	set elementIDs = getFirstColumnArrayListFromQuery(sqlGetData)
	if elementIDs.Count > 0 then
		set eaElement = getEAElementForElementID(elementIDs(0))
	end if
	'return
	set getEAElementForGUID = eaElement
end function

function getEAElementForElementID(elementID)
	dim eaElement
	set eaElement = nothing 'initialize
	dim elements
	set elements = getEAElementsForElementIDs(Array(elementID))
	if elements.Exists(elementID) then
		set eaElement = elements(elementID)
	end if
	'return
	set getEAElementForElementID = eaElement
end function

function getEAElementsForElementIDs(elementIDs)
	'initialize if needed
	if EAElementCache is nothing then
		set EAElementCache = CreateObject("Scripting.Dictionary")
	end if
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
		'debug
		'Session.Output "getting elements for the additional " & newElementIDs.Count & " first ID: " & newElementIDs(0)
		dim sqlGetdata
		on error resume next
		sqlGetdata = "select " & EAElementColumns & "  from t_object o where o.Object_ID in (" & Join(newElementIDs.ToArray(), ",") & ")"
		if not Err.number = 0 then
			dim debugID
			if isObject(newElementIDs(0)) then
				set debugID = newElementIDs(0)
				Session.Output "debugID.Count: " & debugID.Count
				dim field0
				field0 = debugID(0)
				Session.Output "field0: " & field0
			else
				debugID = newElementIDs(0)
			end if
			Session.Output "Wat nu?"
		end if
		on error goto 0
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
		if EAElementCache.Exists(elementID) then
			elementsDictionary.Add elementID, EAElementCache(elementID)
		end if
	next
	'return
	set getEAElementsForElementIDs = elementsDictionary
end function

function getAllEAElementsForPackageIDs(packageIDstring)
	dim elementIDs
	'get all elementID's for the elements in the given package
	set elementIDs = getEAElementIDsforPackageIDs(packageIDstring)
	set getAllEAElementsForPackageIDs = getEAElementsForElementIDs(elementIDs)
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
	set package = Repository.GetPackageByGuid("{EEF07879-A6AB-4c06-9D32-0E6F9D04C817}")
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

