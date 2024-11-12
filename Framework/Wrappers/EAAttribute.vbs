'[path=\Framework\Wrappers]
'[group=Wrappers]

!INC Utils.Include
!INC Local Scripts.EAConstants-VBScript
' Author: Geert Bellekens
' Purpose: A wrapper class for a EA.Attribute
' Date: 2023-05-09

'"static" property propertyNames
dim EAEAttributePropertyNames
set EAEAttributePropertyNames = nothing

'static global cache
dim EAAttributeCache
set EAAttributeCache = nothing

dim EAAttributesCached
EAAttributesCached = true

'initializes the metadata for EA elements (containing all columnNames of t_object
function initializeEAEAttributePropertyNames()
	dim result
	set result = getArrayListFromQueryWithHeaders("select top 1 * from t_attribute")
	set EAEAttributePropertyNames = result(0) 'get the headers
end function

Class EAAttribute
	Private m_properties
	
	'constructor
	Private Sub Class_Initialize
		set m_properties = Nothing
		if EAEAttributePropertyNames is nothing then
			initializeEAEAttributePropertyNames
		end if
		if EAAttributeCache is nothing then
			set EAAttributeCache = CreateObject("Scripting.Dictionary")
		end if
	end sub
	
	public default function Item (propertyName)
		Item = me.Properties.Item(propertyName)
	end function
	
	Public Property Get Properties
		set Properties = m_properties
	End Property

	Public function initializeProperties(propertyList)
		'initialize with new Dictionary
		set m_properties = CreateObject("Scripting.Dictionary")
		dim i
		i = 0
		dim propertyName
		for each propertyName in EAEAttributePropertyNames
			'fill the dictionary
			m_properties.Add propertyName, propertyList(i)
			'add the counter
			i =  i + 1
		next
	end function
end class

function getEAAttributesForElementIDs(elementIDString)
	'initialize if needed
	if EAAttributeCache is nothing then
		set EAAttributeCache = CreateObject("Scripting.Dictionary")
	end if
	dim attributeIDs
	'get all attributeIDs's for the elementIDs
	set attributeIDs = getEAAttributeIDsForElementIDs(elementIDString)
	'remove the attributeIDs's already in the cache
	dim newAttributeIDs
	set newAttributeIDs = CreateObject("System.Collections.ArrayList") 
	dim attributeID
	for each attributeID in attributeIDs
		if not EAAttributeCache.Exists(attributeID) then
			newAttributeIDs.Add attributeID
		end if
	next
	if newAttributeIDs.Count > 0 then
		dim sqlGetdata
		sqlGetdata = "select * from t_attribute a where a.ID in (" & Join(newAttributeIDs.ToArray(), ",") & ")"
		dim queryResults
		set queryResults = getArrayListFromQuery(sqlGetdata)
		dim row
		for each row in queryResults
			dim newAttribute
			set newAttribute = New EAAttribute
			newAttribute.initializeProperties row
			'add to dictionary based on ID
			EAAttributeCache.Add newAttribute("ID"), newAttribute
		next
	end if
	dim attributesDictionary
	set attributesDictionary = CreateObject("Scripting.Dictionary")
	'get the total list from the cache
	for each attributeID in attributeIDs
		attributesDictionary.Add attributeID, EAAttributeCache(attributeID)
	next
	'return
	set getEAAttributesForElementIDs = attributesDictionary
end function

'function getEAAttributesForElementID(elementID)
'	dim attributesDictionary
'	set attributesDictionary = CreateObject("Scripting.Dictionary")
'	dim attribute
'	for each attribute in EAAttributeCache.Items
'		if attribute("Object_ID") = elementID then
'			attributesDictionary.Add attribute("ID"), attribute
'		end if
'	next
'	'return 
'	set getEAAttributesForElementID = attributesDictionary
'end function

function getEAAttributeIDsForElementIDs(elementIDString)
	dim attributeIDs
	set attributeIDs = CreateObject("System.Collections.ArrayList") 'initialize
	dim sqlGetData 
	sqlGetData = "select A.ID from t_attribute a where a.Object_ID in (" & elementIDString & ")"
	dim result
	set result = getVerticalArrayListFromQuery(sqlGetData)
	if result.Count > 0 then
		set attributeIDs = result(0)
	end if
	'return
	set getEAAttributeIDsForElementIDs = attributeIDs
end function


function test
	dim elements
	set elements = getEAAttributesForElementID("478645")
	dim element
	for each element in elements.Items
		Session.Output now() & " AttributeName: " & element("Name")
	next
end function
'test