'[path=\Framework\EAWrappers]
'[group=EAWrappers]

' Author: Geert Bellekens
' Purpose: A wrapper class for a EA.Attribute
' Date: 2023-05-09

'"static" property propertyNames
dim EAEAttributePropertyNames
set EAEAttributePropertyNames = nothing

'static global cache
dim EAAttributeCache
set EAAttributeCache = nothing


'initializes the metadata for EA elements (containing all columnNames of t_object
function initializeEAEAttributePropertyNames()
	dim result
	set result = getArrayListFromQueryWithHeaders("select top 1 * from t_attribute")
	dim headersRow 
	set headersRow = result(0)
	set EAEAttributePropertyNames = CreateObject("Scripting.Dictionary")
	dim i
	for i = 0 to headersRow.Count -1
		EAEAttributePropertyNames.Add lcase(headersRow(i)), i
	next
end function

Class EAAttribute
	Private m_properties
	Private m_typeElement
	Private m_taggedValues
	
	'constructor
	Private Sub Class_Initialize
		set m_properties = Nothing
		set m_typeElement = nothing
		set m_taggedValues = nothing
		if EAEAttributePropertyNames is nothing then
			initializeEAEAttributePropertyNames
		end if
		if EAAttributeCache is nothing then
			set EAAttributeCache = CreateObject("Scripting.Dictionary")
		end if
	end sub
	
	public default function Item (propertyName)
		dim index
		index = EAEAttributePropertyNames(lcase(propertyName))
		Item = me.Properties.Item(index)
	end function
	
	Public Property Get ObjectType
		ObjectType = "EAAttribute"
	End Property
	
	Public Property Get Properties
		set Properties = m_properties
	End Property
	
	Public Property Get Name
		Name = me("Name")
	End Property
		
	Public Property Get Stereotype
		Stereotype = me("Stereotype")
	End Property
	
	Public Property Get UpperBound
		UpperBound = me("UpperBound")
	end Property
	
	Public Property Get LowerBound
		LowerBound = me("LowerBound")
	end Property
	
	Public Property Get AttributeGUID
		AttributeGUID = me("ea_guid")
	end Property
	
	Public Property Get ParentID
		ParentID = me("Object_ID")
	end Property
	
	
	Public Property Get Alias
		Alias = me("Style")
	end Property	
	
	Public Property Get TypeElement
		if m_typeElement is nothing _
		  and isnumeric(me("Classifier")) _
		  and me("Classifier") <> "0" then
			set m_typeElement = getEAElementForElementID(me("Classifier"))
		end if
		set TypeElement = m_typeElement
	End Property
	
	public Property Get TaggedValues
		if m_taggedValues is nothing then
			set m_taggedValues = getEATaggedValuesForElementID(me("ID"), otAttribute)
		end if
		TaggedValues = m_taggedValues.Items()
	End Property

	Public function initializeProperties(propertyList)
		set m_properties = propertyList
		'add myself to the parent element
		if not EAElementCache.Exists(me("Object_ID")) then
			getEAElementsForElementIDs(Array(me("Object_ID")))
		end if
		if EAElementCache.Exists(me("Object_ID")) then
			dim ownerElement
			set ownerElement = EAElementCache(me("Object_ID"))
			ownerElement.AddAttribute(me)
		end if
	end function
end class

function getEAAttributesForPackageIDs(packageIDString)
	dim sqlGetData 
	sqlGetData = "select a.ID from t_attribute a                        " & vbNewLine & _
				" inner join t_object o on o.Object_ID = a.Object_ID   " & vbNewLine & _
				" where o.Package_ID in (" & packageIDString & ")      " & vbNewLine & _
				" order by a.Object_ID, a.Pos, a.Name                  "
	dim attributeIDs
	'get all attributeIDs for the packageIDString
	set attributeIDs = getFirstColumnArrayListFromQuery(sqlGetData)
	'return
	set getEAAttributesForPackageIDs = getEAAttributesForAttributeIDs(attributeIDs)
end function

function getEAAttributeForGUID(attributeGUID)
	dim attribute
	set attribute = nothing 'initialiaze
	dim sqlGetData 
	sqlGetData = "select A.ID from t_attribute a where a.ea_guid = '" & attributeGUID & "'"
	dim attributes
	set attributes = getEAAttributesFromQuery(sqlGetData)
	'get the first one
	if attributes.Count > 0 then
		set attribute = attributes.Items()(0)
	end if
	'return
	set getEAAttributeForGUID = attribute
end function

function getEAAttributesFromQuery(sqlGetData)
	dim attributeIDs
	'get all attributeIDs
	set attributeIDs = getFirstColumnArrayListFromQuery(sqlGetData)
	'return
	set getEAAttributesFromQuery = getEAAttributesForAttributeIDs(attributeIDs)
end function

function getEAAttributesForElementIDs(elementIDString)
	dim sqlGetData 
	sqlGetData = "select A.ID from t_attribute a where a.Object_ID in (" & elementIDString & ") " & vbNewLine & _
				" order by a.Object_ID, a.Pos, a.Name                                           "
	'return
	set getEAAttributesForElementIDs = getEAAttributesFromQuery(sqlGetData)
end function

function getEAAttributesForAttributeIDs(attributeIDs)
	'initialize if needed
	if EAAttributeCache is nothing then
		set EAAttributeCache = CreateObject("Scripting.Dictionary")
	end if
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
	set getEAAttributesForAttributeIDs = attributesDictionary
end function



function getEAAttributeIDsForPackageIDs(packageIDString)
 	dim attributeIDs
	set attributeIDs = CreateObject("System.Collections.ArrayList") 'initialize
	dim sqlGetData 
	sqlGetData = "select a.ID from t_attribute a                        " & vbNewLine & _
				" inner join t_object o on o.Object_ID = a.Object_ID   " & vbNewLine & _
				" where o.Package_ID in (" & packageIDString & ")      "
	dim result
	set result = getVerticalArrayListFromQuery(sqlGetData)
	if result.Count > 0 then
		set attributeIDs = result(0)
	end if
	'return
	set getEAAttributeIDsForPackageIDs = attributeIDs
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