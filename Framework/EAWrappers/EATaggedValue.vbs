'[path=\Framework\EAWrappers]
'[group=EAWrappers]

!INC Utils.Include
!INC Local Scripts.EAConstants-VBScript
' Author: Geert Bellekens
' Purpose: A wrapper class for a all EATaggedValues
' Date: 2023-05-09

'"static" property propertyNames
dim EATaggedValuePropertyNames
set EATaggedValuePropertyNames = nothing

'initializes the metadata for EA elements (containing all columnNames of t_object
function initializeEATaggedValuePropertyNames()
	dim result
	set result = getArrayListFromQueryWithHeaders("select top 1 * from t_attributetag")
	dim headersRow 
	set headersRow = result(0)
	set EATaggedValuePropertyNames = CreateObject("Scripting.Dictionary")
	dim i
	for i = 0 to headersRow.Count -1
		EATaggedValuePropertyNames.Add lcase(headersRow(i)), i
	next
end function

Class EATaggedValue
	Private m_properties
	
	'constructor
	Private Sub Class_Initialize
		set m_properties = Nothing
		if EATaggedValuePropertyNames is nothing then
			initializeEATaggedValuePropertyNames
		end if
	end sub
	
	public default function Item (propertyName)
		dim index
		index = EATaggedValuePropertyNames(lcase(propertyName))
		Item = m_properties.Item(index)
	end function
	
	Public Property Get Properties
		set Properties = m_properties
	End Property
	
	Public Property Get ObjectType
		ObjectType = "EATaggedValue"
	End Property
	
	Public Property Get Name
		Name = me("Property")
	End Property
	
	Public Property Get Value
		Value = me("Value")
	End Property
	
	Public Property Get Notes
		Notes = me("Notes")
	End Property

	Public function initializeProperties(propertyList)
		set m_properties = propertyList
	end function
end class

function getEATaggedValuesForElementID(elementID, ownerType)
	dim taggedValuesDictionary
	set taggedValuesDictionary = CreateObject("Scripting.Dictionary")
	dim sqlGetdata
	select case ownerType
		case otElement
			sqlGetdata = "select * from t_objectProperties tv where tv.Object_ID = " & elementID
		case otAttribute
			sqlGetdata = "select * from t_attributeTag tv where tv.ElementID = " & elementID
		case otConnector
			sqlGetdata = "select * from t_connectorTag tv where tv.ElementID = " & elementID
		case otMethod
			sqlGetdata = "select * from t_operationTag tv where tv.ElementID = " & elementID
	end select
	dim queryResults
	set queryResults = getArrayListFromQuery(sqlGetdata)
	dim row
	for each row in queryResults
		dim newTaggedValue
		set newTaggedValue = New EATaggedValue
		newTaggedValue.initializeProperties row
		'add to dictionary based on ID
		taggedValuesDictionary.Add newTaggedValue("ea_guid"), newTaggedValue
	next
	'return
	set getEATaggedValuesForElementID = taggedValuesDictionary
end function


function test
	dim elements
	set elements = getEATaggedValuesForElementID("478645", otElement)
	dim element
	for each element in elements.Items
		Session.Output now() & " TaggedValueProperty: " & element("Property")
	next
end function
'test