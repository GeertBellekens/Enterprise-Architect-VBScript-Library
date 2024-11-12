'[path=\Framework\Wrappers]
'[group=Wrappers]

!INC Utils.Include
!INC Local Scripts.EAConstants-VBScript
' Author: Geert Bellekens
' Purpose: A wrapper class for a EA.Attribute
' Date: 2023-05-09

'"static" property propertyNames
dim EAEConnectorPropertyNames
set EAEConnectorPropertyNames = nothing

const EAConnectorColumns = " Connector_ID, Name, Direction, Notes, Connector_Type, SubType, SourceCard, SourceAccess, DestCard, SourceIsAggregate, DestIsAggregate, Start_Object_ID, End_Object_ID, SeqNo,Stereotype, PDATA1, PDATA2, PDATA3, PDATA4, PDATA5, DiagramID, ea_guid,  SourceIsNavigable, DestIsNavigable, StyleEx, SourceStyle, DestStyle"

'initializes the metadata for EA elements (containing all columnNames of t_object
function initializeEAEConnectorPropertyNames()
	dim result
	set result = getArrayListFromQueryWithHeaders("select top 1 " & EAConnectorColumns & " from t_connector")
	set EAEConnectorPropertyNames = result(0) 'get the headers
end function

Class EAConnector
	Private m_properties
	
	'constructor
	Private Sub Class_Initialize
		set m_properties = Nothing
		if EAEConnectorPropertyNames is nothing then
			initializeEAEConnectorPropertyNames
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
		for each propertyName in EAEConnectorPropertyNames
			'fill the dictionary
			m_properties.Add propertyName, propertyList(i)
			'add the counter
			i =  i + 1
		next
	end function
end class

function getEAConnectorsForElementID(elementID, connectortype)
	dim connectorsDictionary
	set connectorsDictionary = CreateObject("Scripting.Dictionary")
	dim sqlGetdata
	sqlGetdata = "Select " & EAConnectorColumns & " from t_connector c            " & vbNewLine & _
				" where " & elementID & " in (c.Start_Object_ID, c.End_Object_ID) "
	if len(connectortype) > 0 then
		sqlGetdata = sqlGetdata & vbNewLine & " and c.Connector_Type in ('" & connectortype & "')"
	end if
		
	dim queryResults
	set queryResults = getArrayListFromQuery(sqlGetdata)
	dim row
	for each row in queryResults
		dim newConnector
		set newConnector = New EAConnector
		newConnector.initializeProperties row
		'add to dictionary based on ID
		connectorsDictionary.Add newConnector("Connector_ID"), newConnector
	next
	'return
	set getEAConnectorsForElementID = connectorsDictionary
end function

function test
	dim elements
	set elements = getEAConnectorsForElementID("478645", "Association','Aggregation")
	dim element
	for each element in elements.Items
		Session.Output now() & " ConnectorID: " & element("ea_guid")
	next
end function
'test