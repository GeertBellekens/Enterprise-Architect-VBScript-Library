'[path=\Framework\EAWrappers]
'[group=EAWrappers]

' Author: Geert Bellekens
' Purpose: A wrapper class for a EA.Attribute
' Date: 2023-05-09

'"static" property propertyNames
dim EAEConnectorPropertyNames
set EAEConnectorPropertyNames = nothing

'static global cache
dim EAConnectorCache
set EAConnectorCache = nothing

const EAConnectorColumns = " c.Connector_ID, c.Name, c.Direction, c.Notes, c.Connector_Type, c.SubType, c.SourceCard, c.SourceAccess, c.DestCard, c.SourceRole, c.SourceIsAggregate, c.DestRole, c.DestIsAggregate, c.Start_Object_ID, c.End_Object_ID, c.SeqNo,Stereotype, c.PDATA1, c.PDATA2, c.PDATA3, c.PDATA4, c.PDATA5, c.DiagramID, c.ea_guid, c. SourceIsNavigable, c.DestIsNavigable, c.StyleEx, c.SourceStyle, c.DestStyle"

'initializes the metadata for EA elements (containing all columnNames of t_object
function initializeEAEConnectorPropertyNames()
	dim result
	set result = getArrayListFromQueryWithHeaders("select top 1 " & EAConnectorColumns & " from t_connector c")
	dim headersRow 
	set headersRow = result(0)
	set EAEConnectorPropertyNames = CreateObject("Scripting.Dictionary")
	dim i
	for i = 0 to headersRow.Count -1
		EAEConnectorPropertyNames.Add lcase(headersRow(i)), i
	next
end function

Class EAConnector
	Private m_properties
	Private m_taggedValues
	
	'constructor
	Private Sub Class_Initialize
		set m_properties = Nothing
		set m_taggedValues = Nothing
		if EAEConnectorPropertyNames is nothing then
			initializeEAEConnectorPropertyNames
		end if
	end sub
	
	public default function Item (propertyName)
		dim index
		index = EAEConnectorPropertyNames(lcase(propertyName))
		Item = m_properties.Item(index)
	end function
	
	Public Property Get ObjectType
		ObjectType = "EAConnector"
	End Property
	
	Public Property Get Properties
		set Properties = m_properties
	End Property
	
	Public Property Get ConnectorGUID
		ConnectorGUID = me("ea_guid")
	end Property
	
	Public Property Get Name
		Name = me("Name")
	End Property	
	
	Public Property Get SourceIsAggregate
		SourceIsAggregate = me("SourceIsAggregate")
	End Property
	
	Public Property Get ClientID
		ClientID = me("Start_Object_ID")
	End Property
	
	Public Property Get SupplierID
		SupplierID = me("End_Object_ID")
	End Property
	
	Public Property Get DestCard
		DestCard = me("DestCard")
	End Property
	

	
	Public Property Get TaggedValues
		if m_taggedValues is nothing then
			set m_taggedValues = getEATaggedValuesForElementID(me("Connector_ID"), otConnector)
		end if
		'return only the tagged values themselves
		TaggedValues = m_taggedValues.Items()
	end Property

	Public function initializeProperties(propertyList)
		set m_properties = propertyList
		'add to source and target element, if they are in the cache
		addToElement me("Start_Object_ID")
		addToElement me("End_Object_ID")
	end function
	
	private function addToElement(elementID)
		if EAElementCache.Exists(elementID) then
			dim ownerElement
			set ownerElement = EAElementCache(elementID)
			ownerElement.AddConnector(me)
		end if
	end function
end class

function getEAConnectorForGUID(connectorGUID)
	dim connector
	set connector = nothing 'initialize
	dim connectorID
	connectorID = ""
	dim sqlGetData 
	sqlGetData = "select c.Connector_ID from t_connector c where c.ea_guid = '" & connectorGUID & "'"
	dim result
	set result = getVerticalArrayListFromQuery(sqlGetData)
	if result.Count > 0 then
		if result(0).count > 0 then
			connectorID = result(0)(0)
			dim connectors
			set connectors = getEAConnectorsForConnectorIDs(Array(connectorID))
			if connectors.Count > 0 then
				set connector = connectors.Items()(0)
			end if
		end if
	end if
	'return
	set getEAConnectorForGUID = connector
end function

function getEAConnectorsForPackageIDs(packageIDString, connectortype)
	dim connectorIDs
	'get all connectorIDs's for the elementIDs
	'Session.Output now & " Getting connectorIDs"
	set connectorIDs = getEAConnectorIDsForPackageIDs(packageIDString, connectortype)
	'Session.Output now & " ConnectorIDs.Count: " & connectorIDs.Count
	'return
	set getEAConnectorsForPackageIDs = getEAConnectorsForConnectorIDs(connectorIDs)
end function

function getEAConnectorsForElementIDs(elementIDString, connectortype)
	dim connectorIDs
	'get all connectorIDs's for the elementIDs
	set connectorIDs = getEAConnectorIDsForElementIDs(elementIDString, connectortype)
	'return
	set getEAConnectorsForElementIDs = getEAConnectorsForConnectorIDs(connectorIDs)
end function

function getEAConnectorsForConnectorIDs(connectorIDs)
	'initialize if needed
	if EAConnectorCache is nothing then
		set EAConnectorCache = CreateObject("Scripting.Dictionary")
	end if
	'remove the connectorIDs's already in the cache
	dim newConnectorIDs
	set newConnectorIDs = CreateObject("System.Collections.ArrayList") 
	'Session.Output now & " Getting new Connector IDs"
	dim connectorID
	for each connectorID in connectorIDs
		if not EAConnectorCache.Exists(connectorID) then
			newConnectorIDs.Add connectorID
		end if
	next
	'Session.Output now & " Getting Data for new ids count:  " & newConnectorIDs.Count
	if newConnectorIDs.Count > 0 then
		dim sqlGetdata
		sqlGetdata = "select " & EAConnectorColumns & " from (t_connector c                          " & vbNewLine & _
					" left join t_connectortag tv on (tv.ElementID = c.Connector_ID                  " & vbNewLine & _
					"           						and tv.Property = 'sequencingKey'))          " & vbNewLine & _
					" where c.Connector_ID in (" & Join(newConnectorIDs.ToArray(), ",") & ")         " & vbNewLine & _
					" order by c.Start_Object_ID, tv.VALUE                                           "	
		dim queryResults
		set queryResults = getArrayListFromQuery(sqlGetdata)
		dim row
		'Session.Output now & " creating connectors  " & newConnectorIDs.Count
		for each row in queryResults
			dim newConnector
			set newConnector = New EAConnector
			newConnector.initializeProperties row
			'add to dictionary based on ID
			EAConnectorCache.Add newConnector("Connector_ID"), newConnector
		next
	end if
	dim connectorsDictionary
	set connectorsDictionary = CreateObject("Scripting.Dictionary")
	'Session.Output now & " Looping connectors the get the ones we need"
	'get the total list from the cache
	for each connectorID in connectorIDs
		connectorsDictionary.Add connectorID, EAConnectorCache(connectorID)
	next
	'Session.Output now & " Connectors filtered"
	'return
	set getEAConnectorsForConnectorIDs = connectorsDictionary
end function

function getEAConnectorIDsForElementIDs(elementIDString, connectortype)
	dim connectorIDs
	set connectorIDs = CreateObject("System.Collections.ArrayList") 'initialize
	dim sqlGetData 
	sqlGetData = "select c.Connector_ID from t_connector c where c.Start_Object_ID in (" & elementIDString & ") " & vbNewLine
	if len(connectortype) > 0 then
		sqlGetdata = sqlGetdata & vbNewLine & " and c.Connector_Type in ('" & connectortype & "')"
	end if						
	sqlGetData = sqlGetData  & " union " & vbNewLine & _
				" select c.Connector_ID from t_connector c where c.End_Object_ID in (" & elementIDString & ") " & vbNewLine
	if len(connectortype) > 0 then
		sqlGetdata = sqlGetdata & vbNewLine & " and c.Connector_Type in ('" & connectortype & "')"
	end if		
	dim result
	set result = getVerticalArrayListFromQuery(sqlGetData)
	if result.Count > 0 then
		set connectorIDs = result(0)
	end if
	'return
	set getEAConnectorIDsForElementIDs = connectorIDs
end function

function getEAConnectorIDsForPackageIDs(packageIDString, connectortype)
 	dim connectorIDs
	set connectorIDs = CreateObject("System.Collections.ArrayList") 'initialize
	dim sqlGetData 
	sqlGetData = "select distinct c.Connector_ID                                                 " & vbNewLine & _
				" from t_connector c                                                             " & vbNewLine & _
				" inner join t_object o on o.Object_ID in (c.Start_Object_ID, c.End_Object_ID)   " & vbNewLine & _
				" where o.Package_ID in (" & packageIDString & ")                                "
	if len(connectortype) > 0 then
		sqlGetdata = sqlGetdata & vbNewLine & " and c.Connector_Type in ('" & connectortype & "')"
	end if		
	
	dim result
	set result = getVerticalArrayListFromQuery(sqlGetData)
	if result.Count > 0 then
		set connectorIDs = result(0)
	end if
	'return
	set getEAConnectorIDsForPackageIDs = connectorIDs
end function

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



'function test
'	dim elements
'	set elements = getEAConnectorsForElementID("478645", "Association','Aggregation")
'	dim element
'	for each element in elements.Items
'		Session.Output now() & " ConnectorID: " & element("ea_guid")
'	next
'end function
'test