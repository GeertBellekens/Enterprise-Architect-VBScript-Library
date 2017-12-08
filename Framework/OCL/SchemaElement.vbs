'[path=\Framework\OCL]
'[group=OCL]


'Author: Geert Bellekens
'Date: 2017-12-06
'Purpose: Class representing a Schema Element

Class SchemaElement 
	'private variables
	Private m_Source
	Private m_Name
	Private m_IsRedefiniton
	Private m_Properties
	Private m_Redefines
	
	'constructor
	Private Sub Class_Initialize
		m_Name = ""
		set m_Source = nothing
		set m_Properties = CreateObject("Scripting.Dictionary")
		set m_Redefines = CreateObject("Scripting.Dictionary")
	End Sub
	
	' Source property. (EA.Element)
	Public Property Get Source
		set Source = m_Source
	End Property
	Public Property Let Source(value)
		set m_Source = value
		if m_Name = "" then
			m_Name = me.Source.Name
		end if
	End Property
	
	' Name property.
	Public Property Get Name
		Name = m_Name
	End Property
	Public Property Let Name(value)
		m_Name = value
	End Property
	
	' IsRedefinition property.
	Public Property Get IsRedefinition
		IsRedefinition = m_IsRedefinition
	End Property
	Public Property Let IsRedefinition(value)
		m_IsRedefinition = value
	End Property
	
	' Guid property.
	Public Property Get GUID
		GUID = me.Source.ElementGUID
	End Property
	
	' Properties property
	Public Property Get Properties
		set Properties = m_Properties
	End Property
	
	' Redefines property
	Public Property Get Redefines
		set Redefines = m_Redefines
	End Property

	
	public function getProperty(identifierPart)
		'initialize null
		set getProperty = nothing
		'first check if there is an attribute on the localContext with the given name
		dim sqlGetAttribute
		sqlGetAttribute = "select a.ID from t_attribute a " & _
						" where a.Object_ID = " & me.Source.ElementID  & _
						" and a.Name = '" & identifierPart & "' "
		'get the attribute
		dim attributes
		set attributes = getattributesFromQuery(sqlGetAttribute)
		if attributes.Count > 0 then
			'return the first attribute
			set getProperty = me.addAttributeProperty(attributes(0))
		else
			'get association end
			dim associationEnd
			set associationEnd = nothing
			'first check target end
			dim sqlGetTargetEnd
			sqlGetTargetEnd	= "select c.Connector_ID from t_connector c" & _
							" where c.Start_Object_ID = " & me.Source.ElementID  & _
							" and c.DestRole = '" & identifierPart & "' "
			dim connector as EA.Connector
			dim connectors
			set connectors = getConnectorsFromQuery(sqlGetTargetEnd)
			if connectors.Count > 0 then
				set connector = connectors(0)
				set associationEnd = connector.SupplierEnd
			else
				'then source end
				dim sqlGetSourceEnd
				sqlGetSourceEnd = "select c.Connector_ID from t_connector c" & _
							" where c.End_Object_ID = " & me.Source.ElementID  & _
							" and c.SourceRole = '" & identifierPart & "' "
				set connectors = getConnectorsFromQuery(sqlGetSourceEnd)
				if connectors.Count > 0 then
					set connector = connectors(0)
					set associationEnd = connector.ClientEnd
				end if
			end if
			'create the SchemeProperty
			if not associationEnd is nothing then
				set getProperty = me.addConnectorEndProperty(associationEnd, connector)
			end if
		end if
	end function
		
	' Add Attribute poperty to properties
	Public function addAttributeProperty(newAttribute)
		if me.Properties.Exists(newAttribute.AttributeGUID) then
			'return existing item
			set addAttributeProperty = me.Properties.Item(newAttribute.AttributeGUID)
		else
			'create new item
			dim newProperty
			set newProperty = new SchemaProperty
			newProperty.Source = newAttribute
			me.Properties.Add newProperty.GUID, newProperty
			'return new item
			set addAttributeProperty = newProperty
		end if
	End function
	
	' Add Attribute poperty to properties
	Public function addConnectorEndProperty(newConnectorEnd, newConnector)
		if me.Properties.Exists(newConnector.ConnectorGUID) then
			'return existing item
			set addConnectorEndProperty = me.Properties.Item(newConnector.ConnectorGUID)
		else
			'create new item
			dim newProperty
			set newProperty = new SchemaProperty
			newProperty.Source = newConnectorEnd
			newProperty.Connector = newConnector
			me.Properties.Add newProperty.GUID, newProperty
			'return new item
			set addConnectorEndProperty = newProperty
		end if
	End function
	
	' Add a redefines  element
	Public function addRedefine(newRedefine)
		if not m_Redefines.Exists(newRedefine.Name) then
			m_Redefines.Add newRedefine.Name, newRedefine
		end if
	End function

end Class