'[path=\Framework\OCL]
'[group=OCL]

!INC Local Scripts.EAConstants-VBScript

'Author: Geert Bellekens
'Date: 2017-12-06
'Purpose: Class representing a Schema Property


Class SchemaProperty
	'private variables
	Private m_Source
	Private m_Connector
	Private m_minOccurs
	Private m_maxOccurs
	Private m_Redefines
	Private m_Restricted
	Private m_Classifier
	private m_Owner
	
	
	'constructor
	Private Sub Class_Initialize
		set m_Source = nothing
		set m_Connector = nothing
		m_Restricted = false
		set m_Classifier = nothing
		m_sourceCardinality = ""
		m_targetCardinality = ""
	End Sub
	
	'Properties
	' Source property. (EA.Attribute or EA.ConnectorEnd)
	Public Property Get Source
	  set Source = m_Source
	End Property
	Public Property Let Source(value)
	  set m_Source = value
	  set m_Classifier = nothing
	End Property
	
	' Connector property. (EA.Connector)
	Public Property Get Connector
	  set Connector = m_Connector
	End Property
	Public Property Let Connector(value)
	  set m_Connector = value
	End Property
	
	' PropertyType property.
	Public Property Get PropertyType
		if me.Source.ObjectType = otAttribute then
			PropertyType = "Attribute"
		else
			PropertyType = "Association"
		end if
	End Property
	
	' GUID property
	Public Property Get GUID
		if me.PropertyType = "Attribute" then
			GUID = me.Source.AttributeGUID
		else
			GUID = me.Connector.ConnectorGUID
		end if
	End Property
	
	' Name property.
	Public Property Get Name
		if me.PropertyType = "Attribute" then
			Name = me.Source.Name
		else
			Name = me.Source.Role
		end if
	End Property
	
	' minOccurs property
	Public Property Get minOccurs
	  minOccurs = m_minOccurs
	End Property
	Public Property Let minOccurs(value)
	  m_minOccurs = value
	End Property
	
	' maxOccurs property
	Public Property Get maxOccurs
	  maxOccurs = m_maxOccurs
	End Property
	Public Property Let maxOccurs(value)
	  m_maxOccurs = value
	End Property
	
	' Indicates whether this property uses a redefined element
	Public Property Get Redefines
	  Redefines = me.Classifier.IsRedefinition
	End Property

	' Restricted property
	Public Property Get Restricted
	  Restricted = m_Restricted
	End Property
	Public Property Let Restricted(value)
	  m_Restricted = value
	End Property
	
	' Classifier property (SchemaElement)
	Public Property Get Classifier
		'lazy loading
		if m_Classifier is nothing then
			if me.PropertyType = "Attribute" then
				set m_Classifier = Repository.GetElementByID(me.Source.ClassifierID)
			else
				dim e as EA.ConnectorEnd
				if me.Source.End = "Supplier" then
					set m_Classifier = Repository.GetElementByID(me.Connector.SupplierID)
				else
					set m_Classifier = Repository.GetElementByID(me.Connector.ClientID)
				end if
			end if
		end if
		set Classifier = m_Classifier
	End Property
	
	'Functions
	
	
end Class