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
	Private m_Classifier
	Private m_ClassifierSchemaElement
	Private m_Owner
	
	
	'constructor
	Private Sub Class_Initialize
		set m_Source = nothing
		set m_Connector = nothing
		set m_Classifier = nothing
		set m_ClassifierSchemaElement= nothing
		set m_Owner = nothing
		m_minOccurs = ""
		m_maxOccurs = ""
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
		if len(m_minOccurs) > 0 then
			minOccurs = m_minOccurs
		else
			minOccurs = getSourceLowerBound()
		end if
	End Property
	Public Property Let minOccurs(value)
	  m_minOccurs = value
	End Property
	
	' maxOccurs property
	Public Property Get maxOccurs
		if len(m_maxOccurs) > 0 then
			maxOccurs = m_maxOccurs
		else
			maxOccurs = getSourceUpperBound()
		end if
	End Property
	Public Property Let maxOccurs(value)
	  m_maxOccurs = value
	End Property
	
	' return the name of the redefined schema element
	Public Property Get Redefines
		Redefines = ""
		if not me.ClassifierSchemaElement is nothing then
			if me.ClassifierSchemaElement.IsRedefinition then
				Redefines = me.ClassifierSchemaElement.Name
			end if
		end if
	End Property

	' Restricted property (Boolean)
	Public Property Get IsRestricted
		'default
		IsRestricted = false
		'check if lowerbound is filled in and different
		if len(m_minOccurs) > 0 _
		  and m_minOccurs <> getSourceLowerBound() then
			IsRestricted = true
	    end if
		'check if upperbound is filled in and different
		if len(m_maxOccurs) > 0 _
		  and m_maxOccurs <> getSourceUpperBound() then
			IsRestricted = true
	    end if
		'check if uses Redefined element
		if len(me.Redefines) > 0 then
			IsRestricted = true
		end if
	End Property
	
	' Classifier property (EA.Element)
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
	
	' ClassifierSchemaElement property. (SchemaElement)
	Public Property Get ClassifierSchemaElement
	  set ClassifierSchemaElement = m_ClassifierSchemaElement
	End Property
	Public Property Let ClassifierSchemaElement(value)
	  set m_ClassifierSchemaElement = value
	  'add me as referencing property to the schema element
	  m_ClassifierSchemaElement.addReferencingProperty me
	End Property
	
	' Owner property (SchemaElement)
	Public Property Get Owner
	  set Owner = m_Owner
	End Property
	Public Property Let Owner(value)
	  set m_Owner = value
	End Property
	

	'Public Functions
	'Delete this property by removing it from its owner and from the referencing properties lis of the classifier
	public function Delete
		'debug
		'Session.Output "Deleting property: " & me.Name
		me.Owner.deleteProperty me
		if not me.ClassifierSchemaElement is nothing then
			me.ClassifierSchemaElement.removeReferencingProperty me
		end if
	end function
	
	'Private Functions
	private function getSourceLowerBound()
		if me.PropertyType = "Attribute" then
			getSourceLowerBound = me.Source.LowerBound
			if len(getSourceLowerBound) = 0 then
				'default 1 for attributes
				getSourceLowerBound = "1"
			end if			
		else
			getSourceLowerBound = getAssociationEndLowerBound()
		end if
	end function
	private function getSourceUpperBound()
		if me.PropertyType = "Attribute" then
			getSourceUpperBound = me.Source.UpperBound
			if len(getSourceUpperBound) = 0 then
				'default 1 for attributes
				getSourceUpperBound = "1"
			end if
		else
			getSourceUpperBound = getAssociationEndUpperBound()
		end if
	end function
	
	private function getAssociationEndLowerBound()
		dim cardinalityParts
		cardinalityParts = split(me.Source.Cardinality,"..")
		if Ubound(cardinalityParts) = 1 then
			getAssociationEndLowerBound = cardinalityParts(0)
		else
			getAssociationEndLowerBound = me.Source.Cardinality
		end if
		'default lowerbound = 0
		if len(getAssociationEndLowerBound) = 0 then
			getAssociationEndLowerBound = "0"
		end if
	end function
	private function getAssociationEndUpperBound()
		dim cardinalityParts
		cardinalityParts = split(me.Source.Cardinality,"..")
		if Ubound(cardinalityParts) = 1 then
			getAssociationEndUpperBound = cardinalityParts(1)
		else
			getAssociationEndUpperBound = me.Source.Cardinality
		end if
		'default upperbound = *
		if len(getAssociationEndUpperBound) = 0 then
			getAssociationEndUpperBound	= "*"
		end if
	end function
	
end Class
