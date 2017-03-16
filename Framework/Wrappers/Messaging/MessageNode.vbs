'[path=\Framework\Wrappers\Messaging]
'[group=Messaging]

!INC Utils.Include
!INC Local Scripts.EAConstants-VBScript
' Author: Geert Bellekens
' Purpose: A wrapper class for a message node in a messaging structure
' Date: 2017-03-14

Class MessageNode 
	'private variables
	Private m_Name
	Private m_TypeElement
	Private m_TypeName
	Private m_Multiplicity
	Private m_ParentNode
	Private m_ChildNodes
	Private m_SourceAttribute
	Private m_SourceAssociationEnd
	Private m_SourceElement
	Private m_ValidationRule

	'constructor
	Private Sub Class_Initialize
		m_Name = ""
		set m_TypeElement = nothing
		m_TypeName = ""
		m_Multiplicity = ""
		set m_ParentNode = nothing
		set m_ChildNodes = CreateObject("System.Collections.ArrayList")
		set m_Attribute = nothing
		set m_AssociationEnd = nothing
		set m_Element = nothing
		set m_ValidationRule = nothing
	End Sub
	
	'public properties
	
	' Name property.
	Public Property Get Name
		Name = m_Name
	End Property
	Public Property Let Name(value)
		m_Name = value
	End Property
	
	' TypeElement property.
	Public Property Get TypeElement
		TypeElement = m_TypeElement
	End Property
	Public Property Let TypeElement(value)
		m_TypeElement = value
	End Property

	' ElementID property.
	Public Property Get ElementID
		if not me.TypeElement is noting then
			ElementID = me.TypeElement.ElementID
		else
			ElementID = 0
		end if
	End Property

	' TypeName property
	
	Public Property Get TypeName
		if not me.TypeElement is nothing then
			TypeName = me.TypeElement.Name
		else
			TypeName = m_TypeName
		end if 
	End Property
	Public Property Let TypeName(value)
		m_TypeName = value
		'if the typename if different from the TypeElement name then we remove the type Element
		if not me.TypeElement is nothing then
			if value <> me.TypeElement then
				me.TypeElement = nothing
			end if
		end if
	End Property
	
	' Multiplicity property.
	' only directly used if the source is element, else we use the Attribute or AssociationEnd multiplicity
	Public Property Get Multiplicity
		dim connectorEnd as EA.ConnectorEnd
		connectorEnd.Cardinality
		dim lower
		dim upper
		dim returnedMultiplicity
		if not me.SourceElement is nothing then
			returnedMultiplicity = m_Multiplicity
		elseif not me.sourceAttribute is nothing then
			returnedMultiplicity determineMultiplicity(me.sourceAttribute.LowerBound,me.sourceAttribute.UpperBound, "1", "1")
		elseif not me.sourceAssociationEnd is nothing then
			returnedMultiplicity = sourceAssociationEnd.Cardinality
		end if
		'return the actual value
		Multiplicity = returnedMultiplicity
	End Property
	Public Property Let Multiplicity(value)
		if not me.SourceElement is nothing then
			m_Multiplicity = value
		end if
	End Property
	
	private function determineMultiplicity(lower,upper,defaultLower, defaultUpper)
		'check to make sur the values are filled in and replace them with the default values if not the case
		if len(lower) = 0 then
			lower = defaultLower
		end if
		if len(upper) = 0 then
			upper = defaultUpper
		end if
		'create the multiplicity string
		determineMultiplicity = lower & ".." & upper
	end function
	' ParentNode property.
	Public Property Get ParentNode
		ParentNode = m_ParentNode
	End Property
	Public Property Let ParentNode(value)
		m_ParentNode = value
	End Property

	' ChildNodes property.
	Public Property Get ChildNodes
		ChildNodes = m_ChildNodes
	End Property
	Public Property Let ChildNodes(value)
		m_ChildNodes = value
	End Property
	
	' SourceAttribute property.
	Public Property Get SourceAttribute
		SourceAttribute = m_SourceAttribute
	End Property
	Public Property Let SourceAttribute(value)
		m_SourceAttribute = value
	End Property

	' SourceAssociationEnd property.
	Public Property Get SourceAssociationEnd
		SourceAssociationEnd = m_SourceAssociationEnd
	End Property
	Public Property Let SourceAssociationEnd(value)
		m_SourceAssociationEnd = value
	End Property
	
	' SourceElement property.
	Public Property Get SourceElement
		SourceElement = m_SourceElement
	End Property
	Public Property Let SourceElement(value)
		m_SourceElement = value
	End Property
	
	' ValidationRule property.
	Public Property Get ValidationRule
		ValidationRule = m_ValidationRule
	End Property
	Public Property Let ValidationRule(value)
		m_ValidationRule = value
	End Property	
	
	'public functions
	public function intitializeWithSource(source,sourceConnector,in_multiplicity,in_validationRule,in_parentNode)
		'set validationRule
		if not in_validationRule is nothing then
			me.ValidationRule = in_validationRule
		end if
		'set parentNode
		if not in_parentNode is nothing then
			me.ParentNode = in_parentNode
		end if
		'check if source is Element, Atttribute, or AssociationEnd
		select case source.ObjectType
			case otElement
				me.SourceElement = source
				me.Name = source.Name
				me.TypeElement = source
				me.Multiplicity = in_multiplicity
			case otAttribute
				me.SourceAttribute = source
				me.Name = source.Name
				dim attributeTypeObject
				set attributeTypeObject = nothing
				if Repositorysource.ClassifierID > 0 then
					attributeTypeObject = Repository.GetElementByID(source.ClassifierID)
				else
					me.TypeName = source.Type
				end if
				me.TypeElement = attributeTypeObject
			case otConnectorEnd
				me.SourceAssociationEnd = source
				if len(source.Role) > 0 then
					me.Name = source.Role
				else
					dim endObject as EA.Element
					'get the end object and use that name
					if source.End = "Supplier" then
						set endObject = Repository.GetElementByID(sourceConnector.SupplierID)
					else
						set endObject = Repository.GetElementByID(sourceConnector.ClientID)
					end if
					if not endObject is nothing then
						me.Name = endObject.Name
						me.TypeElement = endObject
					end if
				end if 
		end select
		'then load the child nodes
		loadChildNodes
	end function
	
	'Loads the child nodes for this message (resursively until we have reached all the leaves)
	public function loadChildNodes()
		'first remember the list of parent elements
		dim parents
		set parents = getParents(nothing)
		'TODO: load in correct order?
		'load attributes
		loadAllAttributeNodes parents 
		'load associations
		'load nested classes?
	end function
	'returns a list of all generalized elements of this elemnt
	private function getParents(childElement)
		dim directParents 
		dim sqlGetParents
		dim childElementID
		if not childElement is nothing then
			childElementID = childElement.ElementID
		else
			childElementID = me.ElementID
		end if
		sqlGetParents = "select c.End_Object_ID as Object_ID from t_connector c			 "  & _
						" where c.Connector_Type in ('Generalization','Generalisation')	 "  & _
						" and c.Start_Object_ID =" & childElementID
		set directParents = getElementsFromQuery(sqlGetParents)
		'loop the parent and get their parents
		dim parent
		for each parent in parents
			directParents.AddRange(getParents(parent))
		next
		'return
		set getParents = directParents
	end function
	'loads all Attribute notes both from this eleemnt as from its parents
	private function loadAllAttributeNodes(parents)
		'first load fro this element
		dim allAttributeNodes
		set allAttributeNodes = loadAttributeChildNodes(nothing)
		'then the one from the parents
		dim parent
		for each parent in parents
			allAttributeNodes.AddRange loadAttributeChildNodes(parent)
		next
	end function
	private function loadAttributeChildNodes(currentElement)
		dim ownerElementID
		if not childElement is nothing then
			ownerElementID = currentElement.ElementID
		else
			ownerElementID = me.ElementID
		end if
		'get attributes in the correct order
		dim SQLGetAttributes
		SQLGetAttributes = 	"select a.ID from t_attribute a " & _
							" where a.Object_ID = "& ownerElementID & _
							" order by a.Pos, a.Name		"
		dim attributes
		set attributes = getattributesFromQuery(SQLGetAttributes)
		'loop the attributes
		dim attribute
		for each attribute in attributes
			'create the next messageNode
			dim newMessageNode
			set newMessageNode = new MessageNode
			'initialize
			newMessageNode.intitializeWithSource attribute, nothing, "", nothing, me
			'add to the childnodes list
			me.ChildNodes.Add newMessageNode
		next
	end function
	
end Class